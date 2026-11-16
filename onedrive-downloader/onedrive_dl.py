#!/usr/bin/env python3
"""
onedrive_dl.py - Baixa arquivos/pastas do OneDrive (pessoal ou corporativo)
direto pela API Microsoft Graph, sem instalar o cliente do OneDrive.

Feito para arquivos enormes e conexoes instaveis:
  * retoma downloads interrompidos (HTTP Range) a partir do arquivo .part
  * novas tentativas automaticas com espera progressiva
  * renova o token e o link de download quando expiram no meio do caminho
  * confere tamanho e hash (SHA-256 / SHA-1 / QuickXorHash) ao final
  * pode ser interrompido (Ctrl+C) e executado de novo: continua de onde parou

Usa apenas a biblioteca padrao do Python 3.8+ (nada para instalar).
"""

import argparse
import base64
import errno
import hashlib
import http.client
import json
import os
import re
import socket
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime, timezone

# App publico da Microsoft ("Microsoft Graph Command Line Tools"), que aceita
# login por codigo de dispositivo. Pode ser trocado por um app proprio com
# --client-id (veja o README).
DEFAULT_CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e"
DEFAULT_TENANT = "common"
SCOPES = "Files.Read.All offline_access"

LOGIN_URL = os.environ.get("ONEDRIVE_DL_LOGIN_URL", "https://login.microsoftonline.com").rstrip("/")
GRAPH_URL = os.environ.get("ONEDRIVE_DL_GRAPH_URL", "https://graph.microsoft.com/v1.0").rstrip("/")

CONFIG_DIR = os.path.join(os.path.expanduser("~"), ".onedrive-dl")
TOKEN_FILE = os.path.join(CONFIG_DIR, "token.json")

CHUNK = 1024 * 1024  # 1 MiB
NETWORK_ERRORS = (
    urllib.error.URLError,
    http.client.HTTPException,
    socket.timeout,
    TimeoutError,
    ConnectionError,
)


class FatalError(Exception):
    """Erro que nao adianta tentar de novo."""


# --------------------------------------------------------------------------
# Utilidades
# --------------------------------------------------------------------------

def human(n):
    n = float(n)
    for unit in ("B", "KB", "MB", "GB", "TB"):
        if n < 1024 or unit == "TB":
            return f"{n:.0f} {unit}" if unit == "B" else f"{n:.2f} {unit}"
        n /= 1024


def human_time(sec):
    if sec is None or sec != sec or sec == float("inf"):
        return "--:--"
    sec = int(sec)
    h, rem = divmod(sec, 3600)
    m, s = divmod(rem, 60)
    return f"{h}h{m:02d}m" if h else f"{m:02d}m{s:02d}s"


def log(msg=""):
    sys.stdout.write("\r\033[K" + msg + "\n" if sys.stdout.isatty() else msg + "\n")
    sys.stdout.flush()


def status_line(msg):
    if sys.stdout.isatty():
        sys.stdout.write("\r\033[K" + msg[:150])
        sys.stdout.flush()


def backoff(attempt, retry_after=None):
    if retry_after:
        try:
            return min(float(retry_after), 600.0)
        except ValueError:
            pass
    return min(2.0 ** min(attempt, 6), 60.0)


def safe_name(name):
    """Nome valido no Windows e em outros sistemas."""
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name).rstrip(" .")
    if re.fullmatch(r"(?i)(con|prn|aux|nul|com[1-9]|lpt[1-9])(\..*)?", name or ""):
        name = "_" + name
    return name or "_"


def local_path(path):
    """Caminho absoluto; no Windows usa o prefixo de caminho longo (>260)."""
    path = os.path.abspath(path)
    if os.name == "nt" and not path.startswith("\\\\?\\"):
        if path.startswith("\\\\"):
            return "\\\\?\\UNC\\" + path[2:]
        return "\\\\?\\" + path
    return path


class NoRedirect(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, *args, **kwargs):
        return None


# --------------------------------------------------------------------------
# QuickXorHash (hash usado pelo OneDrive for Business / SharePoint)
# --------------------------------------------------------------------------

class QuickXorHash:
    WIDTH = 160
    SHIFT = 11

    def __init__(self):
        self._acc = 0      # XOR dos blocos de 160 bytes ja processados
        self._tail = b""   # bytes que ainda nao completaram um bloco
        self._len = 0

    def update(self, data):
        self._len += len(data)
        data = self._tail + bytes(data)
        usable = len(data) - (len(data) % self.WIDTH)
        acc = self._acc
        mv = memoryview(data)
        for i in range(0, usable, self.WIDTH):
            acc ^= int.from_bytes(mv[i:i + self.WIDTH], "little")
        self._acc = acc
        self._tail = data[usable:]

    def digest(self):
        cells = bytearray(self._acc.to_bytes(self.WIDTH, "little"))
        for i, b in enumerate(self._tail):
            cells[i] ^= b
        mask = (1 << self.WIDTH) - 1
        h = 0
        for i, b in enumerate(cells):
            if b:
                shifted = b << ((i * self.SHIFT) % self.WIDTH)
                h ^= (shifted & mask) | (shifted >> self.WIDTH)
        h ^= (self._len & 0xFFFFFFFFFFFFFFFF) << (self.WIDTH - 64)
        return (h & mask).to_bytes(20, "little")

    def b64digest(self):
        return base64.b64encode(self.digest()).decode()


# --------------------------------------------------------------------------
# Autenticacao (OAuth2 device code flow)
# --------------------------------------------------------------------------

class Auth:
    def __init__(self, client_id, tenant):
        self.client_id = client_id
        self.tenant = tenant
        self.token = None
        self._load()

    def _endpoint(self, name):
        return f"{LOGIN_URL}/{self.tenant}/oauth2/v2.0/{name}"

    def _post(self, url, fields):
        data = urllib.parse.urlencode(fields).encode()
        req = urllib.request.Request(url, data=data)
        try:
            with urllib.request.urlopen(req, timeout=60) as r:
                return json.load(r)
        except urllib.error.HTTPError as e:
            try:
                return json.load(e)
            except Exception:
                raise FatalError(f"Falha na autenticacao: HTTP {e.code}")

    def _load(self):
        try:
            with open(TOKEN_FILE, encoding="utf-8") as f:
                tok = json.load(f)
            if tok.get("client_id") == self.client_id and tok.get("tenant") == self.tenant:
                self.token = tok
        except (OSError, ValueError):
            pass

    def _save(self, resp):
        self.token = {
            "client_id": self.client_id,
            "tenant": self.tenant,
            "access_token": resp["access_token"],
            "refresh_token": resp.get("refresh_token") or (self.token or {}).get("refresh_token"),
            "expires_at": time.time() + int(resp.get("expires_in", 3600)),
        }
        os.makedirs(CONFIG_DIR, exist_ok=True)
        tmp = TOKEN_FILE + ".tmp"
        fd = os.open(tmp, os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            json.dump(self.token, f)
        os.replace(tmp, TOKEN_FILE)

    def _refresh(self):
        rt = (self.token or {}).get("refresh_token")
        if not rt:
            return False
        for attempt in range(5):
            try:
                resp = self._post(self._endpoint("token"), {
                    "client_id": self.client_id,
                    "grant_type": "refresh_token",
                    "refresh_token": rt,
                    "scope": SCOPES,
                })
            except NETWORK_ERRORS:
                time.sleep(backoff(attempt))
                continue
            if "access_token" in resp:
                self._save(resp)
                return True
            return False
        raise FatalError("Sem conexao para renovar o token de acesso.")

    def login(self):
        resp = self._post(self._endpoint("devicecode"),
                          {"client_id": self.client_id, "scope": SCOPES})
        if "device_code" not in resp:
            raise FatalError("Nao foi possivel iniciar o login: "
                             + str(resp.get("error_description") or resp))
        log()
        log("=" * 70)
        log("LOGIN NECESSARIO")
        log(f"  1. Abra no navegador: {resp.get('verification_uri')}")
        log(f"  2. Digite o codigo:   {resp.get('user_code')}")
        log("  3. Entre com a conta do OneDrive e autorize o acesso.")
        log("=" * 70)
        log("Aguardando a autorizacao...")
        interval = int(resp.get("interval", 5))
        deadline = time.time() + int(resp.get("expires_in", 900))
        while time.time() < deadline:
            time.sleep(interval)
            try:
                tok = self._post(self._endpoint("token"), {
                    "client_id": self.client_id,
                    "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
                    "device_code": resp["device_code"],
                })
            except NETWORK_ERRORS:
                continue
            if "access_token" in tok:
                self._save(tok)
                log("Login concluido.")
                return
            err = tok.get("error")
            if err == "authorization_pending":
                continue
            if err == "slow_down":
                interval += 5
                continue
            raise FatalError("Login recusado: " + str(tok.get("error_description") or err))
        raise FatalError("O codigo de login expirou. Execute novamente.")

    def access_token(self, force_refresh=False):
        tok = self.token
        if tok and not force_refresh and tok.get("expires_at", 0) - time.time() > 300:
            return tok["access_token"]
        if not self._refresh():
            self.login()
        return self.token["access_token"]

    @staticmethod
    def logout():
        try:
            os.remove(TOKEN_FILE)
            log("Token removido de " + TOKEN_FILE)
        except FileNotFoundError:
            log("Nenhum token salvo.")


# --------------------------------------------------------------------------
# Cliente Microsoft Graph
# --------------------------------------------------------------------------

class Graph:
    def __init__(self, auth, timeout=60):
        self.auth = auth
        self.timeout = timeout
        self._noredir = urllib.request.build_opener(NoRedirect)

    def _request(self, url, opener=None):
        """GET autenticado com novas tentativas. Retorna (status, headers, body)."""
        if not url.startswith("http"):
            url = GRAPH_URL + url
        force = False
        attempt = 0
        while True:
            req = urllib.request.Request(url, headers={
                "Authorization": "Bearer " + self.auth.access_token(force_refresh=force),
                "Accept": "application/json",
            })
            force = False
            try:
                op = opener.open if opener else urllib.request.urlopen
                with op(req, timeout=self.timeout) as r:
                    return r.status, r.headers, r.read()
            except urllib.error.HTTPError as e:
                body = e.read()
                if e.code in (301, 302, 303, 307, 308):
                    return e.code, e.headers, body
                if e.code == 401 and attempt < 2:
                    force = True
                elif e.code in (429, 500, 502, 503, 504, 509) and attempt < 10:
                    wait = backoff(attempt, e.headers.get("Retry-After"))
                    status_line(f"Servidor ocupado (HTTP {e.code}); aguardando {wait:.0f}s...")
                    time.sleep(wait)
                else:
                    msg = ""
                    try:
                        msg = json.loads(body)["error"]["message"]
                    except Exception:
                        pass
                    raise FatalError(f"Erro da API (HTTP {e.code}): {msg or url}")
            except NETWORK_ERRORS as e:
                if attempt >= 10:
                    raise FatalError(f"Sem conexao com a API do OneDrive: {e}")
                wait = backoff(attempt)
                status_line(f"Falha de rede ({e}); nova tentativa em {wait:.0f}s...")
                time.sleep(wait)
            attempt += 1

    def get(self, url):
        return json.loads(self._request(url)[2])

    @staticmethod
    def item_url(item):
        drive = (item.get("parentReference") or {}).get("driveId")
        if drive:
            return f"/drives/{drive}/items/{item['id']}"
        return f"/me/drive/items/{item['id']}"

    def resolve(self, remote_path):
        remote_path = remote_path.strip().replace("\\", "/").strip("/")
        if not remote_path:
            return self.get("/me/drive/root")
        return self.get("/me/drive/root:/" + urllib.parse.quote(remote_path))

    def children(self, item):
        url = self.item_url(item) + "/children?$top=200"
        while url:
            page = self.get(url)
            yield from page.get("value", [])
            url = page.get("@odata.nextLink")

    def download_url(self, item):
        """Link de download temporario (pre-autenticado) e sempre novo."""
        fresh = self.get(self.item_url(item))
        url = fresh.get("@microsoft.graph.downloadUrl") or fresh.get("@content.downloadUrl")
        if url:
            return url
        status, headers, _ = self._request(self.item_url(item) + "/content", opener=self._noredir)
        if status in (301, 302, 303, 307, 308) and headers.get("Location"):
            return headers["Location"]
        raise FatalError("A API nao forneceu link de download para este item.")


def walk(graph, item, rel=""):
    """Gera (item, caminho_relativo) para todos os arquivos abaixo de `item`."""
    stack = [(item, rel)]
    while stack:
        folder, frel = stack.pop()
        status_line(f"Listando: /{frel}")
        for child in graph.children(folder):
            # Atalhos para pastas/arquivos compartilhados apontam para outro drive.
            remote = child.get("remoteItem")
            target = child
            if remote:
                target = dict(remote)
                target.setdefault("name", child.get("name"))
            name = safe_name(child.get("name") or target.get("name") or target["id"])
            crel = f"{frel}/{name}" if frel else name
            if "folder" in target:
                stack.append((target, crel))
            elif "file" in target:
                yield target, crel
            else:
                log(f"  [ignorado] {crel} (tipo sem download direto, ex.: bloco do OneNote)")


# --------------------------------------------------------------------------
# Download com retomada
# --------------------------------------------------------------------------

class Progress:
    def __init__(self, label, total, start):
        self.label = label
        self.total = total
        self.done = start
        self._t0 = time.time()
        self._b0 = start
        self._last = 0.0
        self._speed = 0.0

    def add(self, n):
        self.done += n
        now = time.time()
        if now - self._last < 0.5:
            return
        dt = now - self._t0
        if dt >= 3:
            self._speed = (self.done - self._b0) / dt
            self._t0, self._b0 = now, self.done
        elif not self._speed and dt > 0:
            self._speed = (self.done - self._b0) / dt
        self._last = now
        pct = f"{100 * self.done / self.total:5.1f}%" if self.total else "  ?  "
        eta = (self.total - self.done) / self._speed if self._speed and self.total else None
        status_line(f"  {pct} {human(self.done)}/{human(self.total)} "
                    f"{human(self._speed)}/s falta {human_time(eta)} | {self.label}")


def expected_hash(item):
    hashes = (item.get("file") or {}).get("hashes") or {}
    for key, algo in (("sha256Hash", "sha256"), ("sha1Hash", "sha1"), ("quickXorHash", "quickxor")):
        if hashes.get(key):
            return algo, hashes[key]
    return None, None


def file_hash(path, algo, label):
    h = QuickXorHash() if algo == "quickxor" else hashlib.new(algo)
    total = os.path.getsize(path)
    done = 0
    last = 0.0
    with open(path, "rb") as f:
        while True:
            buf = f.read(4 * CHUNK)
            if not buf:
                break
            h.update(buf)
            done += len(buf)
            if time.time() - last > 0.5:
                last = time.time()
                status_line(f"  Verificando {100 * done / max(total, 1):5.1f}% | {label}")
    return h.b64digest() if algo == "quickxor" else h.hexdigest().upper()


def set_mtime(path, item):
    stamp = (item.get("fileSystemInfo") or {}).get("lastModifiedDateTime") \
        or item.get("lastModifiedDateTime")
    if not stamp:
        return
    try:
        stamp = re.sub(r"\.\d+", "", stamp).replace("Z", "+00:00")
        ts = datetime.fromisoformat(stamp).astimezone(timezone.utc).timestamp()
        os.utime(path, (ts, ts))
    except (ValueError, OSError):
        pass


def download_file(graph, item, dest, opts):
    """Baixa um arquivo. Retorna 'ok', 'pulado' ou 'aviso'."""
    label = os.path.basename(dest)
    final = local_path(dest)
    part = final + ".part"
    meta = part + ".json"
    size = int(item.get("size") or 0)
    tag = item.get("cTag") or item.get("eTag") or ""

    if os.path.exists(final) and not opts.force:
        if os.path.getsize(final) == size or opts.skip_existing:
            return "pulado"

    os.makedirs(os.path.dirname(final), exist_ok=True)

    # O .part so e reaproveitado se o arquivo remoto nao mudou desde entao.
    if os.path.exists(part):
        try:
            with open(meta, encoding="utf-8") as f:
                saved = json.load(f)
        except (OSError, ValueError):
            saved = {}
        if saved.get("tag") != tag or saved.get("id") != item["id"]:
            os.remove(part)
    try:
        with open(meta, "w", encoding="utf-8") as f:
            json.dump({"id": item["id"], "tag": tag, "size": size}, f)
    except OSError as e:
        raise FatalError(f"Nao foi possivel gravar em {os.path.dirname(dest)}: {e}")

    failures = 0
    while True:
        offset = os.path.getsize(part) if os.path.exists(part) else 0
        if size and offset == size:
            break
        if size and offset > size:
            os.remove(part)
            offset = 0

        progressed = False
        finished = False
        retry_after = None
        error = None
        try:
            url = graph.download_url(item)
            headers = {"Accept-Encoding": "identity"}
            if offset:
                headers["Range"] = f"bytes={offset}-"
            req = urllib.request.Request(url, headers=headers)
            with urllib.request.urlopen(req, timeout=opts.timeout) as resp:
                total = None
                if resp.status == 206:
                    m = re.match(r"bytes (\d+)-\d+/(\d+)", resp.headers.get("Content-Range", ""))
                    if not m or int(m.group(1)) != offset:
                        raise http.client.HTTPException("resposta parcial inesperada do servidor")
                    total = int(m.group(2))
                else:
                    offset = 0  # servidor ignorou o Range: recomeca do zero
                    if resp.headers.get("Content-Length"):
                        total = int(resp.headers["Content-Length"])
                if total is not None and total != size:
                    # Acontece no SharePoint: o tamanho dos metadados nem sempre
                    # bate com o conteudo real. Vale o que o servidor entrega.
                    size = total
                prog = Progress(label, size, offset)
                try:
                    out = open(part, "r+b" if offset else "wb")
                except OSError as e:
                    raise FatalError(f"Nao foi possivel gravar {part}: {e}")
                with out:
                    out.seek(offset)
                    out.truncate()
                    while True:
                        buf = resp.read(CHUNK)
                        if not buf:
                            break
                        try:
                            out.write(buf)
                        except OSError as e:
                            if e.errno == errno.ENOSPC:
                                raise FatalError("Disco cheio. Libere espaco e execute de novo "
                                                 "(o download continua de onde parou).")
                            raise FatalError(f"Erro ao gravar {part}: {e}")
                        progressed = True
                        prog.add(len(buf))
                got = os.path.getsize(part)
                if got >= size:
                    finished = True
                else:
                    error = "conexao encerrada antes do fim"
        except urllib.error.HTTPError as e:
            if e.code == 416:
                # Range invalido: o .part nao corresponde mais ao arquivo remoto.
                if os.path.exists(part):
                    os.remove(part)
            retry_after = e.headers.get("Retry-After")
            error = f"HTTP {e.code}"
        except NETWORK_ERRORS as e:
            error = str(getattr(e, "reason", e)) or type(e).__name__

        if finished:
            break
        # Enquanto houver avanco o contador zera: so desiste apos N falhas seguidas
        # sem receber nenhum byte.
        failures = 1 if progressed else failures + 1
        if failures > opts.retries:
            raise FatalError(f"{label}: desistindo apos {opts.retries} tentativas seguidas "
                             f"sem progresso ({error}). Execute de novo para continuar.")
        wait = backoff(failures, retry_after)
        have = os.path.getsize(part) if os.path.exists(part) else 0
        status_line(f"  Interrompido em {human(have)} ({error}); "
                    f"retomando em {wait:.0f}s [{failures}/{opts.retries}] | {label}")
        time.sleep(wait)

    if not os.path.exists(part):  # arquivo de tamanho zero
        open(part, "wb").close()

    result = "ok"
    algo, want = expected_hash(item)
    if opts.verify and algo:
        got = file_hash(part, algo, label)
        if got.lower() != want.lower() if algo != "quickxor" else got != want:
            log(f"  [AVISO] hash diferente do informado pelo OneDrive: {dest}")
            result = "aviso"

    os.replace(part, final)
    try:
        os.remove(meta)
    except OSError:
        pass
    set_mtime(final, item)
    return result


# --------------------------------------------------------------------------
# Linha de comando
# --------------------------------------------------------------------------

def parse_args(argv=None):
    p = argparse.ArgumentParser(
        description="Baixa arquivos e pastas do OneDrive com retomada automatica.",
        epilog="Exemplos:\n"
               "  python onedrive_dl.py --listar\n"
               '  python onedrive_dl.py "Videos/Projeto" -d D:\\Backup\n'
               "  python onedrive_dl.py -d D:\\OneDrive        (baixa tudo)",
        formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("origem", nargs="?", default="",
                   help="pasta ou arquivo no OneDrive (padrao: tudo)")
    p.add_argument("-d", "--destino", default="OneDrive",
                   help="pasta local de destino (padrao: ./OneDrive)")
    p.add_argument("-l", "--listar", action="store_true",
                   help="apenas lista os arquivos e tamanhos, sem baixar")
    p.add_argument("--tentativas", dest="retries", type=int, default=30,
                   help="falhas seguidas sem progresso antes de desistir de um arquivo (padrao: 30)")
    p.add_argument("--timeout", type=int, default=60,
                   help="segundos sem receber dados ate considerar a conexao travada (padrao: 60)")
    p.add_argument("--sem-verificar", dest="verify", action="store_false",
                   help="nao confere o hash ao final (mais rapido)")
    p.add_argument("--manter-existentes", dest="skip_existing", action="store_true",
                   help="nao baixa de novo arquivos que ja existem, mesmo com tamanho diferente")
    p.add_argument("--forcar", dest="force", action="store_true",
                   help="baixa de novo mesmo que o arquivo ja exista")
    p.add_argument("--client-id", default=os.environ.get("ONEDRIVE_DL_CLIENT_ID", DEFAULT_CLIENT_ID),
                   help="ID de um aplicativo proprio registrado no Entra ID (opcional)")
    p.add_argument("--tenant", default=os.environ.get("ONEDRIVE_DL_TENANT", DEFAULT_TENANT),
                   help="tenant: common, organizations, consumers ou o dominio/ID da empresa")
    p.add_argument("--sair", dest="logout", action="store_true",
                   help="apaga o login salvo nesta maquina e encerra")
    return p.parse_args(argv)


def main(argv=None):
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(errors="replace")
        except Exception:
            pass
    opts = parse_args(argv)
    if opts.logout:
        Auth.logout()
        return 0

    auth = Auth(opts.client_id, opts.tenant)
    graph = Graph(auth)
    auth.access_token()

    root = graph.resolve(opts.origem)
    if "remoteItem" in root:
        root = dict(root["remoteItem"], name=root.get("name"))

    if "folder" in root:
        # Baixar uma subpasta cria essa subpasta dentro do destino.
        base = safe_name(root["name"]) if opts.origem.strip("/\\ ") else ""
        files = walk(graph, root, base)
    elif "file" in root:
        files = iter([(root, safe_name(root["name"]))])
    else:
        raise FatalError("O item informado nao e um arquivo nem uma pasta.")

    if opts.listar:
        count = total = 0
        for item, rel in files:
            log(f"{human(item.get('size', 0)):>12}  {rel}")
            count += 1
            total += int(item.get("size") or 0)
        log(f"\n{count} arquivo(s), {human(total)} no total.")
        return 0

    stats = {"ok": 0, "pulado": 0, "aviso": 0, "erro": 0}
    failed = []
    for item, rel in files:
        dest = os.path.join(opts.destino, *rel.split("/"))
        try:
            result = download_file(graph, item, dest, opts)
        except FatalError as e:
            stats["erro"] += 1
            failed.append(rel)
            log(f"  [ERRO] {e}")
            if "Disco cheio" in str(e):
                break
            continue
        stats[result] += 1
        if result == "pulado":
            status_line(f"  ja existe: {rel}")
        else:
            log(f"  [{'OK' if result == 'ok' else 'AVISO'}] {rel} ({human(item.get('size', 0))})")

    log()
    log(f"Concluido: {stats['ok']} baixado(s), {stats['pulado']} ja existente(s), "
        f"{stats['aviso']} com aviso de hash, {stats['erro']} com erro.")
    if failed:
        log("Arquivos com erro (execute o mesmo comando de novo para continuar):")
        for rel in failed:
            log("  - " + rel)
        return 1
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        log("\nInterrompido. Execute o mesmo comando de novo para continuar de onde parou.")
        sys.exit(130)
    except FatalError as e:
        log(f"\nERRO: {e}")
        sys.exit(1)
