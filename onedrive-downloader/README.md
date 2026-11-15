# onedrive-downloader

Baixa arquivos e pastas do OneDrive (pessoal ou corporativo) direto pela API
Microsoft Graph, **sem instalar o cliente do OneDrive**. Feito para arquivos
enormes em conexões que caem no meio do caminho.

- **Retoma de onde parou**: se a conexão cair, continua do último byte recebido
  (HTTP Range), inclusive entre execuções diferentes do programa.
- **Novas tentativas automáticas**, com espera progressiva; só desiste de um
  arquivo após 30 falhas seguidas *sem receber nenhum byte*.
- **Renova sozinho** o token de acesso e o link de download quando expiram.
- **Confere a integridade** ao final (SHA-256, SHA-1 ou QuickXorHash).
- **Nada para instalar** além do Python 3.8+ — é um único arquivo, só com a
  biblioteca padrão.

## Uso

```
python onedrive_dl.py --listar                    # lista tudo, sem baixar
python onedrive_dl.py "Videos/Projeto" -d D:\Backup   # baixa uma pasta
python onedrive_dl.py "Backups/base.bak" -d D:\Backup # baixa um arquivo
python onedrive_dl.py -d D:\OneDrive              # baixa o OneDrive inteiro
```

Na primeira execução o programa mostra um endereço e um código:

```
LOGIN NECESSARIO
  1. Abra no navegador: https://microsoft.com/devicelogin
  2. Digite o codigo:   ABCD1234
```

Abra o endereço em qualquer navegador (pode ser em outro aparelho), digite o
código e entre com a sua conta. A senha nunca passa pelo programa. O login fica
salvo em `~/.onedrive-dl/token.json`; para apagá-lo, use `--sair`.

Se o download for interrompido (queda de internet, Ctrl+C, computador
desligado), **execute o mesmo comando de novo**: arquivos completos são
pulados e os parciais (`.part`) continuam de onde pararam.

### Opções

| Opção | Efeito |
|---|---|
| `-d`, `--destino PASTA` | pasta local de destino (padrão: `./OneDrive`) |
| `-l`, `--listar` | só lista arquivos e tamanhos |
| `--tentativas N` | falhas seguidas sem progresso antes de desistir (padrão: 30) |
| `--timeout S` | segundos sem dados até considerar a conexão travada (padrão: 60) |
| `--sem-verificar` | não confere o hash ao final |
| `--manter-existentes` | não rebaixa arquivos existentes, mesmo com tamanho diferente |
| `--forcar` | baixa de novo mesmo que o arquivo já exista |
| `--tenant`, `--client-id` | veja abaixo |
| `--sair` | apaga o login salvo |

## Se o login for bloqueado (conta corporativa)

Por padrão o programa usa o aplicativo público da Microsoft *Microsoft Graph
Command Line Tools*. Algumas empresas bloqueiam o consentimento a aplicativos
ou o login por código de dispositivo. Nesse caso:

1. Tente informar o domínio da empresa: `--tenant suaempresa.com.br`.
2. Se ainda for bloqueado, peça ao administrador para liberar, ou registre um
   aplicativo próprio em <https://entra.microsoft.com> → *Registros de
   aplicativo* → *Novo registro*:
   - em *Autenticação*, ative **Permitir fluxos de cliente público**;
   - em *Permissões de API*, adicione a permissão delegada do Microsoft Graph
     **Files.Read.All**;
   - use o *ID do aplicativo (cliente)* com `--client-id <ID>`.

## Observações

- Blocos de anotações do OneNote não têm download direto e são ignorados.
- Em bibliotecas do SharePoint/OneDrive corporativo, alguns arquivos do Office
  são alterados pelo servidor ao serem baixados, então o aviso de "hash
  diferente" pode aparecer para eles mesmo com o download íntegro.
- Nomes com caracteres inválidos no Windows (`: ? * |` etc.) têm esses
  caracteres trocados por `_`.
