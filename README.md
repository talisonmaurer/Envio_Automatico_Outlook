# Envio Automático Outlook (clássico)

Aplicativo Windows em Python para enviar mensagens individualizadas com anexos específicos a cada destinatário pelo Outlook clássico.

## Instalação e uso

1. Instale Python 3 e abra o Outlook clássico.
2. Execute `INICIAR.bat`. Se necessário, o script instalará `pywin32`.
3. Copie `destinatarios_EXEMPLO.csv` para `destinatarios.csv` e preencha seus próprios dados.
4. Selecione a pasta de documentos e informe o endereço remetente exatamente como listado no Outlook.
5. Clique em **CONFERIR ANEXOS**. Só serão processadas linhas com `SIM` na coluna `liberacao`.
6. Use primeiro **RASCUNHOS (teste)** e confira as mensagens. Depois selecione o envio real, com confirmação.

No CSV, separe os e-mails da mesma empresa com ponto e vírgula (`;`) e vários arquivos com barra vertical (`|`). Cada linha gera uma mensagem independente.

## Segurança e privacidade

Este repositório público contém apenas código e exemplos fictícios. Não inclua documentos reais, e-mails de empresas, pastas institucionais ou histórico de envios. O arquivo `.gitignore` ajuda a prevenir adições acidentais.

O registro ENVIADO significa que o Outlook recebeu o comando de envio, não que o destinatário confirmou o recebimento.

## Compatibilidade

Windows + Outlook **clássico** + Python 3 + pywin32. O novo Outlook não oferece a mesma interface COM. Teste o comportamento com sua conta antes de enviar dados reais.
