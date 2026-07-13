# Zapsign Reports Function

Função Python para gerar relatórios de adoções a partir da API do Zapsign, enviar o resultado para um bucket do Google Cloud Storage e encaminhar o arquivo por e-mail.

## O que faz

- consulta documentos do Zapsign pelo período do mês anterior
- extrai informações importantes de cada documento
- monta um relatório em Excel (`.xlsx`)
- faz upload do arquivo para um bucket do Google Cloud Storage
- envia o relatório por e-mail usando SMTP

## Tecnologias

- Python
- Google Cloud Functions
- Google Cloud Storage
- API Zapsign
- SMTP para envio de e-mail
- pandas

## Como usar

1. Configure as variáveis de ambiente:

- `ZAPSIGN_TOKEN`: token Bearer para a API Zapsign
- `EMAIL_ORIGEM`: e-mail de origem para envio
- `EMAIL_SENHA_APP`: senha de app ou credencial SMTP do e-mail de origem
- `EMAIL_DESTINO`: e-mail destino do relatório
- `BUCKET_NAME`: nome do bucket do Google Cloud Storage

2. Instale as dependências:

```bash
pip install -r requirements.txt
```

3. Execute a função localmente ou faça deploy no Google Cloud Functions.

## Estrutura do projeto

- `main.py`: implementação da função e da lógica de relatório
- `requirements.txt`: dependências Python

## Segurança

O código utiliza variáveis de ambiente para carregar credenciais. Não armazene tokens, senhas ou chaves no repositório.

## Observações

O projeto é ideal para provadores de conceito de automatização de relatórios e integração entre APIs, armazenamento em nuvem e envio de e-mail.
