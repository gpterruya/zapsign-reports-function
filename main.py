import functions_framework
import requests
import pandas as pd
import calendar
from datetime import datetime, timedelta
from google.cloud import storage
import smtplib
from email.message import EmailMessage
import os
import tempfile

# ============== CONFIGURAÇÕES ==============
BASE_API_URL = "https://api.zapsign.com.br/api/v1/docs/"
BEARER_TOKEN = os.getenv("ZAPSIGN_TOKEN", "").strip().strip('"').strip("'")
EMAIL_ORIGEM = os.getenv("EMAIL_ORIGEM")
EMAIL_SENHA_APP = os.getenv("EMAIL_SENHA_APP")
EMAIL_DESTINO = os.getenv("EMAIL_DESTINO")
BUCKET_NAME = os.getenv("BUCKET_NAME")

HEADERS = {"Authorization": f"Bearer {BEARER_TOKEN}"}


def fetch_documents(primeiro_dia, ultimo_dia):
    """Busca todos os documentos dentro do período configurado."""
    params = {
        "created_from": primeiro_dia.strftime("%Y-%m-%d"),
        "created_to": ultimo_dia.strftime("%Y-%m-%d"),
        "page": 1
    }
    all_results = []
    while True:
        response = requests.get(BASE_API_URL, headers=HEADERS, params=params)
        response.raise_for_status()
        data = response.json()
        all_results.extend(data.get("results", []))
        if data.get("next"):
            params["page"] += 1
        else:
            break
    return all_results


def fetch_document_details(doc_token):
    url = f"{BASE_API_URL}{doc_token}/"
    response = requests.get(url, headers=HEADERS)
    response.raise_for_status()
    return response.json()


def extract_data(details):
    answers = {a.get("variable"): a.get("value") for a in details.get("answers", [])}
    endereco = ", ".join(filter(None, [
        answers.get("logradouro", ""),
        answers.get("numero", ""),
        answers.get("bairro", ""),
        answers.get("cidade", ""),
        answers.get("estado", ""),
        answers.get("cep", "")
    ]))
    return {
        "Tutor": answers.get("nome_completo", "N/A"),
        "Pet adotado": answers.get("nome_animal", "N/A"),
        "Raça": answers.get("raca", "N/A"),
        "Telefone": answers.get("celular", "N/A"),
        "E-mail": answers.get("email", "N/A"),
        "Data adoção": answers.get("data", "N/A"),
        "Endereço completo": endereco if endereco else "N/A"
    }


def enviar_email_com_anexo(arquivo_path, arquivo_nome, nome_mes, ano):
    msg = EmailMessage()
    msg['Subject'] = f"Relatório de adoções - {nome_mes} {ano}"
    msg['From'] = EMAIL_ORIGEM
    msg['To'] = EMAIL_DESTINO
    msg.set_content(f"Olá!\n\nSegue o relatório de adoções referente a {nome_mes} de {ano}.\n\nAtenciosamente,\nEquipe Automação")

    with open(arquivo_path, 'rb') as f:
        msg.add_attachment(
            f.read(),
            maintype='application',
            subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            filename=arquivo_nome
        )

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(EMAIL_ORIGEM, EMAIL_SENHA_APP)
        smtp.send_message(msg)


@functions_framework.http
def gerar_relatorio(request):
    # Calcula mês anterior
    hoje = datetime.today()
    primeiro_dia_mes_atual = datetime(hoje.year, hoje.month, 1)
    ultimo_dia_mes_anterior = primeiro_dia_mes_atual - timedelta(days=1)
    mes = ultimo_dia_mes_anterior.month
    ano = ultimo_dia_mes_anterior.year
    nome_mes = calendar.month_name[mes]

    print(f"Buscando documentos de {nome_mes} de {ano}...")
    primeiro_dia = datetime(ano, mes, 1)
    ultimo_dia = datetime(ano, mes, calendar.monthrange(ano, mes)[1])

    # Busca e processa
    documents = fetch_documents(primeiro_dia, ultimo_dia)
    print(f"{len(documents)} documentos encontrados entre {primeiro_dia.date()} e {ultimo_dia.date()}.")

    all_data = []

    for i, doc in enumerate(documents, start=1):
        token = doc.get("token")
        print(f"Processando documento {i}/{len(documents)} - Token: {token}")
        try:
            details = fetch_document_details(token)
            data = extract_data(details)
            all_data.append(data)
            print(f"Documento {i} processado com sucesso: {data}")
        except Exception as e:
            print(f"Erro ao processar documento {i} - Token {token}: {e}")

    print(f"Processamento finalizado. {len(all_data)} documentos processados com sucesso.")

    # Cria Excel temporário
    df = pd.DataFrame(all_data)
    arquivo_nome = f"animais_adotados_{nome_mes.lower()}_{ano}.xlsx"
    temp_path = os.path.join(tempfile.gettempdir(), arquivo_nome)
    df.to_excel(temp_path, index=False)

    # Upload para Cloud Storage
    storage_client = storage.Client()
    bucket = storage_client.bucket(BUCKET_NAME)
    blob = bucket.blob(arquivo_nome)
    blob.upload_from_filename(temp_path)
    print(f"Arquivo enviado ao bucket {BUCKET_NAME} com sucesso.")

    # Envia e-mail com o anexo
    enviar_email_com_anexo(temp_path, arquivo_nome, nome_mes, ano)

    return "Relatório gerado e enviado com sucesso!", 200
