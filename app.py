import os
import re
import fitz  # PyMuPDF
import docx
import pandas as pd
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from datetime import datetime

# Função para extrair texto de arquivos PDF
def extract_text_from_pdf(file_path):
    text = ""
    try:
        with fitz.open(file_path) as pdf_document:
            for page_num in range(pdf_document.page_count):
                page = pdf_document.page(page_num)
                text += page.get_text()
    except Exception as e:
        print(f"Erro ao ler PDF {file_path}: {e}")
    return text

# Função para extrair texto de arquivos DOCX
def extract_text_from_docx(file_path):
    text = ""
    try:
        doc = docx.Document(file_path)
        for paragraph in doc.paragraphs:
            text += paragraph.text + "\n"
    except Exception as e:
        print(f"Erro ao ler DOCX {file_path}: {e}")
    return text

# Função para extrair informações relevantes
def extract_information(text, file_name):
    # Sigla (pelo nome do arquivo, entre colchetes)
    sigla_match = re.search(r'\[([A-Z]{2,})\]', file_name)
    sigla = sigla_match.group(1) if sigla_match else "Não encontrado"

    # Nome da pessoa
    nome_match = re.search(r'Autorização para (?:Estagiária - |MARÍLIA SOUSA PEREIRA - )?(.+?)(?: - Rota|\n|$)', text, re.IGNORECASE)
    if not nome_match:
        # Se não encontrar no texto, tentar pelo nome do arquivo
        nome_arquivo_match = re.search(r'\] (.+?) - ', file_name)
        nome = nome_arquivo_match.group(1).strip() if nome_arquivo_match else "Não encontrado"
    else:
        nome = nome_match.group(1).strip()

    # Número da Rota
    rota_match = re.search(r'Rota\s*(\d+)', text, re.IGNORECASE)
    if not rota_match:
        rota_match = re.search(r'ROTA\s*(\d+)', file_name, re.IGNORECASE)
    rota = rota_match.group(1) if rota_match else "Não encontrado"

    # Data de autorização
    data_match = re.search(
        r'(\d{1,2}\s+de\s+\w+\s+de\s+\d{4})', text, re.IGNORECASE)
    data_autorizacao = data_match.group(1).capitalize() if data_match else "Não encontrada"

    return {
        "Sigla": sigla,
        "Nome": nome,
        "Rota": rota,
        "Data de Autorização": data_autorizacao,
        "Arquivo": file_name
    }

# Função para exportar para Google Sheets
def export_to_google_sheets(dataframe):
    try:
        SERVICE_ACCOUNT_FILE = r"C:\Users\bugzln\Desktop\Script\LER OS ARQUIVOS\credentials.json"
        SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

        credentials = Credentials.from_service_account_file(
            SERVICE_ACCOUNT_FILE, scopes=SCOPES)

        service = build('sheets', 'v4', credentials=credentials)

        SPREADSHEET_ID = '1XJNPa5sKe9HQjIGsDJzLUN-0gLXk6Et_gOx7rDcrUhE'  # <<<< Coloca aqui o ID da sua planilha
        RANGE_NAME = 'ROTAS'

        values = [dataframe.columns.tolist()] + dataframe.values.tolist()

        body = {
            'values': values
        }

        service.spreadsheets().values().update(
            spreadsheetId=SPREADSHEET_ID,
            range=RANGE_NAME,
            valueInputOption='RAW',
            body=body
        ).execute()

        print("✔️ Exportação para o Google Sheets concluída com sucesso!")

    except Exception as e:
        print(f"❌ Erro ao exportar para o Google Sheets: {e}")

# Função principal
def main():
    folder_path = r"C:\Users\bugzln\Desktop\Script\LER OS ARQUIVOS\downloads"

    extracted_data = []

    for file_name in os.listdir(folder_path):
        if file_name.endswith(".pdf") or file_name.endswith(".docx"):
            file_path = os.path.join(folder_path, file_name)

            print(f"🔍 Processando arquivo: {file_name}")

            if file_name.endswith(".pdf"):
                text = extract_text_from_pdf(file_path)
            else:
                text = extract_text_from_docx(file_path)

            info = extract_information(text, file_name)
            extracted_data.append(info)

    df = pd.DataFrame(extracted_data)

    # Exporta para Excel
    excel_file = os.path.join(folder_path, "resultado.xlsx")
    try:
        df.to_excel(excel_file, index=False)
        print(f"✔️ Exportação para Excel concluída com sucesso: {excel_file}")
    except Exception as e:
        print(f"❌ Erro ao exportar para Excel: {e}")

    # Exporta para Google Sheets
    export_to_google_sheets(df)

if __name__ == "__main__":
    main()
