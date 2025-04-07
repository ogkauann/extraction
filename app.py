import os
import re
import io
import threading
import tkinter as tk
from tkinter import ttk, messagebox
from pathlib import Path
import pandas as pd
import docx
from PyPDF2 import PdfReader
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
import webbrowser

# Configurações iniciais para Ubuntu usando Path.home()
home = str(Path.home())
# Ajuste os caminhos conforme sua estrutura de pastas no Ubuntu
FOLDER_PATH = f"/home/kauan/extraction/downloads"  # Pasta para armazenar os arquivos baixados
CAMINHO_CREDENCIAIS = f"/home/kauan/extraction/credentials.json"

# ID da planilha e intervalo serão inseridos via interface
# DOWNLOAD_FOLDER será FOLDER_PATH
DOWNLOAD_FOLDER = FOLDER_PATH

# Escopos de autenticação
SCOPES = ['https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/spreadsheets']

# Cria a pasta de downloads se não existir
os.makedirs(DOWNLOAD_FOLDER, exist_ok=True)

# --- FUNÇÕES DE DOWNLOAD DO DRIVE ---
def baixar_arquivos_do_drive(pasta_id, drive_service):
    """Baixa arquivos PDF e DOCX da pasta do Drive, evitando duplicados."""
    print("Iniciando download dos arquivos do Drive...")
    query = f"'{pasta_id}' in parents and (mimeType='application/pdf' or mimeType='application/vnd.openxmlformats-officedocument.wordprocessingml.document')"
    response = drive_service.files().list(q=query,
                                           fields="nextPageToken, files(id, name)").execute()
    arquivos = response.get('files', [])
    baixados = []
    for arquivo in arquivos:
        nome_arquivo = arquivo['name']
        caminho_destino = os.path.join(DOWNLOAD_FOLDER, nome_arquivo)
        if os.path.exists(caminho_destino):
            print(f"Arquivo {nome_arquivo} já existe, pulando download.")
            baixados.append(caminho_destino)
            continue
        request = drive_service.files().get_media(fileId=arquivo['id'])
        with io.FileIO(caminho_destino, 'wb') as fh:
            downloader = MediaIoBaseDownload(fh, request)
            done = False
            while not done:
                status, done = downloader.next_chunk()
                print(f"Baixando {nome_arquivo}: {int(status.progress() * 100)}%")
        print(f"Arquivo salvo: {nome_arquivo}")
        baixados.append(caminho_destino)
    print("Download concluído!")
    return baixados

# --- FUNÇÕES DE EXTRAÇÃO DE TEXTO ---
def extract_text_from_docx(file_path):
    try:
        doc = docx.Document(file_path)
        return "\n".join([para.text for para in doc.paragraphs])
    except Exception as e:
        print(f"Erro ao ler DOCX {file_path}: {e}")
        return ""

def extract_text_from_pdf(file_path):
    try:
        reader = PdfReader(file_path)
        text = ""
        for page in reader.pages:
            extracted = page.extract_text()
            if extracted:
                text += extracted + "\n"
        return text
    except Exception as e:
        print(f"Erro ao ler PDF {file_path}: {e}")
        return ""

# --- FUNÇÃO DE EXTRAÇÃO DE DADOS ---
def extract_info(file_path, text):
    """
    Extrai as informações:
      - Sigla, Nome e Rota: do nome do arquivo
      - Data de Autorização: do conteúdo (texto)
    """
    file_name = os.path.basename(file_path)
    base_name = os.path.splitext(file_name)[0]

    # Extrair sigla: pega o que está entre colchetes
    sigla_match = re.search(r"\[(.*?)\]", base_name)
    sigla = sigla_match.group(1).strip() if sigla_match else "Não encontrado"

    # Extrair nome: tenta pegar do nome do arquivo após o fechamento do colchete até "- Rota"
    nome_match = re.search(r"\]\s*(.*?)\s*-\s*Rota", base_name, re.IGNORECASE)
    nome = nome_match.group(1).strip() if nome_match else "Não encontrado"

    # Extrair rota: do nome do arquivo (após "Rota") ou como fallback do texto
    rota_match = re.search(r"Rota\s*(\d+)", base_name, re.IGNORECASE)
    if not rota_match:
        rota_match = re.search(r"Rota\s*(\d+)", text, re.IGNORECASE)
    rota = rota_match.group(1).strip() if rota_match else "Não encontrado"

    # Extrair data de autorização do conteúdo (prioritário)
    data_match = re.search(r"(\d{1,2}\s+de\s+[a-zA-ZçÇãõÁÉÍÓÚ]+?\s+de\s+\d{4})", text, re.IGNORECASE)
    data_autorizacao = data_match.group(1).strip() if data_match else "Não encontrada"

    return {
        "Sigla": sigla,
        "Nome": nome,
        "Rota": rota,
        "Data de Autorização": data_autorizacao,
        "Arquivo": file_name
    }

# --- FUNÇÕES DE EXPORTAÇÃO ---
def export_to_excel(data, output_file):
    df = pd.DataFrame(data)
    df.to_excel(output_file, index=False)
    print(f"Arquivo Excel salvo em: {output_file}")

def export_to_google_sheets(data, spreadsheet_id, sheet_range):
    df = pd.DataFrame(data)
    values = [df.columns.tolist()] + df.values.tolist()
    body = {"values": values}
    service = build("sheets", "v4", credentials=Credentials.from_service_account_file(CAMINHO_CREDENCIAIS, scopes=SCOPES))
    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=sheet_range,
        valueInputOption="RAW",
        body=body
    ).execute()
    print("Dados exportados para o Google Sheets com sucesso!")

# --- FUNÇÃO PRINCIPAL ---
def processar():
    # Inicializa o serviço do Drive
    creds = Credentials.from_service_account_file(CAMINHO_CREDENCIAIS, scopes=SCOPES)
    drive_service_local = build("drive", "v3", credentials=creds)
    
    # Baixa os arquivos da pasta do Drive (ID informado pela interface)
    pasta_id = entrada_pasta.get().strip()
    if not pasta_id:
        messagebox.showerror("Erro", "Informe o ID da pasta do Drive.")
        return
    arquivos = baixar_arquivos_do_drive(pasta_id, drive_service_local)
    if not arquivos:
        messagebox.showinfo("Aviso", "Nenhum arquivo encontrado para processar.")
        return
    
    # Processa cada arquivo para extrair os dados
    resultados = []
    for file_path in arquivos:
        if file_path.lower().endswith(".docx"):
            text = extract_text_from_docx(file_path)
        elif file_path.lower().endswith(".pdf"):
            text = extract_text_from_pdf(file_path)
        else:
            continue
        info = extract_info(file_path, text)
        resultados.append(info)
        print(f"Processado: {os.path.basename(file_path)}")
    
    if not resultados:
        messagebox.showinfo("Aviso", "Nenhum dado extraído.")
        return
    
    # Exporta para Excel
    # excel_file = os.path.join(DOWNLOAD_FOLDER, "resultado.xlsx")
    # export_to_excel(resultados, excel_file)
    
    # Exporta para Google Sheets (usando os dados da interface)
    planilha_id = entrada_planilha.get().strip()
    sheet_range = entrada_aba.get().strip()  # exemplo: "Página1!A1"
    if not planilha_id or not sheet_range:
        messagebox.showerror("Erro", "Informe o ID da planilha e o intervalo (ex: Página1!A1).")
        return
    export_to_google_sheets(resultados, planilha_id, sheet_range)
    
    messagebox.showinfo("Concluído", "Processamento e exportação concluídos com sucesso!")

# --- INTERFACE TKINTER ---
janela = tk.Tk()
janela.title("Extração de Dados - Drive para Excel & Sheets")
janela.geometry("700x500")

frame = ttk.Frame(janela, padding=10)
frame.pack(fill=tk.BOTH, expand=True)

ttk.Label(frame, text="ID da Pasta do Drive:").pack(anchor=tk.W)
entrada_pasta = ttk.Entry(frame, width=80)
entrada_pasta.pack(fill=tk.X, pady=5)

ttk.Label(frame, text="ID da Planilha do Google Sheets:").pack(anchor=tk.W)
entrada_planilha = ttk.Entry(frame, width=80)
entrada_planilha.pack(fill=tk.X, pady=5)

ttk.Label(frame, text="Intervalo da Aba (ex: Página1!A1):").pack(anchor=tk.W)
entrada_aba = ttk.Entry(frame, width=80)
entrada_aba.pack(fill=tk.X, pady=5)

botao_iniciar = ttk.Button(frame, text="Iniciar Processo", command=lambda: threading.Thread(target=processar).start())
botao_iniciar.pack(pady=10)

janela.mainloop()
