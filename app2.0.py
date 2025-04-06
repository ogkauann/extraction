import os
import re
import io
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
import pandas as pd
import docx
import textract
import PyPDF2
import pytesseract
from PIL import Image
import fitz  # PyMuPDF
import threading
import webbrowser

# Configurações iniciais
SCOPES = ['https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/spreadsheets']
DOWNLOAD_FOLDER = 'downloads'
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Autenticação Google
credentials = Credentials.from_service_account_file('credentials.json', scopes=SCOPES)
drive_service = build('drive', 'v3', credentials=credentials)
sheets_service = build('sheets', 'v4', credentials=credentials)

# Expressões regulares
regex_nome = re.compile(r'Nome\s*[:\-]?\s*(.+)', re.IGNORECASE)
regex_orgao = re.compile(r'(?:\bOrgao|\b\u00d3rgão)\s*[:\-]?\s*(.+)', re.IGNORECASE)
regex_rota = re.compile(r'Rota\s*[:\-]?\s*([\w\s\d]+)', re.IGNORECASE)
regex_ano = re.compile(r'(?:Ano de Autorizacao|Ano da Autorizacao|Ano)\s*[:\-]?\s*(\d{4})', re.IGNORECASE)

# Funções utilitárias
def limpar_nome_arquivo(nome):
    return re.sub(r'[\\/*?%:"<>|]', "_", nome)

def log(text):
    log_text.configure(state='normal')
    log_text.insert(tk.END, text + '\n')
    log_text.configure(state='disabled')
    log_text.yview(tk.END)
    janela.update()

def abrir_planilha(sheet_url):
    webbrowser.open(sheet_url)

# Download dos arquivos do Google Drive
def baixar_arquivos(pasta_id):
    if not os.path.exists(DOWNLOAD_FOLDER):
        os.makedirs(DOWNLOAD_FOLDER)

    page_token = None
    while True:
        response = drive_service.files().list(q=f"'{pasta_id}' in parents and trashed = false",
                                              spaces='drive',
                                              fields='nextPageToken, files(id, name, mimeType)',
                                              pageToken=page_token).execute()
        for file in response.get('files', []):
            if file.get('mimeType') != 'application/vnd.google-apps.folder':
                file_id = file.get('id')
                file_name = limpar_nome_arquivo(file.get('name'))
                request = drive_service.files().get_media(fileId=file_id)
                filepath = os.path.join(DOWNLOAD_FOLDER, file_name)

                fh = io.FileIO(filepath, 'wb')
                downloader = MediaIoBaseDownload(fh, request)
                done = False
                while not done:
                    status, done = downloader.next_chunk()
                log(f'Arquivo baixado: {file_name}')
        page_token = response.get('nextPageToken', None)
        if page_token is None:
            break

# Extração de texto dos arquivos
def extrair_texto_pdf(path):
    texto = ''
    try:
        with open(path, 'rb') as f:
            reader = PyPDF2.PdfReader(f)
            for page in reader.pages:
                texto += page.extract_text() or ''
    except:
        try:
            doc = fitz.open(path)
            for page in doc:
                pix = page.get_pixmap()
                img = Image.open(io.BytesIO(pix.tobytes()))
                texto += pytesseract.image_to_string(img)
        except Exception as e:
            log(f'Erro no OCR PDF: {e}')
    return texto

def extrair_texto_docx(path):
    doc = docx.Document(path)
    return '\n'.join([p.text for p in doc.paragraphs])

def extrair_texto_doc(path):
    try:
        return textract.process(path).decode('utf-8')
    except Exception as e:
        log(f'Erro ao ler .doc: {e}')
        return ''

def extrair_dados(texto, nome_arquivo):
    nome = orgao = rota = ano = ''

    # Primeiro tenta extrair do texto do arquivo
    if match := regex_nome.search(texto):
        nome = match.group(1).strip()
    if match := regex_orgao.search(texto):
        orgao = match.group(1).strip()
    if match := regex_rota.search(texto):
        rota = match.group(1).strip()
    if match := regex_ano.search(texto):
        ano = match.group(1).strip()

    # Se não encontrar no texto, tenta extrair do nome do arquivo baseado nas regras que você criou
    if not any([nome, orgao, rota, ano]):
        nome_arquivo_limpo = nome_arquivo.replace('_', ' ')

        # Quebrar nome do arquivo em partes por hífen
        partes = nome_arquivo_limpo.split('-')

        if len(partes) >= 4:
            # Órgão: parte após o terceiro hífen
            orgao = partes[3].strip()

            # Verificar se há uma vírgula após a sigla do órgão
            restante = '-'.join(partes[4:]).strip()
            if ',' in restante:
                # Divide pelo primeiro uso da vírgula
                antes_virgula, depois_virgula = restante.split(',', 1)

                # Nome do responsável fica na parte após a vírgula
                nome = depois_virgula.strip()

                # Opcional: a parte antes da vírgula pode conter outras informações, ignoramos por enquanto
            else:
                # Se não tiver vírgula, usar o restante como possível nome
                nome = restante.strip()

            # Procurar rota na string inteira
            rota_match = re.search(r'ROTA\s*([A-Za-z0-9]+)', nome_arquivo_limpo, re.IGNORECASE)
            if rota_match:
                rota = rota_match.group(1).strip()

            # Procurar o último número de 4 dígitos como o ano
            ano_match = re.findall(r'\b\d{4}\b', nome_arquivo_limpo)
            if ano_match:
                ano = ano_match[-1]  # Pega o último encontrado, que geralmente é o ano

    return nome, orgao, rota, ano


# Processamento dos arquivos e exportação
def processar_e_exportar():
    pasta_id = entrada_pasta.get().strip()
    aba_nome = entrada_aba.get().strip()
    planilha_id = entrada_planilha.get().strip()

    if not pasta_id or not aba_nome or not planilha_id:
        messagebox.showerror("Erro", "Preencha todos os campos!")
        return

    botao_iniciar.config(state='disabled')
    threading.Thread(target=lambda: executar_processo(pasta_id, aba_nome, planilha_id)).start()

def executar_processo(pasta_id, aba_nome, planilha_id):
    log('Iniciando processo...')

    # Etapa 1: Download
    baixar_arquivos(pasta_id)

    # Etapa 2: Processamento
    dados = []
    for arquivo in os.listdir(DOWNLOAD_FOLDER):
        caminho = os.path.join(DOWNLOAD_FOLDER, arquivo)
        if os.path.isfile(caminho):
            ext = os.path.splitext(arquivo)[1].lower()
            if ext == '.pdf':
                texto = extrair_texto_pdf(caminho)
            elif ext == '.docx':
                texto = extrair_texto_docx(caminho)
            elif ext == '.doc':
                texto = extrair_texto_doc(caminho)
            else:
                continue

            nome, orgao, rota, ano = extrair_dados(texto, arquivo)
            dados.append([nome, orgao, rota, ano])
            log(f'Processado: {arquivo}')

    df = pd.DataFrame(dados, columns=['Nome', 'Órgão', 'Rota', 'Ano de Autorizacao'])
    df = df.drop_duplicates(subset=['Nome', 'Órgão', 'Rota', 'Ano de Autorizacao'])

    try:
        sheets_service.spreadsheets().get(spreadsheetId=planilha_id).execute()
    except:
        spreadsheet = {
            'properties': {'title': 'Planilha Extração Dados'}
        }
        planilha = sheets_service.spreadsheets().create(body=spreadsheet, fields='spreadsheetId').execute()
        planilha_id = planilha.get('spreadsheetId')

    planilha = sheets_service.spreadsheets().get(spreadsheetId=planilha_id).execute()
    abas = [s['properties']['title'] for s in planilha.get('sheets', [])]

    if aba_nome not in abas:
        sheet_body = {
            'requests': [{
                'addSheet': {
                    'properties': {'title': aba_nome}
                }
            }]
        }
        sheets_service.spreadsheets().batchUpdate(
            spreadsheetId=planilha_id,
            body=sheet_body
        ).execute()

    valores = [df.columns.tolist()] + df.values.tolist()

    sheets_service.spreadsheets().values().update(
        spreadsheetId=planilha_id,
        range=f'{aba_nome}!A1',
        valueInputOption='RAW',
        body={'values': valores}
    ).execute()

    url_planilha = f'https://docs.google.com/spreadsheets/d/{planilha_id}'
    log('Processo finalizado!')
    log(f'Planilha atualizada: {url_planilha}')
    abrir_planilha(url_planilha)
    botao_iniciar.config(state='normal')

# Interface gráfica
janela = tk.Tk()
janela.title("Extração de Dados - Google Drive para Planilha")
janela.geometry("700x500")

frame = ttk.Frame(janela, padding=10)
frame.pack(fill=tk.BOTH, expand=True)

ttk.Label(frame, text="ID da Pasta do Drive:").pack(anchor=tk.W)
entrada_pasta = ttk.Entry(frame, width=80)
entrada_pasta.pack(fill=tk.X)

ttk.Label(frame, text="ID da Planilha:").pack(anchor=tk.W)
entrada_planilha = ttk.Entry(frame, width=80)
entrada_planilha.pack(fill=tk.X)

ttk.Label(frame, text="Nome da Aba:").pack(anchor=tk.W)
entrada_aba = ttk.Entry(frame, width=80)
entrada_aba.pack(fill=tk.X)

botao_iniciar = ttk.Button(frame, text="Iniciar Processo", command=processar_e_exportar)
botao_iniciar.pack(pady=10)

log_text = tk.Text(frame, height=15, state='disabled')
log_text.pack(fill=tk.BOTH, expand=True)

janela.mainloop()
