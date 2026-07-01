import os
import time
import pandas as pd
from datetime import datetime
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ==========================================
# CONFIGURAÇÕES E LINKS
# ==========================================
EMAIL_REMETENTE = "mikaellevictoria2017@gmail.com"
EMAIL_DESTINATARIOS = ["santos.micaelle2006@gmail.com"]
LINK_FORM = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRh-7SIMziaShR1rqLpSnBabRJAIceLSZ6dO0zklOcOg_twfc9G6cwdRGQk1vL2y6lniAmH0mSh6Xw1/pub?gid=1314499551&single=true&output=csv"

USER_PORTAL = os.getenv("USER_PORTAL", "")
SENHA_PORTAL = os.getenv("SENHA_PORTAL", "")
SENHA_GMAIL = os.getenv("SENHA_GMAIL", "")

agora_str = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

# ==========================================
# 1. LEITURA DOS DADOS
# ==========================================
df = pd.read_csv(LINK_FORM)
df.columns = [str(c).strip().upper() for c in df.columns]
col_protocolo = [c for c in df.columns if "PROTOCOLO" in c or "NUMER" in c or "PROCESSO" in c][0]
col_ativo = [c for c in df.columns if "ATIVO" in c][0] if any("ATIVO" in c for c in df.columns) else None

# Garante colunas padrão
for col in ["STATUS ATUAL", "STATUS ANTIGO"]:
    if col not in df.columns: df[col] = "Aguardando..."

# ==========================================
# 2. AUTOMAÇÃO NO PORTAL
# ==========================================
options = Options()
options.add_argument("--headless=new")
driver = webdriver.Chrome(options=options)
wait = WebDriverWait(driver, 30)

dados_portal = {}
driver.get("https://santanadeparnaiba.aprova.com.br/login")
time.sleep(10)
# (Se precisar de login, inserir aqui...)
driver.get("https://santanadeparnaiba.aprova.com.br/processos")
time.sleep(15)

linhas = driver.find_elements(By.TAG_NAME, "tr")
for linha in linhas:
    partes = [p.text.strip() for p in linha.find_elements(By.TAG_NAME, "td") if p.text.strip()]
    if len(partes) >= 2:
        dados_portal[partes[0].upper().strip()] = partes
driver.quit()

# ==========================================
# 3. MAPEAMENTO (CONFORME A ESTRUTURA DO PORTAL)
# ==========================================
processos_alterados = []
for index, row in df.iterrows():
    if col_ativo and str(row.get(col_ativo, "")).strip().upper() != "SIM": continue
    proto = str(row[col_protocolo]).strip().upper()
    
    if proto in dados_portal:
        info = dados_portal[proto]
        status_novo = info[6] if len(info) > 6 else "Sem status"
        status_antigo = str(df.at[index, "STATUS ATUAL"])
        
        if status_antigo != status_novo and "Aguardando" not in status_antigo:
            df.at[index, "STATUS ANTIGO"] = status_antigo
            processos_alterados.append({'p': proto, 'a': status_antigo, 'n': status_novo})
        
        df.at[index, "STATUS ATUAL"] = status_novo
        df.at[index, "ASSUNTO / TIPO"] = info[1]
        df.at[index, "REQUERENTE / PROPRIETÁRIO"] = info[2]
        df.at[index, "ENDEREÇO"] = info[3]
        df.at[index, "DATA DE ATUALIZAÇÃO DO PORTAL"] = info[4]
        df.at[index, "MODIFICADO POR ÚLTIMO"] = info[7] if len(info) > 7 else "N/A"
        df.at[index, "CÓDIGO ATUALIZADO EM:"] = agora_str

# ==========================================
# 4. ORDENAÇÃO E LIMPEZA (B a K)
# ==========================================
ordem_fixa = [
    "PROTOCOLO ATIVO?", "NÚMERO DO PROTOCOLO", "ASSUNTO / TIPO", "REQUERENTE / PROPRIETÁRIO", 
    "ENDEREÇO", "STATUS ANTIGO", "STATUS ATUAL", "MODIFICADO POR ÚLTIMO", 
    "DATA DE ATUALIZAÇÃO DO PORTAL", "CÓDIGO ATUALIZADO EM:"
]

# Remove colunas desnecessárias
if "ÚLTIMA AÇÃO" in df.columns: df.drop(columns=["ÚLTIMA AÇÃO"], inplace=True)
if "CARIMBO DE DATA/HORA" in df.columns: df.drop(columns=["CARIMBO DE DATA/HORA"], inplace=True)

df = df.reindex(columns=ordem_fixa)
df.to_html("monitor_protocolos.html", index=False, encoding="utf-8-sig")

# E-mail de alerta...
