import os
import time
import pandas as pd
from datetime import datetime
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from playwright.sync_api import sync_playwright
import subprocess

# ==========================================
# CONFIGURAÇÕES (PREENCHA O LINK DA SUA PLANILHA AQUI!)
# ==========================================
LINK_CSV = "https://docs.google.com/spreadsheets/d/14FmugxfQEPqruykwRqrOzRYBwcljSC-TucACqYSiPxM/edit?usp=sharing" 

USER_PORTAL = os.getenv("USER_PORTAL", "")
SENHA_PORTAL = os.getenv("SENHA_PORTAL", "")
SENHA_GMAIL = os.getenv("SENHA_GMAIL", "")
agora_str = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

# ==========================================
# 1. LEITURA VIA URL (Google Sheets)
# ==========================================
try:
    # Lê direto da nuvem, sem precisar de arquivo .xlsx no GitHub
    df = pd.read_csv(LINK_CSV)
    df.columns = [str(c).strip().upper() for c in df.columns]
    
    # Identifica colunas (ajuste se os nomes na sua planilha forem diferentes)
    col_protocolo = [c for c in df.columns if "PROTOCOLO" in c][0]
    col_status = "STATUS ATUAL" # Nome da coluna que o robô vai atualizar
    if col_status not in df.columns: df[col_status] = "Aguardando..."
    
    print(f"📥 Dados carregados via URL!")
except Exception as e:
    print(f"❌ Erro ao ler planilha online: {e}")
    exit(1)

# ==========================================
# 2. AUTOMAÇÃO (Playwright)
# ==========================================
dados_portal = {}
try:
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        page.goto("https://santanadeparnaiba.aprova.com.br/login")
        
        # Login
        page.get_by_placeholder("E-mail").fill(USER_PORTAL)
        page.get_by_placeholder("Digite sua senha").first.fill(SENHA_PORTAL)
        page.get_by_role("button", name="Entrar").click()
        page.wait_for_load_state("networkidle")
        
        page.goto("https://santanadeparnaiba.aprova.com.br/processos")
        page.wait_for_load_state("networkidle")
        
        # Extração
        linhas = page.locator("tbody tr").all()
        for linha in linhas:
            texto = linha.inner_text().strip()
            if texto:
                partes = [p.strip() for p in texto.split("\n") if p.strip()]
                if len(partes) >= 2:
                    dados_portal[partes[0].upper()] = partes[-1]
        browser.close()
except Exception as e:
    print(f"❌ Erro no robô: {e}")

# ==========================================
# 3. ATUALIZAÇÃO E SALVAMENTO
# ==========================================
for index, row in df.iterrows():
    proto = str(row[col_protocolo]).strip().upper()
    if proto in dados_portal:
        df.at[index, col_status] = dados_portal[proto]

# Salva o resultado em HTML para visualização
df.to_html("monitor_protocolos.html", index=False)

# Git push
subprocess.run(["git", "config", "user.name", "Robot"])
subprocess.run(["git", "config", "user.email", "robot@artesano.com"])
subprocess.run(["git", "add", "monitor_protocolos.html"])
subprocess.run(["git", "commit", "-m", f"🤖 Atualização: {agora_str}"])
subprocess.run(["git", "push"])
print("✨ GitHub atualizado!")
