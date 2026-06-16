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
# CONFIGURAÇÕES E CHAVES
# ==========================================
EMAIL_REMETENTE = "mikaellevictoria2017@gmail.com"
EMAIL_DESTINATARIOS = ["santos.micaelle2006@gmail.com"]
LINK_PLANILHA = "SUA_URL_DA_PLANILHA_AQUI" # Certifique-se de preencher isso

nome_planilha = "monitor_protocolos.xlsx"
nome_aba = "Santana de Parnaíba"

USER_PORTAL = os.getenv("USER_PORTAL", "")
SENHA_PORTAL = os.getenv("SENHA_PORTAL", "")
SENHA_GMAIL = os.getenv("SENHA_GMAIL", "")

agora_str = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
print(f"===== INICIANDO MONITORAMENTO BLINDADO VIA PLAYWRIGHT: {agora_str} =====")

# ==========================================
# 1. LEITURA DA PLANILHA (Sua lógica robusta)
# ==========================================
try:
    df = pd.read_excel(nome_planilha, sheet_name=nome_aba)
    df.columns = [str(c).strip().upper() for c in df.columns]
    
    col_protocolo = [c for c in df.columns if "PROTOCOLO" in c or "NUMER" in c or "PROCESSO" in c][0]
    col_status = [c for c in df.columns if "STATUS" in c or "SITUA" in c][0] if any("STATUS" in c or "SITUA" in c for c in df.columns) else "STATUS ATUAL"
    col_modificado = [c for c in df.columns if "MODIF" in c or "DATA" in c][0] if any("MODIF" in c or "DATA" in c for c in df.columns) else "MODIFICADO EM"
    col_ativo = [c for c in df.columns if "ATIVO" in c][0] if any("ATIVO" in c for c in df.columns) else None

    print(f"📋 Protocolos carregados da planilha.")
except Exception as e:
    print(f"❌ Erro ao ler Excel: {e}")
    exit(1)

# ==========================================
# 2. AUTOMAÇÃO (Motor Playwright)
# ==========================================
dados_portal = {}

try:
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        
        print("🌐 Acessando o portal...")
        page.goto("https://santanadeparnaiba.aprova.com.br/login")
        
        # Login Preciso e Estável
        print("🔑 Realizando Login...")
        page.get_by_placeholder("E-mail").fill(USER_PORTAL)
        page.get_by_placeholder("Digite sua senha").first.fill(SENHA_PORTAL)
        page.get_by_role("button", name="Entrar").click()
        page.wait_for_load_state("networkidle")
        
        print("📂 Navegando para processos...")
        page.goto("https://santanadeparnaiba.aprova.com.br/processos")
        page.wait_for_load_state("networkidle")
        
        # Extração inteligente (Substitui o Selenium)
        print("🔍 Capturando dados da tabela...")
        linhas = page.locator("tbody tr").all()
        for linha in linhas:
            texto = linha.inner_text().strip()
            if texto:
                partes = [p.strip() for p in texto.split("\n") if p.strip()]
                if len(partes) >= 2:
                    dados_portal[partes[0].upper()] = partes[-1] # Pega o último elemento como status
        
        browser.close()
        print(f"✅ Extraídos {len(dados_portal)} registros.")

except Exception as e:
    print(f"❌ Falha no robô: {e}")

# ==========================================
# 3. COMPARAÇÃO E ATUALIZAÇÃO (Sua lógica)
# ==========================================
processos_alterados = []
for index, row in df.iterrows():
    if col_ativo and str(row.get(col_ativo, "")).strip().upper() != "SIM":
        continue
    
    protocolo = str(row[col_protocolo]).strip().upper()
    status_antigo = str(row.get(col_status, "")).strip()
    status_novo = dados_portal.get(protocolo)
    
    if status_novo and status_antigo != status_novo:
        processos_alterados.append({'protocolo': protocolo, 'antigo': status_antigo, 'novo': status_novo})
        df.at[index, col_status] = status_novo
        df.at[index, col_modificado] = agora_str

# ==========================================
# 4. SALVAMENTO E GITHUB (Sua lógica)
# ==========================================
if processos_alterados:
    with pd.ExcelWriter(nome_planilha, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df.to_excel(writer, sheet_name=nome_aba, index=False)
    
    # Git push (Mantido conforme seu código original)
    subprocess.run(["git", "config", "user.name", "Automated Robot"])
    subprocess.run(["git", "config", "user.email", "robot@artesano.com"])
    subprocess.run(["git", "add", nome_planilha])
    subprocess.run(["git", "commit", "-m", f"🤖 Atualização automática: {agora_str}"])
    subprocess.run(["git", "push"])
    
    # (Adicione aqui a sua lógica de E-mail que já funciona!)
    print("✨ Sistema atualizado e GitHub sincronizado!")
