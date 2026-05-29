import os
import time
import io
import requests
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
from selenium.webdriver.common.action_chains import ActionChains

# ==========================================
# CONFIGURAÇÕES E LINKS
# ==========================================
EMAIL_REMETENTE = "mikaellevictoria2017@gmail.com"
EMAIL_DESTINATARIOS = ["santos.micaelle2006@gmail.com"]

# 🎯 INSIRA SEU LINK DO ONEDRIVE/SHAREPOINT AQUI DENTRO DAS ASPAS:
LINK_PLANILHA_NUVEM = "https://artesanourbanismo-my.sharepoint.com/:x:/g/personal/mvitoria_artesanourbanismo_com_br/IQDgXvD3n6RTRZsZo63IiGBXAeSMCTvv1qBTTDNAD3d1_jE?e=ODW12H"

nome_aba = "Santana de Parnaíba"

# Chaves de segurança guardadas no GitHub Secrets
USER_PORTAL = os.getenv("USER_PORTAL", "")
SENHA_PORTAL = os.getenv("SENHA_PORTAL", "")
SENHA_GMAIL = os.getenv("SENHA_GMAIL", "")

agora_str = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
print(f"===== INICIANDO MONITORAMENTO VIA LINK DIRECT: {agora_str} =====")

# ==========================================
# 1. LEITURA DA PLANILHA VIA REQUESTS (BURLANDO O 403)
# ==========================================
try:
    # Ajusta o link para forçar o download do arquivo bruto se for link comum
    url_download = LINK_PLANILHA_NUVEM
    if "sharepoint.com" in url_download and "download=1" not in url_download:
        if "?" in url_download:
            url_download += "&download=1"
        else:
            url_download += "?download=1"

    # Cabeçalho para o site achar que é uma pessoa comum acessando
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    }
    
    print("📥 Baixando arquivo da nuvem...")
    resposta = requests.get(url_download, headers=headers)
    resposta.raise_for_status() # Dispara erro caso não consiga acessar
    
    # Carrega os bytes do arquivo baixado para o Pandas
    df = pd.read_excel(io.BytesIO(resposta.content), sheet_name=nome_aba)
    
    colunas_originais = list(df.columns)
    df.columns = [str(c).strip().upper() for c in df.columns]
    print(f"📥 Planilha carregada com sucesso! Colunas: {list(df.columns)}")
    
    col_protocolo = [c for c in df.columns if "PROTOCOLO" in c or "NUMER" in c or "PROCESSO" in c][0]
    col_ativo = [c for c in df.columns if "ATIVO" in c][0] if any("ATIVO" in c for c in df.columns) else None
    col_status = [c for c in df.columns if "STATUS" in c or "SITUA" in c][0] if any("STATUS" in c for c in df.columns) else "STATUS ATUAL"
    col_modificado = [c for c in df.columns if "MODIF" in c or "DATA" in c][0] if any("MODIF" in c for c in df.columns) else "MODIFICADO EM"
    col_acao = [c for c in df.columns if "AÇÃO" in c or "ACAO" in c][0] if any("AÇÃO" in c for c in df.columns) else "ÚLTIMA AÇÃO"

    protocolos_verificar = df[col_protocolo].dropna().astype(str).tolist()
    print(f"🔍 Protocolos localizados para checagem: {protocolos_verificar}")

except Exception as e:
    print(f"❌ Erro ao ler a planilha direto do link: {e}")
    exit(1)

# ==========================================
# 2. RASPAGEM DOS DADOS NO PORTAL
# ==========================================
dados_portal = {}
options = Options()
options.add_argument("--headless=new") 
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--window-size=1920,1080")

driver = webdriver.Chrome(options=options)
wait = WebDriverWait(driver, 35)
actions = ActionChains(driver)

try:
    print("🌐 Acessando o portal de Santana de Parnaíba...")
    driver.get("https://santanadeparnaiba.aprova.com.br/login")
    time.sleep(7)
    
    inputs = wait.until(EC.presence_of_all_elements_located((By.TAG_NAME, "input")))
    if len(inputs) >= 2:
        inputs[0].send_keys(USER_PORTAL)
        inputs[1].send_keys(SENHA_PORTAL)
        botao = driver.find_element(By.CSS_SELECTOR, "button[type='submit']")
        actions.move_to_element(botao).click().perform()
        
    time.sleep(15)
    driver.get("https://santanadeparnaiba.aprova.com.br/processos")
    time.sleep(10)
    
    linhas = driver.find_elements(By.XPATH, "//tbody/tr | //tr | //div[contains(@class, 'linha')]")
    palavras_status = ['ANÁLISE', 'DEFERIDO', 'INDEFERIDO', 'COMUNIQUE-SE', 'AGUARDANDO', 'TRIAGEM', 'CONCLUÍDO', 'EMITIDO', 'CORREÇÃO', 'PENDENTE', 'PROCESSO']
    
    for linha in linhas:
        texto_linha = linha.text.strip()
        if texto_linha:
            partes = [p.strip() for p in texto_linha.split("\n") if p.strip()]
            if len(partes) >= 2:
                protocolo_web = partes[0].upper()
                if any(char.isdigit() for char in protocolo_web):
                    status_web = ""
                    for parte in reversed(partes[1:]):
                        if any(termo in parte.upper() for termo in palavras_status):
                            status_web = parte
                            break
                    if not status_web:
                        status_web = partes[-1]
                    dados_portal[protocolo_web] = status_web

    print(f"✅ Extraídos {len(dados_portal)} registros do site.")
except Exception as e:
    print(f"❌ Falha na automação do portal: {e}")
finally:
    driver.quit()

# ==========================================
# 3. ATUALIZAÇÃO RESTRITA DA PLANILHA
# ==========================================
processos_alterados = []

for index, row in df.iterrows():
    if col_ativo and str(row.get(col_ativo, "")).strip().upper() != "SIM":
        continue
        
    protocolo_planilha = str(row[col_protocolo]).strip().upper()
    status_antigo = str(row.get(col_status, "")).strip()
    
    status_novo = None
    for k, v in dados_portal.items():
        if protocolo_planilha in k or k in protocolo_planilha:
            status_novo = v
            break
            
    if status_novo and status_antigo != status_novo:
        print(f"⚠️ MUDANÇA DETECTADA NO PROTOCOLO: {protocolo_planilha}")
        processos_alterados.append({'protocolo': protocolo_planilha, 'antigo': status_antigo, 'novo': status_novo})
        
        df.at[index, col_status] = status_novo
