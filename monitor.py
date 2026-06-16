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
from selenium.webdriver.common.action_chains import ActionChains

# ==========================================
# CONFIGURAÇÕES E LINKS
# ==========================================
EMAIL_REMETENTE = "mikaellevictoria2017@gmail.com"
EMAIL_DESTINATARIOS = ["santos.micaelle2006@gmail.com"]
LINK_ENTRADA_GOOGLE_FORMS = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRh-7SIMziaShR1rqLpSnBabRJAIceLSZ6dO0zklOcOg_twfc9G6cwdRGQk1vL2y6lniAmH0mSh6Xw1/pub?gid=1314499551&single=true&output=csv"

USER_PORTAL = os.getenv("USER_PORTAL", "")
SENHA_PORTAL = os.getenv("SENHA_PORTAL", "")
SENHA_GMAIL = os.getenv("SENHA_GMAIL", "")

agora_str = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
print(f"===== INICIANDO MONITORAMENTO: {agora_str} =====")

# ==========================================
# 1. LEITURA DOS DADOS DO FORMULÁRIO
# ==========================================
try:
    df = pd.read_csv(LINK_ENTRADA_GOOGLE_FORMS)
    df.columns = [str(c).strip().upper() for c in df.columns]
    
    col_protocolo = [c for c in df.columns if "PROTOCOLO" in c or "NUMER" in c or "PROCESSO" in c][0]
    col_ativo = [c for c in df.columns if "ATIVO" in c][0] if any("ATIVO" in c for c in df.columns) else None
    
    if "STATUS ATUAL" not in df.columns: df["STATUS ATUAL"] = "Aguardando primeira checagem..."
    if "ÚLTIMA AÇÃO" not in df.columns: df["ÚLTIMA AÇÃO"] = "Nenhuma"
    if "MODIFICADO EM" not in df.columns: df["MODIFICADO EM"] = agora_str
        
    col_status = "STATUS ATUAL"
    col_acao = "ÚLTIMA AÇÃO"
    col_modificado = "MODIFICADO EM"
    protocolos_verificar = df[col_protocolo].dropna().astype(str).tolist()
except Exception as e:
    print(f"❌ Erro na leitura do Forms: {e}")
    exit(1)

# ==========================================
# 2. AUTOMAÇÃO NO PORTAL DA PREFEITURA
# ==========================================
dados_portal = {}
options = Options()
options.add_argument("--headless=new") 
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")

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
    
    linhas = driver.find_elements(By.XPATH, "//tbody/tr | //tr")
    for linha in linhas:
        texto_linha = linha.text.strip()
        if texto_linha:
            partes = [p.strip() for p in texto_linha.split("\n") if p.strip()]
            if len(partes) >= 1:
                protocolo_web = partes[0].upper().strip()
                dados_portal[protocolo_web] = partes
except Exception as e:
    print(f"❌ Falha na automação: {e}")
finally:
    driver.quit()

# ==========================================
# 3. ATUALIZAÇÃO DA BASE DE DADOS
# ==========================================
processos_alterados = []

for index, row in df.iterrows():
    if col_ativo and str(row.get(col_ativo, "")).strip().upper() != "SIM": continue
    protocolo_planilha = str(row[col_protocolo]).strip().upper()
    status_antigo = str(row.get(col_status, "")).strip()
    
    dados_novos = None
    for k, v in dados_portal.items():
        if protocolo_planilha in k or k in protocolo_planilha:
            dados_novos = v  
            break
            
    if dados_novos:
        status_novo = dados_novos[1] if len(dados_novos) > 1 else dados_novos[0]
        df.at[index, col_status] = status_novo
        df.at[index, col_acao] = status_novo
        df.at[index, col_modificado] = agora_str
        
        # Insere os dados na sequência exata das colunas do portal
        for i, valor in enumerate(dados_novos):
            df.at[index, f"COLUNA_{i+1}"] = valor
        
        if status_antigo != status_novo and "Aguardando" not in status_antigo:
            processos_alterados.append({'protocolo': protocolo_planilha, 'antigo': status_antigo, 'novo': status_novo})

# ==========================================
# 4. SALVAMENTO EM FORMATO TABELA HTML
# ==========================================
# Transforma a tabela de dados num formato HTML que o Google Sheets lê perfeitamente
df.to_html("monitor_protocolos.html", index=False, encoding="utf-8-sig")
print("💾 Base salva com sucesso em monitor_protocolos.html!")

if processos_alterados and SENHA_GMAIL:
    try:
        msg = MIMEMultipart('alternative')
        msg['From'] = EMAIL_REMETENTE
        msg['To'] = ", ".join(EMAIL_DESTINATARIOS)
        msg['Subject'] = "⚠️ Alerta: Status de Protocolo Atualizado!"
        blocos = "".join([f"<li><strong>Protocolo:</strong> {p['protocolo']} | <strong>Novo Status:</strong> {p['novo']}</li>" for p in processos_alterados])
        corpo_html = f"<html><body><p>Olá! O relatório foi atualizado:</p><ul>{blocos}</ul></body></html>"
        msg.attach(MIMEText(corpo_html, 'html'))
        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()
            server.login(EMAIL_REMETENTE, SENHA_GMAIL)
            server.sendmail(EMAIL_REMETENTE, EMAIL_DESTINATARIOS, msg.as_string())
        print("✉️ E-mail enviado!")
    except Exception as e:
        print(f"❌ Erro e-mail: {e}")
