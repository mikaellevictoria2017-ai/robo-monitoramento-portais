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

# 🎯 Link oficial que vai no botão azul da planilha
LINK_PLANILHA = "https://github.com/mikaellevictoria2017-ai/robo-monitoramento-portais/blob/main/monitor_protocolos.xlsx"

nome_planilha = "monitor_protocolos.xlsx"
nome_aba = "Santana de Parnaíba"

# Chaves de segurança do GitHub
USER_PORTAL = os.getenv("USER_PORTAL", "")
SENHA_PORTAL = os.getenv("SENHA_PORTAL", "")
SENHA_GMAIL = os.getenv("SENHA_GMAIL", "")

agora_str = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
data_verificacao = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

print(f"===== INICIANDO MONITORAMENTO SELETIVO: {agora_str} =====")

# ==========================================
# 1. LEITURA DA PLANILHA DO REPOSITÓRIO
# ==========================================
try:
    df = pd.read_excel(nome_planilha, sheet_name=nome_aba)
    colunas_originais = list(df.columns)
    colunas_maiusculas = [str(c).strip().upper() for c in df.columns]
    
    idx_protocolo = colunas_maiusculas.index([c for c in colunas_maiusculas if "PROTOCOLO" in c or "NUMER" in c][0])
    idx_ultima_acao = colunas_maiusculas.index([c for c in colunas_maiusculas if "AÇÃO" in c or "ACAO" in c][0])
    idx_status_atual = colunas_maiusculas.index([c for c in colunas_maiusculas if "STATUS" in c or "SITUA" in c][0])
    idx_modificado_em = colunas_maiusculas.index([c for c in colunas_maiusculas if "MODIF" in c or "DATA" in c][-1])
    
    idx_ativo = colunas_maiusculas.index([c for c in colunas_maiusculas if "ATIVO" in c][0]) if any("ATIVO" in c for c in colunas_maiusculas) else None

    protocolos_verificar = df.iloc[:, idx_protocolo].dropna().astype(str).tolist()
    print(f"📥 Planilha carregada! Buscando atualizações para os protocolos: {protocolos_verificar}")

except Exception as e:
    print(f"❌ Erro ao estruturar as colunas da planilha: {e}")
    exit(1)

# ==========================================
# 2. AUTOMAÇÃO E EXTRAÇÃO DOS DADOS DO PORTAL
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
    
    for linha in linhas:
        texto_linha = linha.text.strip()
        if texto_linha:
            partes = [p.strip() for p in texto_linha.split("\n") if p.strip()]
            if len(partes) >= 2:
                protocolo_web = partes[0].upper()
                if any(char.isdigit() for char in protocolo_web):
                    # Captura o texto completo do status para não cortar palavras compridas
                    status_web = partes[-1]
                    acao_recente_web = partes[-1]
                    
                    # Se houver mais partes, analisa de trás para frente para pegar a descrição correta
                    if len(partes) > 2:
                        for parte in reversed(partes[1:]):
                            if any(t in parte.upper() for t in ['PROCESSO', 'ENCAMINHADO', 'ANÁLISE', 'DEFERIDO', 'COMUNIQUE-SE']):
                                status_web = parte
                                break
                    
                    dados_portal[protocolo_web] = {"status": status_web, "ultima_acao": acao_recente_web}

    print(f"✅ Extraídos {len(dados_portal)} registros do site.")
except Exception as e:
    print(f"❌ Falha na raspagem de dados: {e}")
finally:
    driver.quit()

# ==========================================
# 3. ATUALIZAÇÃO SELETIVA (APENAS AS 3 COLUNAS)
# ==========================================
processos_alterados = []

for index, row in df.iterrows():
    if idx_ativo is not None and str(df.iloc[index, idx_ativo]).strip().upper() != "SIM":
        continue
        
    protocolo_planilha = str(df.iloc[index, idx_protocolo]).strip().upper()
    status_antigo = str(df.iloc[index, idx_status_atual]).strip()
    
    status_novo = None
    acao_nova = None
    
    for k, v in dados_portal.items():
        if protocolo_planilha in k or k in protocolo_planilha:
            status_novo = v["status"]
            acao_nova = v["ultima_acao"]
            break
            
    if status_novo and status_antigo != status_novo:
        print(f"⚠️ ATUALIZANDO: {protocolo_planilha} -> Novo Status: {status_novo}")
        processos_alterados.append({
            'protocolo': protocolo_planilha, 
            'antigo': status_antigo, 
            'novo': status_novo
        })
        
        # 🎯 Atualiza estritamente as 3 colunas desejadas
        df.iloc[index, idx_status_atual] = status_novo
        df.iloc[index, idx_ultima_acao] = acao_nova
        df.iloc[index, idx_modificado_em] = agora_str

df.columns = colunas_originais

# ==========================================
# 4. SALVAMENTO E ENVIO DO E-MAIL ESTILIZADO
# ==========================================
if procesos_alterados:
    print("💾 Gravando dados atualizados no Excel...")
    with pd.ExcelWriter(nome_planilha, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name=nome_aba, index=False)
    
    if SENHA_GMAIL:
        try:
            msg = MIMEMultipart('alternative')
            msg['From'] = EMAIL_REMETENTE
            msg['To'] = ", ".join(EMAIL_DESTINATARIOS)
            msg['Subject'] = "⚠️ Status de Protocolo Alterado no SharePoint!"
            
            # Monta os blocos de processos idênticos ao layout do print da Artesano
            blocos_processos = ""
            for p in processos_alterados:
                blocos_processos += f"""
                <div style="border-left: 4px solid #0078d4; padding-left: 15px; margin-bottom: 20px;">
                    <p style="margin: 5px 0; font-size: 16px; font-weight: bold; color: #333;">
                        🏢 PORTAL DE SANTANA DE PARNAÍBA
                    </p>
                    <ul style="margin: 5px 0; padding-left: 20px; color: #555;">
                        <li><strong>Protocolo:</strong> <span style="background-color: #e1dfdd; padding: 2px 6px; border-radius: 3px;">{p['protocolo']}</span></li>
                        <li>
