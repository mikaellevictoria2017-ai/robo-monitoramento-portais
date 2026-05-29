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

# 🎯 Substitua pelo link direto da sua planilha do SharePoint para o e-mail:
LINK_PLANILHA = "https://artesanourbanismo-my.sharepoint.com/:x:/g/personal/mvitoria_artesanourbanismo_com_br/IQDgXvD3n6RTRZsZo63IiGBXAeSMCTvv1qBTTDNAD3d1_jE?e=5WIK4X"

nome_planilha = "monitor_protocolos.xlsx"
nome_aba = "Santana de Parnaíba"

# Chaves de segurança guardadas no GitHub
USER_PORTAL = os.getenv("USER_PORTAL", "")
SENHA_PORTAL = os.getenv("SENHA_PORTAL", "")
SENHA_GMAIL = os.getenv("SENHA_GMAIL", "")

agora_str = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
print(f"===== INICIANDO MONITORAMENTO REESCRITA TOTAL: {agora_str} =====")

# ==========================================
# 1. LEITURA DA PLANILHA QUE ESTÁ NO GITHUB
# ==========================================
try:
    df = pd.read_excel(nome_planilha, sheet_name=nome_aba)
    df.columns = [str(c).strip().upper() for c in df.columns]
    print(f"📥 Planilha carregada com sucesso! Colunas encontradas: {list(df.columns)}")
    
    col_protocolo = [c for c in df.columns if "PROTOCOLO" in c or "NUMER" in c or "PROCESSO" in c][0]
    col_ativo = [c for c in df.columns if "ATIVO" in c][0] if any("ATIVO" in c for c in df.columns) else None
    col_status = [c for c in df.columns if "STATUS" in c or "SITUA" in c][0] if any("STATUS" in c for c in df.columns) else "STATUS ATUAL"
    col_modificado = [c for c in df.columns if "MODIF" in c or "DATA" in c][0] if any("MODIF" in c for c in df.columns) else "MODIFICADO EM"

    protocolos_verificar = df[col_protocolo].dropna().astype(str).tolist()
    print(f"🔍 Protocolos localizados na sua planilha: {protocolos_verificar}")

except Exception as e:
    print(f"❌ Erro ao ler a planilha: {e}")
    exit(1)

# ==========================================
# 2. AUTOMAÇÃO NO PORTAL (ROBÔ BUSCANDO STATUS)
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
    palavras_status = ['ANÁLISE', 'DEFERIDO', 'INDEFERIDO', 'COMUNIQUE-SE', 'AGUARDANDO', 'TRIAGEM', 'CONCLUÍDO', 'EMITIDO', 'CORREÇÃO', 'PENDENTE']
    
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
    print(f"❌ Falha na automação: {e}")
finally:
    driver.quit()

# ==========================================
# 3. ANÁLISE E REESCRITA DA LINHA
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
        
        # 🧹 Limpa o texto antigo de todas as colunas daquela linha antes de escrever
        for col in df.columns:
            df.at[index, col] = None
            
        # ✍️ Preenche as colunas com as novas informações
        df.at[index, col_protocolo] = protocolo_planilha
        df.at[index, col_status] = status_novo
        df.at[index, col_modificado] = agora_str
        if col_ativo:
            df.at[index, col_ativo] = "SIM"

# ==========================================
# 4. SALVAMENTO AUTOMÁTICO NA PLANILHA
# ==========================================
if processos_alterados:
    print("💾 Gravando alterações na planilha de monitoramento...")
    with pd.ExcelWriter(nome_planilha, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name=nome_aba, index=False)
    
    # Envia e-mail de alerta avisando que mudou
    if SENHA_GMAIL:
        try:
            msg = MIMEMultipart('alternative')
            msg['From'] = EMAIL_REMETENTE
            msg['To'] = ", ".join(EMAIL_DESTINATARIOS)
            msg['Subject'] = "⚠️ Alerta: Linha Alterada na Planilha Oficial!"
            
            linhas_tabela = "".join([f"<tr><td>{p['protocolo']}</td><td style='color:red;'>{p['antigo']}</td><td style='color:green;'>{p['novo']}</td></tr>" for p in processos_alterados])
            
            corpo_html = f"""
            <html>
            <body>
                <h2>O robô detectou alterações de status e já preencheu a planilha!</h2>
                <table border="1" cellpadding="5" style="border-collapse: collapse;">
                    <tr bgcolor="#f2f2f2"><th>Protocolo</th><th>Status Antigo</th><th>Status Novo</th></tr>
                    {linhas_tabela}
                </table>
                <br>
                <p>👉 Acesse a sua planilha oficial aqui: <a href="{LINK_PLANILHA}">Abrir Planilha SharePoint</a></p>
            </body>
            </html>
            """
            msg.attach(MIMEText(corpo_html, 'html'))
            with smtplib.SMTP('smtp.gmail.com', 587) as server:
                server.starttls()
                server.login(EMAIL_REMETENTE, SENHA_GMAIL)
                server.sendmail(EMAIL_REMETENTE, EMAIL_DESTINATARIOS, msg.as_string())
            print("✉️ E-mail de alerta enviado!")
        except Exception as e:
            print(f"❌ Erro ao enviar e-mail: {e}")
else:
    print("🦥 Varredura finalizada. Nenhuma linha precisou ser alterada hoje.")
