import os, time, pandas as pd
from datetime import datetime
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Configurações essenciais
EMAIL_REMETENTE = "mikaellevictoria2017@gmail.com"
EMAIL_DESTINATARIOS = ["santos.micaelle2006@gmail.com"]
LINK_PLANILHA = "https://artesanourbanismo-my.sharepoint.com/:x:/g/personal/mvitoria_artesanourbanismo_com_br/IQBI7DqicZMtSIrLcPwTnM2SAamaiye_3EPs-HAEKli1mZo?e=Zjekvn"
nome_planilha, nome_aba = "monitor_protocolos.xlsx", "Santana de Parnaíba"

USER_PORTAL = os.getenv("USER_PORTAL", "")
SENHA_PORTAL = os.getenv("SENHA_PORTAL", "")
SENHA_GMAIL = os.getenv("SENHA_GMAIL", "")
agora_str = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

print(f"===== INICIANDO MONITORAMENTO: {agora_str} =====")

# 1. LEITURA DA PLANILHA
try:
    df = pd.read_excel(nome_planilha, sheet_name=nome_aba, skiprows=1)
    df.columns = [str(c).strip().upper() for c in df.columns]
    print(f"📊 Planilha carregada: {len(df)} linhas identificadas.")
except Exception as e:
    print(f"❌ Erro ao ler Excel: {e}"); exit(1)

# 2. ACESSO AO PORTAL E EXTRAÇÃO (WEB SCRAPING)
dados_portal = {}
options = Options()
options.add_argument("--headless")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--window-size=1920,1080")
driver = webdriver.Chrome(options=options)
wait = WebDriverWait(driver, 25)

try:
    print("🌐 Abrindo a tela de login global do portal...")
    driver.get("https://aprova.com.br/login")
    time.sleep(5)  
    
    print("🔑 Preenchendo os dados de acesso...")
    campo_email = wait.until(EC.element_to_be_clickable((By.ID, "email")))
    campo_email.send_keys(USER_PORTAL)
    
    campo_senha = driver.find_element(By.ID, "password")
    campo_senha.send_keys(SENHA_PORTAL)
    
    print("Mpost 🖱️ Efetuando o clique de login...")
    botao_entrar = driver.find_element(By.XPATH, "//button[@type='submit' or contains(., 'Entrar')]")
    botao_entrar.click()
    time.sleep(8)
    
    print("📂 Acessando a aba de Processos de Santana de Parnaíba...")
    driver.get("https://santanadeparnaiba.aprova.com.br/processos")
    
    # Verifica se realmente entrou ou se foi barrado no login
    if "login" in driver.current_url:
        print("❌ Erro: O portal não passou da tela de login. Verifique se USER_PORTAL ou SENHA_PORTAL estão corretos nas Secrets.")
        exit(1)
        
    wait.until(EC.presence_of_element_located((By.XPATH, "//tbody/tr")))
    time.sleep(4)
    
    linhas = driver.find_elements(By.XPATH, "//tbody/tr")
    for linha in list(linhas):
        texto = linha.text.strip()
        if texto:
            linhas_texto = texto.split("\n")
            if len(linhas_texto) >= 2:
                protocolo_web = linhas_texto[0].strip().upper()
                status_web = linhas_texto[-1].strip()
                dados_portal[protocolo_web] = status_web
    print(f"✅ Varredura concluída. {len(dados_portal)} processos mapeados no portal.")
except Exception as e:
    print(f"❌ Erro na navegação web: {e}")
finally:
    driver.quit()

# 3. COMPARAÇÃO E ATUALIZAÇÃO
processos_alterados = []
col_status = [c for c in df.columns if "STATUS" in c][0] if any("STATUS" in c for c in df.columns) else "STATUS"
col_modificado = [c for c in df.columns if "MODIFICADO" in c][0] if any("MODIFICADO" in c for c in df.columns) else "MODIFICADO"

for index, row in df.iterrows():
    if str(row.get("ATIVO", "")).strip().upper() != "SIM":
        continue
        
    protocolo_planilha = str(row["PROTOCOLO"]).strip().upper()
    status_antigo = str(row.get(col_status, "")).strip()
    
    status_novo = None
    for k, v in dados_portal.items():
        if protocolo_planilha in k or k in protocolo_planilha:
            status_novo = v
            break
            
    if status_novo and status_antigo != status_novo:
        print(f"⚠️ ALTERAÇÃO: {protocolo_planilha} foi para '{status_novo}'")
        processos_alterados.append({'protocolo': protocolo_planilha, 'antigo': status_antigo, 'novo': status_novo})
        df.at[index, col_status] = status_novo
        df.at[index, col_modificado] = agora_str

# 4. SALVAMENTO E ENVIO DE E-MAIL
if procesos_alterados:
    try:
        with pd.ExcelWriter(nome_planilha, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name=nome_aba, index=False)
        print("💾 Planilha salva com sucesso.")
        
        if SENHA_GMAIL:
            msg = MIMEMultipart()
            msg['From'], msg['To'], msg['Subject'] = EMAIL_REMETENTE, ", ".join(EMAIL_DESTINATARIOS), "⚠️ Alteração de Status detectada"
            corpo = f"O robô detectou mudanças nos processos ativos:\n\n" + "\n".join([f"- {x['protocolo']}: {x['antigo']} -> {x['novo']}" for x in processos_alterados]) + f"\n\nLink: {LINK_PLANILHA}"
            msg.attach(MIMEText(corpo, 'plain'))
            with smtplib.SMTP('smtp.gmail.com', 587) as server:
                server.starttls()
                server.login(EMAIL_REMETENTE, SENHA_GMAIL)
                server.sendmail(EMAIL_REMETENTE, EMAIL_DESTINATARIOS, msg.as_string())
            print("✉️ E-mail de alerta enviado!")
    except Exception as e:
        print(f"❌ Erro ao salvar dados ou enviar e-mail: {e}")
else:
    print("🦥 Nenhum processo ativo sofreu alteração no portal. Tudo atualizado!")
