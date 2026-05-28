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
# CONFIGURAÇÕES E CHAVES (SEGURANÇA TOTAL)
# ==========================================
EMAIL_REMETENTE = "mikaellevictoria2017@gmail.com"
EMAIL_DESTINATARIOS = ["santos.micaelle2006@gmail.com"]
LINK_PLANILHA = "https://artesanourbanismo-my.sharepoint.com/:x:/g/personal/mvitoria_artesanourbanismo_com_br/IQBI7DqicZMtSIrLcPwTnM2SAamaiye_3EPs-HAEKli1mZo?e=Zjekvn"

nome_planilha = "monitor_protocolos.xlsx"
nome_aba = "Santana de Parnaíba"

USER_PORTAL = os.getenv("USER_PORTAL", "")
SENHA_PORTAL = os.getenv("SENHA_PORTAL", "")
SENHA_GMAIL = os.getenv("SENHA_GMAIL", "")

agora_str = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
print(f"===== INICIANDO MONITORAMENTO DEFINITIVO: {agora_str} =====")

# ==========================================
# 1. LEITURA DA PLANILHA
# ==========================================
try:
    df = pd.read_excel(nome_planilha, sheet_name=nome_aba, skiprows=1)
    df.columns = [str(c).strip().upper() for c in df.columns]
    print(f"📥 Planilha carregada! Encontradas {len(df)} linhas para análise.")
except Exception as e:
    print(f"❌ Erro crítico ao ler a planilha Excel: {e}")
    exit(1)

# ==========================================
# 2. AUTOMAÇÃO WEB (MÉTODO INJEÇÃO JS - ANTI-TRAVAMENTO)
# ==========================================
dados_portal = {}

options = Options()
options.add_argument("--headless")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--window-size=1920,1080")

driver = webdriver.Chrome(options=options)
wait = WebDriverWait(driver, 30)

try:
    print("🌐 Acessando o portal de Santana de Parnaíba...")
    driver.get("https://santanadeparnaiba.aprova.com.br/login")
    
    # Aguarda o elemento essencial de login existir na estrutura da página
    wait.until(EC.presence_of_element_located((By.ID, "email")))
    time.sleep(3)
    
    print("🔑 Injetando dados de acesso via JavaScript para burlar bloqueios visuais...")
    # Injeta o e-mail e a senha direto no valor dos campos, contornando erros de "interactable"
    driver.execute_script(f"document.getElementById('email').value = '{USER_PORTAL}';")
    driver.execute_script(f"document.getElementById('password').value = '{SENHA_PORTAL}';")
    time.sleep(1)
    
    print("🔬 Disparando o envio do formulário de login...")
    # Encontra o formulário e envia os dados diretamente
    form_login = driver.find_element(By.XPATH, "//form")
    driver.execute_script("arguments[0].submit();", form_login)
    print("⏳ Aguardando processamento do login...")
    time.sleep(8)
    
    print("📂 Acessando a aba de processos...")
    driver.get("https://santanadeparnaiba.aprova.com.br/processos")
    
    # Aguarda a tabela aparecer em tela
    wait.until(EC.presence_of_element_located((By.XPATH, "//tbody/tr")))
    time.sleep(4)
    
    print("🔍 Varrendo a tabela de processos ativos...")
    linhas = driver.find_elements(By.XPATH, "//tbody/tr")
    for linha in linhas:
        texto_linha = linha.text.strip()
        if texto_linha:
            partes = texto_linha.split("\n")
            if len(partes) >= 2:
                protocolo_web = partes[0].strip().upper()
                status_web = partes[-1].strip()
                dados_portal[protocolo_web] = status_web
                
    print(f"✅ Varredura concluída. {len(dados_portal)} processos localizados no portal.")

except Exception as e:
    print(f"❌ Erro na navegação web: {e}")
finally:
    driver.quit()
    print("🔒 Navegador encerrado com segurança.")

# ==========================================
# 3. COMPARAÇÃO DOS STATUS
# ==========================================
processos_alterados = []
col_status = [c for c in df.columns if "STATUS" in c][0] if any("STATUS" in c for c in df.columns) else "STATUS"
col_modificado = [c for c in df.columns if "MODIFICADO" in c][0] if any("MODIFICADO" in c for c in df.columns) else "MODIFICADO"

print("⚖️ Iniciando validação de atualizações de status...")
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
        print(f"⚠️ ALTERAÇÃO: {protocolo_planilha} mudou para '{status_novo}'")
        processos_alterados.append({'protocolo': protocolo_planilha, 'antigo': status_antigo, 'novo': status_novo})
        df.at[index, col_status] = status_novo
        df.at[index, col_modificado] = agora_str

# ==========================================
# 4. SALVAMENTO E ENVIO DO E-MAIL (COM EMOJIS)
# ==========================================
if processos_alterados:
    try:
        print("💾 Gravando novos status na planilha...")
        with pd.ExcelWriter(nome_planilha, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name=nome_aba, index=False)
        print("📊 Planilha atualizada com sucesso.")
        
        if SENHA_GMAIL:
            print("✉️ Enviando e-mail com os alertas coloridos...")
            msg = MIMEMultipart()
            msg['From'] = EMAIL_REMETENTE
            msg['To'] = ", ".join(EMAIL_DESTINATARIOS)
            msg['Subject'] = "⚠️ Alteração de Status Detectada nos Protocolos"
            
            corpo = (
                f"Olá,\n\n"
                f"O robô identificou que houve alteração de status nos seguintes processos ativos:\n\n"
            )
            for x in processos_alterados:
                corpo += f"• Protocolo: {x['protocolo']}\n"
                corpo += f"  🔴 Status Anterior: {x['antigo']}\n"
                corpo += f"  🟢 Novo Status Atualizado: {x['novo']}\n\n"
                
            corpo += (
                f"A planilha de monitoramento já foi atualizada automaticamente.\n"
                f"Para verificar os detalhes completos, acesse o link do documento:\n"
                f"{LINK_PLANILHA}\n\n"
                f"Atenciosamente,\n"
                f"Robô de Monitoramento de Protocolos\n"
                f"Verificação efetuada em: {agora_str}"
            )
            
            msg.attach(MIMEText(corpo, 'plain'))
            
            with smtplib.SMTP('smtp.gmail.com', 587) as server:
                server.starttls()
                server.login(EMAIL_REMETENTE, SENHA_GMAIL)
                server.sendmail(EMAIL_REMETENTE, EMAIL_DESTINATARIOS, msg.as_string())
            print("✉️ Alerta enviado para o e-mail corporativo com sucesso!")
        else:
            print("⚠️ Envio de e-mail ignorado: Chave 'SENHA_GMAIL' ausente nas Secrets.")
    except Exception as e:
        print(f"❌ Falha ao salvar arquivo ou transmitir e-mail: {e}")
else:
    print("🦥 Varredura finalizada. Nenhum processo sofreu alterações. Tudo em ordem!")
