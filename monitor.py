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
# CONFIGURAÇÕES E CHAVES
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
print(f"===== INICIANDO MONITORAMENTO DESBLOQUEADOR DE FILTROS: {agora_str} =====")

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
# 2. AUTOMAÇÃO WEB
# ==========================================
dados_portal = {}

options = Options()
options.add_argument("--headless=new") 
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--window-size=1920,1080")
options.add_argument("--disable-blink-features=AutomationControlled")
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option('useAutomationExtension', False)
options.add_argument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36")

driver = webdriver.Chrome(options=options)
driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
    "source": "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
})

wait = WebDriverWait(driver, 35)
actions = ActionChains(driver)

try:
    print("🌐 Acessando o portal de Santana de Parnaíba...")
    driver.get("https://santanadeparnaiba.aprova.com.br/login")
    time.sleep(7)
    
    print("🔍 Localizando os campos de entrada...")
    inputs = wait.until(EC.presence_of_all_elements_located((By.TAG_NAME, "input")))
    
    if len(inputs) >= 2:
        campo_email = inputs[0]
        campo_senha = inputs[1]
        
        print("✍️ Preenchendo campo de E-mail...")
        actions.move_to_element(campo_email).click().perform()
        campo_email.send_keys(USER_PORTAL)
        time.sleep(1)
        
        print("✍️ Preenchendo campo de Senha...")
        actions.move_to_element(campo_senha).click().perform()
        campo_senha.send_keys(SENHA_PORTAL)
        time.sleep(1)
        
        print("🚀 Acionando botão de login...")
        botao = driver.find_element(By.CSS_SELECTOR, "button[type='submit']")
        actions.move_to_element(botao).click().perform()
    else:
        raise Exception("Não foi possível renderizar os campos de entrada.")
        
    print("⏳ Aguardando validação dos cookies...")
    time.sleep(15)
    
    print("📂 Navegando para a URL interna de processos...")
    driver.get("https://santanadeparnaiba.aprova.com.br/processos")
    time.sleep(10)
    
    # --- PASSO CHAVE: DESBLOQUEAR OS FILTROS DO PRINT ---
    try:
        print("🧹 Detetado botão de filtros ativos. Tentando clicar em 'Limpar filtros'...")
        # Procura pelo botão que contém o texto de limpar filtros conforme o print
        botao_limpar = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Limpar filtros')] | //a[contains(., 'Limpar filtros')]")))
        actions.move_to_element(botao_limpar).click().perform()
        print("✨ Filtros limpos com sucesso! Aguardando recarregamento da tabela completa...")
        time.sleep(10)
    except Exception as f_err:
        print(f"ℹ️ Botão 'Limpar filtros' não precisou ser acionado ou não foi encontrado: {f_err}")

    print("📜 Rolando a página para garantir a renderização...")
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    time.sleep(3)
    driver.execute_script("window.scrollTo(0, 0);")
    time.sleep(5)  
    
    print("🔍 Capturando dados das linhas reais da tabela...")
    linhas = driver.find_elements(By.XPATH, "//tbody/tr | //tr | //div[contains(@class, 'linha') or contains(@class, 'row')]")
    
    for linha in linhas:
        texto_linha = linha.text.strip()
        if texto_linha:
            partes = texto_linha.split("\n")
            if len(partes) >= 2:
                protocolo_web = partes[0].strip().upper()
                status_web = partes[-1].strip()
                if any(char.isdigit() for char in protocolo_web) and len(protocolo_web) < 35:
                    dados_portal[protocolo_web] = status_web

    print(f"✅ Mapeamento concluído. {len(dados_portal)} processos reais extraídos do site.")

except Exception as e:
    print(f"❌ Falha durante a execução da navegação automatizada: {e}")
finally:
    driver.quit()
    print("🔒 Instância do navegador encerrada com segurança.")

# ==========================================
# 3. COMPARAÇÃO DOS STATUS MAPEADOS
# ==========================================
processos_alterados = []

# Mapeia as colunas exatas com base no print do cabeçalho do Excel
col_status = "STATUS ATUAL" if "STATUS ATUAL" in df.columns else ([c for c in df.columns if "STATUS" in c][0] if any("STATUS" in c for c in df.columns) else "STATUS ATUAL")
col_modificado = "MODIFICADO EM" if "MODIFICADO EM" in df.columns else ([c for c in df.columns if "MODIF" in c][0] if any("MODIF" in c for c in df.columns) else "MODIFICADO EM")

print(f"⚖️ Comparando registros utilizando as colunas: [{col_status}] e [{col_modificado}]...")
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
        print(f"⚠️ ALTERAÇÃO ENCONTRADA: Processo {protocolo_planilha} mudou de '{status_antigo}' para '{status_novo}'")
        processos_alterados.append({'protocolo': protocolo_planilha, 'antigo': status_antigo, 'novo': status_novo})
        df.at[index, col_status] = status_novo
        df.at[index, col_modificado] = agora_str

# ==========================================
# 4. SALVAMENTO E ENVIO DO E-MAIL
# ==========================================
if procesos_alterados:
    try:
        print("💾 Gravando atualizações na planilha de controle...")
        with pd.ExcelWriter(nome_planilha, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name=nome_aba, index=False)
        print("📊 Planilha atualizada com sucesso.")
        
        if SENHA_GMAIL:
            print("✉️ Estruturando e-mail de alerta...")
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
            print("✉️ E-mail disparado com sucesso!")
        else:
            print("⚠️ Envio de e-mail cancelado: Variável 'SENHA_GMAIL' não configurada.")
    except Exception as e:
        print(f"❌ Falha ao salvar arquivo Excel ou disparar e-mail: {e}")
else:
    print("🦥 Varredura finalizada. Nenhum processo sofreu modificações válidas hoje.")
