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
import subprocess

# ==========================================
# CONFIGURAÇÕES E CHAVES
# ==========================================
EMAIL_REMETENTE = "mikaellevictoria2017@gmail.com"
EMAIL_DESTINATARIOS = ["santos.micaelle2006@gmail.com"]

# ⚠️ LEMBRE-SE DE COLOCAR O LINK DA SUA PLANILHA ENTRE AS ASPAS ABAIXO:
LINK_PLANILHA = "https://artesanourbanismo-my.sharepoint.com/:x:/g/personal/mvitoria_artesanourbanismo_com_br/IQDgXvD3n6RTRZsZo63IiGBXAeSMCTvv1qBTTDNAD3d1_jE?e=PWsFlb"

nome_planilha = "monitor_protocolos.xlsx"
nome_aba = "Santana de Parnaíba"

USER_PORTAL = os.getenv("USER_PORTAL", "")
SENHA_PORTAL = os.getenv("SENHA_PORTAL", "")
SENHA_GMAIL = os.getenv("SENHA_GMAIL", "")

agora_str = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
print(f"===== INICIANDO MONITORAMENTO REESCRITA TOTAL: {agora_str} =====")

# ==========================================
# 1. LEITURA DA PLANILHA (MÉTODO BLINDADO)
# ==========================================
try:
    df = pd.read_excel(nome_planilha, sheet_name=nome_aba)
    
    if df.shape[0] > 0 and not any(any(x in str(c).upper() for x in ['PROTOCOLO', 'NUMER']) for c in df.columns):
        df_teste = pd.read_excel(nome_planilha, sheet_name=nome_aba, skiprows=1)
        if any(any(x in str(c).upper() for x in ['PROTOCOLO', 'NUMER']) for c in df_teste.columns):
            df = df_teste

    df.columns = [str(c).strip().upper() for c in df.columns]
    print(f"📥 Planilha carregada com sucesso! Colunas encontradas: {list(df.columns)}")
    
    col_protocolo = [c for c in df.columns if "PROTOCOLO" in c or "NUMER" in c or "PROCESSO" in c][0]
    col_ativo = [c for c in df.columns if "ATIVO" in c][0] if any("ATIVO" in c for c in df.columns) else None
    col_status = [c for c in df.columns if "STATUS" in c or "SITUA" in c][0] if any("STATUS" in c or "SITUA" in c for c in df.columns) else "STATUS ATUAL"
    col_modificado = [c for c in df.columns if "MODIF" in c or "DATA" in c][0] if any("MODIF" in c or "DATA" in c for c in df.columns) else "MODIFICADO EM"

    print(f"🎯 Colunas mapeadas -> Protocolo: [{col_protocolo}] | Status: [{col_status}]")
    lista_protocolos_planilha = df[col_protocolo].dropna().astype(str).str.strip().str.upper().tolist()
    print(f"📋 Protocolos localizados na sua planilha: {lista_protocolos_planilha}")

except Exception as e:
    print(f"❌ Erro crítico ao processar a estrutura da planilha Excel: {e}")
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
        
    print("⏳ Aguardando validação...")
    time.sleep(15)
    
    print("📂 Navegando para a URL interna de processos...")
    driver.get("https://santanadeparnaiba.aprova.com.br/processos")
    time.sleep(10)
    
    try:
        print("🧹 Detectado botão de filtros ativos. Limpando...")
        botao_limpar = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Limpar filtros')] | //a[contains(., 'Limpar filtros')]")))
        actions.move_to_element(botao_limpar).click().perform()
        time.sleep(10)
    except Exception:
        print(f"ℹ️ Filtros limpos ou não encontrados.")

    print("📜 Rolando a página...")
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    time.sleep(3)
    driver.execute_script("window.scrollTo(0, 0);")
    time.sleep(5)  
    
    print("🔍 Capturando dados das linhas reais da tabela...")
    linhas = driver.find_elements(By.XPATH, "//tbody/tr | //tr | //div[contains(@class, 'linha') or contains(@class, 'row')]")
    
    palavras_status = [
        'ANÁLISE', 'DEFERIDO', 'INDEFERIDO', 'COMUNIQUE-SE', 'AGUARDANDO', 
        'TRIAGEM', 'CONCLUÍDO', 'EMITIDO', 'CORREÇÃO', 'PENDENTE', 'EMISSÃO',
        'PROCESSO', 'VALIDAÇÃO', 'REVISÃO', 'SOLICITADO', 'DESPACHO', 'ENCAMINHADO'
    ]
    
    for linha in linhas:
        texto_linha = linha.text.strip()
        if texto_linha:
            partes = [p.strip() for p in texto_linha.split("\n") if p.strip()]
            if len(partes) >= 2:
                protocolo_web = partes[0].upper()
                
                if any(char.isdigit() for char in protocolo_web) and len(protocolo_web) < 35:
                    status_web = ""
                    for parte in reversed(partes[1:]):
                        if any(termo in parte.upper() for termo in palavras_status):
                            status_web = parte
                            break
                    
                    if not status_web:
                        for parte in reversed(partes[1:]):
                            if parte.count(" ") <= 3 and not any(char.isdigit() for char in parte):
                                status_web = parte
                                break
                    
                    if not status_web:
                        status_web = partes[-1]
                        
                    dados_portal[protocolo_web] = status_web

    print(f"✅ Extraídos {len(dados_portal)} registros do site.")

except Exception as e:
    print(f"❌ Falha na execução da automação: {e}")
finally:
    driver.quit()

# ==========================================
# 3. COMPARAÇÃO E SUBSTITUIÇÃO DA LINHA TODA
# ==========================================
processos_alterados = []

print(f"⚖️ Iniciando análise de reescrita de linhas...")
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
        print(f"⚠️ MUDANÇA TOTAL NA LINHA: {protocolo_planilha} de '{status_antigo}' -> '{status_novo}'")
        processos_alterados.append({'protocolo': protocolo_planilha, 'antigo': status_antigo, 'novo': status_novo})
        
        # ⚡ OPÇÃO 2: Limpa a linha inteira da planilha e redefine apenas os dados essenciais novos
        for col in df.columns:
            df.at[index, col] = None # Limpa o conteúdo antigo da linha toda
            
        # Preenche os dados novos limpos do portal
        df.at[index, col_protocolo] = protocolo_planilha
        df.at[index, col_status] = status_novo
        df.at[index, col_modificado] = agora_str
        if col_ativo:
            df.at[index, col_ativo] = "SIM"

# ==========================================
# 4. SALVAMENTO AUTOMÁTICO NO GITHUB E ENVIO DO E-MAIL
# ==========================================
if processos_alterados:
    try:
        print("💾 Gravando atualizações na planilha de controle...")
        with pd.ExcelWriter(nome_planilha, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name=nome_aba, index=False)
        
        print("🚀 Enviando alterações automaticamente para o repositório GitHub...")
        subprocess.run(["git", "config", "user.name", "Automated Robot"], check=True)
        subprocess.run(["git", "config", "user.email", "robot@artesano.com"], check=True)
        subprocess.run(["git", "add", nome_planilha], check=True)
        subprocess.run(["git", "commit", "-m", f"🤖 Substituição total de linhas alteradas: {agora_str}"], check=True)
        subprocess.run(["git", "push"], check=True)
        print("✨ GitHub atualizado com sucesso de forma autónoma!")
        
        if SENHA_GMAIL:
            print("✉️ Enviando e-mail de notificação...")
            msg = MIMEMultipart('alternative')
            msg['From'] = EMAIL_REMETENTE
            msg['To'] = ", ".join(EMAIL_DESTINATARIOS)
            msg['Subject'] = "⚠️ Alerta: Linhas de Protocolos Reescritas!"
            
            corpo_html = f"""
            <html>
            <body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333333;">
                <p>Olá, Artesano!</p>
                <p>O robô identificou mudanças e <strong>substituiu a linha inteira</strong> dos seguintes processos ativos:</p>
                <div style="background-color: #f8f9fa; border-left: 4px solid #dc3545; padding: 12px; margin-bottom: 20px;">
                    <ul style="margin-top: 8px; padding-left: 20px;">
            """
            for x in processos_alterados:
                corpo_html += f"""
                        <li style="margin-bottom: 12px;">
                            <strong>Protocolo:</strong> <code style="background-color: #e9ecef; padding: 2px 6px; border-radius: 4px;">{x['protocolo']}</code><br>
                            <span style="color: #dc3545;">❌ Registro Antigo Apagado</span><br>
                            <span style="color: #28a745;">✅ Nova Linha Criada com Status:</span> <strong>{x['novo']}</strong>
                        </li>
                """
            corpo_html += f"""
                    </ul>
                </div>
                <p>Para ver a planilha atualizada no repositório, acesse o link:</p>
                <p><a href="{LINK_PLANILHA}" target="_blank" style="color: #007bff; font-weight: bold;">🔗 monitor_protocolos</a></p>
                <br>
                <p style="font-size: 12px; color: #6c757d;">Executado de forma autônoma em: {agora_str}</p>
            </body>
            </html>
            """
            msg.attach(MIMEText(corpo_html, 'html'))
            with smtplib.SMTP('smtp.gmail.com', 587) as server:
                server.starttls()
                server.login(EMAIL_REMETENTE, SENHA_GMAIL)
                server.sendmail(EMAIL_REMETENTE, EMAIL_DESTINATARIOS, msg.as_string())
            print("✉️ Notificação enviada!")
    except Exception as e:
        print(f"❌ Falha no salvamento autônomo: {e}")
else:
    print("🦥 Varredura finalizada. Nenhuma linha precisou ser alterada hoje.")
