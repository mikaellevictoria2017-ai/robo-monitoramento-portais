import os
import time
import pandas as pd
from datetime import datetime
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# ==========================================
# CONFIGURAÇÕES INICIAIS E VARIÁVEIS DE AMBIENTE
# ==========================================

# Ajuste com as suas informações reais de e-mail e links:
EMAIL_REMETENTE = "mikaellevictoria2017@gmail.com"
EMAIL_DESTINATARIOS = ["santos.micaelle2006@gmail.com"]
LINK_PLANILHA = "https://artesanourbanismo-my.sharepoint.com/:x:/g/personal/mvitoria_artesanourbanismo_com_br/IQBI7DqicZMtSIrLcPwTnM2SAamaiye_3EPs-HAEKli1mZo?e=Zjekvn"

nome_planilha = "monitor_protocolos.xlsx"
nome_aba = "Santana de Parnaíba"

# Puxando as credenciais secretas do GitHub Actions
USER_PORTAL = os.getenv("USER_PORTAL", "")
SENHA_PORTAL = os.getenv("SENHA_PORTAL", "")
SENHA_GMAIL = os.getenv("SENHA_GMAIL", "")

agora = datetime.now()
agora_str = agora.strftime("%d/%m/%Y %H:%M:%S")

print(f"===== INICIANDO VERIFICAÇÃO: {agora_str} =====")

# ==========================================
# FUNÇÃO PARA ENVIAR E-MAIL DE ALERTA
# ==========================================
def enviar_email_alerta(lista_alteracoes):
    if not lista_alteracoes:
        return
        
    msg = MIMEMultipart()
    msg['From'] = EMAIL_REMETENTE
    msg['To'] = ", ".join(EMAIL_DESTINATARIOS)
    msg['Subject'] = f"⚠️ [Alerta] Alteração de Status de Protocolo - {agora.strftime('%d/%m/%Y')}"
    
    corpo_html = f"""
    <html>
    <body style="font-family: Arial, sans-serif; color: #333;">
        <h2 style="color: #d9534f;">Atenção! O robô detectou mudanças no portal:</h2>
        <table border="1" cellpadding="10" cellspacing="0" style="border-collapse: collapse; width: 100%;">
            <tr style="background-color: #f2f2f2;">
                <th>Protocolo</th>
                <th>Status Antigo</th>
                <th>Status Novo</th>
            </tr>
    """
    for alt in lista_alteracoes:
        corpo_html += f"""
            <tr>
                <td><b>{alt['protocolo']}</b></td>
                <td style="color: #777;">{alt['antigo']}</td>
                <td style="color: #2b78e4; font-weight: bold;">{alt['novo']}</td>
            </tr>
        """
    corpo_html += f"""
        </table>
        <br>
        <p>📊 <b>Link da Planilha para consulta:</b> <a href="{LINK_PLANILHA}">Clique aqui para acessar</a></p>
        <hr style="border: 0; border-top: 1px solid #eee;">
        <p style="font-size: 11px; color: #999;">Este é um e-mail automático gerado pelo Monitor de Protocolos da Artesano Urbanismo.</p>
    </body>
    </html>
    """
    
    msg.attach(MIMEText(corpo_html, 'html'))
    
    with smtplib.SMTP('smtp.gmail.com', 557) as server:
        server.starttls()
        server.login(EMAIL_REMETENTE, SENHA_GMAIL)
        server.sendmail(EMAIL_REMETENTE, EMAIL_DESTINATARIOS, msg.as_string())

# ==========================================
# LEITURA DA PLANILHA LOCAL
# ==========================================
if not os.path.exists(nome_planilha):
    print(f"❌ Erro Crítico: O arquivo '{nome_planilha}' não foi encontrado no repositório.")
    exit(1)

try:
    # Localiza a linha correta do cabeçalho de forma inteligente
    linha_correta = 0
    with pd.ExcelFile(nome_planilha) as xls:
        df_teste = pd.read_excel(xls, sheet_name=nome_aba, nrows=10)
        for i, row in df_teste.iterrows():
            valores = [str(v).strip().upper() for v in row.values if pd.notna(v)]
            if "PROTOCOLO" in valores or "STATUS" in valores:
                linha_correta = i + 1
                break

    df = pd.read_excel(nome_planilha, sheet_name=nome_aba, skiprows=linha_correta)
    df.columns = [str(c).strip().upper() for c in df.columns]
    print(f"📊 Aba '{nome_aba}' carregada com sucesso! Encontradas {len(df)} linhas.")
except Exception as e:
    print(f"❌ Erro ao ler a planilha Excel: {e}")
    exit(1)

# Validação das colunas obrigatórias
colunas_necessarias = ["PROTOCOLO", "ATIVO"]
for col in colunas_necessarias:
    if col not in df.columns:
        print(f"❌ Erro: A coluna obrigatória '{col}' não foi encontrada na planilha. Colunas atuais: {list(df.columns)}")
        exit(1)

col_status = [c for c in df.columns if "STATUS" in c][0] if any("STATUS" in c for c in df.columns) else None
col_modificado = [c for c in df.columns if "MODIFICADO" in c][0] if any("MODIFICADO" in c for c in df.columns) else "MODIFICADO"

if not col_status:
    df["STATUS"] = ""
    col_status = "STATUS"
if col_modificado not in df.columns:
    df[col_modificado] = ""

# ==========================================
# EXTRAÇÃO DE DADOS DO PORTAL (WEB SCRAPING)
# ==========================================
dados_portal = {}
processos_alterados = []
houve_alteracao = False

chrome_options = Options()
chrome_options.add_argument("--headless")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--window-size=1920,1080")

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
wait = WebDriverWait(driver, 15)

try:
    print("🌐 Abrindo a tela de login do portal...")
    driver.get("https://santanadeparnaiba.aprova.com.br/login")
    
    print("🔑 Preenchendo os dados de acesso...")
    wait.until(EC.presence_of_element_located((By.ID, "email"))).send_keys(USER_PORTAL)
    driver.find_element(By.ID, "password").send_keys(SENHA_PORTAL)
    
    print("🖱️ Clicando no botão Entrar...")
    driver.find_element(By.XPATH, "//button[@type='submit' or contains(., 'Entrar')]").click()
    
    print("🔓 Login efetuado! Aguardando o carregamento da área interna...")
    time.sleep(5)
    
    print("📂 Navegando até a aba de Processos...")
    driver.get("https://santanadeparnaiba.aprova.com.br/processos")
    time.sleep(6)
    
    print("Iniciando varredura em busca dos protocolos na tabela...")
    # Coleta todas as linhas da tabela de processos na tela
    linhas_tabela = driver.find_elements(By.XPATH, "//tbody/tr")
    print(f"✅ {len(linhas_tabela)} processos encontrados na tela.")
    
    for linha in linhas_tabela:
        try:
            celulas = linha.find_elements(By.XPATH, "./td")
            if len(celulas) >= 2:
                # Procura de forma inteligente em qual célula está o número do protocolo
                protocolo_texto = ""
                for c in celulas:
                    texto_celula = c.text.strip()
                    if any(char.isdigit() for char in texto_celula) and "-" in texto_celula:
                        protocolo_texto = texto_celula.split('\n')[0].strip().upper()
                        break
                
                if not protocolo_texto:
                    protocolo_texto = celulas[1].text.strip().split('\n')[0].strip().upper()
                
                # O status do processo geralmente fica nas últimas colunas
                status_texto = celulas[-2].text.strip() if len(celulas) >= 3 else celulas[-1].text.strip()
                
                if protocolo_texto:
                    dados_portal[protocolo_texto] = {"status": status_texto}
        except Exception:
            continue

except Exception as e:
    print(f"❌ Erro durante a navegação no portal web: {e}")
finally:
    driver.quit()
    print("🔌 Navegador fechado com segurança pelo sistema.")

# ==========================================
# SEÇÃO DE COMPARAÇÃO DOS DADOS
# ==========================================
print("🔮 Comparando dados da planilha com o portal...")
for index, text_linha in df.iterrows():
    if str(text_linha["ATIVO"]).strip().upper() != "SIM":
        continue  # Pula os processos inativos de forma limpa

    protocolo = str(text_linha["PROTOCOLO"]).strip().upper()
    status_antigo = str(text_linha[col_status]).strip()

    linha_encontrada = None
    for dado in dados_portal:
        if protocolo in dado or dado in protocolo:
            linha_encontrada = dados_portal[dado]
            break

    if not linha_encontrada:
        print(f"❓ {protocolo}: Não foi localizado na página atual do portal.")
        continue

    status_novo = linha_encontrada["status"]

    if status_antigo != status_novo:
        print(f"⚠️ ALTERAÇÃO DETECTADA! {protocolo} mudou de '{status_antigo}' para '{status_novo}'")
        processos_alterados.append({
            'protocolo': protocolo, 'antigo': status_antigo, 'novo': status_novo
        })
        df.at[index, col_status] = str(status_novo)
        df.at[index, col_modificado] = str(agora_str)
        houve_alteracao = True
    else:
        print(f"✅ {protocolo}: Status igual ao do portal ('{status_antigo}').")

# Remove textos nulos antes de salvar
for col in df.columns:
    df[col] = df[col].astype(str).replace('nan', '')

# ==========================================
# SALVAMENTO NA PLANILHA COM LAYOUT AUTOMÁTICO
# ==========================================
salvo_com_sucesso = False
tentativas = 0

if houve_alteracao:
    while not salvo_com_sucesso and tentativas < 3:
        try:
            with pd.ExcelFile(nome_planilha) as reader:
                abas_existentes = {sheet: reader.parse(sheet) for sheet in reader.sheet_names}
            
            abas_existentes[nome_aba] = df
            
            with pd.ExcelWriter(nome_planilha, engine='openpyxl') as writer:
                for sheet, dados_aba in abas_existentes.items():
                    if sheet == nome_aba:
                        start_row = linha_correta if 'linha_correta' in locals() else 0
                        dados_aba.to_excel(writer, sheet_name=sheet, startrow=start_row, index=False)
                        worksheet = writer.sheets[sheet]
                        for col in worksheet.columns:
                            max_len = 0
                            col_letter = col[0].column_letter
                            for cell in col:
                                if cell.value:
                                    max_len = max(max_len, len(str(cell.value)))
                            worksheet.column_dimensions[col_letter].width = max(max_len + 4, 12)
                    else:
                        dados_aba.to_excel(writer, sheet_name=sheet, index=False)
            
            print(f"💾 Planilha salva. Aba '{nome_aba}' atualizada com layout perfeito!")
            salvo_com_sucesso = True
        except PermissionError:
            print("⚠️ AVISO: Arquivo Excel ocupado, tentando novamente em 5 segundos...")
            time.sleep(5)
            tentativas += 1
        except Exception as e:
            print(f"❌ Erro ao salvar planilha: {e}")
            break

    # Envio do e-mail de alerta
    if processos_alterados:
        if not SENHA_GMAIL:
            print("⚠️ Alerta do Sistema: A variável 'SENHA_GMAIL' veio vazia da nuvem. O e-mail não pôde ser enviado.")
        else:
            try:
                enviar_email_alerta(processos_alterados)
                print("✉️ E-mail de alerta enviado com sucesso!")
            except Exception as e:
                print(f"❌ Erro ao enviar e-mail: {e}")
else:
    print("🦥 Nenhuma alteração de status encontrada nos processos ativos. Tudo atualizado!")
