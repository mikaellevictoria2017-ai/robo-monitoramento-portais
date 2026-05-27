import os
import time
from datetime import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ==========================================
# 🌐 CONFIGURAÇÕES GLOBAIS (Ajustadas)
# ==========================================
# Definido seu e-mail como remetente e destinatário oficial dos alertas
EMAIL_REMETENTE = "mikaellevictoria2017@gmail.com"  
EMAIL_DESTINATARIOS = ["mikaellevictoria2017@gmail.com"]  

# Puxa a senha secreta (as 16 letras que você vai cadastrar no GitHub Secrets)
SENHA_REMETENTE = os.environ.get("ihfxftkgihyuniob")

# Link direto para a sua planilha no GitHub
LINK_PLANILHA = "https://github.com/mikaellevictoria2017-ai/robo-monitoramento-portais/blob/main/monitor_protocolos.xlsx"

# ==========================================
# ✉️ FUNÇÃO DE ENVIO DE E-MAIL (Com validação de segurança)
# ==========================================
def enviar_email_alerta(processos_alterados):
    try:
        # Verifica se a senha foi carregada pela nuvem
        if not SENHA_REMETENTE:
            print("⚠️ Alerta do Sistema: A variável 'SENHA_GMAIL' veio vazia da nuvem. O e-mail não pôde ser enviado.")
            return

        msg = MIMEMultipart('alternative')
        msg['From'] = EMAIL_REMETENTE
        msg['To'] = ", ".join(EMAIL_DESTINATARIOS)
        msg['Subject'] = f"📢 [Aviso] Mudança de Status em Processos - {datetime.now().strftime('%d/%m/%Y')}"

        html_corpo = f"""
        <html>
        <body style="font-family: Arial, sans-serif; font-size: 14px; color: #333333; line-height: 1.6;">
            <p style="margin-bottom: 15px;">Olá,</p>
            <p style="margin-bottom: 20px;">O robô de monitoramento identificou mudanças de status nos seguintes processos:</p>
            <p style="margin-bottom: 15px;"><strong>◆ SANTANA DE PARNAÍBA</strong></p>
        """

        agora_str = datetime.now().strftime('%d/%m/%Y %H:%M')

        for proc in processos_alterados:
            html_corpo += f"""
            <div style="margin-bottom: 20px;">
                <p style="margin-top: 0px; margin-bottom: 5px; font-weight: bold;">◆ Protocolo: {proc['protocolo']}</p>
                <p style="margin-top: 0px; margin-bottom: 3px; margin-left: 25px;">🔴 Status Antigo: {proc['antigo']}</p>
                <p style="margin-top: 0px; margin-bottom: 3px; margin-left: 25px;">🟢 Status Novo: {proc['novo']}</p>
                <p style="margin-top: 0px; margin-bottom: 45px; margin-left: 25px; color: #666666; font-size: 13px;">Verificado em: {agora_str}</p>
            </div>
            """

        html_corpo += f"""
            <p style="margin-top: 20px; margin-bottom: 20px;">
                A planilha <a href="{LINK_PLANILHA}" style="color: #1a73e8; text-decoration: underline; font-weight: bold;">'monitor_protocolos.xlsx'</a> foi atualizada.
            </p>
            <p style="margin-bottom: 0px;">Atenciosamente,</p>
            <p style="margin-top: 0px;">Robô.</p>
        </body>
        </html>
        """

        msg.attach(MIMEText(html_corpo, 'html'))

        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(EMAIL_REMETENTE, str(SENHA_REMETENTE))
        server.sendmail(EMAIL_REMETENTE, EMAIL_DESTINATARIOS, msg.as_string())
        server.quit()
        
        print("✉️ E-mail de alerta enviado com sucesso!")
    except Exception as e:
        print(f"❌ Erro ao enviar e-mail: {e}")


# ==========================================
# 🤖 EXECUÇÃO PRINCIPAL DO ROBÔ
# ==========================================
def executar_robo():
    print(f"\n===== INICIANDO VERIFICAÇÃO: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')} =====")
    nome_planilha = "monitor_protocolos.xlsx"
    nome_aba = "Santana de Parnaíba"

    df = None
    for linha_cabecalho in [0, 1]:
        try:
            temp_df = pd.read_excel(nome_planilha, sheet_name=nome_aba, header=linha_cabecalho, dtype=str)
            temp_df.columns = [str(col).strip().upper() for col in temp_df.columns]
            if "PROTOCOLO" in temp_df.columns and "ATIVO" in temp_df.columns:
                df = temp_df
                break
        except Exception:
            continue

    if df is None:
        print(f"❌ Erro crítico: Não encontrei as colunas na aba '{nome_aba}'. Verifique os cabeçalhos!")
        return

    df = df.fillna("")
    print(f"📊 Aba '{nome_aba}' carregada com sucesso! Encontradas {len(df)} linhas.")

    driver = None
    try:
        print("🔧 Configurando as opções do Chrome para a nuvem...")
        opcoes = webdriver.ChromeOptions()
        opcoes.add_argument("--headless=new")
        opcoes.add_argument("--no-sandbox")
        opcoes.add_argument("--disable-dev-shm-usage")
        opcoes.add_argument("--disable-gpu")
        opcoes.add_argument("--window-size=1920,1080")
        
        print("🌐 Tentando abrir o navegador Chrome...")
        driver = webdriver.Chrome(options=opcoes)
        wait = WebDriverWait(driver, 30)
        print("✅ Navegador aberto com sucesso na nuvem!")

        print("Abrindo a tela de login do portal...")
        driver.get("https://santanadeparnaiba.aprova.com.br/login")
        
        print("Preenchendo os dados de acesso para consulta...")
        campo_email = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@type='email' or @name='email' or @id='email']")))
        # E-mail mantido unicamente para a autenticação interna do robô na página
        campo_email.send_keys("caroline@artesanourbanismo.com.br")
        
        campo_senha = driver.find_element(By.XPATH, "//input[@type='password' or @name='password' or @id='password']")
        campo_senha.send_keys("Artesan@2026")
        
        print("Clicando no botão Entrar...")
        botao_entrar = driver.find_element(By.XPATH, "//button[@type='submit' or contains(., 'Entrar')]")
        driver.execute_script("arguments[0].click();", botao_entrar)
        
        print("Login efetuado! Aguardando o carregamento...")
        time.sleep(15)
        
        print("Navegando até a aba 'Processos'...")
        driver.get("https://santanadeparnaiba.aprova.com.br/processos")
        time.sleep(15)
        
        print("Iniciando varredura inteligente de linhas da tabela...")
        linhas_tabela = driver.find_elements(By.XPATH, "//tbody/tr")
        print(f"✅ {len(linhas_tabela)} processos encontrados na tela.")
        
        dados_portal = []
        for linha_tab in linhas_tabela:
            try:
                celulas = [c.text.strip() for c in linha_tab.find_elements(By.XPATH, "./td") if c.text.strip()]
                if celulas:
                    dados_portal.append({
                        "texto_completo": " ".join(celulas).upper(),
                        "lista_celulas": celulas
                    })
            except Exception:
                continue

        # 🔄 Sincronização e Comparação
        houve_alteracao = False
        agora_str = datetime.now().strftime('%d/%m/%Y %H:%M')
        processos_alterados = []

        col_status = [c for c in df.columns if "STATUS" in c][0]
        col_modificado = [c for c in df.columns if "MODIFICADO" in c][0] if any("MODIFICADO" in c for c in df.columns) else "MODIFICADO"

     print("🔎 Comparando dados da planilha com o portal...")
        for index, text_linha in df.iterrows():
            if str(text_linha["ATIVO"]).strip().upper() == "SIM":
                protocolo = str(text_linha["PROTOCOLO"]).strip().upper()
                status_antigo = str(text_linha[col_status]).strip()

                # PASSO 3: Limpa os espaços do protocolo da planilha para comparar com segurança
                protocolo_planilha_limpo = protocolo.replace(" ", "")

                linha_encontrada = None
                for dado in dados_portal:
                    # Limpa os espaços do texto capturado no portal
                    texto_portal_limpo = dado["texto_completo"].replace(" ", "")
                    
                    if protocolo_planilha_limpo in texto_portal_limpo:
                        linha_encontrada = dado
                        break

                if linha_encontrada:
                    celulas = linha_encontrada["lista_celulas"]
                    status_novo = celulas[-2] if len(celulas) >= 2 else celulas[-1]
                    
                    if status_novo == "" and len(celulas) >= 3:
                        status_novo = celulas[-3]

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
                else:
                    print(f"❓ {protocolo}: Não foi localizado na página atual do portal.")

        for col in df.columns:
            df[col] = df[col].astype(str).replace('nan', '')

        if houve_alteracao:
            print("💾 Salvando alterações na planilha...")
            with pd.ExcelWriter(nome_planilha, engine='openpyxl', mode='w') as writer:
                df.to_excel(writer, sheet_name=nome_aba, index=False)
            print("🎉 Planilha atualizada com sucesso!")
            
            # CORREÇÃO DO PASSO 2: Ajustado com dois "SS" para conversar com a função
            enviar_email_alerta(processos_alterados)
        else:
            print("☕ Nenhuma alteração encontrada. Tudo atualizado!")
