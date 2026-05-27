import os
import time
from datetime import datetime
import pandas as pd
from selenium import webdriver
# Se você usa o email:
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# ==================== CONFIGURAÇÕES DE ACESSO E LINKS ====================
EMAIL_REMETENTE = "mikaellevictoria2017@gmail.com"
SENHA_REMETENTE = "ihfxftkgihyuniob"  # 👈 Suas 16 letras da senha de app da Google
EMAIL_DESTINATARIOS = ["mikaellevictoria2017@gmail.com"]

# Cole aqui o link de compartilhamento da sua planilha (Google Drive, OneDrive ou SharePoint)
LINK_PLANILHA = "https://artesanourbanismo-my.sharepoint.com/:x:/g/personal/mvitoria_artesanourbanismo_com_br/IQDpdOR5HECpRJENC4oXLY81ATUwHlYq0zlcKp7o2ueXTrw?e=dZAqpQ" 
# =========================================================================

def enviar_email_alerta(processos_alterados):
    try:
        msg = MIMEMultipart('alternative')
        msg['From'] = EMAIL_REMETENTE
        msg['To'] = ", ".join(EMAIL_DESTINATARIOS)
        msg['Subject'] = f"📢 [Aviso] Mudança de Status em Processos - {datetime.now().strftime('%d/%m/%Y')}"
        
        # Montagem do corpo do e-mail seguindo rigorosamente o seu layout desejado
        html_corpo = f"""
        <html>
        <body style="font-family: Arial, sans-serif; font-size: 14px; color: #333333; line-height: 1.6;">
            <p style="margin-bottom: 15px;">Olá Artesano,</p>
            
            <p style="margin-bottom: 20px;">O robô de monitoramento identificou mudanças de status nos seguintes processos:</p>
            
            <p style="margin-bottom: 15px;"><strong>🔹SANTANA DE PARNAÍBA</strong></p>
        """
        
        agora_str = datetime.now().strftime('%d/%m/%Y %H:%M')
        
        for proc in processos_alterados:
            html_corpo += f"""
            <div style="margin-bottom: 20px;">
                <p style="margin: 0px 0px 5px 0px;">🔹 Protocolo: {proc['protocolo']}</p>
                <p style="margin: 0px 0px 3px 25px;">🔴Status Antigo: {proc['antigo']}</p>
                <p style="margin: 0px 0px 3px 25px;">🟢Status Novo: {proc['novo']}</p>
                <p style="margin: 0px 0px 0px 45px; color: #666666; font-size: 13px;">Verificado em: {agora_str}</p>
            </div>
            """
            
        html_corpo += f"""
            <p style="margin-top: 20px; margin-bottom: 20px;">
                A planilha <a href="{LINK_PLANILHA}" style="color: #1a73e8; text-decoration: underline; font-weight: bold;">'monitor_protocolos.xlsx'</a> já foi atualizada automaticamente.
            
            <p style="margin-bottom: 0px;">Atenciosamente,</p>
            <p style="margin-top: 0px;">Robô.</p>
        </body>
        </html>
        """
        
        msg.attach(MIMEText(html_corpo, 'html'))
        
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(EMAIL_REMETENTE, SENHA_REMETENTE)
        server.sendmail(EMAIL_REMETENTE, EMAIL_DESTINATARIOS, msg.as_string())
        server.quit()
        print("📧 E-mail de alerta enviado com o novo layout de sucesso!")
    except Exception as e:
        print(f"❌ Erro ao enviar e-mail: {e}")

def executar_robo():
    print(f"\n===== INICIANDO VERIFICAÇÃO: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')} =====")
    nome_planilha = "monitor_protocolos.xlsx"
    nome_aba = "Santana de Parnaíba" # 👉 Define a aba alvo

    df = None
    linha_correta = 0
    for linha_cabecalho in [0, 1]:
        try:
            # Lendo especificamente a aba de Santana de Parnaíba
            temp_df = pd.read_excel(nome_planilha, sheet_name=nome_aba, header=linha_cabecalho, dtype=str)
            temp_df.columns = [str(col).strip().upper() for col in temp_df.columns]
            if "PROTOCOLO" in temp_df.columns and "ATIVO" in temp_df.columns:
                df = temp_df
                linha_correta = linha_cabecalho
                break
        except Exception as e:
            continue

  if df is None:
        print(f"❌ Erro crítico: Não encontrei as colunas na aba '{nome_aba}'. Verifique os cabeçalhos!")
        return

    df = df.fillna("")
    print(f"📊 Aba '{nome_aba}' carregada com sucesso! Encontradas {len(df)} linhas.")

    # 🛡️ === INÍCIO DO BLOQUEADOR DE ERROS ===
    driver = None  # Deixa o espaço do navegador reservado
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
        print("✅ Navegador aberto com sucesso na nuvem!")

        # #️⃣ LOGIN NO PORTAL
        print("Abrindo a tela de login do portal...")
        driver.get("https://santanadeparnaiba.aprova.com.br/login")
        
        print("Preenchendo os dados de acesso...")
        campo_email = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@type='email' or @name='email' or @id='email']")))
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
        
        print("Mapeando colunas do portal...")
        colunas_cabecalho = driver.find_elements(By.XPATH, "//thead//th | //tr[th or td]//th")
        
        indices_portal = {
            "DOCUMENTO": -1, "REQUERENTE": -1, "PROPRIETARIO": -1, "CRIADO": -1, "ACAO": -1, "STATUS": -1
        }
        
        for idx, col in enumerate(colunas_cabecalho):
            txt = col.text.strip().upper()
            if "DOCUMENTO" in txt or "REQUERENTE" in txt: indices_portal["DOCUMENTO"] = idx
            elif "REQUERENTE" in txt or "REMETENTE" in txt: indices_portal["REQUERENTE"] = idx
            elif "PROPRIETÁRIO" in txt or "DESTINATÁRIO" in txt: indices_portal["PROPRIETARIO"] = idx
            elif "CRIADO" in txt: indices_portal["CRIADO"] = idx
            elif "AÇÃO" in txt: indices_portal["ACAO"] = idx
            elif "STATUS" in txt: indices_portal["STATUS"] = idx
        
        print("Iniciando varredura em busca dos protocolos (PMSP)...")
        linhas_tabela = driver.find_elements(By.XPATH, "//tr[contains(., 'PMSP')]")
        print(f"✅ {len(linhas_tabela)} processos encontrados na tela.")
        
        dados_portal = {}
        # (Aqui o robô coleta os dados da tabela do site)

        # 🔄 === Sincronização dos dados com a planilha ===
        houve_alteracao = False
        agora_str = datetime.now().strftime('%d/%m/%Y %H:%M')
        processos_alterados = []

        col_doc = [c for c in df.columns if "REQUERIMENTO" in c or "DOCUMENTO" in c][0]
        col_req = [c for c in df.columns if "REQUERENTE" in c or "REMETENTE" in c][0]
        col_prop = [c for c in df.columns if "PROPRIETÁRIO" in c or "DESTINATÁRIO" in c][0]
        col_criado = [c for c in df.columns if "CRIADO" in c][0]
        col_acao = [c for c in df.columns if "AÇÃO" in c][0]
        col_status = [c for c in df.columns if "STATUS" in c][0]
        col_modificado = [c for c in df.columns if "MODIFICADO" in c][0] if any("MODIFICADO" in c for c in df.columns) else "MODIFICADO"

        for index, text_linha in df.iterrows():
            if str(text_linha["ATIVO"]).strip().upper() == "SIM":
                protocolo = str(text_linha["PROTOCOLO"]).strip()
                status_antigo = str(text_linha[col_status]).strip()

                dados_proc = None
                for k in dados_portal.keys():
                    if protocolo in k or k in protocolo:
                        dados_proc = dados_portal[k]
                        break

                if dados_proc:
                    df.at[index, col_doc] = str(dados_proc["doc"])
                    df.at[index, col_req] = str(dados_proc["req"])
                    df.at[index, col_prop] = str(dados_proc["prop"])
                    df.at[index, col_criado] = str(dados_proc["criado"])
                    df.at[index, col_acao] = str(dados_proc["acao"])

                    status_novo = dados_proc["status"]

                    if status_antigo != status_novo:
                        print(f"⚠️ ALTERAÇÃO DETECTADA! {protocolo} mudou para '{status_novo}'")
                        processos_alterados.append({
                            'protocolo': protocolo, 'antigo': status_antigo, 'novo': status_novo
                        })
                        df.at[index, col_status] = str(status_novo)
                        df.at[index, col_modificado] = str(agora_str)
                        houve_alteracao = True
                    else:
                        print(f"✅ {protocolo}: Status igual ao do portal.")

        for col in df.columns:
            df[col] = df[col].astype(str).replace('nan', '')

    except Exception as erro:
        mensagem_erro = f"Atenção! O robô falhou.\nDetalhe do erro: {erro}"
        print(f"🚨 {mensagem_erro}")

    finally:
        if driver is not None:
            driver.quit()
            print("Navegador fechado com segurança pelo sistema de proteção.")

# 🤖 INICIALIZAÇÃO DO SCRIPT NA NUVEM
if __name__ == "__main__":
    print("🤖 ROBÔ MULTI-ABAS ATIVADO!")
    executar_robo()
    print("✅ Execução concluída!")
