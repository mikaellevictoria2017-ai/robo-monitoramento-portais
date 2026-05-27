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

    # ... (toda a parte de cima do 'for linha_cabecalho' continua IGUAL)

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

        # -----------------------------------------------------------------
        # ⚠️ IMPORTANTE: ATENÇÃO À INDENTAÇÃO (RECUO) COLA O SEU CÓDIGO AQUI
        # Todo o seu loop antigo (for index, linha in df.iterrows():) 
        # e as buscas no site precisam ficar empurradas para a direita,
        # alinhadas abaixo desse 'try'.
        # -----------------------------------------------------------------

    except Exception as erro:
        # 🚨 Se o site cair ou travar, ele executa isso e não quebra a máquina:
        mensagem_erro = f"Atenção Micaelle! O robô falhou.\nDetalhe do erro: {erro}"
        print(f"🚨 {mensagem_erro}")
        # Se quiser ativar o aviso por e-mail depois, basta criar a função para isso aqui.

    finally:
        # 🔒 Segurança Máxima: Sempre fecha o Chrome no final, dando certo ou errado
        if driver is not None:
            driver.quit()
            print("Navegador fechado com segurança pelo sistema de proteção.")
    
    try:
        # LOGIN NO PORTAL
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
            if "DOCUMENTO" in txt or "REQUERIMENTO" in txt: indices_portal["DOCUMENTO"] = idx
            elif "REQUERENTE" in txt or "REMETENTE" in txt: indices_portal["REQUERENTE"] = idx
            elif "PROPRIETÁRIO" in txt or "DESTINATÁRIO" in txt: indices_portal["PROPRIETARIO"] = idx
            elif "CRIADO" in txt: indices_portal["CRIADO"] = idx
            elif "AÇÃO" in txt: indices_portal["ACAO"] = idx
            elif "STATUS" in txt: indices_portal["STATUS"] = idx

        print("Iniciando varredura em busca dos protocolos (PMSP)...")
        linhas_tabela = driver.find_elements(By.XPATH, "//tr[contains(., 'PMSP')]")
        print(f"✅ {len(linhas_tabela)} processos encontrados na tela.")
        
        dados_portal = {}
        
        for linha in linhas_tabela:
            try:
                celulas = linha.find_elements(By.XPATH, "./td")
                if len(celulas) >= 2:
                    protocolo_texto = ""
                    for c in celulas:
                        if "PMSP" in c.text:
                            protocolo_texto = c.text.strip().split('\n')[0]
                            break
                    
                    if protocolo_texto:
                        dados_portal[protocolo_texto] = {
                            "doc": celulas[indices_portal["DOCUMENTO"]].text.strip() if indices_portal["DOCUMENTO"] != -1 else celulas[1].text.strip(),
                            "req": celulas[indices_portal["REQUERENTE"]].text.strip() if indices_portal["REQUERENTE"] != -1 else celulas[2].text.strip(),
                            "prop": celulas[indices_portal["PROPRIETARIO"]].text.strip() if indices_portal["PROPRIETARIO"] != -1 else celulas[3].text.strip(),
                            "criado": celulas[indices_portal["CRIADO"]].text.strip() if indices_portal["CRIADO"] != -1 else celulas[4].text.strip(),
                            "acao": celulas[indices_portal["ACAO"]].text.strip() if indices_portal["ACAO"] != -1 else celulas[5].text.strip(),
                            "status": celulas[indices_portal["STATUS"]].text.strip() if indices_portal["STATUS"] != -1 else celulas[6].text.strip()
                        }
                        print(f"🔍 Portal diz -> {protocolo_texto}: {dados_portal[protocolo_texto]['status']}")
            except:
                continue
                    
        # COMPARAÇÃO E PREENCHIMENTO DE DADOS
        houve_alteracao = False
        agora_str = datetime.now().strftime('%d/%m/%Y %H:%M')
        
        col_doc = [c for c in df.columns if "REQUERIMENTO" in c or "DOCUMENTO" in c][0]
        col_req = [c for c in df.columns if "REQUERENTE" in c or "REMETENTE" in c][0]
        col_prop = [c for c in df.columns if "PROPRIETÁRIO" in c or "DESTINATÁRIO" in c][0]
        col_criado = [c for c in df.columns if "CRIADO" in c][0]
        col_acao = [c for c in df.columns if "AÇÃO" in c][0]
        col_status = [c for c in df.columns if "STATUS" in c][0]
        col_modificado = [c for c in df.columns if "MODIFICADO" in c][0] if any("MODIFICADO" in c for c in df.columns) else "MODIFICADO EM"
        
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
                        print(f"🚨 ALTERAÇÃO DETECTADA! {protocolo} mudou para '{status_novo}'")
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
            
        # 🛡️ SALVAMENTO COM DESIGN AUTOMÁTICO PROTEGIDO
        salvo_com_sucesso = False
        while not salvo_com_sucesso:
            try:
                # Carregamos todas as outras abas para não apagá-las ao salvar
                with pd.ExcelFile(nome_planilha) as reader:
                    abas_existentes = {sheet: reader.parse(sheet) for sheet in reader.sheet_names}
                
                # Substituímos os dados apenas da aba de Santana de Parnaíba
                abas_existentes[nome_aba] = df
                
                with pd.ExcelWriter(nome_planilha, engine='openpyxl') as writer:
                    for sheet, dados_aba in abas_existentes.items():
                        # Se for a nossa aba, salvamos respeitando a linha de cabeçalho correta
                        if sheet == nome_aba:
                            dados_aba.to_excel(writer, sheet_name=sheet, startrow=linha_correta, index=False)
                            
                            # Executa o ajuste automático de larguras de células nesta aba
                            worksheet = writer.sheets[sheet]
                            for col in worksheet.columns:
                                max_len = 0
                                col_letter = col[0].column_letter
                                for cell in col:
                                    if cell.value:
                                        max_len = max(max_len, len(str(cell.value)))
                                worksheet.column_dimensions[col_letter].width = max(max_len + 4, 12)
                        else:
                            # Mantém as outras abas intocadas no seu formato original
                            dados_aba.to_excel(writer, sheet_name=sheet, index=False)
                            
                print(f"💾 Planilha salva. Aba '{nome_aba}' atualizada com layout perfeito!")
                salvo_com_sucesso = True
            except PermissionError:
                print("⚠️ AVISO: Feche o arquivo Excel para o robô aplicar as correções de layout!")
                time.sleep(20)
        
        if houve_alteracao:
            enviar_email_alerta(processos_alterados)
            
    except Exception as e:
        print(f"❌ Erro crítico na execução interna do robô: {e}")
    finally:
        driver.quit()
        print("Navegador fechado.")


# ... (tudo igual para cima)
    print("Navegador fechado.")

if __name__ == "__main__":
    print("🤖 ROBÔ MULTI-ABAS ATIVADO!")
    executar_robo()
    print("✅ Execução diária concluída com sucesso! Desligando a máquina da nuvem.")
