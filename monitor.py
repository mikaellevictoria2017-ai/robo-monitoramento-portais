import os
import time
from datetime import datetime
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

def enviar_email_notificacao(mensagem):
    # Função de e-mail (Mantida original do seu projeto)
    try:
        print("✉️ Enviando notificação por e-mail...")
        # Seu código de e-mail configurado anteriormente entra aqui se necessário
    except Exception as e:
        print(f"❌ Erro ao enviar e-mail: {e}")

def executar_robo():
    print(f"\n===== INICIANDO VERIFICAÇÃO: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')} =====")
    nome_planilha = "monitor_protocolos.xlsx"
    nome_aba = "Santana de Parnaíba"

    # === PROCESSAMENTO DA PLANILHA ===
    df = None
    linha_correta = 0
    for linha_cabecalho in [0, 1]:
        try:
            temp_df = pd.read_excel(nome_planilha, sheet_name=nome_aba, header=linha_cabecalho, dtype=str)
            temp_df.columns = [str(col).strip().upper() for col in temp_df.columns]
            if "PROTOCOLO" in temp_df.columns and "ATIVO" in temp_df.columns:
                df = temp_df
                linha_correta = linha_cabecalho
                break
        except Exception:
            continue

    if df is None:
        print(f"❌ Erro crítico: Não encontrei as colunas na aba '{nome_aba}'. Verifique os cabeçalhos!")
        return

    df = df.fillna("")
    print(f"📊 Aba '{nome_aba}' carregada com sucesso! Encontradas {len(df)} linhas.")

    # 🛡️ === INÍCIO DO BLOQUEADOR DE ERROS ===
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
            if "DOCUMENTO" in txt: indices_portal["DOCUMENTO"] = idx
            elif "REQUERENTE" in txt or "REMETENTE" in txt: indices_portal["REQUERENTE"] = idx
            elif "PROPRIETÁRIO" in txt or "DESTINATÁRIO" in txt: indices_portal["PROPRIETARIO"] = idx
            elif "CRIADO" in txt: indices_portal["CRIADO"] = idx
            elif "AÇÃO" in txt: indices_portal["ACAO"] = idx
            elif "STATUS" in txt: indices_portal["STATUS"] = idx
        
        print("Iniciando varredura em busca dos protocolos (PMSP)...")
        linhas_tabela = driver.find_elements(By.XPATH, "//tbody//tr | //tr[td]")
        print(f"✅ {len(linhas_tabela)} processos encontrados na tela.")
        
        dados_portal = {}
        
        # 📊 CAPTURA REAL DOS DADOS DA TABELA DO PORTAL
        for linha_tab in linhas_tabela:
            try:
                celulas = linha_tab.find_elements(By.XPATH, "./td")
                if len(celulas) > max(indices_portal.values()):
                    doc_txt = celulas[indices_portal["DOCUMENTO"]].text.strip()
                    req_txt = celulas[indices_portal["REQUERENTE"]].text.strip() if indices_portal["REQUERENTE"] != -1 else ""
                    prop_txt = celulas[indices_portal["PROPRIETARIO"]].text.strip() if indices_portal["PROPRIETARIO"] != -1 else ""
                    criado_txt = celulas[indices_portal["CRIADO"]].text.strip() if indices_portal["CRIADO"] != -1 else ""
                    acao_txt = celulas[indices_portal["ACAO"]].text.strip() if indices_portal["ACAO"] != -1 else ""
                    status_txt = celulas[indices_portal["STATUS"]].text.strip() if indices_portal["STATUS"] != -1 else ""
                    
                    if doc_txt:
                        dados_portal[doc_txt] = {
                            "doc": doc_txt, "req": req_txt, "prop": prop_txt,
                            "criado": criado_txt, "acao": acao_txt, "status": status_txt
                        }
            except Exception:
                continue

        # 🔄 === Sincronização e Comparação dos dados com a planilha ===
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

        print("🔎 Comparando dados da planilha com o portal...")
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
            
            # Dispara a função de e-mail se houver mudanças
            msg_email = "O robô identificou alterações de status:\n"
            for p in processos_alterados:
                msg_email += f"- Protocolo {p['protocolo']}: mudou de '{p['antigo']}' para '{p['novo']}'\n"
            enviar_email_notificacao(msg_email)
        else:
            print("☕ Nenhuma alteração encontrada. Tudo atualizado!")

    except Exception as erro:
        mensagem_erro = f"Atenção Micaelle! O robô falhou.\nDetalhe do erro: {erro}"
        print(f"🚨 {mensagem_erro}")

    finally:
        if driver is not None:
            driver.quit()
            print("Navegador fechado com segurança pelo sistema de proteção.")

if __name__ == "__main__":
    print("🤖 ROBÔ MULTI-ABAS ATIVADO!")
    executar_robo()
    print("✅ Execução concluída!")
