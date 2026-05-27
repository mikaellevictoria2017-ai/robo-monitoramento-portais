import os
import time
from datetime import datetime
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

def enviar_email_notificacao(mensagem):
    try:
        print("✉️ Enviando notificação por e-mail...")
        # Seu código de e-mail entra aqui se configurado
    except Exception as e:
        print(f"❌ Erro ao enviar e-mail: {e}")

def executar_robo():
    print(f"\n===== INICIANDO VERIFICAÇÃO: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')} =====")
    nome_planilha = "monitor_protocolos.xlsx"
    nome_aba = "Santana de Parnaíba"

    # === PROCESSAMENTO DA PLANILHA ===
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
        
        print("Iniciando varredura inteligente de linhas da tabela...")
        linhas_tabela = driver.find_elements(By.XPATH, "//tbody/tr")
        print(f"✅ {len(linhas_tabela)} processos encontrados na tela.")
        
        # Mapeamento do conteúdo bruto das linhas
        dados_portal = []
        for linha_tab in linhas_tabela:
            try:
                celulas = [c.text.strip() for c in linha_tab.find_elements(By.XPATH, "./td") if c.text.strip()]
                if celulas:
                    # Guardamos o texto completo da linha e as células para análise posterior
                    dados_portal.append({
                        "texto_completo": " ".join(celulas).upper(),
                        "lista_celulas": celulas
                    })
            except Exception:
                continue

        # 🔄 === Sincronização e Comparação dos dados com a planilha ===
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

                linha_encontrada = None
                for dado in dados_portal:
                    # Se o número do protocolo faz parte do texto dessa linha da tabela, achamos!
                    if protocolo in dado["texto_completo"]:
                        linha_encontrada = dado
                        break

                if linha_encontrada:
                    # O status no Aprova Digital costuma ser uma das últimas colunas da linha
                    # Pegamos a penúltima ou última célula que contém o texto do status (ex: "Análise", "Deferido")
                    celulas = linha_encontrada["lista_celulas"]
                    status_novo = celulas[-2] if len(celulas) >= 2 else celulas[-1]
                    
                    # Se a última célula for um botão de ação vazio, usamos a anterior
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
            
            msg_email = "O robô identificou alterações de status:\n"
            for p in processos_alterados:
                msg_email += f"- Protocolo {p['protocolo']}: mudou de '{p['antigo']}' para '{p['novo']}'\n"
            enviar_email_notificacao(msg_email)
        else:
            print("☕ Nenhuma alteração encontrada. Tudo atualizado!")

    except Exception as erro:
        print(f"🚨 Atenção! O robô falhou. Detalhe do erro: {erro}")

    finally:
        if driver is not None:
            driver.quit()
            print("Navegador fechado com segurança pelo sistema de proteção.")

if __name__ == "__main__":
    print("🤖 ROBÔ MULTI-ABAS ATIVADO!")
    executar_robo()
    print("✅ Execução concluída!")
