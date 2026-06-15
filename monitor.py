import os
import time
import pandas as pd
from datetime import datetime
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from playwright.sync_api import sync_playwright

# ==========================================
# CONFIGURAÇÕES E LINKS
# ==========================================
EMAIL_REMETENTE = "mikaellevictoria2017@gmail.com"
EMAIL_DESTINATARIOS = ["santos.micaelle2006@gmail.com"]

# LINK DO SEU GOOGLE FORMS
LINK_ENTRADA_GOOGLE_FORMS = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRh-7SIMziaShR1rqLpSnBabRJAIceLSZ6dO0zklOcOg_twfc9G6cwdRGQk1vL2y6lniAmH0mSh6Xw1/pub?gid=1314499551&single=true&output=csv"

# Variáveis secretas do GitHub
USER_PORTAL = os.getenv("USER_PORTAL", "")
SENHA_PORTAL = os.getenv("SENHA_PORTAL", "")
SENHA_GMAIL = os.getenv("SENHA_GMAIL", "")

agora_str = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
print(f"===== INICIANDO MONITORAMENTO VIA GOOGLE FORMS: {agora_str} =====")

# ==========================================
# 1. LEITURA DOS PROTOCOLOS DA PLANILHA
# ==========================================
try:
    df = pd.read_csv(LINK_ENTRADA_GOOGLE_FORMS)
    df.columns = [str(c).strip().upper() for c in df.columns]
    print(f"📥 Dados carregados! Colunas: {list(df.columns)}")
    
    col_protocolo = [c for c in df.columns if "PROTOCOLO" in c or "NUMER" in c or "PROCESSO" in c][0]
    
    # Prepara colunas de controle se não existirem
    if "STATUS ATUAL" not in df.columns:
        df["STATUS ATUAL"] = "Aguardando..."
    if "ÚLTIMA AÇÃO" not in df.columns:
        df["ÚLTIMA AÇÃO"] = "Nenhuma"
    if "MODIFICADO EM" not in df.columns:
        df["MODIFICADO EM"] = agora_str

    protocolos_verificar = df[col_protocolo].dropna().astype(str).tolist()
    print(f"🔍 Protocolos para checagem: {protocolos_verificar}")

except Exception as e:
    print(f"❌ Erro ao ler os dados do Google Forms: {e}")
    exit(1)

# ==========================================
# 2. AUTOMAÇÃO NO PORTAL (PLAYWRIGHT) E COMPARAÇÃO
# ==========================================
processos_alterados = [] # Lista que vai guardar quem mudou de status

try:
    with sync_playwright() as p:
        navegador = p.chromium.launch(headless=True)
        pagina = navegador.new_page()

        print("🌐 Acessando o portal de Santana de Parnaíba...")
        pagina.goto("https://santanadeparnaiba.aprova.com.br/")
        pagina.get_by_text("Acessar minha conta").click()
        time.sleep(2) 

        print("🔑 Fazendo login...")
        pagina.locator("input[type='email']").fill(USER_PORTAL)
        pagina.locator("input[type='password']").fill(SENHA_PORTAL)
        pagina.get_by_role("button", name="Entrar").click()

        pagina.wait_for_load_state("networkidle")
        time.sleep(3)
        pagina.get_by_text("Processos", exact=True).click()
        time.sleep(3) # Aguarda tabela carregar

        # Agora o robô pesquisa protocolo por protocolo da sua planilha
        for index, row in df.iterrows():
            protocolo_planilha = str(row[col_protocolo]).strip().upper()
            status_antigo = str(row["STATUS ATUAL"]).strip()
            
            print(f"🔎 Pesquisando: {protocolo_planilha}")
            busca_input = pagina.locator("input[placeholder='Buscar aqui']").first
            busca_input.fill(protocolo_planilha)
            busca_input.press("Enter")
            time.sleep(3)

            try:
                # Extrai o novo status da tabela
                linha_processo = pagina.locator(f"tr:has-text('{protocolo_planilha}')")
                status_novo = linha_processo.locator("td").nth(6).inner_text().strip()
                
                # A MÁGICA: Compara o status do site com o da planilha
                if status_novo != status_antigo and status_antigo != "Aguardando...":
                    print(f"⚠️ MUDANÇA DETECTADA! {protocolo_planilha}: {status_antigo} -> {status_novo}")
                    processos_alterados.append({
                        "protocolo": protocolo_planilha,
                        "velho": status_antigo,
                        "novo": status_novo
                    })

                # Atualiza a planilha (Dataframe) com o status novo
                df.at[index, "STATUS ATUAL"] = status_novo
                df.at[index, "MODIFICADO EM"] = agora_str

            except Exception as e:
                print(f"⚠️ Não encontrou status para o protocolo {protocolo_planilha}.")
                
            busca_input.fill("") # Limpa a barra de pesquisa para o próximo
            
        navegador.close()

except Exception as e:
    print(f"❌ Falha crítica no robô: {e}")

# ==========================================
# 3. SALVAR RELATÓRIO E ENVIAR E-MAIL
# ==========================================
print("💾 Gravando base de dados atualizada...")
# O erro de sintaxe (parêntese) foi corrigido aqui:
df.to_html("monitor_protocolos.html", index=False, border=1, classes='tabela-processos')

# Só envia e-mail se a lista 'processos_alterados' tiver coisas dentro
if processos_alterados and SENHA_GMAIL:
    try:
        msg = MIMEMultipart('alternative')
        msg['From'] = EMAIL_REMETENTE
        msg['To'] = ", ".join(EMAIL_DESTINATARIOS)
        msg['Subject'] = "⚠️ Alerta: Status de Protocolo Atualizado!"
        
        blocos = ""
        for p in processos_alterados:
            blocos += f"<li><strong>Protocolo:</strong> {p['protocolo']}<br>De: <em>{p['velho']}</em> ➡️ Para: <strong>{p['novo']}</strong></li><br>"
        
        corpo_html = f"<html><body><p>Olá! O sistema detectou movimentações nos processos abaixo:</p><ul>{blocos}</ul></body></html>"
        msg.attach(MIMEText(corpo_html, 'html'))
        
        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()
            server.login(EMAIL_REMETENTE, SENHA_GMAIL)
            server.sendmail(EMAIL_REMETENTE, EMAIL_DESTINATARIOS, msg.as_string())
        print("✉️ E-mail de notificação enviado com sucesso!")
    except Exception as e:
        print(f"❌ Erro ao enviar e-mail: {e}")
else:
    print("🦥 Varredura finalizada. Nenhuma movimentação nova detectada.")
