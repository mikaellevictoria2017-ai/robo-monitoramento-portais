```python
import os
import time
import pandas as pd
from datetime import datetime
import smtplib
import subprocess

from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ==========================================
# CONFIGURAÇÕES
# ==========================================
EMAIL_REMETENTE = "mikaellevictoria2017@gmail.com"
EMAIL_DESTINATARIOS = ["santos.micaelle2006@gmail.com"]

LINK_PLANILHA = "https://artesanourbanismo-my.sharepoint.com/:x:/g/personal/mvitoria_artesanourbanismo_com_br/IQDgXvD3n6RTRZsZo63IiGBXAeSMCTvv1qBTTDNAD3d1_jE?e=YfzG2P"

NOME_PLANILHA = "monitor_protocolos.xlsx"
NOME_ABA = "Santana de Parnaíba"

USER_PORTAL = os.getenv("USER_PORTAL")
SENHA_PORTAL = os.getenv("SENHA_PORTAL")
SENHA_GMAIL = os.getenv("SENHA_GMAIL")

agora_str = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

print(f"\n===== INICIANDO MONITORAMENTO =====\n")

# ==========================================
# FUNÇÃO NORMALIZAR
# ==========================================
def normalizar(texto):
    return (
        str(texto)
        .strip()
        .upper()
        .replace("-", "")
        .replace("/", "")
        .replace(" ", "")
    )

# ==========================================
# LEITURA PLANILHA
# ==========================================
try:

    df = pd.read_excel(NOME_PLANILHA, sheet_name=NOME_ABA)

    df.columns = [str(c).strip().upper() for c in df.columns]

    col_protocolo = [c for c in df.columns if "PROTOCOLO" in c or "PROCESSO" in c][0]

    col_status = (
        [c for c in df.columns if "STATUS" in c or "SITUA" in c][0]
        if any("STATUS" in c or "SITUA" in c for c in df.columns)
        else "STATUS"
    )

    col_modificado = (
        [c for c in df.columns if "MODIF" in c or "DATA" in c][0]
        if any("MODIF" in c or "DATA" in c for c in df.columns)
        else "MODIFICADO EM"
    )

    col_ativo = (
        [c for c in df.columns if "ATIVO" in c][0]
        if any("ATIVO" in c for c in df.columns)
        else None
    )

    print("✅ Planilha carregada")

except Exception as e:
    print(f"❌ Erro ao ler planilha: {e}")
    exit()

# ==========================================
# SELENIUM
# ==========================================
options = Options()

options.add_argument("--headless=new")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--window-size=1920,1080")

driver = webdriver.Chrome(options=options)

wait = WebDriverWait(driver, 30)

dados_portal = {}

try:

    print("🌐 Abrindo portal...")

    driver.get("https://santanadeparnaiba.aprova.com.br/login")

    wait.until(EC.presence_of_element_located((By.TAG_NAME, "input")))

    inputs = driver.find_elements(By.TAG_NAME, "input")

    inputs[0].send_keys(USER_PORTAL)
    inputs[1].send_keys(SENHA_PORTAL)

    driver.find_element(By.CSS_SELECTOR, "button[type='submit']").click()

    print("🔐 Login realizado")

    time.sleep(8)

    driver.get("https://santanadeparnaiba.aprova.com.br/processos")

    time.sleep(10)

    print("📋 Capturando processos...")

    linhas = driver.find_elements(By.XPATH, "//tbody/tr")

    for linha in linhas:

        texto = linha.text.strip()

        if not texto:
            continue

        partes = [p.strip() for p in texto.split("\n") if p.strip()]

        if len(partes) < 2:
            continue

        protocolo = partes[0]
        status = partes[-1]

        dados_portal[protocolo] = status

    print(f"✅ {len(dados_portal)} processos capturados")

    for k, v in dados_portal.items():
        print(f"{k} -> {v}")

except Exception as e:
    print(f"❌ Erro Selenium: {e}")

finally:
    driver.quit()

# ==========================================
# COMPARAÇÃO
# ==========================================
processos_alterados = []

print("⚖️ Comparando protocolos...")

for index, row in df.iterrows():

    if col_ativo:

        ativo = str(row.get(col_ativo, "")).strip().upper()

        if ativo != "SIM":
            continue

    protocolo_planilha = str(row[col_protocolo]).strip()

    status_antigo = str(row.get(col_status, "")).strip()

    for protocolo_site, status_site in dados_portal.items():

        if normalizar(protocolo_planilha) == normalizar(protocolo_site):

            status_novo = str(status_site).strip()

            if normalizar(status_antigo) != normalizar(status_novo):

                print(f"\n⚠️ ALTERAÇÃO DETECTADA")
                print(f"📄 {protocolo_planilha}")
                print(f"📌 ANTIGO: {status_antigo}")
                print(f"✅ NOVO: {status_novo}")

                processos_alterados.append({
                    "protocolo": protocolo_planilha,
                    "antigo": status_antigo,
                    "novo": status_novo
                })

                # ==========================================
                # LIMPA LINHA INTEIRA
                # ==========================================
                for col in df.columns:
                    df.at[index, col] = ""

                # ==========================================
                # REESCREVE LINHA
                # ==========================================
                df.at[index, col_protocolo] = protocolo_planilha
                df.at[index, col_status] = status_novo
                df.at[index, col_modificado] = agora_str

                if col_ativo:
                    df.at[index, col_ativo] = "SIM"

            break

# ==========================================
# SALVAR EXCEL
# ==========================================
if processos_alterados:

    try:

        print("\n💾 Salvando planilha...")

        with pd.ExcelWriter(
            NOME_PLANILHA,
            engine="openpyxl",
            mode="a",
            if_sheet_exists="replace"
        ) as writer:

            df.to_excel(writer, sheet_name=NOME_ABA, index=False)

        print("✅ Planilha atualizada")

        # ==========================================
        # GITHUB
        # ==========================================
        print("🚀 Atualizando GitHub...")

        subprocess.run(["git", "config", "user.name", "Automated Robot"], check=True)

        subprocess.run(["git", "config", "user.email", "robot@artesano.com"], check=True)

        subprocess.run(["git", "add", NOME_PLANILHA], check=True)

        subprocess.run([
            "git",
            "commit",
            "-m",
            f"🤖 Atualização automática {agora_str}"
        ], check=True)

        subprocess.run(["git", "push"], check=True)

        print("✅ GitHub atualizado")

        # ==========================================
        # E-MAIL
        # ==========================================
        if SENHA_GMAIL:

            print("✉️ Enviando e-mail...")

            msg = MIMEMultipart("alternative")

            msg["From"] = EMAIL_REMETENTE
            msg["To"] = ", ".join(EMAIL_DESTINATARIOS)
            msg["Subject"] = "⚠️ Protocolos Atualizados"

            corpo_html = f"""
            <html>
            <body style="font-family: Arial;">

                <h2>⚠️ Alterações Detectadas</h2>

                <ul>
            """

            for p in processos_alterados:

                corpo_html += f"""
                    <li>
                        <b>Protocolo:</b> {p['protocolo']}<br>
                        <b>Antigo:</b> {p['antigo']}<br>
                        <b>Novo:</b> {p['novo']}<br><br>
                    </li>
                """

            corpo_html += f"""
                </ul>

                <p>
                    <a href="{LINK_PLANILHA}">
                        Abrir Planilha
                    </a>
                </p>

                <br>

                <small>
                    Executado em {agora_str}
                </small>

            </body>
            </html>
            """

            msg.attach(MIMEText(corpo_html, "html"))

            with smtplib.SMTP("smtp.gmail.com", 587) as server:

                server.starttls()

                server.login(
                    EMAIL_REMETENTE,
                    SENHA_GMAIL
                )

                server.sendmail(
                    EMAIL_REMETENTE,
                    EMAIL_DESTINATARIOS,
                    msg.as_string()
                )

            print("✅ E-mail enviado")

    except Exception as e:
        print(f"❌ Erro salvamento/email: {e}")

else:

    print("\n🦥 Nenhuma alteração encontrada.")
```

