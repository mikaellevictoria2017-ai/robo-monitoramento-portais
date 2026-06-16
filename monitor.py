import pandas as pd
from playwright.sync_api import sync_playwright
import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime

# Configurações
USER = os.getenv("USER_PORTAL")
SENHA = os.getenv("SENHA_PORTAL")
GMAIL = os.getenv("SENHA_GMAIL")
AGORA = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

def main():
    # 1. Carrega dados locais
    df = pd.read_csv("protocolos.csv")
    
    # 2. Scrape Completo do Portal
    portal_data = {}
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        page.goto("https://santanadeparnaiba.aprova.com.br/login")
        page.get_by_placeholder("E-mail").fill(USER)
        page.get_by_placeholder("Digite sua senha").first.fill(SENHA)
        page.get_by_role("button", name="Entrar").click()
        page.wait_for_load_state("networkidle")
        page.goto("https://santanadeparnaiba.aprova.com.br/processos")
        page.wait_for_load_state("networkidle")
        
        # Lê a tabela inteira
        linhas = page.locator("tbody tr").all()
        for linha in linhas:
            t = linha.inner_text().split("\n")
            if len(t) >= 2: portal_data[t[0].strip().upper()] = t[-1].strip()
        browser.close()

    # 3. Sincronização e Mudanças
    mudancas = []
    for proto, status in portal_data.items():
        if proto in df["PROTOCOLO"].values:
            idx = df[df["PROTOCOLO"] == proto].index[0]
            if str(df.at[idx, "STATUS ATUAL"]).strip() != status:
                mudancas.append(f"{proto}: {df.at[idx, 'STATUS ATUAL']} -> {status}")
                df.at[idx, "STATUS ATUAL"] = status
                df.at[idx, "MODIFICADO EM"] = AGORA
        else:
            nova_linha = {"PROTOCOLO": proto, "STATUS ATUAL": status, "MODIFICADO EM": AGORA}
            df = pd.concat([df, pd.DataFrame([nova_linha])], ignore_index=True)
            mudancas.append(f"NOVO PROCESSO: {proto} ({status})")

    df.to_csv("protocolos.csv", index=False)
    
    # 4. E-mail
    if mudancas and GMAIL:
        msg = MIMEMultipart()
        msg['Subject'] = f"🔄 Atualização: {len(mudancas)} alterações"
        msg.attach(MIMEText("\n".join(mudancas), 'plain'))
        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()
            server.login("mikaellevictoria2017@gmail.com", GMAIL)
            server.sendmail("mikaellevictoria2017@gmail.com", "santos.micaelle2006@gmail.com", msg.as_string())

if __name__ == "__main__":
    main()
