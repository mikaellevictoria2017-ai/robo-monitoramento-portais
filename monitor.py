import os
import smtplib
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd

# ==========================================
# ⚙️ CONFIGURAÇÕES GLOBAIS
# ==========================================
# Puxa a senha direto do cofre de segurança do GitHub Secrets
SENHA_REMETENTE = os.environ.get("ihfxftkgihyuniob")

# Ajuste com as suas informações reais de e-mail e links:
EMAIL_REMETENTE = "mikaellevictoria2017@gmail.com" 
EMAIL_DESTINATARIOS = ["santos.micaelle2006@gmail.com"]
LINK_PLANILHA = "https://artesanourbanismo-my.sharepoint.com/:x:/g/personal/mvitoria_artesanourbanismo_com_br/IQBI7DqiCzMtSIrLcPwTnM2SAamaiye_3EPs-HAEKli1mZo?e=Zjekvn"

nome_planilha = "monitor_protocolos.xlsx"
nome_aba = "Santana de Parnaíba"

# ==========================================
# ✉️ FUNÇÃO DE ENVIO DE E-MAIL
# ==========================================
def enviar_email_alerta(processos_alterados):
    try:
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
# 🤖 EXECUÇÃO DO MONITORAMENTO
# ==========================================
def executar_monitoramento(dados_portal):
    agora_str = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    print(f"\n===== INICIANDO VERIFICAÇÃO: {agora_str} =====")

    try:
        df = pd.read_excel(nome_planilha, sheet_name=nome_aba)
        print(f"📊 Aba '{nome_aba}' carregada com sucesso! Encontradas {len(df)} linhas.")
    except Exception as e:
        print(f"❌ Erro ao ler a planilha: {e}")
        return

    processos_alterados = []
    houve_alteracao = False

    col_status = [c for c in df.columns if "STATUS" in c][0]
    col_modificado = [c for c in df.columns if "MODIFICADO" in c][0] if any("MODIFICADO" in c for c in df.columns) else "MODIFICADO"

    print("🔎 Comparando dados da planilha com o portal...")
    for index, text_linha in df.iterrows():
        if str(text_linha["ATIVO"]).strip().upper() == "SIM":
            protocolo = str(text_linha["PROTOCOLO"]).strip().upper()
            status_antigo = str(text_linha[col_status]).strip()

            protocolo_planilha_limpo = protocolo.replace(" ", "")

            linha_encontrada = None
            for dado in dados_portal:
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
        print("🎉 Planilha updated com sucesso!")
        
        enviar_email_alerta(processos_alterados)
    else:
        print("☕ Nenhuma alteração encontrada. Tudo atualizado!")

# Simulador para rodar o script principal (substitua pela chamada real do seu portal se necessário)
if __name__ == "__main__":
    # Aqui entra a lista de dados capturados do seu portal (exemplo estrutural)
    # dados_do_portal_capturados = capturar_dados_do_portal() 
    # executar_monitoramento(dados_do_portal_capturados)
    pass
