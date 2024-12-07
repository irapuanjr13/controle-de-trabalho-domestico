from datetime import datetime, timedelta
from fpdf import FPDF
import calendar
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import time
from dotenv import load_dotenv
import os
import logging

# Carregar variáveis de ambiente
load_dotenv()

# Configurações de e-mail
EMAIL_REMETENTE = os.getenv("EMAIL_REMETENTE")  # Nome da variável no arquivo .env
SENHA_EMAIL = os.getenv("SENHA_EMAIL")      # Nome da variável no arquivo .env

# Verificar se as variáveis foram carregadas corretamente
if not EMAIL_REMETENTE or not SENHA_EMAIL:
    raise ValueError("As variáveis de ambiente EMAIL_REMETENTE ou SENHA_EMAIL não foram configuradas corretamente.")

# Configuração de log
logging.basicConfig(filename='errors.log', level=logging.ERROR)

# Caminho para salvar o PDF
PDF_PATH = "recibo_vale_transporte.pdf"

# Custo diário fixo do vale-transporte
CUSTO_DIARIO_TRANSPORTE = 21.0

def gerar_pdf_domestica(dias_trabalhados, folgas, custo_transporte, fechamento_data):
    try:
        vigencia_inicio = fechamento_data + timedelta(days=5)
        ano_vigencia = vigencia_inicio.year
        mes_vigencia = vigencia_inicio.month
        dias_mes = calendar.monthrange(ano_vigencia, mes_vigencia)[1]

        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=12)

        pdf.cell(200, 10, txt="RECIBO DE VALE-TRANSPORTE", ln=True, align='C')
        pdf.ln(10)

        texto = f"""
        Eu, AMANDA KAROLINE RAMOS SILVA,
        declaro que recebi de ELLEN APARECIDA DE SOUZA DUARTE,
        a quantia de R$ {custo_transporte:.2f} (valor por extenso: ___________________),
        referente ao benefício do Vale-Transporte com vigência de {vigencia_inicio.strftime('%d/%m/%Y')}
        a {datetime(ano_vigencia, mes_vigencia, dias_mes).strftime('%d/%m/%Y')}.

        Dados do Benefício:
        - Dias trabalhados: {dias_trabalhados}
        - Dias de folga: {', '.join(map(str, folgas))}
        - Valor diário do transporte: R$ {CUSTO_DIARIO_TRANSPORTE:.2f}
        """

        for linha in texto.split("\n"):
            pdf.cell(200, 10, txt=linha.strip(), ln=True)

        pdf.output(PDF_PATH)
        return PDF_PATH, ano_vigencia, mes_vigencia
    except Exception as e:
        logging.error(f"Erro ao gerar PDF: {e}")
        raise

def enviar_email_com_retentativas(destinatarios, assunto, mensagem, anexo, max_tentativas=3, intervalo_retentativas=5):
    for tentativa in range(1, max_tentativas + 1):
        try:
            for destinatario in destinatarios:
                msg = MIMEMultipart()
                msg['From'] = EMAIL_REMETENTE
                msg['To'] = destinatario
                msg['Subject'] = assunto
                msg.attach(MIMEText(mensagem, 'plain'))

                with open(anexo, "rb") as f:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(f.read())
                    encoders.encode_base64(part)
                    part.add_header('Content-Disposition', f"attachment; filename={os.path.basename(anexo)}")
                    msg.attach(part)

                with smtplib.SMTP("smtp.mail.yahoo.com", 587) as smtp:
                    smtp.starttls()
                    smtp.login(EMAIL_REMETENTE, SENHA_EMAIL)
                    smtp.send_message(msg)
                    print(f"Email enviado para {destinatario} com sucesso!")
            return  # Se tudo der certo, sai da função
        except Exception as e:
            logging.error(f"Tentativa {tentativa} falhou: {e}")
            if tentativa < max_tentativas:
                print(f"Tentativa {tentativa} falhou. Tentando novamente em {intervalo_retentativas} segundos...")
                time.sleep(intervalo_retentativas)
            else:
                print(f"Todas as {max_tentativas} tentativas falharam.")
                raise

def agendar_envio():
    hoje = datetime.now()
    fechamento_data = datetime(hoje.year, hoje.month, 26)

    try:
        anexo, ano_vigencia, mes_vigencia = gerar_pdf_domestica(
            dias_trabalhados=20,
            folgas=[12, 27],
            custo_transporte=20 * CUSTO_DIARIO_TRANSPORTE,
            fechamento_data=fechamento_data
        )

        destinatario = ["ellen.asduarte@yahoo.com.br", "irapuanjunior13@gmail.com"]
        assunto = f"Recibo de Vale-Transporte - Vigência {calendar.month_name[mes_vigencia]} {ano_vigencia}"
        mensagem = f"Segue o recibo de vale-transporte com vigência {calendar.month_name[mes_vigencia]} de {ano_vigencia}."

        nviar_email_com_retentativas(destinatarios, assunto, mensagem, anexo)
    except Exception as e:
        logging.error(f"Erro no agendamento: {e}")
        raise

if __name__ == "__main__":
    try:
        print("Enviando e-mail de teste...")
        agendar_envio()
    except Exception as e:
        print(f"Erro no envio de e-mail: {e}")
