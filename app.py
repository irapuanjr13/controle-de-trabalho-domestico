from datetime import datetime, timedelta
from fpdf import FPDF
import calendar
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import schedule
import time
from dotenv import load_dotenv
import os
import logging

# Carregar variáveis de ambiente
load_dotenv()

# Configurações de email
EMAIL_REMETENTE = os.getenv("EMAIL_REMETENTE")
SENHA_EMAIL = os.getenv("SENHA_EMAIL")

# Configuração de log
logging.basicConfig(filename='errors.log', level=logging.ERROR)

# Caminho para salvar o PDF
PDF_PATH = "recibo_vale_transporte.pdf"

# Custo diário fixo do vale-transporte
CUSTO_DIARIO_TRANSPORTE = 21.0

# Função para calcular a vigência e gerar o recibo
def gerar_pdf_domestica(dias_trabalhados, folgas, custo_transporte, fechamento_data):
    try:
        # Calcula o mês seguinte ao fechamento
        vigencia_inicio = fechamento_data + timedelta(days=5)
        ano_vigencia = vigencia_inicio.year
        mes_vigencia = vigencia_inicio.month
        dias_mes = calendar.monthrange(ano_vigencia, mes_vigencia)[1]  # Total de dias no mês

        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=12)

        # Cabeçalho
        pdf.cell(200, 10, txt="RECIBO DE VALE-TRANSPORTE", ln=True, align='C')
        pdf.ln(10)  # Espaçamento

        # Texto do recibo
        texto = f"""
        Eu, AMANDA KAROLINE RAMOS SILVA,
        declaro para os devidos fins que recebi de ELLEN APARECIDA DE SOUZA DUARTE,
        a quantia de R$ {custo_transporte:.2f} (valor por extenso: __________________________),
        referente ao benefício do Vale-Transporte com vigência de {vigencia_inicio.strftime('%d/%m/%Y')}
        a {datetime(ano_vigencia, mes_vigencia, dias_mes).strftime('%d/%m/%Y')}.

        Dados do Benefício:
        - Dias trabalhados: {dias_trabalhados}
        - Dias de folga: {', '.join(map(str, folgas))}
        - Valor diário do transporte: R$ {CUSTO_DIARIO_TRANSPORTE:.2f}

        Comprometo-me a utilizar o valor recebido exclusivamente para deslocamento entre minha residência
        e o local de trabalho, conforme as normas acordadas.

        Assinatura do Recebedor: _________________________________________________
        Data: ____/____/____
        """

        # Adicionar texto ao PDF
        for linha in texto.split("\n"):
            pdf.cell(200, 10, txt=linha.strip(), ln=True)

        # Salvar o PDF
        pdf.output(PDF_PATH)

        return PDF_PATH, ano_vigencia, mes_vigencia
    except Exception as e:
        logging.error(f"Erro ao gerar PDF: {e}")
        raise

# Função para enviar email com o PDF
def enviar_email(destinatario, assunto, mensagem, anexo):
    try:
        # Configuração do email
        msg = MIMEMultipart()
        msg['From'] = EMAIL_REMETENTE
        msg['To'] = destinatario
        msg['Subject'] = assunto
        
        # Mensagem do email
        msg.attach(MIMEText(mensagem, 'plain'))
        
        # Anexo
        with open(anexo, "rb") as f:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f"attachment; filename={os.path.basename(anexo)}")
            msg.attach(part)
        
        # Configurações do servidor SMTP
        servidor = "smtp.office365.com"
        porta = 587

        with smtplib.SMTP(servidor, porta) as smtp:
            smtp.starttls()  # Inicializa a conexão segura
            smtp.login(EMAIL_REMETENTE, SENHA_EMAIL)  # Autentica
            smtp.send_message(msg)  # Envia o email
            print(f"Email enviado para {destinatario} com sucesso!")
    except Exception as e:
        logging.error(f"Erro ao enviar email: {e}")
        raise

# Agendamento para rodar mensalmente
def agendar_envio():
    try:
        hoje = datetime.now()
        fechamento_data = datetime(hoje.year, hoje.month, 26)
        
        # Gera recibo com base no fechamento do mês
        anexo, ano_vigencia, mes_vigencia = gerar_pdf_domestica(
            dias_trabalhados=20,
            folgas=[12, 27],
            custo_transporte=20 * CUSTO_DIARIO_TRANSPORTE,
            fechamento_data=fechamento_data
        )
        
        # Email e mensagem
        destinatario = "ellen.asduarte@yahoo.com.br"
        assunto = f"Recibo de Vale-Transporte - Vigência {calendar.month_name[mes_vigencia]} {ano_vigencia}"
        mensagem = f"Segue em anexo o recibo de vale-transporte com vigência {calendar.month_name[mes_vigencia]} de {ano_vigencia}."
        enviar_email(destinatario, assunto, mensagem, anexo)
    except Exception as e:
        logging.error(f"Erro ao agendar envio: {e}")
        raise

if __name__ == "__main__":
    # Envio de email imediato para teste
    agendar_envio()
    
    # Configuração do agendamento regular
    schedule.every().month.at("08:00").do(agendar_envio)
    
    print("Sistema de envio automático iniciado...")
    
    # Executa o agendamento
    while True:
        schedule.run_pending()
        time.sleep(1)
