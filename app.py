from fpdf import FPDF
from datetime import datetime, timedelta
import calendar
import holidays
from dotenv import load_dotenv
from num2words import num2words
import schedule
import time
import os
import io
from io import BytesIO
import logging
from email import encoders

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

# Configurações gerais
FERIADOS = holidays.Brazil()  # Lista de feriados no Brasil
VALOR_PASSAGEM = 5.25         # Valor unitário da passagem
CONDUCOES_POR_DIA = 4         # Número de conduções diárias (ida e volta)


def calcular_dias_uteis(ano, mes):
    """
    Calcula os dias úteis no mês, excluindo finais de semana e feriados.
    """
    _, total_dias_mes = calendar.monthrange(ano, mes)
    inicio_mes = datetime(ano, mes, 1)

    # Lista de todos os dias do mês
    dias_do_mes = [inicio_mes + timedelta(days=i) for i in range(total_dias_mes)]

    # Filtra os dias úteis (segunda a sexta, excluindo feriados)
    dias_uteis = [dia for dia in dias_do_mes if dia.weekday() < 5 and dia not in FERIADOS]
    return dias_uteis


def calcular_passagens_e_custo(ano, mes, folgas):
    """
    Calcula a quantidade de passagens e o custo total para o mês.
    """
    dias_uteis = calcular_dias_uteis(ano, mes)

    # Remove os dias de folga
    dias_trabalhados = len(dias_uteis) - len(folgas)

    # Calcula a quantidade de passagens e o custo total
    passagens_total = dias_trabalhados * CONDUCOES_POR_DIA
    custo_total = passagens_total * VALOR_PASSAGEM
    return passagens_total, custo_total, dias_trabalhados


def ajustar_folgas_para_sexta(folgas):
    """
    Ajusta as folgas para a sexta-feira mais próxima.
    """
    folgas_ajustadas = []
    for folga in folgas:
        if folga.weekday() != 4:  # Se não for sexta-feira
            diferenca = 4 - folga.weekday() if folga.weekday() < 4 else -1
            folgas_ajustadas.append(folga + timedelta(days=diferenca))
        else:
            folgas_ajustadas.append(folga)
    return folgas_ajustadas


def converter_valor_por_extenso(valor):
    """
    Converte um valor monetário em reais para extenso.
    """
    reais, centavos = divmod(int(valor * 100), 100)
    reais_extenso = f"{num2words(reais, lang='pt_BR').capitalize()} reais"
    centavos_extenso = f"{num2words(centavos, lang='pt_BR')} centavos" if centavos > 0 else ""
    return f"{reais_extenso} e {centavos_extenso}".strip()


def gerar_recibo(passagens_total, custo_total, mes, ano):
    """
    Gera um recibo no formato solicitado.
    """
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)

    # Cabeçalho
    pdf.set_font("Arial", size=16, style="B")
    pdf.cell(0, 10, txt="RECIBO", ln=True, align="C")
    pdf.cell(0, 10, txt="Entrega Vale-Transporte", ln=True, align="C")
    pdf.ln(20)

    # Corpo do recibo
    pdf.set_font("Arial", size=12)
    pdf.cell(0, 10, txt="Empregador(a): Ellen Aparecida de Souza Duarte", ln=True)
    pdf.cell(0, 10, txt="Empregado(a): Amanda Karoline Ramos da Silva", ln=True)
    pdf.ln(10)

    valor_extenso = converter_valor_por_extenso(custo_total)

    pdf.multi_cell(0, 10, txt=(
        f"Recebi {passagens_total} vales-transporte, referentes ao mês de "
        f"{calendar.month_name[mes]} de {ano}, pelo que firmo o presente."
    ))
    pdf.ln(10)

    pdf.multi_cell(0, 10, txt=f"O valor total é de R$ {custo_total:.2f} ({valor_extenso}).")
    pdf.ln(20)

    # Local e data
    data_atual = datetime.now().strftime("%d/%m/%Y")
    pdf.cell(0, 10, txt=f"Belo Horizonte,", ln=True)
    pdf.cell(0, 10, txt=f"Data: {data_atual}", ln=True)
    pdf.ln(20)

    # Observações
    pdf.cell(0, 10, txt="Obs:", ln=True)
    pdf.ln(20)

    # Assinatura
    pdf.cell(0, 10, txt="________________________________________", ln=True, align="C")
    pdf.cell(0, 10, txt="Amanda Karoline Ramos da Silva", ln=True, align="C")

    # Salvar o recibo
    nome_pdf = "recibo_vale_transporte.pdf"
    pdf.output(nome_pdf)
    return nome_pdf


def verificar_data():
    """
    Verifica se é dia 26 e executa a função se for o caso.
    """
    hoje = datetime.now()
    if hoje.day == 26:  # Verifica se é dia 26
        executar_geracao_recibo()
    else:
        print(f"Hoje não é dia 26. Data atual: {hoje.strftime('%d/%m/%Y')}")


def executar_geracao_recibo():
    """
   Calcula as passagens e gera o recibo para o mês posterior ao agendamento.
    """
    hoje = datetime.now()
    proximo_mes = (hoje.month % 12) + 1
    ano_proximo_mes = hoje.year + (1 if proximo_mes == 1 else 0)

    # Determina as folgas (10º dia útil ajustado para sexta-feira)
    dias_uteis = calcular_dias_uteis(ano_proximo_mes, proximo_mes)
    folgas = ajustar_folgas_para_sexta(dias_uteis[9::10])

    # Calcula as passagens e o custo total
    passagens_total, custo_total, dias_trabalhados = calcular_passagens_e_custo(ano_proximo_mes, proximo_mes, folgas)

    print(f"Dias úteis trabalhados: {dias_trabalhados}")
    print(f"Total de passagens: {passagens_total}")
    print(f"Custo total: R$ {custo_total:.2f}")

    # Gera o recibo em PDF
    recibo_pdf = gerar_recibo(passagens_total, custo_total, proximo_mes, ano_proximo_mes)
    print(f"Recibo gerado: {recibo_pdf}")

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
    """
    Agenda a execução para o dia 26 de cada mês.
    """
     # Configura o agendamento para verificar diariamente às 08:00
    schedule.every().day.at("08:00").do(verificar_data)

    print("Agendamento configurado. O script está em execução...")

    while True:
        schedule.run_pending()
        time.sleep(60)


if __name__ == "__main__":
    agendar_envio()
