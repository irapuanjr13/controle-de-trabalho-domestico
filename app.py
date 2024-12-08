from fpdf import FPDF
from datetime import datetime, timedelta
import calendar
import holidays

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

    pdf.multi_cell(0, 10, txt=(
        f"Recebi {passagens_total} vales-transporte, referentes ao mês de "
        f"{calendar.month_name[mes]} de {ano}, pelo que firmo o presente."
    ))
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


def executar_geracao_recibo():
    """
    Calcula as passagens e gera o recibo.
    """
    hoje = datetime.now()
    mes_atual = hoje.month
    ano_atual = hoje.year

    # Determina as folgas (10º dia útil ajustado para sexta-feira)
    dias_uteis = calcular_dias_uteis(ano_atual, mes_atual)
    folgas = ajustar_folgas_para_sexta(dias_uteis[9::10])

    # Calcula as passagens e o custo total
    passagens_total, custo_total, dias_trabalhados = calcular_passagens_e_custo(ano_atual, mes_atual, folgas)

    print(f"Dias úteis trabalhados: {dias_trabalhados}")
    print(f"Total de passagens: {passagens_total}")
    print(f"Custo total: R$ {custo_total:.2f}")

    # Gera o recibo em PDF
    recibo_pdf = gerar_recibo(passagens_total, custo_total, mes_atual, ano_atual)
    print(f"Recibo gerado: {recibo_pdf}")


if __name__ == "__main__":
    executar_geracao_recibo()
