import tabula
import numpy as np
import warnings
from datetime import datetime
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Frame, Image, Spacer
)
from reportlab.lib import colors
import bot_cc_refatorado as bot_cc

# Função para extrair DRE do arquivo PDF
def extrair_dre_boa(file_path, dre):
    warnings.simplefilter(action='ignore', category=FutureWarning)

    # Lê o PDF e converte em uma lista de DataFrames
    dfs = tabula.read_pdf(file_path, pages='1,6', multiple_tables=True, stream=True, encoding='ISO-8859-1')

    # Extrai o DRE do DataFrame
    dre_info = dfs[0].to_numpy().astype(str)[5][1]  # "DRE:" está em [5][1]
    dre_numero = dre_info.split("SITUAÇÃO ATUAL:")[0].strip()

    return dre_numero

# Função para verificar critérios de elegibilidade
def verifica_criterios(file_path, dre):
    warnings.simplefilter(action='ignore', category=FutureWarning)

    # Lê o PDF e converte em uma lista de DataFrames
    dfs = tabula.read_pdf(file_path, pages='1,6', multiple_tables=True, stream=True, encoding='ISO-8859-1')

    materias_quarto = [
        "ICP131", "ICP132", "ICP133", "ICP134", "ICP135", "ICP136",
        "ICP141", "ICP142", "ICP143", "ICP144", "ICP145", "MAE111",
        "ICP115", "ICP116", "ICP237", "ICP238", "ICP239", "MAE992",
        "ICP246", "ICP248", "ICP249", "ICP489", "MAD243"
    ]

    def verifica_disciplinas():
        """Verifica se o aluno está apto em todas as disciplinas até o 4º período."""
        result = 1
        for item in dfs[0].to_numpy():
            for disciplina in materias_quarto:
                if disciplina in item:
                    if item[5] in ("Cursando", "Inscrição Vedada", "Inscrição Facultada"):
                        return 0
        return result

    def verifica_cra():
        """Verifica se o CRA é maior ou igual a 6."""
        array_str = dfs[1].to_numpy().astype(str)
        indice = np.where(np.char.find(array_str, "CR acumulado") != -1)

        cra_info = array_str[indice][0]
        cra_valor = float(cra_info.split(": ")[1])

        return 1 if cra_valor >= 6 else 0

    apto = verifica_disciplinas() + verifica_cra()

    if apto == 2:
        print("Liberado para Estagiar")
        return True
    else:
        print("Não liberado para Estagiar")
        return False

# Função para gerar parecer em PDF
def gerar_parecer_pdf_boa(file_path, nome_arquivo, situacao):
    warnings.simplefilter(action='ignore', category=FutureWarning)

    # Lê o PDF e converte em uma lista de DataFrames
    dfs = tabula.read_pdf(file_path, pages='1,6', multiple_tables=True, stream=True, encoding="iso-8859-1")

    data_atual = datetime.now().strftime("%d/%m/%Y")
    validade = bot_cc.calcular_validade()

    array_str = dfs[0].to_numpy().astype(str)
    aluno_nome = array_str[4][1].split("CURSO ATUAL:")[0].strip()
    dre_numero = array_str[5][1].split("SITUAÇÃO ATUAL:")[0].strip()

    doc = SimpleDocTemplate(nome_arquivo, pagesize=A4)
    elements = []

    if situacao:
        imagem_path = "assets/Minerva_Oficial_UFRJ_(Orientação_Horizontal).png"
        image = Image(imagem_path, 350, 140)
        image.hAlign = 'CENTER'
        elements.append(image)
        elements.append(Spacer(1, 80))

        title_style = getSampleStyleSheet()['Title']
        title_style.fontName = "Helvetica-Bold"
        title_style.alignment = 1
        elements.append(Paragraph("<b>Parecer de Autorização para Estágio</b>", title_style))
        elements.append(Spacer(1, 40))

        body_style = getSampleStyleSheet()['Normal']
        body_style.fontSize = 12
        body_style.alignment = 4
        elements.append(Paragraph(f"O aluno <b>{aluno_nome}</b> de DRE <b>{dre_numero}</b> está autorizado a estagiar.", body_style))
        elements.append(Spacer(1, 10))
        elements.append(Paragraph(f"Este documento tem validade de 3 meses contando a partir de {data_atual} até {validade}.", body_style))

    doc.build(elements)

# Função para gerar parecer em PDF
def reemitir_parecer_pdf_boa(file_path, nome_arquivo, situacao,validade):
    warnings.simplefilter(action='ignore', category=FutureWarning)

    # Lê o PDF e converte em uma lista de DataFrames
    dfs = tabula.read_pdf(file_path, pages='1,6', multiple_tables=True, stream=True, encoding="iso-8859-1")

    data_atual = datetime.now().strftime("%d/%m/%Y")
    #validade = bot_cc.calcular_validade()

    array_str = dfs[0].to_numpy().astype(str)
    aluno_nome = array_str[4][1].split("CURSO ATUAL:")[0].strip()
    dre_numero = array_str[5][1].split("SITUAÇÃO ATUAL:")[0].strip()

    doc = SimpleDocTemplate(nome_arquivo, pagesize=A4)
    elements = []

    if situacao:
        imagem_path = "assets/Minerva_Oficial_UFRJ_(Orientação_Horizontal).png"
        image = Image(imagem_path, 350, 140)
        image.hAlign = 'CENTER'
        elements.append(image)
        elements.append(Spacer(1, 80))

        title_style = getSampleStyleSheet()['Title']
        title_style.fontName = "Helvetica-Bold"
        title_style.alignment = 1
        elements.append(Paragraph("<b>Parecer de Autorização para Estágio</b>", title_style))
        elements.append(Spacer(1, 40))

        body_style = getSampleStyleSheet()['Normal']
        body_style.fontSize = 12
        body_style.alignment = 4
        elements.append(Paragraph(f"O aluno <b>{aluno_nome}</b> de DRE <b>{dre_numero}</b> está autorizado a estagiar.", body_style))
        elements.append(Spacer(1, 10))
        elements.append(Paragraph(f"Este documento tem validade de 3 meses contando a partir de {data_atual} até {validade}.", body_style))

    doc.build(elements)
