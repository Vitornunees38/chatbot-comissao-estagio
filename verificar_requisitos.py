import tabula,numpy as np
import warnings
from datetime import datetime

from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Paragraph,Frame,Image,Spacer
from reportlab.lib import colors
import bot_cc

def extrair_dre_boa(file_path, dre):
    warnings.simplefilter(action='ignore', category=FutureWarning)

    # Read pdf into list of DataFrame
    dfs = tabula.read_pdf(file_path, pages='1,6', multiple_tables=True, stream=True)

    # A primeira verificação será do DRE
    dre_info = dfs[0].to_numpy().astype(str)[5][1]  # "DRE:" está em [5][1]
    dre_numero = dre_info.split("SITUAÇÃO ATUAL:")[0].strip()

    return dre_numero


def verifica_criterios(file_path, dre):
    warnings.simplefilter(action='ignore', category=FutureWarning)

    # Read pdf into list of DataFrame
    dfs = tabula.read_pdf(file_path, pages='1,6', multiple_tables=True, stream=True)

    materias_quarto = [
        "ICP131", "ICP132", "ICP133", "ICP134", "ICP135", "ICP136",
        "ICP141", "ICP142", "ICP143", "ICP144", "ICP145", "MAE111",
        "ICP115", "ICP116", "ICP237", "ICP238", "ICP239", "MAE992",
        "ICP246", "ICP248", "ICP249", "ICP489", "MAD243"
    ]
    
    apto = 0

    # Verifica se o BOA é do aluno mesmo verificando o DRE
    dre = str(dre)
    
    # Verifica se o aluno tem notas para todas as disciplinas até o quarto período com aproveitamento
    def verifica_disciplinas():
        for item in dfs[0].to_numpy():
            for disciplina in materias_quarto:
                if disciplina in item:
                    # Verifica se a matéria está em uma das situações que não permitem aproveitamento
                    if item[5] in ("Cursando", "Inscrição Vedada", "Inscrição Facultada"):
                        result = 0
                    else:
                        result = 1
        return result
    
    # Verifica CRA >= 6
    def verifica_CRA():
        array_str = dfs[1].to_numpy().astype(str)
        indice = np.where(np.char.find(array_str, "CR acumulado") != -1)
        
        CRA = array_str[indice][0]
        CRA = float(CRA.split(": ")[1])
        
        return 1 if CRA >= 6 else 0

    # Chamando as funções de verificação
    apto += verifica_disciplinas()
    apto += verifica_CRA()

    # Verifica se o aluno está apto a estagiar
    if apto == 2:
        print("Liberado para Estagiar")
        return "Liberado para Estagiar"  # Retorna sucesso para estagiar
    else:
        print("Não liberado para Estagiar")
        return "Não liberado para Estagiar"  # Retorna que não está liberado



def gerar_parecer_pdf_BOA(file_path, nome_arquivo, situacao):

    warnings.simplefilter(action='ignore', category=FutureWarning)

     # pdf to dataframe
    dfs = tabula.read_pdf(file_path, pages='1,6',multiple_tables=True,stream=True, encoding="iso-8859-1")

    # Obter a data atual no formato desejado
    data_atual = datetime.now().strftime("%d/%m/%Y")

    # Garantir que o array seja do tipo str
    array_str = dfs[0].to_numpy().astype(str)

     # Procurando o índice que contém "ALUNO:"
    aluno_info = array_str[4][1] #array_str[4][1] é o índice que contém "ALUNO:"
  
    aluno_nome = aluno_info.split("CURSO ATUAL:")[0].strip()

    # Procurando o índice que contém "DRE:"
    dre_info = array_str[5][1] #array_str[5][1] é o índice que contém "DRE:"
    dre_numero = dre_info.split("SITUAÇÃO ATUAL:")[0].strip()

    validade=bot_cc.calcular_validade()
    

    # Cria o documento
    doc = SimpleDocTemplate(nome_arquivo, pagesize=A4)
    largura, altura = A4  # Dimensões da página A4

    # Obtém o estilo padrão
    style = getSampleStyleSheet()['Normal']
    style.fontName = "Helvetica"
    style.fontSize = 12
    
    elements = []
        
    if situacao:
        imagem_path = "assets/Minerva_Oficial_UFRJ_(Orientação_Horizontal).png"  # Caminho da imagem
        imagem_width = 350  # Largura da imagem
        imagem_height = 140  # Altura da imagem

        # Adiciona a imagem no topo da página
        image = Image(imagem_path, imagem_width, imagem_height)
        image.hAlign = 'CENTER'  # Centraliza horizontalmente
        elements.append(image)
        elements.append(Spacer(1, 80))  # Adiciona espaço após a imagem

        # Adiciona título centralizado
        title_style = getSampleStyleSheet()['Title']
        title_style.fontName = "Helvetica-Bold"
        title_style.alignment = 1  # Centraliza o título
        elements.append(Paragraph("<b>Parecer de Autorização para Estágio</b>", title_style))
        elements.append(Spacer(1, 40))  # Espaço entre título e corpo do texto

        # Corpo do texto, com quebra de linha após o nome
        body_style = getSampleStyleSheet()['Normal']
        body_style.fontSize = 12
        body_style.alignment = 4  # Justifica o texto
        elements.append(Paragraph(f"O aluno <b>{aluno_nome}</b> de DRE <b>{dre_numero}</b> está autorizado a estagiar.", body_style))
        elements.append(Spacer(1, 10))  # Espaço entre parágrafos
        elements.append(Paragraph(f"Este documento tem validade de 3 meses contando a partir de {data_atual} até {validade}.", body_style))
    # Criar o documento com os elementos
    doc.build(elements)

