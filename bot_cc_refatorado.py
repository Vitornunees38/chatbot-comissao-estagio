import os
import logging
from datetime import datetime
from functools import wraps
from dateutil.relativedelta import relativedelta
from telegram import Update
from telegram.ext import (
    Application, CommandHandler, MessageHandler, filters, ConversationHandler, CallbackContext
)
import gspread
from apscheduler.schedulers.asyncio import AsyncIOScheduler
from oauth2client.service_account import ServiceAccountCredentials
import verificar_requisitos_refatorado as verificar_requisitos
import yagmail


#CHAVES




# Configurações e constantes
SPREADSHEET_SCOPE = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive"
]
LOG_FORMAT = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
LOG_LEVEL = logging.INFO

# Estados da conversa
DRE, AVALIACAO_CONTRATO, WAITING_FOR_BOA, EMAIL,WAITING_FOR_BOA_REEMISSAO ,WAITING_FOR_CONTRATO= range(6)

# Variáveis globais



# Inicialização
logging.basicConfig(format=LOG_FORMAT, level=LOG_LEVEL)
credentials = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', SPREADSHEET_SCOPE)
client = gspread.authorize(credentials)

# Funções auxiliares
def calcular_validade():
    """Calcula a validade do parecer (data atual + 3 meses)."""
    return (datetime.now() + relativedelta(months=3)).strftime("%d/%m/%Y")

def limpar_arquivo(caminho):
    """Remove um arquivo temporário, se existente."""
    try:
        os.remove(caminho)
    except Exception as e:
        logging.error(f"Erro ao tentar apagar o arquivo {caminho}: {e}")


async def timeout_handler(context):
    # Ação após o tempo limite (notificação)
    chat_id=context.job.data
    await context.bot.send_message(chat_id=chat_id, text="Sua conversa expirou! Digite /start para reiniciar.")
    return ConversationHandler.END  # Encerra a conversa após o timeout

async def start(update: Update, context: CallbackContext):
    """Inicia a conversa solicitando o DRE."""
    context.job_queue.run_once(timeout_handler, 60, data=update.message.chat_id)
    await update.message.reply_text(
        'Olá! Bem-vindo ao bot da comissão de estágio de Ciência da Computação.\n'
        'Por favor, informe seu DRE (número de matrícula).'
    )
    return DRE


async def coletar_dre(update: Update, context: CallbackContext):
    """Valida o DRE informado pelo usuário e verifica sua existência na planilha."""
    dre = update.message.text.strip()
    context.user_data["dre"] = dre


    if len(dre) == 9 and dre.isdigit() and dre[0] != "0":
        try:
            sheet = client.open_by_key(SPREADSHEET_ID).sheet1
            dre_column = sheet.col_values(1)


            if dre in dre_column:
                await update.message.reply_text(
                    "Aluno já está apto a procurar estágio. O que você deseja fazer?\n"
                    "1. Emitir parecer de que está apto a estagiar (/EmitirParecer).\n"
                    "2. Fazer avaliação de contrato (/Contrato).\n"
                    "3. Encerrar o bot (/Encerrar)."
                    
                )

                return AVALIACAO_CONTRATO
            else:
                await update.message.reply_text("DRE não encontrado. Por favor, envie o seu BOA.")
                return WAITING_FOR_BOA

        except Exception as e:
            logging.error(f"Erro ao acessar a planilha: {e}")
            await update.message.reply_text("Erro ao verificar o DRE. Tente novamente mais tarde.")
            return ConversationHandler.END
    else:
        await update.message.reply_text("Por favor, insira um DRE válido com 9 dígitos numéricos.")
        

    return DRE


async def verificar_boa(update: Update, context: CallbackContext):
    """Recebe o BOA e verifica o DRE com os requisitos."""


    if update.message.document:
        document = update.message.document
        download_dir = os.path.join(os.getcwd(), "downloads")
        os.makedirs(download_dir, exist_ok=True)
        file_path = os.path.join(download_dir, document.file_name)

        new_file = await document.get_file()
        await new_file.download_to_drive(file_path)

        context.user_data["file_path"] = file_path

        dre = context.user_data.get("dre")
        if not dre:
            await update.message.reply_text("DRE não encontrado. Por favor, envie seu DRE antes do BOA.")
            return WAITING_FOR_BOA

        dre_extraido = verificar_requisitos.extrair_dre_boa(file_path, dre)

        if dre_extraido == dre:
            situacao = verificar_requisitos.verifica_criterios(file_path, dre)
            context.user_data["situacao"] = situacao

            if situacao:
                await update.message.reply_text("Liberado para estagiar. Por favor, insira seu e-mail.")
                return EMAIL
            else:
                await update.message.reply_text(
                    "Status: Não autorizado para estagiar.\n"
                    "Caso deseje recorrer, preencha o formulário: https://forms.gle/NBuhtQ9HhHN1W11c6"
                )
                return ConversationHandler.END

        await update.message.reply_text("DRE incompatível com o BOA. Por favor, reenvie seu DRE.")
        return DRE

    await update.message.reply_text("Por favor, envie um arquivo PDF válido.")
    return WAITING_FOR_BOA


async def emitir_parecer(update: Update, context: CallbackContext):
    """Emite o parecer de aptidão para estágio, recebendo um BOA em PDF."""

    dre = context.user_data.get("dre")

    if not dre:
        await update.message.reply_text("DRE não encontrado no contexto. Por favor, inicie novamente com /start.")
        return ConversationHandler.END

    try:
        # Verifica se o usuário enviou um documento
        if not update.message.document:
            await update.message.reply_text("Por favor, envie o BOA como um arquivo PDF para continuar.")
            return WAITING_FOR_BOA_REEMISSAO  # Define um estado para esperar o BOA

        # Processa o documento enviado (BOA)
        document = update.message.document

        # Verifica se é um PDF
        if not document.file_name.endswith(".pdf"):
            await update.message.reply_text("O arquivo enviado não é um PDF. Por favor, envie um arquivo válido.")
            return WAITING_FOR_BOA_REEMISSAO

        # Define o diretório e o caminho do arquivo
        download_dir = os.path.join(os.getcwd(), "downloads")
        os.makedirs(download_dir, exist_ok=True)
        file_path = os.path.join(download_dir, "boa_simulado.pdf")  # Salva como boa_simulado.pdf

        # Faz o download do arquivo
        new_file = await document.get_file()
        await new_file.download_to_drive(file_path)

        # Armazena o caminho do arquivo no contexto
        context.user_data["boa_path"] = file_path

        # Abre planilha
        sheet = client.open_by_key(SPREADSHEET_ID).sheet1

        # Localiza a célula com o valor de busca na primeira coluna
        celula = sheet.find(dre, in_column=1)

        # Pega o valor na mesma linha, na coluna desejada
        validade = sheet.cell(celula.row, 3).value

        # Gera o parecer em PDF após registrar na planilha
        nome_pdf = "parecer_aluno.pdf"
        verificar_requisitos.reemitir_parecer_pdf_boa(file_path, nome_pdf, True,validade)  # Aqui situacao já deve ser booleano

        # Envia o parecer como documento para o usuário
        with open(nome_pdf, "rb") as pdf_file:
            await update.message.reply_document(document=pdf_file, filename=nome_pdf)

        await update.message.reply_text(
            f"Parecer emitido com sucesso!\n\n"
            f"Aluno com DRE {dre} está apto a estagiar.\n"
            f"Validade: {validade}\n"
            f"Se precisar de mais algo, digite /start para reiniciar."
        )

    except Exception as e:
        logging.error(f"Erro ao emitir o parecer: {e}")
        await update.message.reply_text("Erro ao emitir o parecer. Tente novamente mais tarde.")

    return ConversationHandler.END



async def receber_email(update: Update, context: CallbackContext):


    email = update.message.text  # Captura o e-mail enviado pelo usuário
    context.user_data['email'] = email  # Armazenar no user_data para uso posterior


    """Recebe e valida o e-mail, atualizando a planilha com os dados."""
    email = update.message.text.strip()
    if "@" not in email or "." not in email.split("@")[-1]:
        await update.message.reply_text("E-mail inválido. Por favor, envie um e-mail válido.")
        return EMAIL

    
    dre = context.user_data.get("dre")
    situacao = context.user_data.get("situacao")
    file_path = context.user_data.get("file_path")

    if dre and situacao is not None:
        try:
            sheet = client.open_by_key(SPREADSHEET_ID).sheet1
            row = len(sheet.col_values(1)) + 1

            validade = calcular_validade()
            sheet.update_cell(row, 1, dre)
            sheet.update_cell(row, 2, email)
            sheet.update_cell(row, 3, validade)

            await update.message.reply_text("Dados registrados com sucesso!")

            # Gera o parecer em PDF após registrar na planilha
            nome_pdf = "parecer_aluno.pdf"

            # Verifica diretamente a situação, sem usar .lower() se já for um booleano
            verificar_requisitos.gerar_parecer_pdf_boa(file_path,nome_pdf, situacao)  # Aqui situacao já deve ser booleano
            
            with open(nome_pdf, "rb") as pdf_file:
                await update.message.reply_document(document=pdf_file, filename=nome_pdf)

        except Exception as e:
            logging.error(f"Erro ao atualizar planilha: {e}")

    return ConversationHandler.END


async def contrato(update: Update, context: CallbackContext):

    dre = context.user_data.get("dre")
      
    if not dre:
        await update.message.reply_text("DRE não encontrado no contexto. Por favor, inicie novamente com /start.")
        return ConversationHandler.END


    #pegar email do aluno na planilha

    try:
        # Abre a planilha
        sheet = client.open_by_key(SPREADSHEET_ID).sheet1

        # Localiza a célula com o valor de busca na primeira coluna
        celula = sheet.find(dre, in_column=1)

        # Pega o valor na mesma linha, na coluna desejada
        email_aluno = sheet.cell(celula.row, 2).value
       

    except Exception as e:
        logging.error(f"Erro ao acessar a planilha: {e}")
        await update.message.reply_text("Erro ao verificar o email. Tente novamente mais tarde.")
        return WAITING_FOR_CONTRATO


    try:
        # Verifica se o usuário enviou um documento
        if not update.message.document:
            await update.message.reply_text("Por favor, envie o CONTRATO como um arquivo PDF para continuar.")
            return WAITING_FOR_CONTRATO  

        # Processa o documento enviado (contrato)
        document = update.message.document

        # Verifica se é um PDF
        if not document.file_name.endswith(".pdf"):
            await update.message.reply_text("O arquivo enviado não é um PDF. Por favor, envie um arquivo válido.")
            
            return WAITING_FOR_CONTRATO
        

        #baixa o arquivo 
        download_dir = os.path.join(os.getcwd(), "downloads")
        os.makedirs(download_dir, exist_ok=True)
        file_path = os.path.join(download_dir, document.file_name)

        new_file = await document.get_file()
        await new_file.download_to_drive(file_path)

        #context.user_data["file_path"] = file_path

        #enviar email

        yag = yagmail.SMTP("c2botcoaa@gmail.com",oauth2_file="email_sender_credencials.json")
        yag.send(
        to="c2botcoaa@gmail.com",
        subject=f"Contrato de Estagio - {dre} ",
        contents=f''' 
        Olá, segue em anexo o contrato de estágio do aluno de DRE: {dre} para análise.
        
        Enviar resposta para o email: {email_aluno}

        atenciosamente,
        C2BOT
        
        ''', 
        attachments=file_path,)


        # Remove o arquivo local após o envio
        os.remove(file_path)

        await update.message.reply_text("Envio de contrato realizado com sucesso! Aguarde a resposta por email.") 

        
    except Exception as e:
        logging.error(f"Erro ao enviar contrato {e}")
        await update.message.reply_text("Erro ao enviar contrato. Tente novamente mais tarde.")


    return ConversationHandler.END


async def encerrar(update: Update, context: CallbackContext):

    await update.message.reply_text(
            "Obrigado por usar o bot da Comissão de Estágio! Estou encerrando agora. Para falar comigo novamente, use o comando /start."
        )
    
    # Cancelar o job de timeout se ele estiver agendado
    job_timeout = context.job_queue.get_jobs_by_name("timeout_handler")  # Checar jobs agendados
    if job_timeout:
        for job in job_timeout:
            job.remove()  # Remover job de timeout

    # Garantir que o timeout não seja agendado após o comando de encerramento
    context.user_data['timeout_active'] = False  # Flag para indicar que o timeout não deve ser executado


    return ConversationHandler.END


# Função principal
def main():
    application = Application.builder().token(TELEGRAM_TOKEN).build()

    conv_handler = ConversationHandler(
        entry_points=[CommandHandler('start', start)],
        states={
            DRE: [MessageHandler(filters.TEXT & ~filters.COMMAND, coletar_dre)],
            AVALIACAO_CONTRATO: [
                CommandHandler('EmitirParecer', emitir_parecer),
                CommandHandler('Contrato', contrato),
                CommandHandler('Encerrar', encerrar)
            ],
            WAITING_FOR_CONTRATO: [MessageHandler(filters.Document.ALL & ~filters.TEXT, contrato)],
            WAITING_FOR_BOA: [MessageHandler(filters.Document.ALL, verificar_boa)],
            WAITING_FOR_BOA_REEMISSAO: [MessageHandler(filters.Document.ALL, emitir_parecer)],
            EMAIL: [MessageHandler(filters.TEXT & ~filters.COMMAND, receber_email)],
        },
        fallbacks=[MessageHandler(filters.TEXT & ~filters.COMMAND, encerrar)],
        conversation_timeout=60,
    )
    
    application.add_handler(conv_handler)
    application.run_polling()

if __name__ == '__main__':
    main()