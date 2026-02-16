import os
import logging
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ConversationHandler
from telegram import Update
from telegram.ext import CallbackContext
import gspread
import verificar_requisitos
import asyncio
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
from dateutil.relativedelta import relativedelta


#chaves



# Definir o escopo de permissões para o Google Sheets
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]


# Use suas credenciais para autenticar
creds = ServiceAccountCredentials.from_json_keyfile_name(r"C:\Users\phf33\ufrj\botufrj\chatbot_comissao_estagio\credentials.json", scope)
client = gspread.authorize(creds)

# Configuração do gspread com API Key
gc = gspread.api_key(GOOGLE_API_KEY)

# Configuração de logging
logging.basicConfig(format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', level=logging.INFO)

# Definir etapas da conversa
DRE, AVALIACAO_CONTRATO, WAITING_FOR_BOA, EMAIL = range(4)

async def coletar_dre(update: Update, context: CallbackContext):
    dre = update.message.text
    context.user_data["dre"] = dre  # Armazena o DRE no contexto para uso posterior

    # Verificar se o DRE tem exatamente 9 números
    if len(dre) == 9 and dre.isdigit() and dre[0]!= "0":
        try:
            sheet = client.open_by_key(SPREADSHEET_ID).sheet1
            dre_column = sheet.col_values(1)  # Supondo que a coluna DRE seja a primeira (A)

            if dre in dre_column:
                # DRE encontrado
                await update.message.reply_text(
                    "Aluno já está apto a procurar estágio. Deseja fazer avaliação de contrato?\n"
                    "Digite /Contrato caso tenha conseguido estágio e queira tratar do contrato ou /Encerrar para encerrar o bot."
                )
                return AVALIACAO_CONTRATO

            else:
                # DRE não encontrado
                await update.message.reply_text("DRE não encontrado. Por favor, envie o seu BOA.")

                # Mensagem de encerramento após 1 minuto
                asyncio.create_task(timeout(update, context))

                return WAITING_FOR_BOA  # Agora segue para o próximo estado

        except Exception as e:
            await update.message.reply_text(f"Erro ao verificar o DRE: {str(e)}")  # Mensagem de erro
            return ConversationHandler.END
    else:
        await update.message.reply_text("Por favor, insira um DRE válido com 9 dígitos numéricos.")
        asyncio.create_task(timeout(update, context))
        return DRE

# Função para coletar o e-mail
async def coletar_email(update: Update, context: CallbackContext):
    await update.message.reply_text("Por favor, insira seu e-mail:")
    return EMAIL

 



# Função para verificar o BOA e coletar o e-mail
async def verificar_boa(update: Update, context: CallbackContext):
    if update.message.document:

        dre = context.user_data.get('dre')  # Recupera o DRE armazenado no contexto
        if not dre:
            await update.message.reply_text("DRE não encontrado. Por favor, envie seu DRE antes de enviar o BOA.")
            return
    
        document = update.message.document
        download_dir = os.path.join(os.getcwd(), "downloads")
        file_path = os.path.join(download_dir, document.file_name)

        os.makedirs(download_dir, exist_ok=True)
        
        new_file = await document.get_file()
        await new_file.download_to_drive(file_path)
        
        await update.message.reply_text(f"Arquivo PDF recebido e salvo como {document.file_name}!")

        context.user_data["file_path"] = file_path

        # Coletar o DRE do contexto
        dre = context.user_data.get("dre")
        if dre is None:
            await update.message.reply_text("Não foi possível encontrar o seu DRE. Por favor, envie novamente.")
            return WAITING_FOR_BOA

        # Passar o DRE para a função de verificação no arquivo verificar_requisitos
        dre_extraido_do_boa = verificar_requisitos.extrair_dre_boa(file_path, dre)  # Função para extrair o DRE do BOA    

        # Comparar o DRE fornecido com o DRE extraído do BOA
        if dre_extraido_do_boa == dre:
            await update.message.reply_text("DRE compativel com o BOA")

            # recebe a situação do aluno, se pode ou não estagiar considerando os critérios
            situacao = verificar_requisitos.verifica_criterios(file_path, dre)  # Passando o DRE aqui           

            # Salvar a situação no contexto para ser usado depois
            context.user_data['situacao'] = situacao

            if situacao:
                await update.message.reply_text("Liberado para Estagiar. Antes de enviar o parecer, preciso do seu e-mail.")
                return EMAIL  # Solicita o e-mail
            
            else:
                # Envia a mensagem de não autorização e o link para o Google Forms
                link_google_forms = "https://forms.gle/NBuhtQ9HhHN1W11c6"  # Substitua com o link real do seu Google Forms
                mensagem_recurso = (
                    "Status: Não autorizado para estagiar.\n"
                    "Caso deseje recorrer, preencha o formulário no link a seguir:\n"
                    f"{link_google_forms}"
                )
                await update.message.reply_text(mensagem_recurso)

        else:
            # Enviar mensagem de erro e pedir para reescrever o DRE
            await update.message.reply_text(f"DRE incompatível com o BOA. O DRE fornecido é {dre}, mas o DRE no BOA é {dre_extraido_do_boa}. Por favor, reescreva o DRE.")
            return DRE  # Retorna para a etapa de reescrever o DRE
      
        # Limpeza de arquivos temporários
        try:
            os.remove(file_path)
        except Exception as e:
            print(f"Erro ao tentar apagar o arquivo {file_path}: {e}")

        # Mensagem de encerramento após 1 minuto
        asyncio.create_task(timeout(update, context))

    else:
        await update.message.reply_text("Por favor, envie um arquivo PDF.")

# Função para calcular a validade (data atual + 3 meses)
def calcular_validade():
    # Obter a data atual
    data_atual = datetime.now()

    # Adicionar 3 meses à data atual
    data_validade = data_atual + relativedelta(months=3)

    # Formatando a data no formato desejado (DD/MM/YYYY)
    data_validade_formatada = data_validade.strftime("%d/%m/%Y")

    return data_validade_formatada

# Função para lidar com o e-mail após o usuário enviar
async def receber_email(update: Update, context: CallbackContext):
    email = update.message.text  # Captura o e-mail enviado pelo usuário
    context.user_data['email'] = email  # Armazenar no user_data para uso posterior

    # Valida o e-mail (simplesmente checando se contém '@')
    if not email or "@" not in email or "." not in email.split("@")[-1]:
        await update.message.reply_text("E-mail inválido. Por favor, envie um e-mail válido.")
        return EMAIL

    # Agora, adiciona o DRE e o e-mail na planilha
    dre = context.user_data.get("dre")
    situacao = context.user_data.get("situacao")
    file_path = context.user_data.get("file_path")

    if dre and situacao is not None:
        try:

            # Calcular validade(3 meses contando da data atual)
            validade = calcular_validade()

            sheet = client.open_by_key(SPREADSHEET_ID).sheet1
            row = len(sheet.col_values(1)) + 1
            sheet.update_cell(row, 1, dre)  # Adiciona DRE na coluna 1
            sheet.update_cell(row, 2, email)  # Adiciona e-mail na coluna 2
            sheet.update_cell(row,3,validade) # Adiciona validade na coluna 3
            await update.message.reply_text("Seu DRE e e-mail foram registrados com sucesso na planilha!")

            # Gera o parecer em PDF após registrar na planilha
            nome_pdf = "parecer_aluno.pdf"

            # Verifica diretamente a situação, sem usar .lower() se já for um booleano
            verificar_requisitos.gerar_parecer_pdf_BOA(file_path,nome_pdf, situacao)  # Aqui situacao já deve ser booleano
            
            with open(nome_pdf, "rb") as pdf_file:
                await update.message.reply_document(document=pdf_file, filename=nome_pdf)

        except Exception as e:
            await update.message.reply_text(f"Erro ao atualizar a planilha: {str(e)}")

    # Limpeza de arquivos temporários
    try:
        os.remove(nome_pdf)  # Remove o parecer após enviar
    except Exception as e:
        print(f"Erro ao tentar apagar o arquivo {nome_pdf}: {e}")

  # Mensagem de encerramento após 1 minuto
    asyncio.create_task(timeout(update, context))

# Função para o comando /start
async def start(update: Update, context: CallbackContext):
    await update.message.reply_text(
        'Olá! Bem-vindo ao bot da comissão de estágio de Ciência da Computação.\n'
        'Por favor, informe seu DRE (número de matrícula).'
    )

    # Mensagem de encerramento após 1 minuto
    asyncio.create_task(timeout(update, context))

    return DRE  # A conversa entra na etapa DRE

# Função para iniciar verificação de contrato
async def contrato(update: Update, context: CallbackContext):
    await update.message.reply_text("Análise de contrato. Detalhes em breve...")
    return ConversationHandler.END  # Finaliza a conversa ou direciona para outra etapa

# Função para encerrar a conversa
async def encerrar(update: Update, context: CallbackContext):
    await update.message.reply_text("Encerrando o processo. Até logo!") ### ESTA SENDO USADO???
    return ConversationHandler.END  # Finaliza a conversa

# Função assíncrona para o comando /Novo_pedido_Autorizacao
async def pede_boa(update: Update, context: CallbackContext):
    await update.message.reply_text("Envie seu BOA para eu verificar sua situação.")

# Função assíncrona para lidar com comandos inválidos
async def escolha_opcoes(update, context):
    await update.message.reply_text("Comando Inválido \n")
    await start(update, context)


#função para interromper conversa se o usuario demorar 60 segundos para responder
async def timeout(update: Update, context: CallbackContext):
    # Mensagem de encerramento após 1 minuto
    await asyncio.sleep(60)
    await update.message.reply_text("Obrigado por usar o bot da Comissão de Estágio! Estou encerrando agora. Para falar comigo novamente, use o comando /start.")
    return ConversationHandler.END


def main():
    application = Application.builder().token(TELEGRAM_TOKEN).build()

    # Definir os manipuladores de comando e mensagem
    conv_handler = ConversationHandler(
        entry_points=[CommandHandler('start', start)],
        states={
            DRE: [MessageHandler(filters.TEXT & ~filters.COMMAND, coletar_dre)],
            AVALIACAO_CONTRATO: [MessageHandler(filters.TEXT & ~filters.COMMAND, contrato)],
            WAITING_FOR_BOA: [MessageHandler(filters.Document.ALL, verificar_boa)],
            EMAIL: [MessageHandler(filters.TEXT & ~filters.COMMAND, receber_email)],
        },
        fallbacks=[MessageHandler(filters.TEXT & ~filters.COMMAND, escolha_opcoes)],
    )

    application.add_handler(conv_handler)
    application.run_polling()

if __name__ == '__main__':
    main()
