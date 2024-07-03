import os
from sys import prefix
from flask import config
import win32com.client
from selenium import webdriver
from datetime import datetime, date
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
import chromedriver_autoinstaller
import time
from selenium.common.exceptions import ElementClickInterceptedException
from bs4 import BeautifulSoup
import imgkit
import re
import logging
from openai import OpenAI
import json

def carregar_configuracoes():
    """
    Carrega as configurações do arquivo credenciais.json.

    Returns:
        dict: Dicionário contendo as configurações carregadas.
    """
    with open('credenciais.json', 'r') as file:
        config = json.load(file)
    return config

config = carregar_configuracoes()
token_openai = config['api_openAI']['token']
client = OpenAI(api_key= token_openai)

contato = config['nome_contato']['contato']
chromedriver_autoinstaller.install()
path_img = os.getcwd()

#===================================Extração da imagem===================================
def extrair_tabela_html(email_tabela):
    """
    Extrai tabelas HTML do corpo do e-mail.

    Args:
        email_tabela: Objeto de e-mail contendo HTML.

    Returns:
        list: Lista de objetos de tabela extraídos do HTML.
    """
    html_body = email_tabela.HTMLbody
    soup = BeautifulSoup(html_body, 'html.parser')
    tabelas = soup.find_all('table')
    return tabelas

def limpar_html_cid(html_content):
    """
    Remove referências a imagens embutidas (CID) do conteúdo HTML.

    Args:
        html_content: Conteúdo HTML como string.

    Returns:
        str: Conteúdo HTML limpo como string.
    """
    soup = BeautifulSoup(html_content, 'html.parser')
    for img in soup.find_all('img', src=lambda x: x and x.startswith('cid:')):
        img.decompose()  
    return str(soup)

def tabela_img(tabela_html):
    """
    Converte uma tabela HTML em uma imagem.

    Args:
        tabela_html: Conteúdo HTML da tabela como string.

    Returns:
        str: Caminho completo da imagem gerada, ou None em caso de erro.
    """
    try:
        clean_html = limpar_html_cid(tabela_html)  
        config_options = {
            'load-error-handling': 'ignore',
            'width': 600,
            'height': 800,
        }
        timestamp = datetime.now().strftime("%Y%m%d%H%M%S%f")
        img_path = os.path.join(path_img + '/pasta_img',f"tabela_{timestamp}.png")
        full_img_path = os.path.abspath(img_path)
        config = imgkit.config(wkhtmltoimage=r"C:\\Program Files\\wkhtmltopdf\\bin\\wkhtmltoimage.exe")
        
        
        imgkit.from_string(clean_html, full_img_path, options=config_options, config=config)
        return full_img_path
    except OSError as e:
        print(f"Erro ao criar a imagem: {e}")
        return None

def ask_question(email):
    """
    Faz uma pergunta à API do OpenAI para extrair informações específicas do conteúdo do e-mail.

    Args:
        email: Conteúdo do e-mail como string.

    Returns:
        str: Resposta gerada pela API do OpenAI.
    """
    prompt = f"""
    Você é um assistente útil. Abaixo está o conteúdo de um e-mail que contém informações sobre um leilão. Extraia o prazo de envio e validade envio do leilão a partir deste e-mail.

    E-mail:
    {email}

    Por favor, forneça apenas o prazo de envio do leilão, e a validade de envio, e com isso peço que seja conciso ao foncer as informações, apenas datas e horarios, observação, a validade pode aparecer dessa forma: Data e horário limite para envio da Proposta por resposta
    """
    response = client.chat.completions.create(
        model="gpt-3.5-turbo",
        messages=[
            {"role": "system", "content": "Você é um assistente útil."},
            {"role": "user", "content": prompt}
        ],
        max_tokens=300
    )
    return response.choices[0].message.content
#===================================Servidor Selenium===================================
driver = webdriver.Chrome()

driver.get('https://web.whatsapp.com/')

WebDriverWait(driver, 80).until(
    EC.visibility_of_element_located((By.CSS_SELECTOR, "canvas"))
)

WebDriverWait(driver, 80).until(
    EC.visibility_of_element_located((By.ID, "side"))
)

emails_notificados = set()
novas_mensagens = set()
tamanho_minimo = 120000
wait = WebDriverWait(driver, 60)
#===================================Filtro de Leilões e lógica principal===================================
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s', filename='script.log', filemode='a')

ultima_verificacao = datetime.now()

def verificar_emails():
    """
    Verifica os e-mails recebidos no Outlook e processa aqueles que correspondem aos filtros configurados.

    Raises:
        Exception: Se ocorrer um erro ao conectar ao Outlook ou processar os e-mails.

    Returns:
        set: Conjunto de novas mensagens e imagens a serem enviadas via WhatsApp.
    """
    try:
        config = carregar_configuracoes()
        logging.info('Configurações carregadas.')
        assuntos_procurados = config['filtros']['assuntos_procurados']
        prefixos_para_remover = config['filtros']['prefixos_para_remover']

        outlook = win32com.client.Dispatch("outlook.Application")
        inbox = outlook.GetNameSpace("MAPI").GetDefaultFolder(6)

        hoje = date.today().strftime("%d/%m/%Y %H:%M")
        # filtro = "[ReceivedTime] >= '" + hoje + "'"

        filtro = f"[ReceivedTime] >= '{ultima_verificacao.strftime('%d/%m/%Y %H:%M')}'"

        emails_hoje = inbox.Items.Restrict(filtro)

        for email in emails_hoje:
            try:
                assunto_original = email.Subject
                assunto = re.sub(r'\s+', ' ', assunto_original.upper()).strip()  # Normaliza os espaços
                identificador_email = f"{assunto}"

                if identificador_email in emails_notificados:
                    continue

                if re.search(r'\bRFQ\b', assunto) and re.search(r'\bCOPEL\b', assunto):
                    processar_email(email, assunto_original, identificador_email)
                    continue

                if any(prefixo.upper() in assunto for prefixo in prefixos_para_remover):
                    continue

                if "COTAÇÕES" in assunto or "COTACOES" in assunto:
                    continue

                for assunto_procurado in assuntos_procurados:
                    if re.search(rf'\b{re.escape(assunto_procurado.upper())}\b', assunto):
                        tabela_html = extrair_tabela_html(email)
                        html_completo = ''.join([str(tabela) for tabela in tabela_html])
                        imagem = tabela_img(html_completo) if tabela_html else None
                        if imagem:
                            mensagem = f"*Assunto*: {assunto_original}\n\n {ask_question(email.body)}"
                            novas_mensagens.add((mensagem, imagem))
                            emails_notificados.add(identificador_email)
                            logging.info(f'E-mail processado: {assunto_original}')
                        break

                if re.search(r'\bCHAMADA\b', assunto) and re.search(r'\bBTG\b', assunto):
                    processar_email(email, assunto_original, identificador_email)
                    continue

                if re.search(r'\bCEMIG\b', assunto):
                    processar_attachments(email, assunto_original, identificador_email)
                    continue
                
                if re.search(r'\bCHAMADA\b', assunto) and re.search(r'\bCCGNBE\b', assunto):
                    processar_attachments(email, assunto_original, identificador_email)
                    continue	

            except Exception as e:
                logging.error(f"Erro ao processar o e-mail: {e}")

    except Exception as e:
        logging.error(f"Erro ao conectar ao Outlook {e}")

    return novas_mensagens

def processar_email(email, assunto_original, identificador_email):
    """
    Processa um e-mail extraindo suas tabelas HTML e convertendo-as em imagens.

    Args:
        email: Objeto de e-mail.
        assunto_original: Assunto original do e-mail.
        identificador_email: Identificador único do e-mail.

    Returns:
        None
    """
    tabela_html = extrair_tabela_html(email)
    html_completo = ''.join([str(tabela) for tabela in tabela_html])
    imagem = tabela_img(html_completo) if tabela_html else None
    if imagem:
        mensagem = f"*Assunto*: {assunto_original}"
        novas_mensagens.add((mensagem, imagem))
        emails_notificados.add(identificador_email)

def processar_attachments(email, assunto_original, identificador_email):
    """
    Processa os anexos de um e-mail, salvando imagens que atendam aos critérios definidos.

    Args:
        email: Objeto de e-mail.
        assunto_original: Assunto original do e-mail.
        identificador_email: Identificador único do e-mail.

    Returns:
        None
    """
    image_attachments = [attachment for attachment in email.Attachments if attachment.FileName.lower().endswith('.png') and attachment.Size > tamanho_minimo]
    if image_attachments:
        selected_attachment = image_attachments[0]
        save_path = os.path.join(os.getcwd(), 'pasta_img', selected_attachment.FileName)
        selected_attachment.SaveAsFile(save_path)
        mensagem = f"*Assunto*: {assunto_original}"
        novas_mensagens.add((mensagem, save_path))
        emails_notificados.add(identificador_email)

#===================================Interação com o Whatsapp===================================
def enviar_imagem_whatsapp(img_path):
    """
    Envia uma imagem para o chat aberto no WhatsApp Web.

    Args:
        img_path: Caminho da imagem a ser enviada.

    Raises:
        NoSuchElementException: Se não conseguir encontrar os elementos necessários para enviar a imagem.
    """
    try:
        anexo_botao = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'span[data-icon="attach-menu-plus"]')))
        driver.execute_script("arguments[0].scrollIntoView(true);", anexo_botao)
        time.sleep(1)
        anexo_botao.click()
    except ElementClickInterceptedException:
        driver.execute_script("arguments[0].click();", anexo_botao)
    except NoSuchElementException as e:
        logging.error(f"Erro ao clicar no botão de anexo: {e}")
        raise

    try:
        input_imagem = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'input[accept="image/*,video/mp4,video/3gpp,video/quicktime"]')))
        input_imagem.send_keys(img_path)
        time.sleep(2)
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//span[@data-icon="send"]'))
        ).click()
    except NoSuchElementException as e:
        logging.error(f"Erro ao enviar a imagem: {e}")
        raise

def abrir_chat_contato(driver, contato):
    """
    Abre o chat do contato especificado no WhatsApp Web.

    Args:
        driver: Instância do Selenium WebDriver.
        contato: Nome do contato para abrir o chat.

    Raises:
        NoSuchElementException: Se não conseguir encontrar o contato.
    """
    try:
        barra_pesquisa = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, '//div[@contenteditable="true"][@data-tab="3"]'))
        )
        barra_pesquisa.clear()
        barra_pesquisa.send_keys(contato)
        time.sleep(2)
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, f'//span[@title="{contato}"]'))
        ).click()
    except NoSuchElementException as e:
        logging.error(f"Erro ao abrir o chat do contato {contato}: {e}")
        raise

def enviar_mensagem_whatsapp(driver, mensagem):
    """
    Envia uma mensagem de texto para o chat aberto no WhatsApp Web.

    Args:
        driver: Instância do Selenium WebDriver.
        mensagem: Mensagem de texto a ser enviada.

    Raises:
        NoSuchElementException: Se não conseguir encontrar o campo de mensagem ou o botão de enviar.
    """
    try:
        campo_mensagem = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, '//div[@contenteditable="true"][@data-tab="10"]'))
        )
        campo_mensagem.send_keys(mensagem)

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//button[@class="x1c4vz4f x2lah0s xdl72j9 xfect85 x1iy03kw x1lfpgzf"]'))
        ).click()
    except NoSuchElementException as e:
        logging.error(f"Erro ao enviar a mensagem: {e}")
        raise
#===================================Loop principal===================================

abrir_chat_contato(driver, contato)
enviar_mensagem_whatsapp(driver, 'Bom dia! Iniciando monitoramento de email...')

while True:
    mensagens_com_imagens = verificar_emails()
    for mensagem, imagem in mensagens_com_imagens:
        time.sleep(2)
        enviar_mensagem_whatsapp(driver, f"Novo email encontrado:\n\n {mensagem}")
        time.sleep(2)
        enviar_imagem_whatsapp(imagem)
    novas_mensagens = set()
    time.sleep(60)
#===================================fim do código===================================