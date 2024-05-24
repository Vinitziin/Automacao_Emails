import os
import win32com.client
from selenium import webdriver
from datetime import datetime, date
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import chromedriver_autoinstaller
import time
from selenium.common.exceptions import ElementClickInterceptedException
from bs4 import BeautifulSoup
import imgkit
from openai import OpenAI
import json

with open('credenciais.json', 'r') as file:
    config = json.load(file)

token_openai = config['api_openAI']['token']

assuntos_procurados = config['filtros']['assuntos_procurados']

prefixos_para_remover = config['filtros']['prefixos_para_remover']


chromedriver_autoinstaller.install()
path_img = os.getcwd()
client = OpenAI(api_key= token_openai)
#===================================Extração da imagem===================================
def extrair_tabela_html(email_tabela):
    html_body = email_tabela.HTMLbody
    soup = BeautifulSoup(html_body, 'html.parser')
    tabelas = soup.find_all('table')
    return tabelas

def limpar_html_cid(html_content):
    soup = BeautifulSoup(html_content, 'html.parser')
    for img in soup.find_all('img', src=lambda x: x and x.startswith('cid:')):
        img.decompose()  
    return str(soup)

def tabela_img(tabela_html):
    try:
        clean_html = limpar_html_cid(tabela_html)  
        config_options = {
            'load-error-handling': 'ignore',
            'width': 500,
            'height': 500,
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
    prompt = f"""
    Você é um assistente útil. Abaixo está o conteúdo de um e-mail que contém informações sobre um leilão. Extraia o prazo de envio e validade envio do leilão a partir deste e-mail.

    E-mail:
    {email}

    Por favor, forneça apenas o prazo de envio do leilão, e a validade de envio, e com isso peço que seja conciso ao foncer as informações, apenas datas e horarios.
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
tamanho_minimo = 120_000
wait = WebDriverWait(driver, 60)
#===================================Filtro de Leilões e lógica principal===================================
def verificar_emails():
    outlook = win32com.client.Dispatch("outlook.Application")
    inbox = outlook.GetNameSpace("MAPI").GetDefaultFolder(6)

    hoje = date.today().strftime("%d/%m/%Y %H:%M")
    # hoje = datetime.now()
    filtro = "[ReceivedTime] >= '" + hoje + "'"
    emails_hoje = inbox.Items.Restrict(filtro)

    for email in emails_hoje:
        
        assunto_original = email.Subject
        assunto = assunto_original.upper() 
        identificador_email = f"{assunto}"

        if identificador_email in emails_notificados:
            continue
        

        if "RFQ "  in assunto and "COPEL " in assunto:
            tabela_html = extrair_tabela_html(email)
            html_completo = ''.join([str(tabela) for tabela in tabela_html])
            imagem = tabela_img(html_completo)         
            mensagem = f"*Assunto*: {assunto_original}"
            novas_mensagens.add((mensagem, imagem))
            emails_notificados.add(identificador_email)    
            continue
            

        if "CEMIG " in assunto:
                image_attachments = [attachment for attachment in email.Attachments if attachment.FileName.lower().endswith('.png') and attachment.Size > tamanho_minimo]
                if image_attachments:
                    selected_attachment = image_attachments[0]
                    save_path = os.path.join(os.getcwd(), 'pasta_img', selected_attachment.FileName)
                    selected_attachment.SaveAsFile(save_path)
                    mensagem = f"*Assunto*: {assunto_original}"
                    novas_mensagens.add((mensagem, save_path))
                    emails_notificados.add(identificador_email)           
                continue


        else:
            if ("COTAÇÕES" in assunto) or ("COTAÇÃO" in assunto):
                continue
            
            if "AES" in assunto:
                continue

            

            if any(prefixo in assunto for prefixo in prefixos_para_remover):
                continue
            
            for assunto_procurado in assuntos_procurados:
                if (assunto_procurado.upper() in assunto):    

                    tabela_html = extrair_tabela_html(email)
                    if tabela_html:
                        html_completo = ''.join([str(tabela) for tabela in tabela_html])  
                        imagem = tabela_img(html_completo) 
                        mensagem = f"*Assunto*: {assunto_original}\n\n {ask_question(email.body)}"
                        novas_mensagens.add((mensagem, imagem))
                        emails_notificados.add(identificador_email)                           
                    break         
                
    return novas_mensagens

#===================================Interação com o Whatsapp===================================
def enviar_imagem_whatsapp(img_path):
    try:
        anexo_botao = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'span[data-icon="attach-menu-plus"]')))
        driver.execute_script("arguments[0].scrollIntoView(true);", anexo_botao)
        time.sleep(1)  
        anexo_botao.click()
    except ElementClickInterceptedException:
        driver.execute_script("arguments[0].click();", anexo_botao)

    input_imagem = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'input[accept="image/*,video/mp4,video/3gpp,video/quicktime"]')))
    input_imagem.send_keys(img_path)
    time.sleep(2)

    enviar_botao = wait.until(EC.presence_of_element_located((By.XPATH, '//span[@data-icon="send"]')))
    enviar_botao.click()


def abrir_chat_contato(driver, contato):
    barra_pesquisa = WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.XPATH, '//div[@contenteditable="true"][@data-tab="3"]'))
    )
    barra_pesquisa.clear()
    barra_pesquisa.send_keys(contato)
    time.sleep(2)

    contato_selecionado = WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.XPATH, f'//span[@title="{contato}"]'))
    )
    contato_selecionado.click()
    time.sleep(2)

def enviar_mensagem_whatsapp(driver, mensagem):
    
    campo_mensagem = WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.XPATH, '//div[@contenteditable="true"][@data-tab="10"]'))
    )

    campo_mensagem.send_keys(mensagem)

    botao_enviar = driver.find_element(By.XPATH, '//button[@class="x1c4vz4f x2lah0s xdl72j9 xfect85 x1iy03kw x1lfpgzf"]')
    botao_enviar.click()
#===================================Loop principal===================================

abrir_chat_contato(driver, 'Eu')
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