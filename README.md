# Automação de Processamento de E-mails e Envio via WhatsApp

Este projeto automatiza a verificação de e-mails, extração de informações relevantes e envio de mensagens e imagens via WhatsApp. Ele foi desenvolvido para facilitar o monitoramento de e-mails importantes, como leilões e chamadas públicas.

## Funcionalidades

- **Verificação de E-mails**: Verifica e-mails recebidos no Outlook.
- **Filtragem Inteligente**: Filtra e-mails com base em assuntos específicos e remove prefixos indesejados.
- **Extração de Tabelas**: Extrai tabelas do corpo dos e-mails e as converte em imagens.
- **Envio via WhatsApp**: Envia mensagens e imagens extraídas dos e-mails para contatos no WhatsApp.

## Requisitos

- **Python 3.x**
- **Bibliotecas Python**:
  - selenium
  - chromedriver-autoinstaller
  - beautifulsoup4
  - imgkit
  - openai
  - pywin32
- **Outros**:
  - [wkhtmltopdf](https://wkhtmltopdf.org/downloads.html) (para converter HTML em imagens)
  - Conta no OpenAI com API key

## Instalação

1. Clone este repositório:
   ```sh
   git clone https://github.com/Vinitziin
   cd seu-repositorio


## Configure suas credenciais no arquivo credenciais.json:

{
  "api_openAI": {
    "token": "sua_api_key"
  },
  "nome_contato":{
    "contato": "Eu" 
  },
  "filtros": {
    "assuntos_procurados": [],
    "prefixos_para_remover": []
  }
}

## Instale as dependências:

pip install -r requirements.txt

## Uso

1. Inicie o script principal
**python leitura.py**

2.O script abrirá o WhatsApp Web no navegador. Faça login na sua conta do WhatsApp.

3.O script começará a verificar e-mails recebidos e,
ao encontrar um e-mail relevante, 
enviará a mensagem e a imagem correspondente para o contato configurado.


## Estrutura do Projeto

- pasta_img/: Diretório onde as imagens geradas a partir dos e-mails são armazenadas.
- credenciais.json: Arquivo de configuração contendo as credenciais e filtros.
- dev.ipynb: Notebook Jupyter para desenvolvimento e testes.
- Leitura_app.bat: Script batch para iniciar a aplicação no Windows.
- leitura.py: Script principal que integra todas as funcionalidades.
- README.md: Este arquivo, contendo informações sobre o projeto.
- Recomendação_de_uso.txt: Arquivo de texto com recomendações de uso.
- requirements.txt: Lista de dependências do projeto.