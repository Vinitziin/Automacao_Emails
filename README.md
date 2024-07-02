```markdown
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
   git clone https://github.com/Vinitziin/Automacao_Emails
   cd seu-repositorio
   ```

2. Configure suas credenciais no arquivo `credenciais.json`:
   ```json
   {
     "api_openAI": {
       "token": "sua_api_key"
     },
     "filtros": {
       "assuntos_procurados": [
         "LEILÃO",
         "Chamada de Compra e Venda",
         "Chamada Pública",
         "Leilão de Compra e Venda",
         "Venda de Energia Elétrica",
         "Chamada para Compra e Venda de Energia",
         "Venda",
         "Compra",
         "Chamada de Venda Energia Elétrica",
         "Consulta Pública",
         "RFQ COPEL",
         "Compra de Energia",
         "Comunicado de Venda",
         "Comunicado de Compra"
       ],
       "prefixos_para_remover": [
         "ENC: ",
         "RES: ",
         "LEMBRETE: ",
         "Re: ",
         "RE",
         "Assinatura Contrato",
         "Aceita: ",
         "[ENCERRAMENTO]",
         "NÃO VENCEDOR",
         "[LEMBRETE]",
         "DIVULGAÇÃO",
         "I-REC",
         "AVISO DE LICITAÇÃO",
         "Processo de VENDA",
         "[ENCERRADO] ",
         "LEMBRETE",
         "Operação",
         "Transferência",
         "AES",
         "MVE",
         "UISA",
         "CP",
         "CURTO PRAZO",
         "MCP",
         "Cotação",
         "Cotações"
       ]
     }
   }
   ```

3. Instale as dependências:
   ```sh
   pip install -r requirements.txt
   ```

4. Baixe e instale o [wkhtmltopdf](https://wkhtmltopdf.org/downloads.html).

## Uso

1. Inicie o script principal:
   ```sh
   python leitura.py
   ```
2. O script abrirá o WhatsApp Web no navegador. Faça login na sua conta do WhatsApp.

3. O script começará a verificar e-mails recebidos e, ao encontrar um e-mail relevante, enviará a mensagem e a imagem correspondente para o contato configurado.

## Estrutura do Projeto

- **pasta_img/**: Diretório onde as imagens geradas a partir dos e-mails são armazenadas.
- **credenciais.json**: Arquivo de configuração contendo as credenciais e filtros.
- **dev.ipynb**: Notebook Jupyter para desenvolvimento e testes.
- **Leitura_app.bat**: Script batch para iniciar a aplicação no Windows.
- **leitura.py**: Script principal que integra todas as funcionalidades.
- **README.md**: Este arquivo, contendo informações sobre o projeto.
- **Recomendação_de_uso.txt**: Arquivo de texto com recomendações de uso.
- **requirements.txt**: Lista de dependências do projeto.

## Contribuição

Se você quiser contribuir com este projeto, siga estas etapas:

1. Faça um fork deste repositório.
2. Crie uma branch para sua feature:
   ```sh
   git checkout -b minha-feature
   ```
3. Faça commit das suas alterações:
   ```sh
   git commit -m 'Adicionei minha feature'
   ```
4. Faça push para a branch:
   ```sh
   git push origin minha-feature
   ```
5. Abra um Pull Request.

## Licença

Este projeto está licenciado sob a MIT License. Veja o arquivo [LICENSE](LICENSE) para mais detalhes.

### Melhorias Implementadas

1. **Estrutura do Repositório**: Adicionada uma seção de estrutura do projeto para ajudar os usuários a entender a organização dos arquivos.
2. **Passos de Instalação**: Clarificados os passos de instalação, incluindo a configuração do arquivo `credenciais.json` e instalação do `wkhtmltopdf`.
3. **Informações de Uso**: Passos claros para iniciar o script e usar a aplicação.
4. **Contribuição**: Instruções para colaboradores interessados em contribuir com o projeto.

