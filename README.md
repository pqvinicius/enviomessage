# 📊 Disparador de Relatórios de Vendas para WhatsApp

Este é um script em Python para automatizar o envio de relatórios de desempenho de vendas, de forma personalizada, para uma lista de vendedores via WhatsApp.

O projeto lê os dados de uma planilha Excel, formata uma mensagem customizada para cada vendedor com base nesses dados e agenda o envio automático.

---

## ✨ Funcionalidades

- **Leitura de Dados via Planilha:** Carrega a lista de vendedores e seus dados de desempenho a partir de um arquivo `planilha_vendedores.xlsx`.
- **Templates de Mensagem:** Utiliza um arquivo de texto (`mensagem_template.txt`) para a mensagem, permitindo fácil customização sem alterar o código.
- **Mensagens Personalizadas:** Insere dinamicamente os dados de cada vendedor (nome, metas, faturamento, etc.) na mensagem.
- **Ambiente Isolado:** Utiliza um ambiente virtual (`venv`) e um arquivo `requirements.txt` para garantir a reprodutibilidade e facilitar a instalação.
- **Automação de Envio:** Agenda e envia as mensagens através do WhatsApp Web.

---

## 🛠️ Tecnologias Utilizadas

- **Python 3**
- **Pandas:** Para leitura e manipulação da planilha Excel.
- **PyWhatKit:** Para a automação do envio de mensagens no WhatsApp.
- **Openpyxl:** Como motor para o Pandas ler arquivos `.xlsx`.

---

## 📁 Estrutura do Projeto

Para o correto funcionamento, os arquivos devem estar organizados da seguinte forma:

/seu-projeto/
|
|-- /venv/                   # Pasta do ambiente virtual
|-- enviar_mensagens.py      # O script principal
|-- mensagem_template.txt    # O template da mensagem
|-- planilha_vendedores.xlsx # A planilha com os dados dos vendedores
|-- requirements.txt         # As dependências do projeto
`-- README.md                # Este arquivo


---

## 🚀 Instalação e Configuração

Siga os passos abaixo para configurar e rodar o projeto.

### Pré-requisitos

- Python 3.8 ou superior instalado.
- Uma conta do WhatsApp ativa e conectada ao WhatsApp Web no seu navegador padrão.

### Passos para Instalação

1.  **Clone ou baixe o projeto:**
    Baixe todos os arquivos para uma pasta em seu computador.

2.  **Crie e ative o Ambiente Virtual:**
    Abra um terminal na pasta do projeto e execute os comandos:

    ```bash
    # Criar o ambiente virtual
    python -m venv venv

    # Ativar o ambiente virtual (Windows)
    .\venv\Scripts\activate
    ```

3.  **Instale as Dependências:**
    Com o ambiente virtual ativo, instale todas as bibliotecas necessárias com um único comando:

    ```bash
    pip install -r requirements.txt
    ```

4.  **Prepare a Planilha de Dados (`planilha_vendedores.xlsx`):**
    Preencha a planilha com os dados dos vendedores. Ela **deve** conter as seguintes colunas (os nomes devem ser exatos):

    - `Nome_Vendedor`
    - `Telefone` (no formato internacional, ex: `+5531999999999`)
    - `Faturado_mes`
    - `Meta`
    - `Alcance`
    - `Fat_Projetado`
    - `Pct_Projetado`
    - `Meta_diaria`

5.  **Customize a Mensagem (`mensagem_template.txt`):**
    Edite o arquivo de texto para alterar a mensagem conforme desejado. Os placeholders (`{nome_da_coluna}`) serão substituídos pelos dados da planilha.

---

## ▶️ Como Usar

Após a instalação e configuração, siga os passos para executar o script:

1.  Certifique-se de que seu ambiente virtual (`venv`) está **ativo**.
2.  Garanta que os arquivos `planilha_vendedores.xlsx` e `mensagem_template.txt` estão preenchidos e salvos na pasta do projeto.
3.  Execute o script principal no terminal:

    ```bash
    python enviar_mensagens.py
    ```
4.  O script irá ler os dados, processar cada vendedor e começar a agendar as mensagens. Uma aba do navegador será aberta para cada envio no WhatsApp Web.

---

## 👤 Autor

 Vincius Xavier