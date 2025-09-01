# RPA para Coleta de Dados Web com Selenium

![Python](https://img.shields.io/badge/Python-3.7+-blue.svg)
![Libraries](https://img.shields.io/badge/Libraries-Selenium%20%7C%20Pandas-orange.svg)
![Status](https://img.shields.io/badge/Status-Funcional-success.svg)

Este projeto consiste em um robô de automação de processos (RPA) desenvolvido em Python para automatizar a coleta de dados de um sistema web. O script utiliza a biblioteca Selenium para navegação e interação com a página, e Pandas para manipulação de dados de entrada e saída em planilhas Excel.

## Descrição do Projeto

O objetivo principal deste robô é eliminar a tarefa manual e repetitiva de consultar informações de múltiplos "instrumentos" (contratos, processos, etc.) em um portal web. Ele lê uma lista de instrumentos de uma planilha, acessa a página de cada um, extrai dados de diversas abas (como informações financeiras, anexos, ajustes) e consolida tudo em uma única planilha de saída, pronta para análise.

## Funcionalidades Principais

- **Automação de Navegação Web:** Utiliza Selenium para controlar um navegador Chrome, realizar buscas e navegar por menus e abas.
- **Leitura e Escrita de Planilhas:** Usa a biblioteca Pandas para ler a lista de instrumentos de um arquivo `.xlsx` e salvar os dados coletados em outro.
- **Sistema de Checkpoint:** Salva o progresso em um arquivo `checkpoint.json`. Se a execução for interrompida, o robô pode continuar de onde parou, evitando o reprocessamento de dados já coletados.
- **Conexão com Navegador Existente:** O script é projetado para se conectar a uma instância do Chrome já aberta em modo de depuração. Isso é extremamente útil para lidar com logins, CAPTCHAs e autenticação de dois fatores manualmente antes de iniciar a automação.
- **Extração de Dados Complexos:** Capaz de extrair dados de tabelas, incluindo o tratamento de paginação.
- **Tratamento de Erros:** Lida com exceções comuns em web scraping (elementos não encontrados, tempo de espera esgotado) e registra o status na planilha de saída para fácil identificação de falhas.
- **Estrutura Modular:** O código é organizado em funções claras e bem documentadas, facilitando a manutenção e a expansão.

## Pré-requisitos

Antes de executar o projeto, garanta que você tenha os seguintes softwares instalados:

- [Python 3.7](https://www.python.org/downloads/) ou superior
- [Google Chrome](https://www.google.com/chrome/) (navegador web)
- Um editor de código de sua preferência (ex: VS Code, PyCharm).

## Instalação e Configuração

1.  **Clone o repositório:**
    ```bash
    git clone [https://github.com/seu-usuario/seu-repositorio.git](https://github.com/seu-usuario/seu-repositorio.git)
    cd seu-repositorio
    ```

2.  **Crie e ative um ambiente virtual (recomendado):**
    ```bash
    # Para Windows
    python -m venv venv
    .\venv\Scripts\activate

    # Para macOS/Linux
    python3 -m venv venv
    source venv/bin/activate
    ```

3.  **Instale as dependências:**
    Crie um arquivo chamado `requirements.txt` na raiz do projeto com o seguinte conteúdo:
    ```
    pandas
    selenium
    webdriver-manager
    openpyxl
    ```
    Em seguida, instale as bibliotecas com o comando:
    ```bash
    pip install -r requirements.txt
    ```

4.  **Configure os Caminhos dos Arquivos:**
    Abra o arquivo de script (`.py`) e ajuste as variáveis globais no início do código para refletir os caminhos corretos no seu sistema:
    ```python
    # --- Configurações Globais ---
    CAMINHO_PLANILHA_SAIDA = r"C:\caminho\completo\para\saida.xlsx"
    CAMINHO_PLANILHA_ENTRADA = r"C:\caminho\completo\para\pasta1.xlsx"
    CHECKPOINT_FILE = r"C:\caminho\completo\para\checkpoint.json"
    ```

## Como Executar

1.  **Prepare a Planilha de Entrada:**
    Certifique-se de que o arquivo `pasta1.xlsx` (ou o nome que você configurou) exista no caminho especificado e contenha uma coluna chamada `"Instrumento nº"` com a lista de itens a serem pesquisados.

2.  **Inicie o Google Chrome em Modo de Depuração:**
    Esta é a etapa mais importante. Feche todas as janelas do Google Chrome e inicie uma nova instância através do terminal/CMD com o seguinte comando. Isso permite que o Selenium se conecte a uma sessão que você controla.

    * **Para Windows (CMD ou PowerShell):**
        ```bash
        "C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222
        ```
        *(Ajuste o caminho se sua instalação do Chrome for diferente)*

    * **Para macOS:**
        ```bash
        /Applications/Google\ Chrome.app/Contents/MacOS/Google\ Chrome --remote-debugging-port=9222
        ```

3.  **Acesse o Sistema Manualmente:**
    Na janela do Chrome que abriu, navegue até o portal web, faça o login, passe por qualquer verificação de segurança e deixe-o pronto na página inicial.

4.  **Execute o Script Python:**
    Com o navegador aberto e logado, execute o script a partir do seu terminal:
    ```bash
    python nome_do_seu_script.py
    ```

O robô irá se conectar ao navegador, iniciar o processo de coleta e exibir o progresso no terminal. Ao final, a planilha `saida.xlsx` conterá todos os dados extraídos.

## Observações Importantes

> **Fragilidade dos Seletores (XPath):** Este script utiliza seletores XPath absolutos para localizar elementos na página. Esses seletores são **extremamente frágeis** e podem quebrar se a estrutura do site for alterada (mesmo que minimamente). Para uma solução mais robusta e de longo prazo, é altamente recomendável refatorar os seletores para usar alternativas mais estáveis como **IDs**, **nomes de classes**, **atributos `data-*`** ou **XPaths relativos**.

> **Uso Específico:** O robô foi desenvolvido para funcionar em um portal web específico. Ele não funcionará em outros sites sem modificações significativas nos seletores e na lógica de navegação.

## Estrutura do Projeto

```
/seu-repositorio/
├── seu_script.py              # O código principal do robô
├── pasta1.xlsx                # Arquivo de entrada (exemplo)
├── requirements.txt           # Lista de dependências Python
└── README.md                  # Este arquivo

# Arquivos gerados durante a execução:
# ├── saida.xlsx
# └── checkpoint.json
```
