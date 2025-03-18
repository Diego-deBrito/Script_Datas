📌 Full Abas - Automação com Selenium

🛠 Sobre o Projeto

Este projeto é um robô de automação web desenvolvido em Python utilizando Selenium, que acessa um sistema web, navega por diferentes abas e extrai informações importantes. Os dados coletados são armazenados em uma planilha Excel, permitindo uma análise estruturada.

🚀 Principais Funcionalidades

🔹 1. Conexão com o Navegador Existente

O código se conecta a um navegador Google Chrome já aberto utilizando a porta de depuração 9222. Isso evita a necessidade de iniciar um novo navegador toda vez que o robô for executado.

options = webdriver.ChromeOptions()
options.debugger_address = "localhost:9222"
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

🔹 2. Leitura e Manipulação de Planilhas

O código lê um arquivo Excel contendo números de instrumentos que devem ser pesquisados no sistema. A biblioteca pandas é usada para manipular esses dados:

df = pd.read_excel(arquivo, engine="openpyxl")

Além disso, os números de instrumentos são formatados corretamente:

df["Instrumento nº"] = df["Instrumento nº"].astype(str).str.replace(r"\.0$", "", regex=True)

🔹 3. Automação da Navegação no Sistema

O robô interage com o sistema web utilizando comandos do Selenium para:

Clicar em botões e menus

Preencher formulários

Capturar tabelas e textos das páginas

Aguardar elementos carregarem antes de interagir com eles

Exemplo de espera por um elemento:

def esperar_elemento(driver, xpath, tempo=3):
    try:
        return WebDriverWait(driver, tempo).until(EC.presence_of_element_located((By.XPATH, xpath)))
    except:
        print(f"⚠️ Elemento {xpath} não encontrado!")
        return None

🔹 4. Extração de Dados Importantes

O código acessa diversas abas do sistema e coleta informações específicas:

Aba Ajustes do PT: Extrai número e situação do ajuste

Aba TA (Termo de Ajuste): Identifica a solicitação mais recente

Aba Anexos: Encontra a última data de upload de documentos

Aba Esclarecimentos: Verifica respostas pendentes

Exemplo de extração da data mais recente na aba de anexos:

data_uploads = []
for linha in linhas[1:]:
    colunas = linha.find_elements(By.TAG_NAME, "td")
    if len(colunas) >= 3:
        data_texto = colunas[2].text.strip()
        data_uploads.append(datetime.strptime(data_texto, "%d/%m/%Y"))

🔹 5. Salvamento dos Dados Extraídos

Os dados coletados são salvos em uma nova planilha Excel, garantindo que não sejam sobrescritos:

if os.path.exists(arquivo):
    df_existente = pd.read_excel(arquivo, engine="openpyxl")
    df = pd.concat([df_existente, df], ignore_index=True)

df.to_excel(arquivo, index=False)

🔹 6. Execução do Robô

O código principal executa todas as etapas do robô de forma automatizada:

def executar_robo():
    driver = conectar_navegador_existente()
    df_entrada = ler_planilha()

    for index, row in df_entrada.iterrows():
        instrumento = str(row["Instrumento nº"]).strip()
        if navegar_menu_principal(driver, instrumento):
            situacao_ajustes, numero_maior, data_solicitacao_ajustes = processar_aba_ajustes(driver)
            salvar_planilha(df_saida)

🛠 Tecnologias Utilizadas

Python 3

Selenium (para automação web)

Pandas (para manipulação de planilhas)

OpenPyXL (para leitura e escrita de arquivos Excel)

WebDriver Manager (para gerenciamento do ChromeDriver)

🔧 Como Configurar e Rodar o Projeto

1️⃣ Instale as dependências:

pip install -r requirements.txt

2️⃣ Abra o Chrome no modo de depuração:

chrome.exe --remote-debugging-port=9222 --user-data-dir="C:\\ChromeProfile"

3️⃣ Execute o robô:

python Full_Abas.py

4️⃣ Verifique a planilha gerada na pasta data.

📌 Autor

👤 Diego Bruno Santos de Brito

📧 Entre em contato: debrito521@gmail.com

📝 Projeto em constante evolução! Sugestões são bem-vindas! 🚀

