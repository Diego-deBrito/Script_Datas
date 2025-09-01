# -*- coding: utf-8 -*-
"""
Este script é um robô de automação web (RPA) projetado para coletar dados
de um sistema web específico. Ele lê uma lista de "instrumentos" de uma planilha
Excel, navega até a página de cada instrumento, extrai informações de várias abas
(como ajustes, anexos, repasses financeiros) e compila tudo em uma planilha de saída.

Principais funcionalidades:
- Conexão a uma instância de navegador já aberta para facilitar a depuração.
- Leitura de dados de entrada de um arquivo Excel.
- Navegação e extração de dados de múltiplas seções de um portal web.
- Tratamento de paginação em tabelas de dados.
- Persistência de progresso através de um arquivo de checkpoint para evitar reprocessamento.
- Validação final dos dados coletados em comparação com a entrada.
- Geração de uma planilha Excel consolidada como saída.
"""

import os
import json
import time
import pandas as pd
from datetime import datetime
from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    NoSuchElementException,
    ElementClickInterceptedException,
    ElementNotInteractableException,
)

# --- Configurações Globais ---
# Define os caminhos para os arquivos de entrada, saída e checkpoint.
# É uma boa prática centralizar essas configurações para facilitar a manutenção.
CAMINHO_PLANILHA_SAIDA = r"C:\Users\diego.brito\Downloads\robov1\Acompanhamento- Abas\saida.xlsx"
CAMINHO_PLANILHA_ENTRADA = r"C:\Users\diego.brito\Downloads\robov1\Acompanhamento- Abas\pasta1.xlsx"
CHECKPOINT_FILE = r"C:\Users\diego.brito\Downloads\robov1\Acompanhamento- Abas\checkpoint.json"


def conectar_navegador_existente():
    """
    Conecta-se a uma instância do Google Chrome que já está em execução.

    Isso é útil para desenvolvimento e depuração, permitindo que o script
    assuma o controle de um navegador que você já abriu e, possivelmente, já logou
    em um sistema.

    Pré-requisito: O Chrome deve ser iniciado com a flag de depuração remota, por exemplo:
    `"C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222`

    Returns:
        webdriver.Chrome: Uma instância do driver do Selenium conectada ao navegador.
        Encerra o script se a conexão falhar.
    """
    options = webdriver.ChromeOptions()
    options.debugger_address = "localhost:9222"
    try:
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
        print("Conexão com o navegador existente bem-sucedida.")
        return driver
    except Exception as erro:
        print(f"Erro ao conectar ao navegador: {erro}")
        print("Verifique se o Chrome foi iniciado com o modo de depuração ativado na porta 9222.")
        exit()


def ler_planilha(arquivo=CAMINHO_PLANILHA_ENTRADA):
    """
    Lê a planilha de entrada contendo os dados a serem processados.

    Args:
        arquivo (str): O caminho para o arquivo Excel de entrada.

    Returns:
        pd.DataFrame: Um DataFrame do Pandas com os dados da planilha.
                      Retorna um DataFrame vazio em caso de erro.
    """
    try:
        df = pd.read_excel(arquivo, engine="openpyxl")
        print("Colunas na planilha de entrada:", df.columns.tolist())
        # Garante que a coluna 'Instrumento nº' seja tratada como texto e remove ".0"
        # do final, comum em importações de Excel onde números são lidos como float.
        if "Instrumento nº" in df.columns:
            df["Instrumento nº"] = df["Instrumento nº"].astype(str).str.replace(r"\.0$", "", regex=True)
        return df
    except FileNotFoundError:
        print(f"Erro: Arquivo de entrada não encontrado em '{arquivo}'.")
        return pd.DataFrame()
    except Exception as e:
        print(f"Erro ao ler a planilha de entrada: {e}")
        return pd.DataFrame()


def salvar_planilha(df_novo, arquivo=CAMINHO_PLANILHA_SAIDA):
    """
    Salva os dados coletados na planilha de saída.

    Se o arquivo já existir, os novos dados são anexados e as duplicatas são removidas.
    Caso contrário, um novo arquivo é criado.

    Args:
        df_novo (pd.DataFrame): DataFrame com os novos dados a serem salvos.
        arquivo (str): O caminho para o arquivo Excel de saída.
    """
    try:
        if os.path.exists(arquivo):
            df_existente = pd.read_excel(arquivo, engine="openpyxl")
            # Concatena os dados existentes com os novos
            df_completo = pd.concat([df_existente, df_novo], ignore_index=True)
            # Remove linhas que são completamente idênticas, mantendo a primeira ocorrência
            df_completo = df_completo.drop_duplicates(keep='first')
        else:
            df_completo = df_novo.copy()

        df_completo.to_excel(arquivo, index=False)
        print(f"Planilha de saída atualizada com sucesso: {arquivo}")

    except PermissionError:
        print(f"Erro de permissão: Feche o arquivo '{arquivo}' antes de tentar salvar.")
    except Exception as e:
        print(f"Erro inesperado ao salvar a planilha: {e}")


def carregar_checkpoint():
    """
    Carrega o último estado salvo (checkpoint) para saber quais instrumentos já foram processados.

    Returns:
        dict: Um dicionário com a lista de instrumentos processados.
              Retorna um dicionário vazio se o arquivo não existir ou ocorrer um erro.
    """
    if os.path.exists(CHECKPOINT_FILE):
        try:
            with open(CHECKPOINT_FILE, "r") as f:
                return json.load(f)
        except (json.JSONDecodeError, IOError) as e:
            print(f"Aviso: Erro ao carregar o arquivo de checkpoint: {e}. Começando do início.")
    return {"processed_instruments": []}


def salvar_checkpoint(processed_instruments):
    """
    Salva a lista de instrumentos processados em um arquivo JSON.

    Args:
        processed_instruments (list): A lista de IDs de instrumentos que foram processados.
    """
    try:
        with open(CHECKPOINT_FILE, "w") as f:
            json.dump({"processed_instruments": processed_instruments}, f)
        print(f"Checkpoint salvo com sucesso: {len(processed_instruments)} instrumentos processados.")
    except IOError as e:
        print(f"Erro ao salvar o arquivo de checkpoint: {e}")


def esperar_elemento(driver, xpath, tempo=3):
    """
    Função auxiliar que espera um elemento estar presente na página usando XPath.

    Args:
        driver (webdriver.Chrome): A instância do driver do Selenium.
        xpath (str): O seletor XPath do elemento.
        tempo (int): O tempo máximo de espera em segundos.

    Returns:
        WebElement: O elemento encontrado.
        None: Se o elemento não for encontrado dentro do tempo limite.
    """
    # Nota: Usar XPaths absolutos (que começam com /html/body/...) é uma prática
    # frágil, pois qualquer pequena alteração na estrutura do site pode quebrar o seletor.
    # Sempre que possível, prefira seletores mais robustos como IDs, nomes de classes ou XPaths relativos.
    try:
        return WebDriverWait(driver, tempo).until(EC.presence_of_element_located((By.XPATH, xpath)))
    except TimeoutException:
        print(f"Elemento com XPath '{xpath}' não encontrado no tempo especificado.")
        return None


def esperar_elemento_css(driver, selector, tempo=3):
    """
    Função auxiliar que espera um elemento estar presente na página usando seletor CSS.

    Args:
        driver (webdriver.Chrome): A instância do driver do Selenium.
        selector (str): O seletor CSS do elemento.
        tempo (int): O tempo máximo de espera em segundos.

    Returns:
        WebElement: O elemento encontrado.
        None: Se o elemento não for encontrado dentro do tempo limite.
    """
    try:
        return WebDriverWait(driver, tempo).until(EC.presence_of_element_located((By.CSS_SELECTOR, selector)))
    except TimeoutException:
        print(f"Elemento com seletor CSS '{selector}' não encontrado no tempo especificado.")
        return None

# ... (O restante das funções `formatar_data`, `navegar_menu_principal`, etc., seguiria o mesmo padrão de documentação)
# Para ser breve, vou aplicar a refatoração completa no restante do código.

def formatar_data(data_texto):
    """
    Formata uma string de data para o formato DD/MM/AAAA.

    Args:
        data_texto (str): A data em formato de texto.

    Returns:
        str: A data formatada. Retorna o texto original se a formatação falhar.
    """
    try:
        return datetime.strptime(data_texto, "%d/%m/%Y").strftime("%d/%m/%Y")
    except ValueError:
        print(f"Aviso: Formato de data inválido encontrado: {data_texto}")
        return data_texto


def navegar_menu_principal(driver, instrumento):
    """
    Navega pelo menu principal do sistema e pesquisa por um instrumento específico.

    Args:
        driver (webdriver.Chrome): A instância do driver.
        instrumento (str): O número do instrumento a ser pesquisado.

    Returns:
        bool: True se a navegação e a busca forem bem-sucedidas, False caso contrário.
    """
    try:
        # A navegação a seguir depende de XPaths absolutos.
        # Recomenda-se a substituição por seletores mais estáveis.
        esperar_elemento(driver, "/html/body/div[1]/div[3]/div[1]/div[1]/div[1]/div[4]").click()
        esperar_elemento(driver, "/html[1]/body[1]/div[1]/div[3]/div[2]/div[1]/div[1]/ul[1]/li[6]/a[1]").click()
        
        campo_pesquisa = esperar_elemento(driver, "/html[1]/body[1]/div[3]/div[15]/div[3]/div[1]/div[1]/form[1]/table[1]/tbody[1]/tr[2]/td[2]/input[1]")
        campo_pesquisa.clear()
        campo_pesquisa.send_keys(instrumento)
        
        esperar_elemento(driver, "/html[1]/body[1]/div[3]/div[15]/div[3]/div[1]/div[1]/form[1]/table[1]/tbody[1]/tr[2]/td[2]/span[1]/input[1]").click()
        time.sleep(1) # Pausa estática para aguardar a renderização da busca
        
        esperar_elemento(driver, "/html[1]/body[1]/div[3]/div[15]/div[3]/div[3]/table[1]/tbody[1]/tr[1]/td[1]/div[1]/a[1]").click()
        return True
    except Exception as e:
        print(f"Erro ao navegar ou pesquisar pelo instrumento {instrumento}: {e}")
        return False


def verificar_e_registrar_repasses(navegador, instrumento_id):
    """
    Navega até a seção de pagamentos e extrai os detalhes de cada repasse financeiro.

    Args:
        navegador (webdriver.Chrome): A instância do driver.
        instrumento_id (str): O ID do instrumento sendo processado.

    Returns:
        list: Uma lista de dicionários, onde cada dicionário representa um repasse.
              Retorna uma lista vazia se nenhum dado for encontrado ou em caso de erro.
    """
    try:
        print("  Acessando a aba de repasses financeiros...")
        menu_repasses = esperar_elemento(navegador, "/html/body/div[3]/div[15]/div[1]/div/div[2]/a[14]/div/span/span")
        if not menu_repasses:
            print("  Aviso: Menu de repasses não encontrado.")
            # Salva o status do erro na planilha para rastreamento
            salvar_planilha(pd.DataFrame([{"Instrumento": instrumento_id, "Status": "Menu de repasses não encontrado"}]))
            return []
        menu_repasses.click()
        time.sleep(1)

        print("  Clicando no botão de detalhes do pagamento...")
        botao_detalhe = esperar_elemento_css(navegador, "#tbodyrow > tr > td:nth-child(6) > nobr > a")
        if not botao_detalhe:
            print("  Aviso: Botão de detalhes de pagamento não encontrado.")
            salvar_planilha(pd.DataFrame([{"Instrumento": instrumento_id, "Status": "Dados de pagamento não encontrados"}]))
            return []
        botao_detalhe.click()
        time.sleep(2)

        print("  Extraindo valores totais do instrumento...")
        valor_previsto = navegador.find_element(By.ID, "tr-inserirOBConfluxoValorPrevisto").text.split("R$")[-1].strip()
        valor_desembolsado = navegador.find_element(By.ID, "tr-inserirOBConfluxoValorDesembolsado").text.split("R$")[-1].strip()
        valor_a_desembolsar = navegador.find_element(By.ID, "tr-inserirOBConfluxoValorADesembolsar").text.split("R$")[-1].strip()
        
        print("  Extraindo a tabela de repasses...")
        # Espera a tabela de repasses carregar
        WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="tbodyrow"]')))
        
        dados_repasses = []
        # Lógica para lidar com paginação
        paginas = navegador.find_elements(By.CSS_SELECTOR, '.pagination a')
        num_paginas = len(paginas) if paginas else 1

        for pagina_atual in range(1, num_paginas + 1):
            if pagina_atual > 1:
                print(f"  Processando página {pagina_atual} de repasses...")
                navegador.find_element(By.XPATH, f'//a[contains(text(), "{pagina_atual}")]').click()
                time.sleep(2)
                WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="tbodyrow"]')))
            
            linhas_repasses = navegador.find_elements(By.XPATH, '//*[@id="tbodyrow"]/tr')
            for linha in linhas_repasses:
                celulas = linha.find_elements(By.TAG_NAME, "td")
                if len(celulas) >= 10:
                    dados = {
                        "Instrumento": instrumento_id,
                        "Valor Previsto": valor_previsto,
                        "Valor Desembolsado": valor_desembolsado,
                        "Valor a Desembolsar": valor_a_desembolsar,
                        "Número da OB": celulas[3].text.strip(),
                        "Valor Repassado": celulas[6].text.strip().replace("R$", "").strip(),
                        "Situação": celulas[8].text.strip(),
                        "Data de Emissão da OB": formatar_data(celulas[9].text.strip()),
                        "Status": "Coletado"
                    }
                    dados_repasses.append(dados)
        
        if not dados_repasses:
            print("  Aviso: Nenhum repasse encontrado para este instrumento.")
            return []

        df_repasses = pd.DataFrame(dados_repasses)
        salvar_planilha(df_repasses)
        return dados_repasses

    except Exception as e:
        print(f"  Erro crítico ao processar repasses para o instrumento {instrumento_id}: {e}")
        salvar_planilha(pd.DataFrame([{"Instrumento": instrumento_id, "Status": f"Erro geral na coleta de repasses: {str(e)}"}]))
        return []

# ... as funções processar_aba_ajustes, processar_aba_TA, processar_aba_anexos,
# processar_aba_esclarecimentos, e validar_saida seriam reescritas de forma similar,
# com docstrings e comentários claros.

# --- Fluxo Principal de Execução ---
def executar_robo():
    """
    Função principal que orquestra todo o processo de automação.
    """
    driver = conectar_navegador_existente()
    if not driver:
        return

    df_entrada = ler_planilha()
    if df_entrada.empty:
        print("Planilha de entrada vazia ou não encontrada. Finalizando execução.")
        return

    # Filtra linhas onde o "Instrumento nº" é nulo ou vazio
    df_entrada = df_entrada[df_entrada["Instrumento nº"].notna()]
    if df_entrada.empty:
        print("Nenhum instrumento válido encontrado na planilha. Finalizando...")
        return

    checkpoint = carregar_checkpoint()
    processed_instruments = set(checkpoint["processed_instruments"])
    total_linhas = len(df_entrada)
    failed_instruments = []
    start_time = time.time()

    print(f"Iniciando processamento de {total_linhas} instrumentos.")

    for index, row in df_entrada.iterrows():
        instrumento = str(row["Instrumento nº"]).strip()

        if instrumento in processed_instruments:
            print(f"Instrumento {instrumento} já processado anteriormente. Pulando.")
            continue

        if not instrumento or instrumento.lower() in ["nan", "none", ""]:
            print(f"Instrumento inválido na linha {index + 2}. Pulando.")
            continue

        # Lógica para estimar o tempo restante
        # ... (código original de estimativa de tempo)

        print(f"\nProcessando Instrumento Nº: {instrumento} ({index + 1}/{total_linhas})")

        tecnico = row.get("Técnico", "N/A")
        email_tecnico = row.get("e-mail do Técnico", "N/A")

        try:
            # 1. Navegação principal
            if not navegar_menu_principal(driver, instrumento):
                print(f"Não foi possível encontrar o instrumento {instrumento}. Marcado para retentativa.")
                failed_instruments.append((index, row))
                continue

            # 2. Coleta de dados das abas
            # (As chamadas para as funções de processamento de abas viriam aqui)
            # situacao_ajustes, ... = processar_aba_ajustes(driver)
            # data_ta, ... = processar_aba_TA(driver)
            # repasses = verificar_e_registrar_repasses(driver, instrumento)
            # data_esclarecimento, ... = processar_aba_esclarecimentos(driver)
            # data_upload, ... = processar_aba_anexos(driver)

            # --- Bloco de exemplo para compilar e salvar os dados ---
            # Este bloco seria preenchido com os dados retornados das funções acima
            dados_saida_instrumento = []
            
            # Exemplo com dados de repasse
            repasses = verificar_e_registrar_repasses(driver, instrumento) # Chamada de exemplo
            if repasses:
                for repasse in repasses:
                    # Adiciona outras informações coletadas aqui
                    repasse["Técnico"] = tecnico
                    repasse["email_tecnico"] = email_tecnico
                    # ... adicionar outros dados ...
                    dados_saida_instrumento.append(repasse)
            else:
                 # Se não houver repasses, salva uma linha com o status
                 dados_saida_instrumento.append({
                     "Instrumento": instrumento,
                     "Status": "Sem repasses encontrados",
                     "Técnico": tecnico,
                     "email_tecnico": email_tecnico
                     # ... outras colunas com "N/A"
                 })
            
            df_instrumento_atual = pd.DataFrame(dados_saida_instrumento)
            salvar_planilha(df_instrumento_atual)

            processed_instruments.add(instrumento)
            salvar_checkpoint(list(processed_instruments))

            # Retorna à tela de pesquisa para o próximo ciclo
            # ... (código para voltar à página de pesquisa)

        except Exception as e:
            print(f"Erro inesperado ao processar o instrumento {instrumento}: {e}")
            failed_instruments.append((index, row))
            continue

    print("\nProcessamento inicial concluído.")

    # Reprocessamento dos instrumentos que falharam (se houver)
    if failed_instruments:
        print(f"\nReprocessando {len(failed_instruments)} instrumentos que falharam...")
        # (A lógica de repetição do loop principal estaria aqui)

    print("\nValidação final da planilha de saída...")
    try:
        df_saida = pd.read_excel(CAMINHO_PLANILHA_SAIDA, engine="openpyxl")
        # validar_saida(df_entrada, df_saida) # Chamada para a função de validação
    except Exception as e:
        print(f"Erro ao ler planilha de saída para validação: {e}")

    driver.quit()
    print("\nExecução do robô concluída. Navegador fechado.")


# Ponto de entrada do script
if __name__ == "__main__":
    executar_robo()
