import time
import pandas as pd
import os
from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from datetime import datetime
import pandas as pd
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementClickInterceptedException, ElementNotInteractableException







# 🛠 Conectar ao navegador já aberto
def conectar_navegador_existente():
    options = webdriver.ChromeOptions()
    options.debugger_address = "localhost:9222"
    try:
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
        print("✅ Conectado ao navegador existente!")
        return driver
    except Exception as erro:
        print(f"❌ Erro ao conectar ao navegador: {erro}")
        exit()


# 📂 Ler planilha de entrada
def ler_planilha(arquivo=r"C:\Users\diego.brito\Downloads\robov1\pasta1.xlsx"):
    df = pd.read_excel(arquivo, engine="openpyxl")

    # 🛠️ Remover ".0" da coluna "Instrumento nº"
    if "Instrumento nº" in df.columns:
        df["Instrumento nº"] = df["Instrumento nº"].astype(str).str.replace(r"\.0$", "", regex=True)

    return df

# 📤 Salvar planilha de saída sem sobrescrever os dados
def salvar_planilha(df, arquivo=r"C:\Users\diego.brito\Downloads\robov1\resultado_abas_main.xlsx"):
    try:
        if os.path.exists(arquivo):
            df_existente = pd.read_excel(arquivo, engine="openpyxl")
            df = pd.concat([df_existente, df], ignore_index=True)  # Mesclar os dados antigos com os novos

        df.to_excel(arquivo, index=False)
        print(f"📂 Planilha atualizada com sucesso: {arquivo}")
    except PermissionError:
        print(f"⚠️ Erro: Feche o arquivo {arquivo} antes de salvar.")
    except Exception as e:
        print(f"❌ Erro ao salvar a planilha: {e}")


# 🔍 Espera um elemento estar visível
def esperar_elemento(driver, xpath, tempo=3):
    try:
        return WebDriverWait(driver, tempo).until(EC.presence_of_element_located((By.XPATH, xpath)))
    except:
        print(f"⚠️ Elemento {xpath} não encontrado!")
        return None


# 🔄 Navegar no menu principal
def navegar_menu_principal(driver, instrumento):
    try:
        esperar_elemento(driver, "/html/body/div[1]/div[3]/div[1]/div[1]/div[1]/div[4]").click()
        esperar_elemento(driver, "/html[1]/body[1]/div[1]/div[3]/div[2]/div[1]/div[1]/ul[1]/li[6]/a[1]").click()
        campo_pesquisa = esperar_elemento(driver, "/html[1]/body[1]/div[3]/div[15]/div[3]/div[1]/div[1]/form[1]/table[1]/tbody[1]/tr[2]/td[2]/input[1]")
        campo_pesquisa.clear()
        campo_pesquisa.send_keys(instrumento)
        esperar_elemento(driver, "/html[1]/body[1]/div[3]/div[15]/div[3]/div[1]/div[1]/form[1]/table[1]/tbody[1]/tr[2]/td[2]/span[1]/input[1]").click()
        time.sleep(1)
        esperar_elemento(driver, "/html[1]/body[1]/div[3]/div[15]/div[3]/div[3]/table[1]/tbody[1]/tr[1]/td[1]/div[1]/a[1]").click()
        return True
    except:
        print(f"⚠️ Instrumento {instrumento} não encontrado.")
        return False



import time


def processar_aba_ajustes(driver):
    """ Acessa a aba Ajustes do PT, identifica o maior número (antes da barra) e sua situação,
    clica em 'Detalhar' e extrai a 'Data da Solicitação'. """

    situacao_ajustes = "Nenhum registro encontrado"
    numero_maior = None
    data_solicitacao = "Data não encontrada"

    try:
        print("📂 Acessando Aba Ajustes do PT...")

        # 📌 1️⃣ Localizar e clicar na aba principal de Ajustes do PT
        aba_ajustes = WebDriverWait(driver, 0.1).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "div[id='div_-481524888'] span span"))
        )  # Adicionei o parêntese de fechamento aqui
        print("✅ Aba Ajustes do PT encontrada!")
        driver.execute_script("arguments[0].scrollIntoView();", aba_ajustes)
        time.sleep(1)

        try:
            aba_ajustes.click()
        except (ElementNotInteractableException, ElementClickInterceptedException):
            print("⚠️ Clique normal falhou, tentando via JavaScript...")
            driver.execute_script("arguments[0].click();", aba_ajustes)

        time.sleep(2)

        # 📌 2️⃣ Localizar e clicar na aba secundária dentro de Ajustes do PT
        sub_aba_ajustes = WebDriverWait(driver, 0.1).until(
            EC.element_to_be_clickable(
                (By.CSS_SELECTOR, "a[id='menu_link_-481524888_-1293190284'] div[class='inactiveTab'] span span"))
        )
        print("✅ Sub Aba Ajustes do PT encontrada!")

        driver.execute_script("arguments[0].scrollIntoView();", sub_aba_ajustes)
        time.sleep(1)

        try:
            sub_aba_ajustes.click()
        except (ElementNotInteractableException, ElementClickInterceptedException):
            print("⚠️ Clique normal falhou, tentando via JavaScript...")
            driver.execute_script("arguments[0].click();", sub_aba_ajustes)

        time.sleep(2)

        # 📌 3️⃣ Esperar a tabela carregar dentro do caminho fornecido
        tabela_ajustes = WebDriverWait(driver, 0.1).until(
            EC.presence_of_element_located((By.XPATH, "/html/body/div[3]/div[15]/div[3]/div[2]/div[2]"))
        )
        linhas = tabela_ajustes.find_elements(By.TAG_NAME, "tr")

        if not linhas or len(linhas) < 2:  # Verifica se há pelo menos uma linha de dados
            print("⚠️ Nenhuma linha de dados encontrada na tabela de Ajustes do PT.")
            return situacao_ajustes, numero_maior, data_solicitacao

        maior_numero = -1
        situacao_maior = "Desconhecida"
        linha_maior = None  # Armazenar a linha onde o maior número foi encontrado

        # 📌 4️⃣ Encontrar índice das colunas "Número" e "Situação"
        cabecalho = linhas[0].find_elements(By.TAG_NAME, "th")
        indices = {"Número": None, "Situação": None}

        for i, coluna in enumerate(cabecalho):
            texto = coluna.text.strip()
            if "Número" in texto:
                indices["Número"] = i
            elif "Situação" in texto:
                indices["Situação"] = i

        if indices["Número"] is None or indices["Situação"] is None:
            print("❌ Erro: Não foram encontradas as colunas 'Número' e 'Situação'.")
            return situacao_ajustes, numero_maior, data_solicitacao

        # 📌 5️⃣ Identificar o maior número antes da barra ("/") e sua situação correspondente
        for linha in linhas[1:]:  # Ignora cabeçalho
            colunas = linha.find_elements(By.TAG_NAME, "td")

            if len(colunas) > max(indices["Número"], indices["Situação"]):  # Garante que os índices são válidos
                try:
                    numero_texto = colunas[indices["Número"]].text.strip()

                    # 💡 Extraindo apenas a parte antes da barra
                    if "/" in numero_texto:
                        numero_base = int(numero_texto.split("/")[0])  # Pega apenas o número antes da barra
                    else:
                        numero_base = int(numero_texto)  # Caso não tenha barra, converte normalmente

                    situacao = colunas[indices["Situação"]].text.strip()

                    if numero_base > maior_numero:
                        maior_numero = numero_base
                        situacao_maior = situacao
                        linha_maior = linha  # Armazenar a linha do maior número

                except ValueError:
                    print(f"⚠️ Número inválido encontrado: {numero_texto}")

        if maior_numero == -1:
            print("⚠️ Nenhum número válido encontrado.")
            return situacao_ajustes, numero_maior, data_solicitacao

        situacao_ajustes = situacao_maior
        numero_maior = f"{maior_numero}/2024"  # Formata de volta para o formato correto

        print(f"✅ Maior número encontrado: {numero_maior} - Situação: {situacao_ajustes}")

        # 📌 6️⃣ Clicar em "Detalhar" na linha do maior número
        try:
            if linha_maior:
                botao_detalhar = linha_maior.find_element(By.XPATH, ".//nobr/a[contains(text(), 'Detalhar')]")
                botao_detalhar.click()
                print("✅ Botão 'Detalhar' clicado!")
            else:
                print("❌ Erro: Linha do maior número não encontrada.")
                return situacao_ajustes, numero_maior, data_solicitacao
        except Exception as e:
            print(f"❌ Erro ao clicar em 'Detalhar': {e}")
            return situacao_ajustes, numero_maior, data_solicitacao

        # 📌 7️⃣ Extrair a "Data da Solicitação" da página de detalhes
        try:
            data_element = WebDriverWait(driver, 0.1).until(
                EC.presence_of_element_located((By.XPATH, "/html/body/div[3]/div[15]/div[4]/div[1]/div/form/table/tbody/tr[13]/td[2]"))
            )
            data_solicitacao = data_element.text.strip()
            print(f"✅ Data da Solicitação encontrada: {data_solicitacao}")
        except TimeoutException:
            print("⚠️ Data da Solicitação não encontrada.")
        except Exception as e:
            print(f"❌ Erro ao extrair a Data da Solicitação: {e}")

    except TimeoutException:
        print("❌ Erro: Tempo limite ao tentar acessar a Aba Ajustes do PT.")
    except NoSuchElementException:
        print("❌ Erro: Elemento não encontrado. O seletor pode estar incorreto.")
    except Exception as e:
        print(f"❌ Erro ao processar Aba Ajustes do PT: {e}")

    return situacao_ajustes, numero_maior, data_solicitacao  # Retorna os valores extraídos


def processar_aba_TA(driver):
    """Acessa a Aba TA, identifica o maior número da coluna 'Número', extrai sua situação correspondente,
    clica em 'Detalhar' e extrai a 'Data da Solicitação'."""

    situacao_TA = "Tabela não encontrada"
    numero_maior = "Tabela não encontrada"
    data_solicitacao = "Data não encontrada"

    try:
        print("📂 Acessando a Aba TA...")

        # 1️⃣ Clicar na Aba TA principal via JavaScript (usando querySelector)
        try:
            aba_TA = driver.execute_script("return document.querySelector('#menu_link_-481524888_82854 > div')")
            if aba_TA:
                driver.execute_script("arguments[0].scrollIntoView();", aba_TA)
                driver.execute_script("arguments[0].click();", aba_TA)
                print("✅ Aba TA acessada!")
            else:
                print("❌ Erro: Aba TA não encontrada!")
                return numero_maior, situacao_TA, data_solicitacao
        except Exception as e:
            print(f"❌ Erro ao clicar na Aba TA: {e}")
            return numero_maior, situacao_TA, data_solicitacao

        # 2️⃣ Clicar na Sub Aba TA via JavaScript (usando querySelector)
        try:
            sub_aba_TA = driver.execute_script(
                "return document.querySelector('#menu_link_-173460853_82854 > div > span > span');")

            if sub_aba_TA:
                driver.execute_script("arguments[0].scrollIntoView();", sub_aba_TA)  # Garantir visibilidade

                # Tentar clique normal
                try:
                    sub_aba_TA.click()
                except (ElementNotInteractableException, ElementClickInterceptedException):
                    print("⚠️ Clique normal falhou, tentando via JavaScript...")
                    driver.execute_script("arguments[0].click();", sub_aba_TA)

                print("✅ Sub Aba TA acessada!")
            else:
                print("❌ Erro: Sub Aba TA não encontrada!")

        except Exception as e:
            print(f"❌ Erro ao clicar na Sub Aba TA: {e}")
            return numero_maior, situacao_TA, data_solicitacao

        # 3️⃣ Esperar a tabela carregar
        try:
            tabela_TA = WebDriverWait(driver, 0.1).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "#listaSolicitacoes"))
            )
            print("✅ Tabela TA carregada!")
        except TimeoutException:
            print("⚠️ Tabela TA não carregou completamente.")
            return numero_maior, situacao_TA, data_solicitacao

        # 4️⃣ Capturar os índices das colunas "Número" e "Situação"
        cabecalho = tabela_TA.find_elements(By.TAG_NAME, "th")
        indices = {"Número": None, "Situação": None}

        for i, coluna in enumerate(cabecalho):
            texto = coluna.text.strip().lower()
            if "número" in texto:
                indices["Número"] = i
            elif "situação" in texto:
                indices["Situação"] = i

        if indices["Número"] is None or indices["Situação"] is None:
            print("❌ Erro: Colunas 'Número' e 'Situação' não foram encontradas.")
            return numero_maior, situacao_TA, data_solicitacao

        # 5️⃣ Identificar o maior número antes da barra ("/") e sua situação correspondente
        maior_numero = -1
        situacao_maior = "Desconhecida"
        linha_maior = None  # Armazenar a linha onde o maior número foi encontrado
        linhas = tabela_TA.find_elements(By.TAG_NAME, "tr")[1:]  # Ignorar cabeçalho

        for linha in linhas:
            colunas = linha.find_elements(By.TAG_NAME, "td")
            if len(colunas) > max(indices["Número"], indices["Situação"]):  # Garante que os índices são válidos
                try:
                    numero_texto = colunas[indices["Número"]].text.strip()
                    situacao = colunas[indices["Situação"]].text.strip()

                    # 💡 Extraindo apenas a parte antes da barra "/"
                    numero_base = int(numero_texto.split("/")[0]) if "/" in numero_texto else int(numero_texto)

                    if numero_base > maior_numero:
                        maior_numero = numero_base
                        situacao_maior = situacao
                        linha_maior = linha  # Armazenar a linha do maior número

                except ValueError:
                    print(f"⚠️ Número inválido encontrado: {numero_texto}")

        if maior_numero == -1:
            print("⚠️ Nenhum número válido encontrado.")
            return numero_maior, situacao_TA, data_solicitacao

        situacao_TA = situacao_maior
        numero_maior = f"{maior_numero}/2024"  # Formata de volta para o formato correto

        print(f"✅ Maior número encontrado: {numero_maior} - Situação: {situacao_TA}")

        # 6️⃣ Clicar em "Detalhar" na linha do maior número
        try:
            if linha_maior:
                botao_detalhar = linha_maior.find_element(By.XPATH, ".//nobr/a[contains(text(), 'Detalhar')]")
                botao_detalhar.click()
                print("✅ Botão 'Detalhar' clicado!")
            else:
                print("❌ Erro: Linha do maior número não encontrada.")
                return numero_maior, situacao_TA, data_solicitacao
        except Exception as e:
            print(f"❌ Erro ao clicar em 'Detalhar': {e}")
            return numero_maior, situacao_TA, data_solicitacao

        # 7️⃣ Extrair a "Data da Solicitação" da página de detalhes
        try:
            data_element = WebDriverWait(driver, 0.1).until(
                EC.presence_of_element_located((By.XPATH, "/html/body/div[3]/div[15]/div[3]/div[1]/div/form/table/tbody/tr[13]/td[2]"))
            )
            data_solicitacao = data_element.text.strip()
            print(f"✅ Data da Solicitação encontrada: {data_solicitacao}")
        except TimeoutException:
            print("⚠️ Data da Solicitação não encontrada.")
        except Exception as e:
            print(f"❌ Erro ao extrair a Data da Solicitação: {e}")

    except TimeoutException:
        print("❌ Erro: Tempo limite ao tentar acessar a Aba TA.")
    except NoSuchElementException:
        print("❌ Erro: Elemento não encontrado. O seletor pode estar incorreto.")
    except Exception as e:
        print(f"❌ Erro ao processar Aba TA: {e}")

    return numero_maior, situacao_TA, data_solicitacao  # Retorna os valores extraídos



# 📌 Processar Aba Anexos
def processar_aba_anexos(driver):
    """ Acessa a aba de Anexos e extrai a Data Upload mais recente. """

    data_upload_recente = "Nenhum anexo encontrado"  # Valor padrão caso nada seja encontrado
    erro_pesquisa = "Pesquisa não realizada"  # Caso o botão de pesquisa não seja encontrado

    try:
        print("📂 Acessando Aba de Anexos...")

        # 📌 Passo 1: Acessar a aba correta
        aba_anexos_primaria = esperar_elemento(driver, "/html[1]/body[1]/div[3]/div[15]/div[1]/div[1]/div[1]/a[2]/div[1]/span[1]/span[1]")
        if aba_anexos_primaria:
            driver.execute_script("arguments[0].scrollIntoView();", aba_anexos_primaria)
            aba_anexos_primaria.click()
        else:
            print("⚠️ Aba Anexos não encontrada.")
            return data_upload_recente, erro_pesquisa

        aba_anexos_secundaria = esperar_elemento(driver, "/html[1]/body[1]/div[3]/div[15]/div[1]/div[1]/div[2]/a[8]/div[1]/span[1]/span[1]")
        if aba_anexos_secundaria:
            driver.execute_script("arguments[0].scrollIntoView();", aba_anexos_secundaria)
            aba_anexos_secundaria.click()
        else:
            print("⚠️ Aba secundária de Anexos não encontrada.")
            return data_upload_recente, erro_pesquisa

        # 📌 Passo 2: Clicar no botão de pesquisa para carregar a lista de anexos
        botao_pesquisar = esperar_elemento(driver, "/html[1]/body[1]/div[3]/div[15]/div[3]/div[1]/div[1]/form[1]/table[1]/tbody[1]/tr[1]/td[2]/input[2]")
        if botao_pesquisar:
            driver.execute_script("arguments[0].click();", botao_pesquisar)
        else:
            print("⚠️ Botão de pesquisa não encontrado.")
            return data_upload_recente, "Botão de pesquisa não encontrado"

        # 📌 Passo 3: Aguardar a tabela carregar
        tabela_anexos = esperar_elemento(driver, "/html/body/div[3]/div[15]/div[4]/div/div[1]/form/div/div[1]/table")
        if not tabela_anexos:
            print("⚠️ Tabela de anexos não encontrada.")
            return data_upload_recente, "Tabela de anexos não encontrada"

        # 📌 Passo 4: Coletar todas as linhas da tabela
        linhas = tabela_anexos.find_elements(By.TAG_NAME, "tr")
        if not linhas:
            print("⚠️ Nenhum anexo encontrado.")
            return data_upload_recente, "Nenhum anexo na tabela"

        datas_upload = []  # Lista para armazenar as datas encontradas

        for linha in linhas[1:]:  # Ignorar cabeçalho
            colunas = linha.find_elements(By.TAG_NAME, "td")

            if len(colunas) >= 3:  # Garante que há pelo menos 3 colunas
                data_texto = colunas[2].text.strip()  # Pegando a coluna 'Data Upload'

                if data_texto:
                    try:
                        data_formatada = datetime.strptime(data_texto, "%d/%m/%Y")
                        datas_upload.append(data_formatada)
                    except ValueError:
                        print(f"⚠️ Data inválida ignorada: {data_texto}")


        # 📌 Passo 5: Se houver datas, pegar a mais recente
        if datas_upload:
            data_upload_recente = max(datas_upload).strftime("%d/%m/%Y")
            print(f"📅 Data mais recente na coluna 'Data Upload': {data_upload_recente}")
        else:
            print("⚠️ Nenhuma data válida encontrada na coluna 'Data Upload'.")

    except Exception as e:
        print(f"❌ Erro ao processar Aba de Anexos: {e}")
        return "Erro ao processar", "Erro ao processar"

    return data_upload_recente, "Pesquisa realizada com sucesso"




# 📌 Processar Aba Esclarecimentos
def processar_aba_esclarecimentos(driver):
    """Acessa a aba Esclarecimentos, percorre todas as páginas, encontra a Data de Solicitação mais recente e clica em 'Detalhar'."""
    try:
        print("📂 Acessando Aba Esclarecimentos...")

        # 📌 Passo 1: Acessar a aba correta
        aba_esclarecimentos_primaria = WebDriverWait(driver, 0.1).until(
            EC.element_to_be_clickable((By.XPATH, "/html[1]/body[1]/div[3]/div[2]/div[4]/div[1]/div[7]"))
        )
        aba_esclarecimentos_primaria.click()
        time.sleep(1)

        aba_esclarecimentos_secundaria = WebDriverWait(driver, 0.1).until(
            EC.element_to_be_clickable(
                (By.XPATH, "/html[1]/body[1]/div[3]/div[2]/div[5]/div[1]/div[2]/ul[1]/li[1]/a[1]"))
        )
        aba_esclarecimentos_secundaria.click()
        time.sleep(1)

        # 📌 Passo 2: Identificar o número total de páginas
        try:
            paginacao_texto = WebDriverWait(driver, 0.1).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "#esclarecimentos > span:nth-child(1)"))
            ).text

            # Extraindo o número total de páginas do texto (ex: "Página 1 de 5 (81 item(s))")
            total_paginas = int(paginacao_texto.split(" de ")[1].split(" ")[0])
            print(f"📄 Total de páginas encontradas: {total_paginas}")

        except Exception as e:
            print(f"⚠️ Erro ao identificar número de páginas: {e}")
            total_paginas = 1  # Se não conseguir identificar, assume que há apenas uma página

        # Variáveis para armazenar a data mais recente e seu botão correspondente
        data_mais_recente = None
        botao_detalhar_associado = None

        # 📌 Passo 3: Percorrer todas as páginas e encontrar a data mais recente
        for pagina in range(1, total_paginas + 1):
            print(f"➡️ Acessando página {pagina} de {total_paginas}...")

            if pagina > 1:
                try:
                    # Clicar no botão para avançar para a próxima página
                    botao_proxima_pagina = WebDriverWait(driver, 0.1).until(
                        EC.element_to_be_clickable((By.LINK_TEXT, str(pagina)))
                    )
                    driver.execute_script("arguments[0].scrollIntoView();", botao_proxima_pagina)
                    botao_proxima_pagina.click()
                    time.sleep(1)  # Aguarde a nova página carregar
                except Exception as e:
                    print(f"⚠️ Erro ao mudar para a página {pagina}: {e}")
                    break

            try:
                # Localizar a tabela de esclarecimentos
                tabela_esclarecimentos = WebDriverWait(driver, 0.1).until(
                    EC.presence_of_element_located((By.XPATH,
                                                    "/html/body/div[3]/div[15]/div[4]/div[1]/div/form/table/tbody/tr[6]/td/div[1]/table"))
                )
                linhas = tabela_esclarecimentos.find_elements(By.TAG_NAME, "tr")

                for linha in linhas:
                    colunas = linha.find_elements(By.TAG_NAME, "td")

                    if len(colunas) >= 7:
                        try:
                            data_texto = colunas[0].text.strip()
                            botao_detalhar = colunas[6].find_element(By.TAG_NAME, "a")

                            if data_texto and botao_detalhar:
                                try:
                                    data_formatada = datetime.strptime(data_texto, "%d/%m/%Y")

                                    if data_mais_recente is None or data_formatada > data_mais_recente:
                                        data_mais_recente = data_formatada
                                        botao_detalhar_associado = botao_detalhar

                                except ValueError:
                                    print(f"⚠️ Data inválida ignorada: {data_texto}")

                        except Exception as e:
                            print(f"⚠️ Erro ao processar linha da tabela: {e}")

            except Exception as e:
                print(f"⚠️ Erro ao buscar dados na página {pagina}: {e}")

        # 📌 Passo 4: Se encontrou uma data válida, clicar em "Detalhar"
        if botao_detalhar_associado:
            data_esclarecimento = data_mais_recente.strftime("%d/%m/%Y")
            print(f"📅 Data de Esclarecimento mais recente: {data_esclarecimento}")

            driver.execute_script("arguments[0].scrollIntoView();", botao_detalhar_associado)
            ActionChains(driver).move_to_element(botao_detalhar_associado).perform()
            botao_detalhar_associado.click()
            print("✅ Clicou no botão 'Detalhar'!")
            time.sleep(1)
        else:
            print("⚠️ Nenhuma Data de Solicitação encontrada.")
            return "Sem informação", "Nenhum anexo encontrado", "Data não encontrada"

        # 📌 Passo 5: Verificar Respostas
        campo_presente = "NÃO"
        try:
            print("🔍 Verificando se há respostas...")

            campo_especificado = WebDriverWait(driver, 0.1).until(
                EC.presence_of_element_located(
                    (By.XPATH, "/html/body/div[3]/div[15]/div[3]/div/div/form/table/tbody/tr[17]/td[1]"))
            )
            if campo_especificado:
                campo_presente = "SIM"
                print(f"📂 Resposta encontrada: 'SIM'")
            else:
                print("⚠️ Nenhuma resposta encontrada.")
                return "Sem informação", "Nenhum anexo encontrado", "Data não encontrada"

        except Exception as e:
            print(f"⚠️ Erro ao verificar respostas: {e}")
            return "Sem informação", "Nenhum anexo encontrado", "Data não encontrada"

        # 📌 Passo 6: Extrair a Data da Resposta do Esclarecimento
        try:
            data_resposta_esclarecimento = WebDriverWait(driver, 0.1).until(
                EC.presence_of_element_located(
                    (By.XPATH, "/html/body/div[3]/div[15]/div[3]/div/div/form/table/tbody/tr[16]/td[4]"))
            ).text
            print(f"📅 Data da Resposta do Esclarecimento: {data_resposta_esclarecimento}")
        except Exception as e:
            print(f"⚠️ Erro ao extrair a Data da Resposta do Esclarecimento: {e}")
            return "Sem informação", "Nenhum anexo encontrado", "Data não encontrada"

        return data_esclarecimento, campo_presente, data_resposta_esclarecimento

    except Exception as e:
        print(f"⚠️ Erro ao verificar o novo campo: {e}")
        return "Erro ao processar", "Erro ao processar", "Erro ao processar"



# 📂 Caminho da planilha de saída
CAMINHO_PLANILHA_SAIDA = r"C:\Users\diego.brito\Downloads\robov1\saida.xlsx"


# 🚀 Fluxo principal do robô
def executar_robo():
    """ Executa o robô navegando entre as abas e coletando os dados, ignorando campos vazios (NaN). """
    driver = conectar_navegador_existente()
    df_entrada = ler_planilha()

    # 🔹 **Filtrar apenas instrumentos válidos (não NaN)**
    df_entrada = df_entrada[df_entrada["Instrumento nº"].notna()]

    if df_entrada.empty:
        print("⚠️ Nenhum instrumento válido encontrado na planilha. Finalizando...")
        return

    total_linhas = len(df_entrada)  # total de linhas a serem processadas
    dados_saida = []

    print(f"🚀 Iniciando processamento dos instrumentos...({total_linhas} no total)")

    for index, row in df_entrada.iterrows():
        linha_atual = index + 1  # Linha começa do 1
        linhas_restantes = total_linhas - linha_atual

        print(f"📌 Buscando linha {linha_atual}... Restam {linhas_restantes} linhas.")

        instrumento = str(row["Instrumento nº"]).strip()
        tecnico = row["Técnico"].strip() if pd.notna(row["Técnico"]) else "N/A"
        email_tecnico = row["e-mail do Técnico"].strip() if pd.notna(row["e-mail do Técnico"]) else "N/A"

        # 🔹 **Verificar se o campo não está vazio após conversão**
        if not instrumento or instrumento in ["nan", "None", ""]:
            print(f"⚠️ Instrumento inválido encontrado na linha {index + 1}. Pulando...")
            continue

        print(f"\n🔎 Processando Instrumento Nº: {instrumento} ({index + 1}/{len(df_entrada)})")

        try:
            if not navegar_menu_principal(driver, instrumento):
                print(f"⚠️ Instrumento {instrumento} não encontrado. Pulando para o próximo...")
                continue

            # Chamando funções de processamento de cada aba
            situacao_ajustes, numero_maior, data_solicitacao_ajustes = processar_aba_ajustes(driver)
            data_ta, situacao_ta, data_solicitacao_ta = processar_aba_TA(driver)
            data_upload, pesquisa_status = processar_aba_anexos(driver)
            data_esclarecimento, anexo_esclarecimento, data_resposta_esclarecimento = processar_aba_esclarecimentos(driver)

            # Adicionando os dados na lista de saída
            dados_saida.append([
                instrumento, situacao_ajustes, numero_maior, data_solicitacao_ajustes, situacao_ta, data_ta, data_solicitacao_ta,
                data_upload, data_esclarecimento, anexo_esclarecimento, data_resposta_esclarecimento, tecnico, email_tecnico
            ])

            # 📌 Criar DataFrame e salvar após cada instrumento processado
            df_saida = pd.DataFrame(dados_saida, columns=[
                "Instrumento", "Situação Ajustes", "Número Ajustes", "Data Solicitação Ajuste", "Situação TA", "Número TA", "Data Solicitação TA",
                "Aba Anexos", "Data Esclarecimento", "Resposta Esclarecimento", "Data Resposta Esclarecimento", "Técnico", "e-mail do Técnico"
            ])
            df_saida.to_excel(CAMINHO_PLANILHA_SAIDA, index=False)
            print(f"📂 Planilha atualizada: {CAMINHO_PLANILHA_SAIDA}")

            # 📌 Passo Final: Voltar para a pesquisa de instrumentos
            try:
                print("🔄 Voltando para a tela de pesquisa...")

                botao_voltar = WebDriverWait(driver, 0.1).until(
                    EC.element_to_be_clickable((By.XPATH, "/html/body/div[3]/div[2]/div[1]/a"))
                )
                driver.execute_script("arguments[0].scrollIntoView();", botao_voltar)
                botao_voltar.click()
                time.sleep(1)  # Aguarde a página carregar

                print("✅ Retornou para a tela de pesquisa!")

            except Exception as e:
                print(f"⚠️ Erro ao tentar voltar para a tela de pesquisa: {e}")

        except Exception as e:
            print(f"❌ Erro ao processar o instrumento {instrumento}: {e}")
            continue  # Continua para o próximo instrumento mesmo em caso de erro

    print("✅ Processamento concluído! Planilha salva com sucesso.")

# 🔥 Executando o robô
executar_robo()