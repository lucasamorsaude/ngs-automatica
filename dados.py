# medicina.py

import os
import time
from datetime import datetime
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.action_chains import ActionChains
from datetime import datetime, timedelta
from slack_sdk import WebClient
from slack_sdk.errors import SlackApiError

client = WebClient(token=os.getenv("SLACK_TOKEN"))

hoje = datetime.today()

# Primeiro dia do mês atual
primeiro_dia_mes = hoje.replace(day=1)

# Dia de ontem
ontem = hoje - timedelta(days=1)

# Formatando no formato dd/MM/yyyy
data_inicio = primeiro_dia_mes.strftime('%d/%m/%Y')
data_fim = ontem.strftime('%d/%m/%Y')



def coletar_dados():
    # Cria diretórios necessários
    pasta_saida = "personal_dir"
    pasta_cache = "cache"
    os.makedirs(pasta_saida, exist_ok=True)
    os.makedirs(pasta_cache, exist_ok=True)

    data_hoje = datetime.today().strftime('%Y-%m-%d')
    caminho_arquivo = os.path.join(pasta_saida, f"Dados_Medicina_{data_hoje}.xlsx")

    # Configura o Chrome com cache de usuário
    chrome_options = Options()
    chrome_options.add_argument(f"--user-data-dir={os.path.abspath(pasta_cache)}")
    # chrome_options.add_argument("--headless")  # Ativa se quiser rodar sem abrir janela

    driver = webdriver.Chrome(options=chrome_options)

    try:
        url = "https://app.powerbi.com/groups/854db62b-4b1a-4bbf-a835-096847a7ae57/reports/8d6c53c6-9d40-49b0-a25a-3519ecd31bf8/a23b795f7901ac08309c?experience=power-bi"
        driver.get(url)
        print("Acessando Power BI...")

        wait = WebDriverWait(driver, 60)  # até 60 segundos para carregamento

        for label, data in [("Data de início", data_inicio), ("Data de término", data_fim)]:
            campo_data = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, f'input[aria-label^="{label}"]')))
            campo_data.clear()
            campo_data.send_keys(data)

        
        time.sleep(5)

        # region Tela de Performance

        numero_agendamentos = int(((driver.find_element(By.CSS_SELECTOR, 'p[data-sub-selection-object-name="callout-value__medidas.Agendamentos realizados"]')).text).replace('.', ''))

        qca = int(((driver.find_element(By.CSS_SELECTOR, 'p[data-sub-selection-object-name="callout-value__medidas.qca"]')).text).replace('.', ''))

        numero_consultas = int(((driver.find_element(By.CSS_SELECTOR, 'p[data-sub-selection-object-name="callout-value__medidas.Atendimentos realizados"]')).text).replace('.', ''))

        conversao_atendimentos = int(round(float(((driver.find_element(By.CSS_SELECTOR, 'p[data-sub-selection-object-name="callout-value__medidas.Conversão de atendimentos"]')).text).replace(',', '.').replace('%',''))))

        faturamento_medicina = int(round(float(((driver.find_element(By.CSS_SELECTOR, 'p[data-sub-selection-object-name="callout-value__medidas.Faturamento bruto"]')).text).replace('.', '').replace('R$', '').replace(',', '.'))))

        tm_faturamento_medicina = int(round(float(faturamento_medicina/numero_consultas)))



        url = "https://app.powerbi.com/groups/854db62b-4b1a-4bbf-a835-096847a7ae57/reports/8d6c53c6-9d40-49b0-a25a-3519ecd31bf8/ReportSection19646c6494d1d39a877b?experience=power-bi&clientSideAuth=0"
        driver.get(url)

        time.sleep(5)

        valor_exames = int(round(float(WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//div[@class="pivotTableCellWrap tablixAlignRight main-cell " and @column-index="3" and @aria-colindex="5"]'))).text.replace('R$', '').replace('.', '').replace(',', '.'))))

        tm_exames = round(valor_exames / numero_consultas, 2)


        driver.get("https://novobi.webdentalsolucoes.io/pages/login")
        wait = WebDriverWait(driver, 10)

        usuario = "coorlucasnunes"
        senha = "Luc220703*"

        # Preenche o campo de usuário
        usuario_campo = wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/app-layout/app-login/div/div[2]/form/div[1]/input")))
        usuario_campo.send_keys(usuario)
        time.sleep(0.5)

        # Preenche o campo de senha
        senha_campo = wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/app-layout/app-login/div/div[2]/form/div[2]/input")))
        senha_campo.send_keys(senha)
        time.sleep(0.5)

        # Clica nas opções restantes...
        solucoes_button = wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/app-layout/app-login/div/div[2]/form/div[4]/select/option[2]")))
        solucoes_button.click()
        time.sleep(0.5)

        blank_campo = wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/app-layout/app-login/div/div[2]/form/div[5]/ng-select/div")))
        blank_campo.click()

        blank_campo = wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/app-layout/app-login/div/div[2]/form/div[5]/ng-select/select-dropdown/div/div[2]/ul/li/span")))
        blank_campo.click()

        loggin_button = wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/app-layout/app-login/div/div[2]/form/div[6]/div/button")))
        loggin_button.click()
        time.sleep(0.5)


        url = "https://novobi.webdentalsolucoes.io/indicador/finan_caixa_total"
        driver.get(url)

        time.sleep(5)
        # Obter valor do caixa total
        valor_caixa_total = wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/app-layout/div/div/app-indicador/div/section[2]/div/div[1]/div[2]/app-painel-resultado-mensal/div[1]/div/div[1]/h3"))).text
        valor_caixa_total = valor_caixa_total.replace("R$ ", "")

        url = "https://novobi.webdentalsolucoes.io/indicador/finan_total_efetivacoes"
        driver.get(url)

        time.sleep(5)
        # Obter valor da efetivação total
        valor_efetivacao_total = wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/app-layout/div/div/app-indicador/div/section[2]/div/div[1]/div[2]/app-painel-resultado-mensal/div[1]/div/div[1]/h3"))).text
        valor_efetivacao_total = valor_efetivacao_total.replace("R$ ", "")

        url = "https://novobi.webdentalsolucoes.io/indicador/pacientes_cadastrados"
        driver.get(url)

        time.sleep(5)
        # Obter valor da efetivação total
        novos_cadastrados = wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/app-layout/div/div/app-indicador/div/section[2]/div/div[1]/div[2]/app-painel-resultado-mensal/div[1]/div/div[1]/h3"))).text
        novos_cadastrados = novos_cadastrados.replace("R$ ", "").replace(",00","")

        url = "https://novobi.webdentalsolucoes.io/indicador/finan_total_ticketmedio"
        driver.get(url)

        time.sleep(5)
        # Obter valor da efetivação total
        tm_medio_efetivado = wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/app-layout/div/div/app-indicador/div/section[2]/div/div[1]/div[2]/app-painel-resultado-mensal/div[1]/div/div[1]/h3"))).text
        tm_medio_efetivado = tm_medio_efetivado.replace("R$ ", "").replace(",00","")



        # region Sair BI Odonto
        perfil_button = wait.until(EC.presence_of_element_located((By.XPATH,"/html/body/app-layout/div/app-header/header/nav/div[3]/ul/li[2]/a/span")))
        perfil_button.click()

        logoff_button = wait.until(EC.presence_of_element_located((By.XPATH,"/html/body/app-layout/div/app-header/header/nav/div[3]/ul/li[2]/ul/li[2]/div[2]")))
        logoff_button.click()
        # endregion Sair BI Odonto



    finally:
        driver.quit()

        dados = [
            {"Indicador": "NUM_QCA", "Valor": qca},
            {"Indicador": "NUM_AGENDAMENTOS", "Valor": numero_agendamentos},
            {"Indicador": "NUM_CONSULTAS", "Valor": numero_consultas},
            {"Indicador": "CONVERSAO", "Valor": conversao_atendimentos},
            {"Indicador": "R$_FATURAMENTO_MED", "Valor": faturamento_medicina},
            {"Indicador": "R$_TM_MEDIO_MEDICINA", "Valor": tm_faturamento_medicina},
            {"Indicador": "R$_EXAMES_LABORATORIAIS", "Valor": valor_exames},
            {"Indicador": "R$_TM_EXAMES_LABORATORIAIS", "Valor": tm_exames},
            {"Indicador": "R$_CAIXA_TOTAL", "Valor": valor_caixa_total},
            {"Indicador": "R$_EFETIVACAO_TOTAL", "Valor": valor_efetivacao_total},
            {"Indicador": "NOVOS_PACIENTES", "Valor": novos_cadastrados},
        ]

        df = pd.DataFrame(dados)
        df.to_excel("indicadores.xlsx", index=False)
        
        

