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



        # region Sair BI Odonto
        perfil_button = wait.until(EC.presence_of_element_located((By.XPATH,"/html/body/app-layout/div/app-header/header/nav/div[3]/ul/li[2]/a/span")))
        perfil_button.click()

        logoff_button = wait.until(EC.presence_of_element_located((By.XPATH,"/html/body/app-layout/div/app-header/header/nav/div[3]/ul/li[2]/ul/li[2]/div[2]")))
        logoff_button.click()
        # endregion Sair BI Odonto



        print("Agendamentos:", numero_agendamentos)
        print("Consultas:", numero_consultas)
        print(f"Conversão de Atendimento: {conversao_atendimentos}%")
        print(f"Faturamento: R${faturamento_medicina}")
        print(f"T.M. Medicina: R${tm_faturamento_medicina}")
        print(f"Valor de exame: R${valor_exames}")
        print(f"T.M. Exames: R${tm_exames}")
        print(f"Caixa total: {valor_caixa_total}")
        print(f"Efetivação total: {valor_efetivacao_total}")
        print(f"Novos Cadastrados: {novos_cadastrados}")

        

    finally:
        driver.quit()

        mensagem = (
            f"*Agendamentos:* {numero_agendamentos}\n"
            f"*Consultas:* {numero_consultas}\n"
            f"*Conversão de Atendimento:* {conversao_atendimentos}%\n"
            f"*Faturamento:* R${faturamento_medicina}\n"
            f"*T.M. Medicina:* R${tm_faturamento_medicina}\n"
            f"*Valor de exame:* R${valor_exames}\n"
            f"*T.M. Exames:* R${tm_exames}\n"
            f"*Caixa total:* R${valor_caixa_total}\n"
            f"*Efetivação total:* R${valor_efetivacao_total}\n"
            f"*Novos Cadastrados:* {novos_cadastrados}"
        )
        
        try:
            client.chat_postMessage(channel='C07LJHERK1T', text=mensagem)
        except SlackApiError as e:
            print(f"Erro ao enviar mensagem: {e.response['error']}")


coletar_dados()