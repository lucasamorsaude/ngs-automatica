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
    chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_argument("--headless")  # Ativa se quiser rodar sem abrir janela

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

        numero_agendamentos = int(((driver.find_element(By.CSS_SELECTOR, 'p[data-sub-selection-object-name="callout-value__medidas.Agendamentos realizados"]')).text).replace('.', ''))

        numero_consultas = int(((driver.find_element(By.CSS_SELECTOR, 'p[data-sub-selection-object-name="callout-value__medidas.Atendimentos realizados"]')).text).replace('.', ''))

        conversao_atendimentos = int(round(float(((driver.find_element(By.CSS_SELECTOR, 'p[data-sub-selection-object-name="callout-value__medidas.Conversão de atendimentos"]')).text).replace(',', '.').replace('%',''))))

        faturamento_medicina = int(round(float(((driver.find_element(By.CSS_SELECTOR, 'p[data-sub-selection-object-name="callout-value__medidas.Faturamento bruto"]')).text).replace('.', '').replace('R$', '').replace(',', '.'))))

        tm_faturamento_medicina = int(round(float(faturamento_medicina/numero_consultas)))


        print("Agendamentos:", numero_agendamentos)
        print("Consultas:", numero_consultas)
        print(f"Conversão de Atendimento: {conversao_atendimentos}%")
        print(f"Faturamento: R${faturamento_medicina}")
        print(f"T.M. Medicina: R${tm_faturamento_medicina}")

        
        print("✅ Botão clicado com sucesso!")


        input("Finalizado")
    finally:
        driver.quit()


coletar_dados()