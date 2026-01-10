import os
import time
import pandas as pd
from datetime import datetime

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from webdriver_manager.chrome import ChromeDriverManager


# =========================
# CONFIGURAÇÕES
# =========================

URL = "https://consultacpf.space/"
EXCEL_ORIGINAL = r"C:/dev/BotValidaCpfReceita_2/cpfDataNascimento.xlsx"
PASTA_SAIDA = r"C:/dev/BotValidaCpfReceita_2/Relatorios"

os.makedirs(PASTA_SAIDA, exist_ok=True)


# =========================
# UTILIDADES
# =========================

def normalizar_cpf(cpf):
    """Garante CPF com 11 dígitos e zeros à esquerda"""
    return "".join(filter(str.isdigit, str(cpf))).zfill(11)


def carregar_excel():
    df = pd.read_excel(EXCEL_ORIGINAL, dtype=str)
    df = df.dropna(how="all").reset_index(drop=True)

    df["CPF"] = df["CPF"].apply(normalizar_cpf)

    df["Data de Nascimento"] = pd.to_datetime(
        df["Data de Nascimento"], errors="coerce"
    ).dt.strftime("%d/%m/%Y")

    return df


def iniciar_driver():
    options = Options()
    options.add_argument("--start-maximized")

    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=options
    )

    wait = WebDriverWait(driver, 20)
    driver.get(URL)

    return driver, wait


# =========================
# CONSULTA CPF
# =========================

def consultar_cpf(driver, wait, cpf, data_nasc_excel):
    resultado = {
        "CPF": cpf,
        "Data Nascimento Excel": data_nasc_excel,
        "Data Nascimento Site": None,
        "Nome": None,
        "Situação": None,
        "Resultado": None,
        "Mensagem": None,
        "Data/Hora de Processamento": None
    }

    try:
        # Preenche CPF
        input_cpf = wait.until(EC.presence_of_element_located((By.ID, "cpf")))
        input_cpf.clear()
        input_cpf.send_keys(cpf)

        # Consulta
        driver.find_element(By.CSS_SELECTOR, "button.btn").click()

        result_div = wait.until(EC.visibility_of_element_located((By.ID, "result")))
        time.sleep(1)

        # CPF inválido ou erro
        if "error" in result_div.get_attribute("class"):
            resultado["Resultado"] = "ERRO"
            resultado["Mensagem"] = result_div.text.strip()
        else:
            dados = {}
            rows = result_div.find_elements(By.CLASS_NAME, "data-row")

            for row in rows:
                label = (
                    row.find_element(By.CLASS_NAME, "data-label")
                    .text.replace(":", "")
                    .strip()
                    .lower()
                )
                value = row.find_element(By.CLASS_NAME, "data-value").text.strip()
                dados[label] = value

            data_site = dados.get("data de nascimento")
            sexo = dados.get("sexo")

            resultado["Data Nascimento Site"] = data_site
            resultado["Nome"] = dados.get("nome")
            resultado["Situação"] = dados.get("situação")

            # =========================
            # REGRAS DE NEGÓCIO
            # =========================

            if not data_site:
                resultado["Resultado"] = "ERRO"
                resultado["Mensagem"] = "Data de nascimento não retornada"

            elif data_site == data_nasc_excel:
                resultado["Resultado"] = "SUCESSO"
                resultado["Mensagem"] = "Data de nascimento confirmada"

            else:
                # Data divergente
                if not sexo:
                    resultado["Resultado"] = "VERIFICAR"
                    resultado["Mensagem"] = (
                        "Data divergente e campo Sexo não retornado. Verifique resultado."
                    )
                else:
                    resultado["Resultado"] = "ERRO"
                    resultado["Mensagem"] = "Data de nascimento divergente"

    except Exception as e:
        resultado["Resultado"] = "ERRO"
        resultado["Mensagem"] = f"Erro inesperado: {str(e)}"

    # Timestamp
    resultado["Data/Hora de Processamento"] = datetime.now().strftime(
        "%d/%m/%Y %H:%M:%S"
    )

    return resultado


# =========================
# SALVAMENTO INCREMENTAL
# =========================

def salvar_resultado_incremental(arquivo, resultado):
    df_novo = pd.DataFrame([resultado])

    if os.path.exists(arquivo):
        df_existente = pd.read_excel(arquivo, dtype=str)
        df_final = pd.concat([df_existente, df_novo], ignore_index=True)
    else:
        df_final = df_novo

    df_final.to_excel(arquivo, index=False)


# =========================
# PROCESSO PRINCIPAL
# =========================

def processar_cpfs():
    df = carregar_excel()

    data_hora = datetime.now().strftime("%Y%m%d_%H%M%S")
    arquivo_excel = os.path.join(
        PASTA_SAIDA,
        f"Resultado_Consulta_CPF_{data_hora}.xlsx"
    )

    driver, wait = iniciar_driver()

    try:
        total = len(df)

        for i, row in df.iterrows():
            cpf = row["CPF"]
            data_nasc = row["Data de Nascimento"]

            print(f"🔎 {i+1}/{total} - CPF {cpf}")

            resultado = consultar_cpf(driver, wait, cpf, data_nasc)

            salvar_resultado_incremental(arquivo_excel, resultado)

            if resultado["Resultado"] == "SUCESSO":
                print("SUCESSO\n")
            elif resultado["Resultado"] == "VERIFICAR":
                print("VERIFICAR\n")
            else:
                print(f"ERRO: {resultado['Mensagem']}\n")

            time.sleep(2)

    finally:
        driver.quit()

    print("✅ PROCESSO FINALIZADO")
    print(f"📁 Relatório gerado: {arquivo_excel}")


# =========================
# START
# =========================

if __name__ == "__main__":
    processar_cpfs()
