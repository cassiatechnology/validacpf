import pyautogui
import pandas as pd
import time
from datetime import datetime
import os
import pyperclip
import re
import subprocess
import random

# ================= CONFIGURAÇÕES =================

pyautogui.PAUSE = 0.4
pyautogui.FAILSAFE = True

URL_CONSULTA = "https://servicos.receita.fazenda.gov.br/servicos/cpf/consultasituacao/consultapublica.asp"

EXCEL_ORIGINAL = r"C:/dev/BotValidaCpfReceita_2/cpfDataNascimento.xlsx"
PASTA_SAIDA = r"C:/dev/BotValidaCpfReceita_2/Relatorios"

POSICAO_CONTEUDO = (500, 400)

INTERVALO_REINICIO = 6   # reinicia o Chrome
INTERVALO_MOUSE = 3       # move mouse sutilmente

# ================= NAVEGADOR =================

def abrir_chrome():
    chrome_path = os.environ.get("CHROME_PATH", "chrome")

    subprocess.Popen([
        chrome_path,
        "--incognito",
        "--start-maximized",
        URL_CONSULTA
    ])

    time.sleep(8)


def reiniciar_chrome():
    subprocess.run(
        ["taskkill", "/F", "/IM", "chrome.exe"],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL
    )

    time.sleep(3)
    abrir_chrome()


def reabrir_pagina_consulta():
    time.sleep(0.8)

    pyautogui.hotkey('ctrl', 'l')
    time.sleep(0.3)

    pyautogui.hotkey('ctrl', 'a')
    pyautogui.press('backspace')

    pyautogui.write(URL_CONSULTA)
    pyautogui.press('enter')

    time.sleep(3)

# ================= HUMANIZAÇÃO =================

def mover_mouse_sutil():
    x, y = pyautogui.position()

    desloc_x = random.randint(-8, 8)
    desloc_y = random.randint(-8, 8)

    pyautogui.moveTo(
        x + desloc_x,
        y + desloc_y,
        duration=random.uniform(0.1, 0.25)
    )

# ================= EXTRAÇÃO =================

def extrair_dados_texto(texto_pagina: str) -> dict:
    resultado = {
        'nome': '',
        'situacao_cadastral': '',
        'status': 'ERRO',
        'mensagem': 'Situação Cadastral não encontrada'
    }

    match_nome = re.search(r'Nome:\s*(.+)', texto_pagina, re.IGNORECASE)
    if match_nome:
        resultado['nome'] = match_nome.group(1).strip()

    match_situacao = re.search(r'Situação Cadastral:\s*(.+)', texto_pagina, re.IGNORECASE)
    if match_situacao:
        resultado['situacao_cadastral'] = match_situacao.group(1).strip()

    if resultado['nome'] and resultado['situacao_cadastral']:
        resultado['status'] = 'SUCESSO'
        resultado['mensagem'] = 'Extração realizada com sucesso'

    return resultado


def save_resultados(resultados, arquivo_excel):
    if resultados:
        pd.DataFrame(resultados).to_excel(arquivo_excel, index=False)

# ================= MAIN =================

def main():
    print("=== INSTRUÇÕES IMPORTANTES ===")
    print("• Não use mouse ou teclado durante a execução")
    print("• O Chrome será aberto automaticamente")
    print("----------------------------------------")

    input("Pressione ENTER para iniciar...")

    abrir_chrome()

    df = pd.read_excel(EXCEL_ORIGINAL, dtype={'CPF': str, 'Data de Nascimento': str})

    df['Data de Nascimento'] = df['Data de Nascimento'].apply(
        lambda x: pd.to_datetime(x, errors='coerce').strftime('%d/%m/%Y')
        if pd.notnull(x) else ''
    )

    df = df.dropna(how='all').reset_index(drop=True)

    data_hora = datetime.now().strftime("%Y%m%d_%H%M%S")
    arquivo_excel = os.path.join(PASTA_SAIDA, f"Resultado_Consulta_CPF_{data_hora}.xlsx")

    resultados = []
    contador = 0

    for idx, row in df.iterrows():
        contador += 1

        cpf = str(row['CPF']).strip().zfill(11)
        data_nasc = str(row['Data de Nascimento']).strip()

        if not cpf or not data_nasc:
            continue

        print(f"🔎 {contador}/{len(df)} | CPF: {cpf}")

        try:
            if contador % INTERVALO_MOUSE == 0:
                print("🖱️ Movimento sutil do mouse")
                mover_mouse_sutil()

            pyautogui.hotkey('ctrl', 'a')
            pyautogui.press('backspace')
            pyautogui.write(cpf)

            pyautogui.press('tab')

            pyautogui.hotkey('ctrl', 'a')
            pyautogui.press('backspace')
            pyautogui.write(data_nasc)

            pyautogui.press('tab')
            pyautogui.press('space')

            time.sleep(3)

            for _ in range(5):
                pyautogui.press('tab')

            pyautogui.press('space')
            time.sleep(3)

            pyautogui.click(POSICAO_CONTEUDO)
            pyautogui.hotkey('ctrl', 'a')
            pyautogui.hotkey('ctrl', 'c')
            time.sleep(0.5)

            texto = pyperclip.paste()
            dados = extrair_dados_texto(texto)

            resultados.append({
                'CPF': cpf,
                'Data de Nascimento': data_nasc,
                'Nome': dados['nome'],
                'Situação Cadastral': dados['situacao_cadastral'],
                'Status': dados['status'],
                'Observação': dados['mensagem']
            })

            save_resultados(resultados, arquivo_excel)

            if contador % INTERVALO_REINICIO == 0:
                print("♻️ Reiniciando Chrome (modo anônimo)...")
                reiniciar_chrome()
            else:
                reabrir_pagina_consulta()

        except KeyboardInterrupt:
            print("⛔ Execução interrompida pelo usuário.")
            break

        except Exception as e:
            print(f"Erro: {e}")
            reabrir_pagina_consulta()

    print("✅ Processo finalizado.")


if __name__ == "__main__":
    main()