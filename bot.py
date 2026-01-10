import pyautogui
import pandas as pd
import pyperclip
import subprocess
import time
import random
import os
import sys
from datetime import datetime

# ======================================================
# CONFIGURAÇÕES
# ======================================================

URL_CONSULTA = (
    "https://servicos.receita.fazenda.gov.br/"
    "servicos/cpf/consultasituacao/consultapublica.asp"
)

EXCEL_ORIGINAL = r"C:/dev/BotValidaCpfReceita_2/cpfDataNascimento.xlsx"
PASTA_SAIDA = r"C:/dev/BotValidaCpfReceita_2/Relatorios"

pyautogui.PAUSE = 0.4
pyautogui.FAILSAFE = True

# ======================================================
# TEMPOS
# ======================================================

TEMPO_ABERTURA_CHROME = 6
TEMPO_APOS_CHECKBOX = 2
TEMPO_APOS_CONSULTAR = 4
TEMPO_VOLTA = 3

# ======================================================
# LOTES CURTOS
# ======================================================

LOTE_CURTO_QTD = 3
LOTE_CURTO_MIN = 2.5
LOTE_CURTO_MAX = 6.0

# ======================================================
# MOUSE SUTIL
# ======================================================

MOVER_MOUSE_A_CADA = 2
MOUSE_VARIACAO = 15
MOUSE_DURACAO = 0.3

# ======================================================
# FUNÇÕES BASE
# ======================================================

def abrir_chrome():
    print("🌐 Abrindo Chrome...")
    subprocess.Popen(["chrome", URL_CONSULTA])
    time.sleep(TEMPO_ABERTURA_CHROME)


def copiar_texto():
    pyautogui.hotkey("ctrl", "a")
    time.sleep(0.2)
    pyautogui.hotkey("ctrl", "c")
    time.sleep(0.4)
    return pyperclip.paste() or ""


def mover_mouse_sutil():
    x, y = pyautogui.position()
    pyautogui.moveTo(
        x + random.randint(-MOUSE_VARIACAO, MOUSE_VARIACAO),
        y + random.randint(-MOUSE_VARIACAO, MOUSE_VARIACAO),
        duration=MOUSE_DURACAO
    )


def pausa_lote_curto():
    t = random.uniform(LOTE_CURTO_MIN, LOTE_CURTO_MAX)
    print(f"⏳ Pausa de {t:.1f}s (lote curto)")
    time.sleep(t)


# ======================================================
# CAPTCHA
# ======================================================

def texto_praticamente_igual(t1: str, t2: str) -> bool:
    return abs(len(t1.strip()) - len(t2.strip())) < 50


def aguardar_resolucao_captcha_checkbox(texto_base):
    print("\n🚨 CAPTCHA DE IMAGEM DETECTADO")
    print("👉 Resolva o captcha manualmente")
    print("👉 Depois clique em CONSULTAR")
    print("⏸️ Aguardando mudança da página...\n")

    while True:
        time.sleep(2)
        texto_atual = copiar_texto()
        if not texto_praticamente_igual(texto_base, texto_atual):
            print("▶️ Página mudou, seguindo fluxo\n")
            return


# ======================================================
# EXTRAÇÃO
# ======================================================

def extrair_dados(texto: str) -> dict:
    texto = texto.upper()

    resultado = {
        "nome": "",
        "situacao": "",
        "status": "ERRO",
        "mensagem": "Situação Cadastral não encontrada"
    }

    for linha in texto.splitlines():
        if linha.startswith("NOME"):
            resultado["nome"] = linha.split(":", 1)[-1].strip()
        elif linha.startswith("SITUAÇÃO CADASTRAL"):
            resultado["situacao"] = linha.split(":", 1)[-1].strip()

    if resultado["situacao"]:
        resultado["status"] = "SUCESSO"
        resultado["mensagem"] = "Consulta realizada com sucesso"

    return resultado


# ======================================================
# AÇÃO PRINCIPAL
# ======================================================

def preencher_e_consultar(cpf, nascimento):
    # CPF
    pyautogui.click(600, 580)
    pyautogui.hotkey("ctrl", "a")
    pyautogui.press("backspace")
    pyautogui.write(cpf, interval=0.05)

    pyautogui.press("tab")

    # Data
    pyautogui.hotkey("ctrl", "a")
    pyautogui.press("backspace")
    pyautogui.write(nascimento, interval=0.05)

    pyautogui.press("tab")


# ======================================================
# MAIN
# ======================================================

def main():
    abrir_chrome()

    df = pd.read_excel(EXCEL_ORIGINAL, dtype=str)
    df = df.dropna(how="all").reset_index(drop=True)

    df["Data de Nascimento"] = pd.to_datetime(
        df["Data de Nascimento"], errors="coerce"
    ).dt.strftime("%d/%m/%Y")

    data_hora = datetime.now().strftime("%Y%m%d_%H%M%S")
    arquivo_excel = os.path.join(
        PASTA_SAIDA,
        f"Resultado_Consulta_CPF_{data_hora}.xlsx"
    )

    resultados = []

    for idx, row in df.iterrows():
        cpf = row["CPF"].zfill(11)
        nascimento = row["Data de Nascimento"]

        print(f"🔎 {idx+1}/{len(df)} | CPF {cpf} | Data de Nascimento {nascimento}")

        preencher_e_consultar(cpf, nascimento)

        # ===== CHECKBOX =====
        texto_antes_checkbox = copiar_texto()

        pyautogui.press("space")
        time.sleep(TEMPO_APOS_CHECKBOX)

        texto_depois_checkbox = copiar_texto()

        # 🚨 CAPTCHA DISPARADO NO CHECKBOX
        if texto_praticamente_igual(texto_antes_checkbox, texto_depois_checkbox):
            aguardar_resolucao_captcha_checkbox(texto_antes_checkbox)

        # ===== CONSULTAR =====
        for _ in range(5):
            pyautogui.press("tab")
        pyautogui.press("space")

        time.sleep(TEMPO_APOS_CONSULTAR)

        texto_resultado = copiar_texto()
        dados = extrair_dados(texto_resultado)

        resultados.append({
            "CPF": cpf,
            "Data de Nascimento": nascimento,
            "Nome": dados["nome"],
            "Situação Cadastral": dados["situacao"],
            "Status": dados["status"],
            "Observação": dados["mensagem"]
        })

        pd.DataFrame(resultados).to_excel(arquivo_excel, index=False)

        pyautogui.hotkey("alt", "left")
        time.sleep(TEMPO_VOLTA)

        if (idx + 1) % MOVER_MOUSE_A_CADA == 0:
            mover_mouse_sutil()

        if (idx + 1) % LOTE_CURTO_QTD == 0:
            pausa_lote_curto()

    print("\n✅ Finalizado com sucesso")
    print(f"📁 Arquivo gerado: {arquivo_excel}")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n⛔ Execução interrompida")
        sys.exit(0)
