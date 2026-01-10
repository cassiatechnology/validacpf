import pyautogui
import pandas as pd
import time
from datetime import datetime
import os
import pyperclip
import re

# ================= CONFIGURAÇÕES =================

pyautogui.PAUSE = 0.4
pyautogui.FAILSAFE = True

URL_CONSULTA = "https://servicos.receita.fazenda.gov.br/Servicos/CPF/ConsultaSituacao/ConsultaPublica.asp"

EXCEL_ORIGINAL = r"C:/dev/BotValidaCpfReceita_2/cpfDataNascimento.xlsx"
PASTA_SAIDA = r"C:/dev/BotValidaCpfReceita_2/Relatorios"

POSICAO_CONTEUDO = (500, 400)

LIMITE_CONSULTAS = 400
TEMPO_ESPERA_LOTE = 300  # 5 minutos

# ================= FUNÇÕES =================

def reabrir_pagina_consulta():
    time.sleep(0.8)

    pyautogui.hotkey('ctrl', 'l')
    time.sleep(0.3)

    pyautogui.hotkey('ctrl', 'a')
    pyautogui.press('backspace')

    pyautogui.write(URL_CONSULTA)
    pyautogui.press('enter')

    time.sleep(3)


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


def save_resultados(resultados, arquivo_txt, arquivo_excel):
    if resultados:
        pd.DataFrame(resultados).to_excel(arquivo_excel, index=False)

    print("\n📁 Resultados atualizados:")
    print(f"TXT  → {arquivo_txt}")
    print(f"XLSX → {arquivo_excel}")


# ================= MAIN =================

def main():
    print("=== INSTRUÇÕES IMPORTANTES ===")
    print("1. Abra o navegador e acesse:")
    print(f"   {URL_CONSULTA}")
    print("2. MAXIMIZE o navegador")
    print("3. Zoom 100%")
    print("4. Cursor piscando no campo CPF")
    print("5. NÃO use mouse/teclado durante execução")
    print("----------------------------------------")

    input("Pressione ENTER para começar...")

    try:
        df = pd.read_excel(EXCEL_ORIGINAL, dtype={'CPF': str, 'Data de Nascimento': str})
    except Exception as e:
        print("Erro ao ler Excel:", e)
        return

    df['Data de Nascimento'] = df['Data de Nascimento'].apply(
        lambda x: pd.to_datetime(x, errors='coerce').strftime('%d/%m/%Y')
        if pd.notnull(x) else ''
    )

    df = df.dropna(how='all').reset_index(drop=True)

    data_hora = datetime.now().strftime("%Y%m%d_%H%M%S")
    arquivo_txt = os.path.join(PASTA_SAIDA, f"Resultado_Consulta_CPF_{data_hora}.txt")
    arquivo_excel = os.path.join(PASTA_SAIDA, f"Resultado_Consulta_CPF_{data_hora}.xlsx")

    resultados = []

    with open(arquivo_txt, 'w', encoding='utf-8') as log:
        log.write("Linha | CPF | Data Nasc | Nome | Situação | Status | Observação\n")
        log.write("-" * 100 + "\n")

        for idx, row in df.iterrows():
            cpf = str(row['CPF']).strip().zfill(11)
            data_nasc = str(row['Data de Nascimento']).strip()

            if not cpf or not data_nasc:
                continue

            print(f"\n🔎 {idx + 1}/{len(df)} | CPF: {cpf}")

            try:
                pyautogui.hotkey('ctrl', 'a')
                pyautogui.press('backspace')
                pyautogui.write(cpf)

                pyautogui.press('tab')

                pyautogui.hotkey('ctrl', 'a')
                pyautogui.press('backspace')
                pyautogui.write(data_nasc)

                pyautogui.press('tab')
                pyautogui.press('space')

                print("⏳ Resolva o captcha (3s)...")
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

                linha = (
                    f"{idx + 1} | {cpf} | {data_nasc} | "
                    f"{dados['nome']} | {dados['situacao_cadastral']} | "
                    f"{dados['status']} | {dados['mensagem']}"
                )

                print(linha)
                log.write(linha + "\n")

                resultados.append({
                    'CPF': cpf,
                    'Data de Nascimento': data_nasc,
                    'Nome': dados['nome'],
                    'Situação Cadastral': dados['situacao_cadastral'],
                    'Status': dados['status'],
                    'Observação': dados['mensagem']
                })

                save_resultados(resultados, arquivo_txt, arquivo_excel)

                # ⏸️ Pausa a cada 400 consultas
                if (idx + 1) % LIMITE_CONSULTAS == 0:
                    print(f"\n⏸️ {LIMITE_CONSULTAS} consultas realizadas.")
                    print("⏳ Aguardando 5 minutos...")
                    time.sleep(TEMPO_ESPERA_LOTE)
                    print("▶️ Retomando consultas...")
                    reabrir_pagina_consulta()
                else:
                    reabrir_pagina_consulta()

            except KeyboardInterrupt:
                print("\n⛔ Execução interrompida.")
                break

            except Exception as e:
                print(f"Erro na linha {idx + 1}: {e}")
                reabrir_pagina_consulta()
                continue

    print("\n✅ Processo finalizado.")


if __name__ == "__main__":
    main()
