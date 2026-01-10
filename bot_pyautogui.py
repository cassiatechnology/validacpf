import pyautogui
import pandas as pd
import time
from datetime import datetime
import os

# Configurações importantes
pyautogui.PAUSE = 0.4          # tempo entre ações - ajuste se necessário
pyautogui.FAILSAFE = True      # mover mouse pro canto superior-esquerdo = emergência

# Caminhos
EXCEL_ORIGINAL = r"C:/dev/BotValidaCpfReceita_2/cpfDataNascimento.xlsx"
PASTA_SAIDA = r"C:/dev/BotValidaCpfReceita_2/Relatorios"

def main():
    print("=== INSTRUÇÕES IMPORTANTES ===")
    print("1. Abra o navegador e acesse:")
    print("   https://servicos.receita.fazenda.gov.br/Servicos/CPF/ConsultaSituacao/ConsultaPublica.asp")
    print("2. MAXIMIZE a janela do navegador")
    print("3. Zoom 100%")
    print("4. Certifique-se que o cursor já está piscando no campo CPF")
    print("5. NÃO clique em nada depois disso")
    print("6. Tenha o Excel com CPF e Data de Nascimento já aberto")
    print("----------------------------------------")
    print("Quando estiver pronto → pressione ENTER aqui...\n")

    input("Pressione ENTER para começar...")

    # Carrega os dados
    try:
        df = pd.read_excel(EXCEL_ORIGINAL, dtype={'CPF': str, 'Data de Nascimento': str})
    except Exception as e:
        print("Erro ao ler o excel:", e)
        return

    # Tratamento básico da data
    df['Data de Nascimento'] = df['Data de Nascimento'].apply(
        lambda x: pd.to_datetime(x, errors='coerce').strftime('%d/%m/%Y') 
        if pd.notnull(x) else ''
    )

    # Remove linhas completamente vazias no final
    df = df.dropna(how='all').reset_index(drop=True)

    data_hora = datetime.now().strftime("%Y%m%d_%H%M%S")
    arquivo_txt = os.path.join(PASTA_SAIDA, f"Resultado_Consulta_CPF_{data_hora}.txt")
    arquivo_excel = os.path.join(PASTA_SAIDA, f"Resultado_Consulta_CPF_{data_hora}.xlsx")

    resultados = []

    with open(arquivo_txt, 'w', encoding='utf-8') as log:
        log.write("Linha | CPF | Data Nasc | Status | Observação\n")
        log.write("-"*70 + "\n")

        for idx, row in df.iterrows():
            cpf = str(row['CPF']).strip().zfill(11)
            data_nasc = str(row['Data de Nascimento']).strip()

            if not cpf or not data_nasc:
                print(f"Pulando linha {idx+1} → CPF ou data vazia")
                continue

            print(f"\n→ Processando {idx+1}/{len(df)}  |  CPF: {cpf}")

            try:
                # 1. Colar CPF (já estamos no campo)
                pyautogui.hotkey('ctrl', 'a')     # seleciona tudo
                pyautogui.press('backspace')      # limpa
                pyautogui.write(cpf)

                # 2. Tab uma vez → vai pro campo data
                pyautogui.press('tab')

                # 3. Colar data de nascimento
                pyautogui.hotkey('ctrl', 'a')
                pyautogui.press('backspace')
                pyautogui.write(data_nasc)

                # 4. Tab uma vez → vai pro checkbox "Sou humano"
                pyautogui.press('tab')

                # 5. Espaço → marca o checkbox
                pyautogui.press('space')

                # 6. Aguarda 3 segundos (tempo para resolver o captcha visualmente)
                print("   → 3 segundos para resolver captcha manualmente...")
                time.sleep(3.0)

                # 7. Tab 5 vezes até chegar no botão "Consultar"
                for _ in range(5):
                    pyautogui.press('tab')

                time.sleep(0.8)   # pequena pausa antes de confirmar

                # 8. Espaço → aciona o botão Consultar
                pyautogui.press('space')

                # Tempo de espera maior para carregar o resultado
                time.sleep(5.5)

                # Aqui você pode adicionar alguma lógica de captura de resultado
                # (por enquanto só registra que enviou)
                status = "Enviado / aguardando análise manual"
                observacao = "Resultado deve ser conferido manualmente"

                linha_log = f"{idx+1} | {cpf} | {data_nasc} | {status} | {observacao}"
                print(linha_log)
                log.write(linha_log + "\n")

                resultados.append({
                    'CPF': cpf,
                    'Data de Nascimento': data_nasc,
                    'Status': status,
                    'Observação': observacao
                })

                # Volta para nova consulta (clique no link "Voltar" ou "Nova Consulta")
                # Normalmente após o resultado aparece um link "Nova consulta"
                # Você pode tentar automatizar esse passo também com mais tabs + espaço
                print("   → Pressione ENTER quando voltar para a tela inicial...")
                input("   (ou Ctrl+C para interromper) ")

            except KeyboardInterrupt:
                print("\nProcesso interrompido pelo usuário.")
                break
            except Exception as e:
                print(f"Erro na linha {idx+1}: {e}")
                continue

    # Salvando resultados
    if resultados:
        pd.DataFrame(resultados).to_excel(arquivo_excel, index=False)
        print(f"\nResultados salvos em:")
        print(f"   TXT → {arquivo_txt}")
        print(f"   XLSX → {arquivo_excel}")
    else:
        print("\nNenhum resultado foi processado.")

    print("\nFinalizado.")

if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        print("Erro fatal:", e)