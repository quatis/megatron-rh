import pyautogui as bot
from datetime import datetime, date
import time
import pandas as pd
import webbrowser
import sys
import os
import toolkit as tk

BASE_DIR = tk.BASE_DIR
EXCEL_PATH = tk.EXCEL_PATH
LOG_PATH = tk.LOG_PATH
URL = tk.URL

df = pd.read_excel(EXCEL_PATH)
totalPositions = len(df)


changeDate = os.environ.get("CHANGE_DATE")
LANGUAGE = os.environ.get("LANGUAGE", "EN")

if os.path.exists(LOG_PATH):
    df_log = pd.read_excel(LOG_PATH)
else:
    df_log = pd.DataFrame(columns=["Position", "DataHora"])

if not changeDate:
    today = date.today()
    changeDate = today.strftime("%m/%d/%Y") if LANGUAGE == "EN" else today.strftime("%d/%m/%Y")

bot.PAUSE = 0.5
bot.FAILSAFE = True

start_time = time.time()
webbrowser.open(URL)
time.sleep(6)

bot.press('tab', presses=10)
contador = 0
# Loop principal de processamento de posições (mantendo Excel em memória)
for index in range(totalPositions):
    try:
        if len(df) < 1:
            print("Nenhuma posição restante para processar.")
            break

        row = df.iloc[0]  # sempre processa a primeira linha de dados
        positionStartTime = time.time()
        position = str(int(row["Position"]))

        # ------------------ AUTOMAÇÃO ------------------
        bot.write(position)
        time.sleep(3)

        bot.press('down')
        bot.press('enter')
        time.sleep(3)
        bot.press('tab', presses=11)
        time.sleep(1.5)
        bot.press('enter')
        tk.validarTexto("as of today", "Card")
        time.sleep(1)
        bot.press('tab')
        bot.press('enter')
        time.sleep(1.5)

        bot.press('enter')
        time.sleep(1)
        bot.press('tab')
        tk.validarTexto("insert new changes for position:", "confirmação")
        bot.press('enter')
        time.sleep(3)
        tk.validarTexto("planned occupation date", "edição")
        jobFamily = tk.checkJobFamily(tk.copiarTexto())
        time.sleep(1)
        bot.press('tab', presses=13 if jobFamily else 12)
        bot.write(pd.to_datetime(row["Planned Occupation Date"]).strftime("%m/%d/%Y"))
        time.sleep(0.1)
        bot.press('tab', presses=67)
        time.sleep(0.2)
        bot.press('enter')
        time.sleep(3)

        resultadoPopUps = tk.tratar_popups(position)
        time.sleep(0.5)
        if resultadoPopUps == "error":
            raise Exception(f"Erro ao verificar pop-ups na position: {position}")

        time.sleep(2)
        bot.press('enter')
        # ------------------ FIM AUTOMAÇÃO ------------------

        # ------------------ LOG SUCESSO ------------------
        positionEndTime = time.time()
        totalTime = positionEndTime - positionStartTime
        pminutos, psegundos = divmod(totalTime, 60)
        contador += 1
        print(f"Processada position n° {contador}/{totalPositions}: {position}")
        print(f"Tempo: {int(pminutos)} minutos e {int(psegundos)} segundos.")

        data_hora = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        df_log.loc[len(df_log)] = [position, data_hora]
        df_log.to_excel(LOG_PATH, index=False)

        # ------------------ REMOVER LINHA PROCESSADA ------------------
        df = df.iloc[1:].reset_index(drop=True)  # remove linha 2
        df.to_excel(EXCEL_PATH, index=False)

        time.sleep(3)
        bot.click(462, 221)

    except Exception as e:
        print(f"Erro ao processar posição: {position} -> {e}")
        data_hora = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

        # ------------------ LOG DE ERRO ------------------
        if os.path.exists(LOG_PATH):
            with pd.ExcelWriter(LOG_PATH, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
                try:
                    df_err = pd.read_excel(LOG_PATH, sheet_name='erros')
                except:
                    df_err = pd.DataFrame(columns=["Position", "Erro", "DataHora"])
                df_err.loc[len(df_err)] = [position, str(e), data_hora]
                df_err.to_excel(writer, index=False, sheet_name='erros')
        else:
            df_err = pd.DataFrame([[position, str(e), data_hora]], columns=["Position", "Erro", "DataHora"])
            with pd.ExcelWriter(LOG_PATH, engine='openpyxl') as writer:
                df_log.to_excel(writer, index=False, sheet_name='log')
                df_err.to_excel(writer, index=False, sheet_name='erros')

        # ------------------ REMOVER LINHA COM ERRO ------------------
        df = df.iloc[1:].reset_index(drop=True)
        df.to_excel(EXCEL_PATH, index=False)
        continue

# ------------------ TEMPO TOTAL ------------------
end_time = time.time()
total_time = end_time - start_time
minutos, segundos = divmod(total_time, 60)

print(f"{contador} posições processadas em {int(minutos)} minutos e {int(segundos)} segundos.")

