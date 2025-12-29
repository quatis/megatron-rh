from datetime import date
import pyautogui as bot
import time
import pandas as pd
import webbrowser
import sys
import os
import pyperclip
import re
import toolkit as tk

BASE_DIR = tk.BASE_DIR
EXCEL_PATH = tk.EXCEL_PATH
LOG_PATH = tk.LOG_PATH
URL = tk.URL

df = pd.read_excel(EXCEL_PATH)
totalPositions = len(df)

changeDate = os.environ.get("CHANGE_DATE")
LANGUAGE = os.environ.get("LANGUAGE", "EN")
if not changeDate:
    today = date.today()
    changeDate = today.strftime("%m/%d/%Y") if LANGUAGE == "EN" else today.strftime("%d/%m/%Y")

bot.PAUSE = 0.5
bot.FAILSAFE = True

start_time = time.time()
webbrowser.open(URL)
time.sleep(6)

df = pd.read_excel(EXCEL_PATH)
bot.press('tab', presses=10)
contador = 0

for index, row in df.iterrows():
    positionStartTime = time.time()
    position = str(int(row["Position"]))
    bot.write(position)
    time.sleep(3)

    bot.press('down')
    bot.press('enter')
    time.sleep(3)
    bot.press('tab', presses=11)
    bot.press('enter')
    time.sleep(3)
    if not tk.validarTexto("as of today"):
        print("Card não encontrado após 3 tentativas - encerrando.")
        sys.exit(1)
    print("Card OK")
    time.sleep(0.3)
    bot.press('tab')
    bot.press('enter')
    time.sleep(1.5)

    bot.press('enter')
    time.sleep(1)
    #bot.hotkey("ctrl", "a")
    #bot.write(changeDate)
    #time.sleep(1.5)
    bot.press('tab')
    tk.validarTexto("insert new changes for position:", "confirmação")
    bot.press('enter')
    time.sleep(3)
    if not tk.validarTexto("planned occupation date"):
        print("Tela de edicao não encontrada após 3 tentativas - encerrando.")
        sys.exit(1)
    print("Tela de edicao OK")
    jobFamily = tk.checkJobFamily(tk.copiarTexto())
    bot.press('tab', presses=18 if jobFamily else 17)
    bot.press('enter')
    bot.press('down', presses=2)
    bot.press('enter')
    bot.press('tab', presses=62)
    bot.press('enter')
    time.sleep(0.2)
    bot.press('enter')
    time.sleep(3)


    resultadoPopUps = tk.tratar_popups(position)

    if resultadoPopUps == "error":
        print(f"Erro ao verificar pop-ups na position: {position}. Encerrando script.")
        sys.exit(1)


    time.sleep(2)
    bot.press('enter')

    positionEndTime = time.time()
    totalTime = positionEndTime - positionStartTime
    pminutos, psegundos = divmod(totalTime, 60)
    contador += 1
    print(f"Processada position n° {contador}/{totalPositions}: {position}")
    print(f"Tempo: {int(pminutos)} minutos e {int(psegundos)} segundos.")

    time.sleep(3)
    bot.click(462, 221)

end_time = time.time()
total_time = end_time - start_time
minutos, segundos = divmod(total_time, 60)

print(f"{contador} posições processadas em {int(minutos)} minutos e {int(segundos)} segundos.")