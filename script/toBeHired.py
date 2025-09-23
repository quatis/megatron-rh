from datetime import date
import pyautogui as bot
import time
import pandas as pd
import webbrowser
import sys
import os
import pyperclip
import re

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(BASE_DIR, "..","Config.xlsx")
URL = "https://performancemanager5.successfactors.eu/sf/orgchart?type=position&bplte_company=StraumannPROD"
df = pd.read_excel(EXCEL_PATH)
totalPositions = len(df)
changeDate = os.environ.get("CHANGE_DATE")
LANGUAGE = os.environ.get("LANGUAGE", "EN")
if not changeDate:
    today = date.today()
    changeDate = today.strftime("%m/%d/%Y") if LANGUAGE == "EN" else today.strftime("%d/%m/%Y")

def checkJobFamily(text: str) -> bool:
    match = re.search(r"standard job title\s+(.*?)\s+work level", text)
    return bool(match and match.group(1).strip())

def setChangeDate(changeDate):
    changeDate = changeDate
def copiarTexto(delay=0.5):
    bot.hotkey("ctrl", "a")
    time.sleep(0.2)
    bot.hotkey("ctrl", "c")
    time.sleep(delay)
    return pyperclip.paste().replace("\n", " ").replace("\r", " ").replace("\t", " ").lower()

def tratar_popups(position, max_retries=3):
    for tentativa in range(max_retries):
        try:
            texto = copiarTexto()

            if "changes were propagated forward to future records." in texto:
                bot.press('esc')
                print(f"Pop-up de mudança futura detectado e fechado na position: {position}")
                time.sleep(2)

                try:
                    textosub = copiarTexto()
                    if 'please select the correct sti matrix. click "cancel" to continue editing, "no" to exit, or "yes" to save the record.' in textosub:
                        print(f"Pop-up de STI detectado após o de mudança futura na position: {position}")
                        bot.press('tab', presses=2)
                        bot.press('enter')
                        return "bothPopups"
                    else:
                        return "futureOnly"
                except Exception as e:
                    print(f"Erro ao verificar segundo pop-up: {e}")
                    return "futureError"

            elif 'please select the correct sti matrix. click "cancel" to continue editing, "no" to exit, or "yes" to save the record.' in texto:
                print(f"Pop-up de STI tratado na position: {position}")
                bot.press('tab', presses=2)
                bot.press('enter')
                return "stiOnly"

            else:
                print(f"Nenhum pop-up encontrado na position: {position}")
                return "noPopup"

        except Exception as e:
            print(f"Erro ao tratar pop-ups (tentativa {tentativa + 1}/{max_retries}): {e}")

        if tentativa < max_retries - 1:
            time.sleep(1)
            print(f"Falha ao processar pop-ups após {max_retries} tentativas")
            return "error"

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
    texto = copiarTexto()
    if "as of today" in texto:
        print("Card OK")
    else:
        print("Card não encontrado - encerrando.")
        sys.exit(1)
    time.sleep(0.2)
    bot.press('tab')
    bot.press('enter')
    time.sleep(3)

    bot.press('enter')
    time.sleep(2)
    bot.hotkey("ctrl", "a")
    bot.write(changeDate)
    bot.sleep(1)
    bot.press('tab')
    texto = copiarTexto()
    if "insert new changes for position:" in texto:
        print("Tela de confirmacao OK")
    else:
        print("Texto não encontrado - confirmacao, encerrando.")
        sys.exit(1)
    bot.press('enter')
    time.sleep(3)
    texto = copiarTexto()
    if "planned occupation date" in texto:
        print("Tela de edicao OK")
    else:
        print("Texto nao encontrado - edicao, encerrando.")
        sys.exit(1)
    jobFamily = checkJobFamily(texto)
    time.sleep(1)
    bot.press('tab', presses=18 if jobFamily else 17)
    bot.press('enter')
    bot.press('down', presses=2)
    bot.press('enter')
    bot.press('tab', presses=62)
    bot.press('enter')
    texto = copiarTexto()
    if "updated by" not in texto:
        print("Final não encontrado - encerrando.")
        sys.exit(1)
    time.sleep(0.2)
    bot.press('enter')
    time.sleep(3)


    resultadoPopUps = tratar_popups(position)

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