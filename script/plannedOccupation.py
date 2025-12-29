import pyautogui as bot
from datetime import datetime, date
import time
import pandas as pd
import webbrowser
import sys
import os
import pyperclip
import re
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(BASE_DIR, "..","Config.xlsx")
LOG_PATH = os.path.join(BASE_DIR, "..","log.xlsx")
URL = "https://performancemanager5.successfactors.eu/sf/orgchart?type=position&bplte_company=StraumannPROD"
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

def copiarTexto(delay=0.5):
    time.sleep(delay)
    bot.hotkey("ctrl", "a")
    time.sleep(0.2)
    bot.hotkey("ctrl", "c")
    time.sleep(delay)
    return pyperclip.paste().replace("\n", " ").replace("\r", " ").replace("\t", " ").lower()

def checkJobFamily(text: str) -> bool:
    match = re.search(r"standard job title\s+(.*?)\s+work level", text)
    return bool(match and match.group(1).strip())

def tratar_popups(position, max_retries=3):
    for tentativa in range(max_retries):
        try:
            texto = copiarTexto()

            # Primeiro pop-up: Changes were propagated
            if ("changes were propagated forward") in texto:
                bot.press('esc')
                print(f"Pop-up de mudança futura detectado e fechado na position: {position}")
                time.sleep(2)

                # Verificar se apareceu o segundo pop-up (STI Matrix)
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

            # Apenas pop-up STI Matrix
            elif 'please select the correct sti matrix. click "cancel" to continue editing, "no" to exit, or "yes" to save the record.' in texto:
                print(f"Pop-up de STI tratado na position: {position}")
                bot.press('tab', presses=2)
                bot.press('enter')
                return "stiOnly"

            #Nenhum pop-up
            else:
                print(f"Nenhum pop-up encontrado na position: {position}")
                return "noPopup"

        except Exception as e:
            print(f"Erro ao tratar pop-ups (tentativa {tentativa + 1}/{max_retries}): {e}")

        if tentativa < max_retries - 1:
            time.sleep(1)
            print(f"Falha ao processar pop-ups após {max_retries} tentativas")
            return "error"
def validarTexto(expected, max_retries=3, delay=1):
    for tentativa in range(max_retries):
        texto = copiarTexto()
        if expected.lower() in texto:
            return True
        time.sleep(delay)
    return False

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
    phasing = str(row["Planned Occupation Date"])

    bot.write(position)
    time.sleep(3)

    bot.press('down')
    bot.press('enter')
    time.sleep(3)
    bot.press('tab', presses=11)
    bot.press('enter')
    time.sleep(2)
    if not validarTexto("as of today"):
        print("Card não encontrado após 3 tentativas - encerrando.")
        sys.exit(1)
    print("Card OK")
    time.sleep(0.5)
    bot.press('tab')
    bot.press('enter')
    time.sleep(1.5)
    bot.press('enter')
    bot.sleep(1)
    bot.press('enter')
    bot.sleep(0.5)
    #bot.hotkey("ctrl", "a")
    #bot.write(changeDate)
    #time.sleep(1.5)
    bot.press('tab')
    if not validarTexto("insert new changes for position:"):
        print("Tela de confirmacao não encontrada após 3 tentativas - encerrando.")
        sys.exit(1)
    print("Tela de confirmacao OK")
    bot.press('enter')
    time.sleep(3)
    if not validarTexto("planned occupation date"):
        print("Tela de edicao não encontrada após 3 tentativas - encerrando.")
        sys.exit(1)
    print("Tela de edicao OK")
    jobFamily = checkJobFamily(copiarTexto())
    time.sleep(1)
    bot.press('tab', presses=13 if jobFamily else 12)
    bot.write(pd.to_datetime(row["Planned Occupation Date"]).strftime("%m/%d/%Y"))
    time.sleep(0.1)
    bot.press('tab', presses=67)
    time.sleep(0.2)
    bot.press('enter')
    time.sleep(4)
    resultadoPopUps = tratar_popups(position)
    time.sleep(0.5)
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
    data_hora = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    df_log.loc[len(df_log)] = [position, data_hora]
    df_log.to_excel(LOG_PATH, index=False)

    time.sleep(3)
    bot.click(462, 221)

end_time = time.time()
total_time = end_time - start_time
minutos, segundos = divmod(total_time, 60)

print(f"{contador} posições processadas em {int(minutos)} minutos e {int(segundos)} segundos.")