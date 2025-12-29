import os
import time
import pyautogui as bot
import pyperclip
import re
import sys

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(BASE_DIR, "..","Config.xlsx")
LOG_PATH = os.path.join(BASE_DIR, "..","log.xlsx")
URL = "https://performancemanager5.successfactors.eu/sf/orgchart?type=position&bplte_company=StraumannPROD"


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
def validarTexto(expected, mensagem="", max_retries=3, delay=1):
    for tentativa in range(max_retries):
        texto = copiarTexto()
        if expected.lower() in texto:
            if mensagem:
                print(f"{mensagem} OK")
            return True
        time.sleep(delay)
    if mensagem:
        print(f"{mensagem} não encontrado após {max_retries} tentativas - encerrando.")
    else:
        print(f"Texto '{expected}' não encontrado após {max_retries} tentativas - encerrando.")
    sys.exit(1)