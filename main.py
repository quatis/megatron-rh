import subprocess
import os
import sys
from datetime import datetime, date
import time

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

LANGUAGE = "EN"
changeDate = date.today()


class Colors:
    """style"""
    RESET = '\033[0m'
    BOLD = '\033[1m'
    DIM = '\033[2m'

    # Cores básicas
    BLACK = '\033[30m'
    RED = '\033[31m'
    GREEN = '\033[32m'
    YELLOW = '\033[33m'
    BLUE = '\033[34m'
    MAGENTA = '\033[35m'
    CYAN = '\033[36m'
    WHITE = '\033[37m'

    # Cores brilhantes
    BRIGHT_BLACK = '\033[90m'
    BRIGHT_RED = '\033[91m'
    BRIGHT_GREEN = '\033[92m'
    BRIGHT_YELLOW = '\033[93m'
    BRIGHT_BLUE = '\033[94m'
    BRIGHT_MAGENTA = '\033[95m'
    BRIGHT_CYAN = '\033[96m'
    BRIGHT_WHITE = '\033[97m'

    # Backgrounds
    BG_BLACK = '\033[40m'
    BG_RED = '\033[41m'
    BG_GREEN = '\033[42m'
    BG_YELLOW = '\033[43m'
    BG_BLUE = '\033[44m'
    BG_MAGENTA = '\033[45m'
    BG_CYAN = '\033[46m'
    BG_WHITE = '\033[47m'



def getChangeDate():
    """data formatada conforme o idioma do EC"""
    global changeDate, LANGUAGE
    if LANGUAGE == "PT":
        return changeDate.strftime("%d/%m/%Y")
    else:  # EN
        return changeDate.strftime("%m/%d/%Y")


def clear_screen():
    """Limpa a tela do console"""
    os.system('cls' if os.name == 'nt' else 'clear')


def print_header():
    """cabeçalho da aplicação"""
    clear_screen()
    print(f"{Colors.CYAN}{Colors.BOLD}")
    print("╔" + "═" * 60 + "╗")
    print("║" + " " * 60 + "║")
    print("║" + f"{'🤖 ROBOZINHO AUTOMATION SUITE 🤖':^58}" + "║")
    print("║" + " " * 60 + "║")
    print("║" + f"{'SuccessFactors  R O B O T I Z A D O ':^60}" + "║")
    print("║" + " " * 60 + "║")
    print("║" + f"{'      👾  M E G A T R O N   D O   R H  👾    ':^58}" + "║")
    print("║" + " " * 60 + "║")
    print("╚" + "═" * 60 + "╝")
    print(f"{Colors.RESET}")


def print_menu_options():
    """opções do menu principal"""
    global changeDate
    print(f"\n{Colors.BRIGHT_WHITE}{Colors.BOLD}📋 OPERAÇÕES DISPONÍVEIS:{Colors.RESET}")
    print(
        f"{Colors.BRIGHT_CYAN}┌───────────────────────────────────────────────────────────────────────┐{Colors.RESET}")

    options = [
        ("1", "📅 Planned Occupation Date", "Atualiza datas de ocupação planejada", Colors.BRIGHT_GREEN),
        ("2", "👤 To Be Hired", "Ativa o campo To Be Hired no EC     ", Colors.BRIGHT_YELLOW),
        ("3", "🔮 Data da Mudança", "Alterar data efetiva da mudança     ", Colors.BRIGHT_BLUE),
        ("4", "📁 Abrir Arquivo Config", "Abre arquivo Excel                  ", Colors.BRIGHT_MAGENTA),
        ("0", "🚪 Sair", "Encerra a aplicação                 ", Colors.BRIGHT_RED)
    ]

    for num, title, desc, color in options:
        print(
            f"{Colors.BRIGHT_CYAN}│{Colors.RESET} {color}{num}{Colors.RESET} - {Colors.BOLD}{title:<25}{Colors.RESET} {Colors.DIM}│ {desc:<25}{Colors.RESET} {Colors.BRIGHT_CYAN}│{Colors.RESET}")

    print(
        f"{Colors.BRIGHT_CYAN}└───────────────────────────────────────────────────────────────────────┘{Colors.RESET}")

    # Mostrar data atual configurada
    print(f"\n{Colors.DIM}📅 Data de mudança atual: {Colors.BOLD}{getChangeDate()}{Colors.RESET}")
    print(f"\n{Colors.DIM}🌎 Idioma atual: {Colors.BOLD}{LANGUAGE}{Colors.RESET}")



def print_info_box():
    """tooltip"""
    print(f"\n{Colors.YELLOW}{Colors.BOLD}⚠️  INFORMAÇÕES IMPORTANTES:{Colors.RESET}")
    print(f"{Colors.BRIGHT_BLACK}┌────────────────────────────────────────────────────────────┐{Colors.RESET}")
    print(
        f"{Colors.BRIGHT_BLACK}│{Colors.RESET} • Certifique-se de que o arquivo {Colors.BOLD}Config.xlsx{Colors.RESET} está preenchido")
    print(f"{Colors.BRIGHT_BLACK}│{Colors.RESET} • O idioma padrão é EN e a data de execução é a de hoje")
    print(f"{Colors.BRIGHT_BLACK}│{Colors.RESET} • Não interaja com o mouse ou teclado durante a execução")
    print(
        f"{Colors.BRIGHT_BLACK}│{Colors.RESET} • Use {Colors.BOLD}Ctrl+C{Colors.RESET} para interromper em caso de emergência")
    print(f"{Colors.BRIGHT_BLACK}└────────────────────────────────────────────────────────────┘{Colors.RESET}")


def configure_change_date():
    """configurar a data de mudança e idioma"""
    global changeDate, LANGUAGE

    print(f"\n{Colors.BRIGHT_BLUE}{Colors.BOLD}📅 CONFIGURAÇÃO DA DATA DE MUDANÇA{Colors.RESET}")
    print(f"{Colors.BRIGHT_BLACK}{'═' * 60}{Colors.RESET}")

    # Pergunta o idioma primeiro
    print(f"\n{Colors.BRIGHT_WHITE}Selecione o idioma do sistema:{Colors.RESET}")
    print(f"{Colors.BRIGHT_CYAN}1{Colors.RESET} - Inglês (padrão)")
    print(f"{Colors.BRIGHT_CYAN}2{Colors.RESET} - Português")

    while True:
        lang_choice = input(f"{Colors.BRIGHT_WHITE}Escolha o idioma: {Colors.RESET}").strip()
        if lang_choice == "1":
            LANGUAGE = "EN"
            break
        elif lang_choice == "2":
            LANGUAGE = "PT"
            break
        else:
            print(f"{Colors.BRIGHT_RED}❌ Opção inválida! Digite 1 ou 2.{Colors.RESET}")

    print(f"\n{Colors.BRIGHT_GREEN}📅 Data atual configurada: {Colors.BOLD}{getChangeDate()}{Colors.RESET}")
    print(f"{Colors.DIM}Esta data será usada como data efetiva nos processos{Colors.RESET}")
    print(f"\n{Colors.BRIGHT_WHITE}Opções:{Colors.RESET}")
    print(f"{Colors.BRIGHT_CYAN}1{Colors.RESET} - Usar data de hoje ({date.today().strftime('%d/%m/%Y')})")
    print(f"{Colors.BRIGHT_CYAN}2{Colors.RESET} - Definir data personalizada")
    print(f"{Colors.BRIGHT_CYAN}0{Colors.RESET} - Voltar ao menu principal")

    while True:
        choice = input(f"\n{Colors.BRIGHT_WHITE}Escolha uma opção: {Colors.RESET}").strip()

        if choice == "0":
            return
        elif choice == "1":
            changeDate = date.today()
            print(f"{Colors.BRIGHT_GREEN}✅ Data alterada para hoje: {getChangeDate()}{Colors.RESET}")
            break
        elif choice == "2":
            print(f"\n{Colors.BRIGHT_YELLOW}📝 Digite a nova data (formato: dd/mm/yyyy):{Colors.RESET}")
            print(f"{Colors.DIM}Exemplo: 20/10/2025{Colors.RESET}")

            while True:
                new_date = input(f"{Colors.BRIGHT_WHITE}Nova data: {Colors.RESET}").strip()

                if not new_date:
                    print(f"{Colors.BRIGHT_RED}❌ Data não pode ser vazia!{Colors.RESET}")
                    continue

                try:
                    # Valida a data
                    changeDate = datetime.strptime(new_date, '%d/%m/%Y').date()
                    print(f"{Colors.BRIGHT_GREEN}✅ Data alterada para: {getChangeDate()}{Colors.RESET}")
                    break
                except ValueError:
                    print(f"{Colors.BRIGHT_RED}❌ Data inválida! Use o formato dd/mm/yyyy{Colors.RESET}")
            break
        else:
            print(f"{Colors.BRIGHT_RED}❌ Opção inválida! Digite 0, 1 ou 2{Colors.RESET}")

    input(f"\n{Colors.DIM}Pressione Enter para continuar...{Colors.RESET}")



def check_files_status():
    """status dos arquivos necessários"""
    print(f"\n{Colors.BRIGHT_BLUE}{Colors.BOLD}🔍 VERIFICANDO STATUS DO SISTEMA...{Colors.RESET}\n")

    files_to_check = [
        (os.path.join(BASE_DIR, "Config.xlsx"), "📋 Arquivo de Configuração"),
        (os.path.join(BASE_DIR, "script", "plannedOccupation.py"), "📅 Script Planned Occupation"),
        (os.path.join(BASE_DIR, "script", "toBeHired.py"), "👤 Script To Be Hired")
    ]

    all_good = True

    for file_path, description in files_to_check:
        if os.path.exists(file_path):
            size = os.path.getsize(file_path)
            if size > 0:
                print(f"{Colors.BRIGHT_GREEN}✅ {description:<30} {Colors.DIM}({size} bytes){Colors.RESET}")
            else:
                print(f"{Colors.BRIGHT_YELLOW}⚠️  {description:<30} {Colors.DIM}(arquivo vazio){Colors.RESET}")
                all_good = False
        else:
            print(f"{Colors.BRIGHT_RED}❌ {description:<30} {Colors.DIM}(não encontrado){Colors.RESET}")
            all_good = False

    print()
    if all_good:
        print(f"{Colors.BRIGHT_GREEN}{Colors.BOLD}🎉 Todos os arquivos estão prontos!{Colors.RESET}")
    else:
        print(f"{Colors.BRIGHT_RED}{Colors.BOLD}⚠️  Alguns arquivos precisam de atenção{Colors.RESET}")

    input(f"\n{Colors.DIM}Pressione Enter para continuar...{Colors.RESET}")


def open_config_folder():
    """abre o arquivo de config"""
    try:
        script_folder = os.path.join(BASE_DIR, "Config.xlsx")
        if os.name == 'nt':  # Windows
            os.startfile(script_folder)
        else:  # Linux/Mac
            subprocess.run(['xdg-open', script_folder])
        print(f"{Colors.BRIGHT_GREEN}✅ Arquivo de configuração aberta!{Colors.RESET}")
    except Exception as e:
        print(f"{Colors.BRIGHT_RED}❌ Erro ao abrir arquivo: {e}{Colors.RESET}")

    input(f"\n{Colors.DIM}Pressione Enter para continuar...{Colors.RESET}")


def run_script(script_name, script_description, script_emoji):
    print(f"\n{Colors.BRIGHT_BLUE}{Colors.BOLD}{script_emoji} INICIANDO {script_description.upper()}...{Colors.RESET}")
    print(f"{Colors.BRIGHT_BLACK}{'═' * 60}{Colors.RESET}")

    # Mostrar timestamp de início
    start_time = datetime.now()
    print(f"{Colors.DIM}Iniciado em: {start_time.strftime('%d/%m/%Y %H:%M:%S')}{Colors.RESET}")
    print(f"{Colors.DIM}Data de mudança: {getChangeDate()}{Colors.RESET}\n")

    try:
        # Executar o script
        script_path = os.path.join(BASE_DIR, "script", script_name)
        result = subprocess.run([sys.executable, script_path],
                                capture_output=False, text=True)

        end_time = datetime.now()
        duration = end_time - start_time

        print(f"\n{Colors.BRIGHT_BLACK}{'═' * 60}{Colors.RESET}")
        print(f"{Colors.DIM}Finalizado em: {end_time.strftime('%d/%m/%Y %H:%M:%S')}{Colors.RESET}")
        print(f"{Colors.DIM}Duração: {duration}{Colors.RESET}")

        if result.returncode == 0:
            print(f"{Colors.BRIGHT_GREEN}{Colors.BOLD}✅ {script_description} executado com sucesso!{Colors.RESET}")
        else:
            print(
                f"{Colors.BRIGHT_RED}{Colors.BOLD}❌ {script_description} finalizou com erros (código: {result.returncode}){Colors.RESET}")

    except FileNotFoundError:
        print(
            f"{Colors.BRIGHT_RED}❌ Erro: Arquivo '{os.path.join('', script_name)}' não encontrado!{Colors.RESET}")
    except KeyboardInterrupt:
        print(f"\n{Colors.BRIGHT_YELLOW}⚠️  Operação interrompida pelo usuário{Colors.RESET}")
    except Exception as e:
        print(f"{Colors.BRIGHT_RED}❌ Erro inesperado: {e}{Colors.RESET}")

    input(f"\n{Colors.DIM}Pressione Enter para voltar ao menu...{Colors.RESET}")


def get_user_choice():
    """captura a escolha do usuário"""
    while True:
        try:
            print(f"\n{Colors.BRIGHT_WHITE}┌─────────────────────────────────────────────────────────┐{Colors.RESET}")
            choice = input(
                f"{Colors.BRIGHT_WHITE}│{Colors.RESET} {Colors.BOLD}Escolha uma opção:{Colors.RESET} ").strip()
            print(f"{Colors.BRIGHT_WHITE}└─────────────────────────────────────────────────────────┘{Colors.RESET}")

            if choice in ['0', '1', '2', '3', '4']:
                return choice
            else:
                print(f"{Colors.BRIGHT_RED}❌ Opção inválida! Digite apenas: 0, 1, 2, 3 ou 4{Colors.RESET}")
                time.sleep(1)
        except KeyboardInterrupt:
            print(f"\n{Colors.BRIGHT_YELLOW}👋 Saindo...{Colors.RESET}")
            sys.exit(0)


def main():
    while True:
        try:
            print_header()
            print_menu_options()
            print_info_box()

            choice = get_user_choice()

            if choice == "1":
                run_script("plannedOccupation.py", "Planned Occupation Date", "📅")

            elif choice == "2":
                run_script("toBeHired.py", "To Be Hired", "👤")

            elif choice == "3":
                configure_change_date()

            elif choice == "4":
                open_config_folder()

            elif choice == "0":
                clear_screen()
                print(f"{Colors.BRIGHT_CYAN}{Colors.BOLD}")
                print("╔" + "═" * 40 + "╗")
                print("║" + f"{'👋 Obrigado por usar o robozinho!':^40}" + "║")
                print("║" + f"{'Até a próxima! 🤖':^40}" + "║")
                print("╚" + "═" * 40 + "╝")
                print(f"{Colors.RESET}")
                break

        except KeyboardInterrupt:
            print(f"\n{Colors.BRIGHT_YELLOW}👋 Saindo...{Colors.RESET}")
            break
        except Exception as e:
            print(f"{Colors.BRIGHT_RED}❌ Erro inesperado: {e}{Colors.RESET}")
            input(f"{Colors.DIM}Pressione Enter para continuar...{Colors.RESET}")


if __name__ == "__main__":
    main()