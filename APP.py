import pandas as pd
import win32com.client
import sys
import gspread
from datetime import datetime
import subprocess
import time
import re
import getpass
import configparser
import os
import pywintypes 

# --- Classes e Funções de Interface ---
class Cores:
    """Define cores para uma saída de console mais legível."""
    RESET, VERDE, AMARELO, VERMELHO, AZUL, CIANO = '\033[0m', '\033[92m', '\033[93m', '\033[91m', '\033[94m', '\033[96m'

def print_header(texto):
    """Imprime um cabeçalho formatado."""
    print(f"\n{Cores.AZUL}{'='*60}\n {texto.center(58)}\n {'='*60}{Cores.RESET}")

def print_sucesso(texto):
    """Imprime uma mensagem de sucesso."""
    print(f"{Cores.VERDE}[SUCESSO] {texto}{Cores.RESET}")

def print_info(texto):
    """Imprime uma mensagem informativa."""
    print(f"{Cores.CIANO}[INFO]    {texto}{Cores.RESET}")

def print_aviso(texto):
    """Imprime uma mensagem de aviso."""
    print(f"{Cores.AMARELO}[AVISO]   {texto}{Cores.RESET}")

def print_erro(texto):
    """Imprime uma mensagem de erro."""
    print(f"{Cores.VERMELHO}[ERRO]     {texto}{Cores.RESET}")

def aguardar_sap(session):
    """Espera a sessão SAP parar de processar (não estar ocupada)."""
    while session.busy:
        time.sleep(0.2)
    return

def obter_confirmacao(prompt, default_yes=True):
    """Pede uma confirmação (s/n) ao usuário de forma robusta."""
    options = "(S/n)" if default_yes else "(s/N)"
    prompt_full = f"  {Cores.AMARELO}{prompt} {options}: {Cores.RESET}"
    
    while True:
        resposta = input(prompt_full).lower().strip()
        if resposta == '':
            return default_yes
        if resposta in ['s', 'sim', 'y', 'yes','SIM']:
            return True
        if resposta in ['n', 'nao', 'no','não','NÃO']:
            return False
        print_erro("Resposta inválida. Por favor, digite 's' para sim ou 'n' para não.")

# --- Funções de Gerenciamento da Sessão SAP ---

def is_session_valid(session):
    """
    Verifica se a sessão SAP ainda está ativa e responsiva.
    Esta é a verificação principal para evitar problemas com sessões desconectadas.
    """
    if session is None:
        return False
    try:
        # Tenta acessar uma propriedade simples da sessão.
        # Se a sessão foi fechada ou desconectada, isso gerará um erro.
        session.findById("wnd[0]")
        return True
    except (pywintypes.com_error, Exception):
        # pywintypes.com_error é o erro específico que ocorre quando o objeto COM (SAP) não é mais válido.
        return False

def sap_login_handler(config):
    """
    Gerencia a conexão SAP. Tenta se conectar a uma sessão existente
    ou abre uma nova e faz o login se nenhuma for encontrada.
    """
    try:
        print_info("Procurando por uma sessão SAP GUI...")
        sap_gui_auto = win32com.client.GetObject("SAPGUI")
        if not sap_gui_auto:
            return open_and_login_sap(config)

        application = sap_gui_auto.GetScriptingEngine
        if not application:
            return open_and_login_sap(config)

        # Tenta encontrar uma conexão existente
        if application.Connections.Count > 0:
            connection = application.Connections(0)
            if connection.Sessions.Count > 0:
                # Itera sobre as sessões abertas para encontrar uma válida
                for i in range(connection.Sessions.Count):
                    session = connection.Sessions(i)
                    if is_session_valid(session):
                        print_sucesso(f"Sessão SAP ativa encontrada (Sessão {i}).")
                        return session
        
        # Se nenhuma sessão válida foi encontrada, abre uma nova
        print_aviso("Nenhuma sessão SAP válida encontrada. Iniciando uma nova conexão...")
        return open_and_login_sap(config)

    except (pywintypes.com_error, Exception):
        print_aviso("Nenhuma instância do SAP GUI encontrada. Iniciando processo de login...")
        return open_and_login_sap(config)


def open_and_login_sap(config):
    """Abre o SAP Logon, conecta-se ao sistema e realiza o login de forma robusta."""
    try:
        sap_path = config['SAP']['caminho_logon']
        sap_system = config['SAP']['sistema']
        
        print_info(f"Abrindo SAP Logon de '{sap_path}'...")
        try:
            subprocess.Popen(sap_path)
            time.sleep(5) # Espera o processo saplogon.exe iniciar
        except FileNotFoundError:
            print_erro(f"Caminho para o SAP Logon não encontrado: '{sap_path}'. Verifique o config.ini.")
            return None

        sap_gui_auto = win32com.client.GetObject("SAPGUI")
        application = sap_gui_auto.GetScriptingEngine
        
        print_info(f"Abrindo conexão com o sistema '{sap_system}'...")
        connection = application.OpenConnection(sap_system, True)
        session = connection.Children(0)

        print_info("Aguardando a janela do SAP ficar pronta...")
        start_time = time.time()
        timeout = 30
        while time.time() - start_time < timeout:
            try:
                if session.findById("wnd[0]").text:
                    break
            except:
                time.sleep(0.5)
        
        main_window = session.findById("wnd[0]")
        
        # LÓGICA PRINCIPAL DE LOGIN
        # Primeiro, tenta encontrar o campo de usuário. Se não achar, assume que já está logado.
        try:
            main_window.findById("usr/txtRSYST-BNAME") # Tenta acessar o campo de usuário
            
            # Se o código chegou aqui, a tela de login está ativa.
            print_info("Tela de login detectada. Preenchendo credenciais...")
            print_header("Login no SAP")
            
            user = config.get('SAP', 'usuario', fallback=None)
            password = config.get('SAP', 'senha', fallback=None)

            if not user or not password:
                print_erro("Login automático falhou: usuário/senha não encontrados no config.ini.")
                return None

            main_window.findById("usr/txtRSYST-BNAME").text = user
            main_window.findById("usr/pwdRSYST-BCODE").text = password
            main_window.sendVKey(0)
            aguardar_sap(session)

            # Lida com pop-up de "licença múltipla"
            try:
                session.findById("wnd[1]").sendVKey(0)
                aguardar_sap(session)
            except Exception:
                pass

            # Verificação final após a tentativa de login
            if "easy access" in session.findById("wnd[0]").text.lower() or "menú" in session.findById("wnd[0]").text.lower():
                print_sucesso("Login no SAP realizado com sucesso!")
                return session
            else:
                status_text = session.findById("sbar").text
                error_message = status_text if status_text else "Credenciais inválidas ou tela inesperada."
                print_erro(f"Falha no login: {error_message}")
                return None

        except (pywintypes.com_error, Exception):
            # Se não encontrou o campo de usuário, provavelmente já está logado.
            print_info("Campo de login não encontrado. Verificando se a sessão já está na tela principal...")
            if "easy access" in main_window.text.lower() or "menú" in main_window.text.lower():
                print_sucesso("Login automático/sessão existente detectado.")
                return session
            else:
                print_erro("Erro inesperado: não foi possível identificar a tela de login nem a tela principal do SAP.")
                print_erro(f"Título da janela atual: '{main_window.text}'")
                return None

    except Exception as e:
        print_erro(f"Ocorreu um erro crítico irrecuperável durante o processo de login: {str(e)}")
        return None

# ==============================================================================
# FUNÇÃO PRINCIPAL MODIFICADA - Substitua a sua função novamente por esta
# ==============================================================================
def processar_lote_de_rc(session, lote_de_itens):
    """
    Processa um lote de itens para criar uma Requisição de Compra no SAP,
    com tratamento de erros e pop-ups integrado diretamente no fluxo.
    Retorna (numero_rc, mensagem_status) ou (None, mensagem_erro).
    """
    if lote_de_itens.empty:
        return None, "Lote de itens estava vazio."

    try:
        print_header(f"Processando Lote de {len(lote_de_itens)} itens")
        
        # --- PASSO 1: NAVEGAÇÃO E PREENCHIMENTO DA GRADE ---
        print_info("Navegando para a transação ME51N...")
        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[0]/okcd").text = "/NME51N"
        session.findById("wnd[0]").sendVKey(0)
        aguardar_sap(session)
        
        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:3327/cmbMEREQ_TOPLINE-BSART").key = "ZRT"
        session.findById("wnd[0]").sendVKey(0)
        aguardar_sap(session)
        
        print_info("Preenchendo a grade de itens...")
        grid = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell")

        for _, item in lote_de_itens.iterrows():
            item_index = int(item['Validador'])
            data_obj = datetime.strptime(str(item['DATA REMESSA']), '%d/%m/%Y')
            data_remessa_formatada = data_obj.strftime('%d.%m.%Y')
            
            grid.modifyCell(item_index, "MATNR", str(item['PN']))
            grid.modifyCell(item_index, "MENGE", str(item['QTD']).replace(',', '.'))
            grid.modifyCell(item_index, "RESWK", str(item['ORIGEM']))
            grid.modifyCell(item_index, "EEIND", data_remessa_formatada)
            grid.modifyCell(item_index, "EPSTP", "U")
            grid.modifyCell(item_index, "NAME1", str(item['DESTINO']))
            grid.modifyCell(item_index, "EKGRP", "P04")
            grid.modifyCell(item_index, "TXZ01", str(item['TEXTO']))
        
        print_info("Validando itens da grade (pressionando Enter)...")
        session.findById("wnd[0]").sendVKey(0)
        aguardar_sap(session)
        
        # --- MELHORIA: Tratamento do pop-up de aviso (wnd[1]) ---
        try:
            # Tenta encontrar a janela de pop-up (wnd[1])
            popup_window = session.findById("wnd[1]")
            popup_text = getattr(popup_window, 'text', 'Pop-up sem texto')
            print_aviso(f"Pop-up detectado: '{popup_text}'. Pressionando Enter para continuar.")
            popup_window.sendVKey(0) # 0 = VKey for Enter
            aguardar_sap(session)
        except Exception:
            # Se não encontrou wnd[1], significa que não havia pop-up. O que é normal.
            print_info("Nenhum pop-up de aviso encontrado. Prosseguindo...")
            pass

        # --- Verificação de erros APÓS a validação e o tratamento do pop-up ---
        status_bar = session.findById("wnd[0]/sbar")
        if status_bar.messageType in ("E", "A"): # 'E' para Erro, 'A' para Abortar
            error_message = status_bar.text
            print_erro(f"Erro impeditivo detectado após validação: {error_message}")
            return None, error_message
        
        print_sucesso("Grade principal preenchida e validada.")

        # --- PASSO 2: INSERINDO DEPÓSITOS ---
        print_header("Passo 2: Inserindo Depósitos")
        # (Seu código para preencher os depósitos continua o mesmo)
        grid.setCurrentCell(int(lote_de_itens.iloc[0]['Validador']), "MATNR")
        aguardar_sap(session)
        for i, (_, item) in enumerate(lote_de_itens.iterrows()):
            item_index = int(item['Validador'])
            print_info(f"Preenchendo depósito para o item {item_index + 1}...")
            session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT16").select()
            session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT16/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsTABREITER1/tabpTRANS").select()
            depot_field = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT16/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsTABREITER1/tabpTRANS/ssubSUBBILD1:SAPLXM02:0114/ctxtEBAN-ZZDEP_FORNEC")
            depot_field.text = str(item['DEPOSITO'])
            if i < len(lote_de_itens) - 1:
                session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB1:SAPLMEGUI:6000/btn%#AUTOTEXT002").press()
                aguardar_sap(session)

        # --- PASSO 3: SALVANDO E VERIFICANDO O RESULTADO ---
        print_header("Passo 3: Tentando salvar a Requisição")
        session.findById("wnd[0]/tbar[0]/btn[11]").press() # Botão Salvar
        aguardar_sap(session)
        
        # --- MELHORIA: Adicionado tratamento de pop-up também APÓS salvar ---
        # Às vezes, o SAP pede uma confirmação final em um pop-up.
        try:
            popup_window = session.findById("wnd[1]")
            popup_text = getattr(popup_window, 'text', 'Pop-up sem texto')
            print_aviso(f"Pop-up detectado após salvar: '{popup_text}'. Pressionando Enter/Sim.")
            popup_window.sendVKey(0)
            aguardar_sap(session)
        except Exception:
            pass # É normal não ter pop-up aqui.

        # --- LÓGICA DE VERIFICAÇÃO FINAL (SUCESSO OU ERRO) ---
        status_message = session.findById("wnd[0]/sbar").text
        status_type = session.findById("wnd[0]/sbar").messageType

        match = re.search(r'(\d{10,})', status_message)
        if match:
            rc_number = match.group(0)
            print_sucesso(f"Requisição criada com sucesso! Número: {rc_number}")
            return rc_number, status_message
        
        if status_type in ("E", "A"):
            print_erro(f"Falha ao salvar a requisição. Mensagem SAP: {status_message}")
            return None, status_message
        
        if "criada" in status_message.lower() or "created" in status_message.lower():
            print_sucesso(f"Requisição criada, mas não foi possível extrair o número. Mensagem: {status_message}")
            return "Criada com sucesso", status_message

        # Se a mensagem de status ainda estiver vazia, é porque o processo está preso.
        if not status_message:
            error_msg = "Processo interrompido. Provavelmente por um pop-up ou tela inesperada que não foi tratada."
            print_erro(error_msg)
            return None, error_msg

        print_aviso(f"Resultado inesperado ao salvar. Status SAP: '{status_message}'")
        return None, f"Status inesperado: {status_message}"

    except Exception as e:
        # Captura de erro crítico de automação (ex: objeto não encontrado, sessão caiu)
        error_msg = f"Erro crítico na automação SAP: {str(e)}"
        print_erro(error_msg)
        try:
            session.findById("wnd[0]/tbar[0]/okcd").text = "/N"
            session.findById("wnd[0]").sendVKey(0)
        except:
            pass
        return None, error_msg

def get_base_path():
    """Retorna o caminho base, seja rodando como script ou como executável."""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))

def main():
    """
    Função principal que opera em ciclos manuais, permitindo ao usuário
    processar a planilha sob demanda, com reconexão automática ao SAP.
    """
    base_path = get_base_path()
    config_path = os.path.join(base_path, 'config.ini')
    session = None

    # Carrega configurações
    try:
        config = configparser.ConfigParser()
        config.read(config_path, encoding='utf-8')
    except Exception as e:
        print_erro(f"Não foi possível ler o arquivo 'config.ini'. Verifique se ele existe. Erro: {e}")
        time.sleep(10)
        return

    print_header("Robô de Criação de Requisição de Compra no SAP")
    
    if not obter_confirmacao("Deseja iniciar a automação?", default_yes=True):
        print_info("Automação cancelada pelo usuário.")
        return

    while True:
        # 1. VERIFICAR E GARANTIR SESSÃO SAP A CADA CICLO
        if not is_session_valid(session):
            print_aviso("Sessão SAP inválida ou inexistente. Tentando conectar...")
            session = sap_login_handler(config)
            if not session:
                print_erro("Não foi possível estabelecer uma sessão com o SAP.")
                if not obter_confirmacao("Deseja tentar a conexão novamente?", default_yes=True):
                    break # Sai do loop principal se o usuário não quiser tentar de novo
                else:
                    continue # Volta para o início do loop para tentar a conexão de novo

        # 2. CONECTAR À PLANILHA E PROCESSAR
        try:
            print_header("INICIANDO VERIFICAÇÃO DA PLANILHA")
            print_info("Conectando à Planilha Google...")
            credentials_file = os.path.join(base_path, config['GOOGLE']['credenciais'])
            gc = gspread.service_account(filename=credentials_file)
            spreadsheet = gc.open(config['GOOGLE']['planilha'])
            worksheet = spreadsheet.worksheet(config['GOOGLE']['aba'])
            print_sucesso("Conexão com a planilha estabelecida.")

            headers = worksheet.row_values(1)
            status_col_index = headers.index("Status") + 1
            req_col_index = headers.index("REQUISIÇÃO") + 1
            
            df = pd.DataFrame(worksheet.get_all_records())
            if df.empty:
                print_aviso("A planilha está vazia. Nenhuma ação necessária.")
            else:
                df['linha_planilha'] = df.index + 2
                df_para_processar = df[df['Status'] == ''].copy()

                if not df_para_processar.empty:
                    print_info(f"Encontradas {len(df_para_processar)} linhas novas para processar.")
                    df_para_processar['Validador'] = pd.to_numeric(df_para_processar['Validador'])
                    df_para_processar['lote_id'] = (df_para_processar['Validador'] == 0).cumsum()
                    
                    for lote_id, lote_df in df_para_processar.groupby('lote_id'):
                        if not is_session_valid(session):
                            print_erro("Sessão SAP perdida durante o processamento. O ciclo será reiniciado.")
                            session = None
                            break 

                        numero_rc, msg_status = processar_lote_de_rc(session, lote_df)
                        
                        for linha in lote_df['linha_planilha']:
                            try:
                                worksheet.update_cell(linha, status_col_index, msg_status)
                                if numero_rc and numero_rc != "N/A":
                                    worksheet.update_cell(linha, req_col_index, numero_rc)
                                time.sleep(1.1)
                            except gspread.exceptions.APIError as api_err:
                                print_erro(f"Erro de API do Google ao atualizar linha {linha}: {api_err}.")
                                time.sleep(10)
                else:
                    print_aviso("Nenhuma linha nova para processar (todas já possuem status).")

        except FileNotFoundError:
            print_erro(f"Arquivo de credenciais '{config['GOOGLE']['credenciais']}' não encontrado. Verifique o config.ini.")
        except Exception as e:
            print_erro(f"Erro crítico no processamento do ciclo: {e}")
            if "session" in str(e).lower() or "pywintypes.com_error" in str(e):
                session = None

        # 3. PERGUNTAR AO USUÁRIO SE DESEJA CONTINUAR
        print_header("FIM DO CICLO")
        if not obter_confirmacao("Deseja verificar a planilha e processar novos itens novamente?", default_yes=False):
            break 
        
        print_info("Reiniciando o processo...")

    print_header("Robô Encerrado")

if __name__ == "__main__":
    main()