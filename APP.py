# -*- coding: utf-8 -*-
import pandas as pd
import win32com.client
import sys
import gspread
from datetime import datetime
import subprocess
import time
import re
import os
import configparser
import pywintypes # Import necessário para capturar erros específicos de COM

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
    """Verifica se a sessão SAP ainda está ativa e responsiva."""
    if session is None:
        return False
    try:
        session.findById("wnd[0]")
        return True
    except (pywintypes.com_error, Exception):
        return False

def sap_login_handler(config):
    """Gerencia a conexão SAP."""
    try:
        print_info("Procurando por uma sessão SAP GUI...")
        sap_gui_auto = win32com.client.GetObject("SAPGUI")
        application = sap_gui_auto.GetScriptingEngine
        if application.Connections.Count > 0:
            connection = application.Connections(0)
            if connection.Sessions.Count > 0:
                for i in range(connection.Sessions.Count):
                    session = connection.Sessions(i)
                    if is_session_valid(session):
                        print_sucesso(f"Sessão SAP ativa encontrada (Sessão {i}).")
                        return session
        print_aviso("Nenhuma sessão SAP válida encontrada. Iniciando uma nova conexão...")
        return open_and_login_sap(config)
    except (pywintypes.com_error, Exception):
        print_aviso("Nenhuma instância do SAP GUI encontrada. Iniciando processo de login...")
        return open_and_login_sap(config)

def open_and_login_sap(config):
    """Abre o SAP Logon e faz o login."""
    try:
        sap_path = config['SAP']['caminho_logon']
        sap_system = config['SAP']['sistema']
        print_info(f"Abrindo SAP Logon de '{sap_path}'...")
        subprocess.Popen(sap_path)
        time.sleep(5)
        sap_gui_auto = win32com.client.GetObject("SAPGUI")
        application = sap_gui_auto.GetScriptingEngine
        print_info(f"Abrindo conexão com o sistema '{sap_system}'...")
        connection = application.OpenConnection(sap_system, True)
        session = connection.Children(0)
        aguardar_sap(session)
        main_window = session.findById("wnd[0]")
        try:
            main_window.findById("usr/txtRSYST-BNAME")
            print_info("Tela de login detectada. Preenchendo credenciais...")
            user = config.get('SAP', 'usuario', fallback=None)
            password = config.get('SAP', 'senha', fallback=None)
            main_window.findById("usr/txtRSYST-BNAME").text = user
            main_window.findById("usr/pwdRSYST-BCODE").text = password
            main_window.sendVKey(0)
            aguardar_sap(session)
            try:
                session.findById("wnd[1]").sendVKey(0)
                aguardar_sap(session)
            except Exception: pass
            if "easy access" in session.findById("wnd[0]").text.lower() or "menú" in session.findById("wnd[0]").text.lower():
                print_sucesso("Login no SAP realizado com sucesso!")
                return session
            else:
                print_erro(f"Falha no login: {session.findById('sbar').text}")
                return None
        except (pywintypes.com_error, Exception):
            print_sucesso("Sessão existente detectada. Pulando login.")
            return session
    except Exception as e:
        print_erro(f"Ocorreu um erro crítico durante o processo de login: {str(e)}")
        return None

# ==============================================================================
# FASE 1: FUNÇÃO DE VALIDAÇÃO (ITEM A ITEM)
# ==============================================================================
def validar_lote_na_rc(session, lote_de_itens):
    """Valida itens na ME51N um por um e retorna 'OK' ou a mensagem de erro."""
    if lote_de_itens.empty: return []
    resultados_finais = []
    try:
        print_header(f"Validando Lote de {len(lote_de_itens)} itens (um por um)")
        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[0]/okcd").text = "/NME51N"
        session.findById("wnd[0]").sendVKey(0)
        aguardar_sap(session)
        time.sleep(1)
        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:3327/cmbMEREQ_TOPLINE-BSART").key = "ZRT"
        session.findById("wnd[0]").sendVKey(0)
        aguardar_sap(session)
        grid = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell")
        for _, item in lote_de_itens.iterrows():
            grid_index = int(item['Validador'])
            linha_planilha = item['linha_planilha']
            status_item = "OK"
            print_header(f"Processando Linha {grid_index + 1} (PN: {item['PN']})")
            try:
                data_obj = datetime.strptime(str(item['DATA REMESSA']), '%d/%m/%Y')
                data_remessa_formatada = data_obj.strftime('%d.%m.%Y')
                grid.modifyCell(grid_index, "MATNR", str(item['PN']))
                grid.modifyCell(grid_index, "MENGE", str(item['QTD']).replace(',', '.'))
                grid.modifyCell(grid_index, "RESWK", str(item['ORIGEM']))
                grid.modifyCell(grid_index, "EEIND", data_remessa_formatada)
                grid.modifyCell(grid_index, "EPSTP", "U")
                grid.modifyCell(grid_index, "NAME1", str(item['DESTINO']))
                grid.modifyCell(grid_index, "EKGRP", "P04")
                grid.modifyCell(grid_index, "TXZ01", str(item['TEXTO']))
                print_info("Pressionando Enter para validar o item atual...")
                session.findById("wnd[0]").sendVKey(0)
                aguardar_sap(session)
                time.sleep(1.5)
                try:
                    session.findById("wnd[1]").sendVKey(0)
                    aguardar_sap(session)
                except Exception: pass
                status_bar = session.findById("wnd[0]/sbar")
                message_type = status_bar.messageType
                message_text = status_bar.text
                if message_type in ('E', 'A') or "não está atualizado no centro" in message_text:
                    status_item = message_text
                    print_erro(f"Resultado: [ERRO] {status_item}")
                else:
                    status_item = "OK"
                    if message_text: print_aviso(f"Resultado: [AVISO, tratado como OK] - {message_text}")
                    else: print_sucesso("Resultado: OK")
            except Exception as e:
                status_item = f"Falha crítica ao processar a linha. Erro: {str(e)}"
                print_erro(status_item)
            resultados_finais.append({'linha_planilha': linha_planilha, 'status': status_item, 'numero_rc': '' if status_item == 'OK' else 'ERRO'})
        return resultados_finais
    finally:
        print_info("Validação do lote concluída. Encerrando a transação.")
        try:
            if is_session_valid(session):
                session.findById("wnd[0]/tbar[0]/okcd").text = "/N"
                session.findById("wnd[0]").sendVKey(0)
        except Exception: pass

# ==============================================================================
# FASE 2: FUNÇÃO DE CRIAÇÃO DA RC
# ==============================================================================
def criar_rc_para_lote_ok(session, lote_de_itens_ok):
    """Pega um lote de itens já validados como OK e cria a RC."""
    if lote_de_itens_ok.empty:
        return None, "Lote de itens OK estava vazio."
    try:
        print_header(f"Criando Requisição para {len(lote_de_itens_ok)} itens validados")
        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[0]/okcd").text = "/NME51N"
        session.findById("wnd[0]").sendVKey(0)
        aguardar_sap(session)
        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:3327/cmbMEREQ_TOPLINE-BSART").key = "ZRT"
        session.findById("wnd[0]").sendVKey(0)
        aguardar_sap(session)
        grid = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell")
        
        lote_de_itens_ok = lote_de_itens_ok.reset_index(drop=True)
        for i, item in lote_de_itens_ok.iterrows():
            data_obj = datetime.strptime(str(item['DATA REMESSA']), '%d/%m/%Y')
            data_remessa_formatada = data_obj.strftime('%d.%m.%Y')
            grid.modifyCell(i, "MATNR", str(item['PN']))
            grid.modifyCell(i, "MENGE", str(item['QTD']).replace(',', '.'))
            grid.modifyCell(i, "RESWK", str(item['ORIGEM']))
            grid.modifyCell(i, "EEIND", data_remessa_formatada)
            grid.modifyCell(i, "EPSTP", "U")
            grid.modifyCell(i, "NAME1", str(item['DESTINO']))
            grid.modifyCell(i, "EKGRP", "P04")
            grid.modifyCell(i, "TXZ01", str(item['TEXTO']))
        
        session.findById("wnd[0]").sendVKey(0)
        aguardar_sap(session)
        
        print_info("Inserindo Depósitos...")
        for i, item in lote_de_itens_ok.iterrows():
            grid.setCurrentCell(i, "MATNR")
            aguardar_sap(session)
            session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT16").select()
            session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT16/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsTABREITER1/tabpTRANS").select()
            # ID CORRIGIDO ABAIXO, USANDO A VERSÃO DO SEU SCRIPT ORIGINAL
            depot_field = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT16/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsTABREITER1/tabpTRANS/ssubSUBBILD1:SAPLXM02:0114/ctxtEBAN-ZZDEP_FORNEC")
            depot_field.text = str(item['DEPOSITO'])
            if i < len(lote_de_itens_ok) - 1:
                session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB1:SAPLMEGUI:6000/btn%#AUTOTEXT002").press()
                aguardar_sap(session)

        print_info("Salvando a Requisição de Compra...")
        session.findById("wnd[0]/tbar[0]/btn[11]").press()
        aguardar_sap(session)
        status_message = session.findById("wnd[0]/sbar").text
        match = re.search(r'(\d{10,})', status_message)
        if match:
            rc_number = match.group(0)
            print_sucesso(f"Requisição criada com sucesso! Número: {rc_number}")
            return rc_number, status_message
        else:
            print_erro(f"Falha ao salvar a requisição. Mensagem SAP: {status_message}")
            return None, status_message
    except Exception as e:
        error_msg = f"Erro crítico na criação da RC: {str(e)}"
        print_erro(error_msg)
        return None, error_msg

def get_base_path():
    """Retorna o caminho base, seja rodando como script ou como executável."""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

# ==============================================================================
# FUNÇÃO PRINCIPAL REESTRUTURADA PARA FLUXO CONTÍNUO
# ==============================================================================
def main():
    """Função principal que opera em duas fases automáticas: Validação e Criação."""
    base_path = get_base_path()
    config_path = os.path.join(base_path, 'config.ini')
    session = None
    try:
        config = configparser.ConfigParser()
        config.read(config_path, encoding='utf-8')
    except Exception as e:
        print_erro(f"Não foi possível ler 'config.ini'. Erro: {e}"); time.sleep(10); return

    print_header("Robô de Requisição de Compra no SAP (Valida e Cria)")
    if not obter_confirmacao("Deseja iniciar a automação?", default_yes=True):
        print_info("Automação cancelada."); return

    while True:
        if not is_session_valid(session):
            print_aviso("Sessão SAP inválida. Tentando conectar/reconectar...")
            session = sap_login_handler(config)
            if not session:
                print_erro("Não foi possível estabelecer uma sessão com o SAP.")
                if not obter_confirmacao("Tentar novamente?", default_yes=True): break
                else: continue
        try:
            # --- CONEXÃO COM GOOGLE SHEETS ---
            print_header("CONECTANDO À PLANILHA")
            credentials_file = os.path.join(base_path, config['GOOGLE']['credenciais'])
            gc = gspread.service_account(filename=credentials_file)
            spreadsheet = gc.open(config['GOOGLE']['planilha'])
            worksheet = spreadsheet.worksheet(config['GOOGLE']['aba'])
            print_sucesso("Conexão com a planilha estabelecida.")
            headers = worksheet.row_values(1)
            status_col_index = headers.index("Status") + 1
            req_col_index = headers.index("REQUISIÇÃO") + 1
            
            df = pd.DataFrame(worksheet.get_all_records())
            df['linha_planilha'] = df.index + 2
            df_para_processar = df[df['Status'] == ''].copy()

            if not df_para_processar.empty:
                print_info(f"Encontradas {len(df_para_processar)} linhas novas para o fluxo de Valida/Cria.")
                df_para_processar['Validador'] = pd.to_numeric(df_para_processar['Validador'])
                df_para_processar['lote_id'] = (df_para_processar['Validador'] == 0).cumsum()
                
                for lote_id, lote_df in df_para_processar.groupby('lote_id'):
                    if not is_session_valid(session): print_erro("Sessão SAP perdida."); session=None; break
                    
                    # --- FASE 1: VALIDAÇÃO DO LOTE ATUAL ---
                    resultados_validacao = validar_lote_na_rc(session, lote_df)
                    
                    # Atualiza a planilha com os resultados da validação
                    print_header("Atualizando Planilha com Resultados da Validação")
                    linhas_ok = []
                    for res in resultados_validacao:
                        try:
                            print_info(f"Atualizando linha {res['linha_planilha']}: Status='{res['status']}'")
                            worksheet.update_cell(res['linha_planilha'], status_col_index, res['status'])
                            worksheet.update_cell(res['linha_planilha'], req_col_index, str(res['numero_rc']))
                            if res['status'] == 'OK':
                                linhas_ok.append(res['linha_planilha'])
                            time.sleep(1.2)
                        except gspread.exceptions.APIError as api_err:
                            print_erro(f"Erro de API Google na linha {res['linha_planilha']}: {api_err}. Aguardando...")
                            time.sleep(10)

                    # --- FASE 2: CRIAÇÃO AUTOMÁTICA PARA ITENS OK DO LOTE ATUAL ---
                    if not linhas_ok:
                        print_aviso(f"Lote {lote_id} não possui itens 'OK'. Pulando para o próximo lote.")
                        continue

                    print_header(f"FASE 2: CRIAÇÃO AUTOMÁTICA PARA LOTE {lote_id}")
                    lote_df_ok = lote_df[lote_df['linha_planilha'].isin(linhas_ok)].copy()
                    
                    if not is_session_valid(session): print_erro("Sessão SAP perdida."); session=None; break
                    
                    numero_rc, msg_status = criar_rc_para_lote_ok(session, lote_df_ok)
                    
                    print_header("Atualizando Planilha com a RC Criada")
                    for linha in lote_df_ok['linha_planilha']:
                        try:
                            worksheet.update_cell(linha, status_col_index, msg_status)
                            if numero_rc: worksheet.update_cell(linha, req_col_index, numero_rc)
                            time.sleep(1.2)
                        except gspread.exceptions.APIError as api_err:
                            print_erro(f"Erro de API Google na linha {linha}: {api_err}. Aguardando...")
                            time.sleep(10)
            else:
                print_aviso("Nenhuma linha nova para validar e criar.")
                
        except FileNotFoundError:
            print_erro(f"Arquivo de credenciais '{config['GOOGLE']['credenciais']}' não encontrado.")
        except Exception as e:
            print_erro(f"Erro crítico no ciclo principal: {e}")
            session = None

        print_header("FIM DO CICLO")
        if not obter_confirmacao("Verificar a planilha novamente?", default_yes=False): break
        print_info("Reiniciando o processo...")

    print_header("Robô Encerrado")

if __name__ == "__main__":
    main()

