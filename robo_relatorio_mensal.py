# =========================IMPORTAÇÕES DE BIBLIOTECAS E COMPONENTES========================
from components.importacao_diretorios_windows import listagem_arquivos, procura_pasta_cliente
from components.importacao_caixa_dialogo import DialogBox
from components.checar_ativacao_google_drive import checa_google_drive
from components.configuracao_db import configura_db, ler_sql
from components.importacao_automacao_excel_openpyxl import carrega_excel
from components.procura_cliente import procura_cliente, procura_clientes_por_regiao
from components.procura_valores import procura_valores, procura_todos_valores_ano
from components.enviar_emails import enviar_email_com_anexos
from components.aws_parameters import get_ssm_parameter
import boto3
from botocore.exceptions import ClientError
from google.auth.transport.requests import Request
from google.auth import identity_pool
from googleapiclient.discovery import build
import mysql.connector
import tkinter as tk
from pathlib import Path
from openpyxl.styles import Border, Side, NamedStyle
from shutil import copy
import win32com.client as win32
from dotenv import load_dotenv
import os
import json
from time import sleep
import locale
import calendar
import pythoncom
import pandas as pd
from flask import Flask, request
from flask_restful import Resource, Api, reqparse

locale.setlocale(locale.LC_ALL, 'pt_BR.utf8')
load_dotenv()

# =====================CONFIGURAÇÂO DO BANCO DE DADOS===========================
db_conf = configura_db()

# =============CHECANDO SE O GOOGLE FILE STREAM ESTÁ INICIADO NO SISTEMA========
checa_google_drive()

# ============= MÉTODOS AUXILIARES =============================================

def get_secret():
    secret_name = "GoogleFederationConfig"
    region_name = "sa-east-1"

    # Create a Secrets Manager client
    session = boto3.session.Session()
    client = session.client(
        service_name='secretsmanager',
        region_name=region_name
    )
    try:
        get_secret_value_response = client.get_secret_value(
            SecretId=secret_name
        )
    except ClientError as e:
        # For a list of exceptions thrown, see
        # https://docs.aws.amazon.com/secretsmanager/latest/apireference/API_GetSecretValue.html
        raise e

    secret = get_secret_value_response['SecretString']
    return json.loads(secret)

def carregar_credenciais():
    try:
        secret_name = "GoogleFederationConfig"
        secret_json = get_secret(secret_name)
        
        credentials = identity_pool.Credentials.from_info(secret_json)

        SCOPES = [get_ssm_parameter('/human/API_SCOPES')]
        credentials = credentials.with_scopes(SCOPES)

        credentials.refresh(Request())
        return credentials
    except Exception as error:
        print(error)

def autenticacao_google_drive():
    try:
        service_name = get_ssm_parameter('/human/API_NAME')
        service_version = get_ssm_parameter('/human/API_VERSION')
        credentials = carregar_credenciais()
        drive_service = build(service_name, service_version, credentials=credentials)
        return drive_service
    except Exception as error:
        print(error)

driver_service = autenticacao_google_drive()

def lista_pastas_em_diretorio(folder_id):
    try:
        query = f"'{folder_id}' in parents and trashed=false"
        results = driver_service.files().list(q=query, pageSize=80, fields="files(id, name)").execute()
        items = results.get('files', [])
        return items
    except Exception as error:
        print(error)

def lista_pastas_subpastas_em_diretorio(folder_id):
    try:
        all_files = []
        folders_to_process = [folder_id]
    except Exception as error:
        print(error)

def procura_pasta_drive_por_nome(folder_name: str) -> list[dict] | None:
    """
    Procura pastas no Google Drive pelo nome fornecido.

    Args:
        folder_name (str): Nome da pasta a ser pesquisada. Caracteres '/' são tratados como separadores de diretórios e 
                           "S/S" é substituído por "S S" para evitar problemas na pesquisa.

    Returns:
        list[dict]: Lista de dicionários contendo os dados das pastas encontradas. Cada dicionário inclui pelo menos as chaves 'id' e 'name'.
        None: Caso nenhuma pasta seja encontrada.

    Note:
        - A função realiza uma busca pela pasta com o nome exato fornecido e retorna todas as pastas que correspondem ao nome especificado.
        - Se várias pastas com o mesmo nome forem encontradas, todas serão retornadas na lista.

    Examples:
        1. Busca de uma pasta pelo nome:
            >>> procura_pasta_drive_por_nome("Documentos")
        # A pasta buscada será "Documentos".
        
        2. Busca de uma pasta com o nome incluindo separadores de diretórios:
            >>> procura_pasta_drive_por_nome("Relatórios/2023")
        # A pasta buscada será "Relatórios/2023".
        
        3. Caso especial onde "S/S" é convertido para "S S":
            >>> procura_pasta_drive_por_nome("Empresa S/S/relatorios/2024")
        # A pasta buscada será "Empresa S S/relatorios/2024".
    """
    try:
        # Substitui "S/S" por "S S" e mantém a estrutura de diretórios com "/"
        folder_name = folder_name.replace("S/S", "S S")
        query = f"name = '{folder_name}' and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
        found_items = []
        page_token = None
        while True:
            # Executa a busca e coleta os resultados
            results = driver_service.files().list(
                q=query, 
                pageSize=80, 
                fields="files(id, name), nextPageToken", 
                pageToken=page_token
            ).execute()
            items = results.get('files', [])
            found_items.extend(items)
            page_token = results.get('nextPageToken')
            if not page_token:
                break

        if not found_items:
            print("Nenhuma pasta encontrada.")
            return None

        print("Pastas encontradas:")
        for item in found_items:
            print(f'{item["name"]} ({item["id"]})')
        return found_items

    except Exception as error:
        print(f"Erro ao procurar pasta no Google Drive: {error}")
        return None

def procura_subpasta_drive_por_nome(parent_folder_name: str, subfolder_name: str, intermediate_folders: list[str] =None, recursive=False) -> list[dict] | None:
    """
    Procura uma subpasta no Google Drive pelo nome da pasta pai e nome da subpasta.
    
    Args:
        parent_folder_name (str): Nome da pasta pai onde a busca começará. Caracteres '/' são tratados como separadores de diretórios.
        subfolder_name (str): Nome da subpasta a ser pesquisada. Caracteres '/' são tratados como separadores de diretórios.
        intermediate_folders (list of str, optional): Lista de nomes de diretórios intermediários a serem pesquisados sequencialmente antes da subpasta final. Cada nome deve corresponder a uma pasta que deve ser encontrada na ordem especificada. Caracteres '/' são tratados como separadores de diretórios.
        recursive (bool, optional): Se verdadeiro, a busca será feita recursivamente em todas as subpastas, não apenas no nível atual.

    Returns:
        list[dict]: Lista de dicionários contendo os dados das subpastas encontradas. Cada dicionário inclui pelo menos as chaves 'id' e 'name'.
        None: Caso nenhuma subpasta seja encontrada.

    Note:
        - Se `recursive` for verdadeiro, a função irá buscar a subpasta em todas as subpastas do diretório pai especificado ou da última pasta intermediária encontrada.
        - Se `intermediate_folders` for fornecido, a função procurará cada pasta intermediária na ordem fornecida antes de buscar pela subpasta final. Caso algum diretório intermediário não seja encontrado, a função retorna `None`.

    Examples:
        1. Busca simples (diretório pai é diretamente pesquisado):
            >>> procura_subpasta_drive_por_nome("Cliente", "2023")
        # A função irá buscar na pasta "Cliente/2023".
        
        2. Busca com diretórios intermediários (procura na pasta 'Cliente', em seguida 'Economia Mensal', e finalmente 'Relatórios' para encontrar '2023'):
            >>> procura_subpasta_drive_por_nome("Cliente", "2023", intermediate_folders=["Economia Mensal", "Relatórios"])
        # A função irá buscar na estrutura "Cliente/Economia Mensal/Relatórios/2023".
        
        3. Busca recursiva (procura '2023' em todas as subpastas a partir da pasta 'Cliente'):
            >>> procura_subpasta_drive_por_nome("Cliente", "2023", recursive=True)
        # A função irá buscar recursivamente na estrutura "Cliente/*/2023".
        
        4. Busca com diretórios intermediários e recursiva (procura '2023' em 'Cliente', 'Economia Mensal', e subpastas):
            >>> procura_subpasta_drive_por_nome("Cliente", "2023", intermediate_folders=["Economia Mensal"], recursive=True)
        # A função irá buscar na estrutura "Cliente/Economia Mensal/*/2023".
        
        5. Busca em pastas com nome contendo "S/S" (o caractere '/' é tratado como separador de diretórios e "S/S" é substituído por "S S"):
            >>> procura_subpasta_drive_por_nome("EMPRESA TESTE S/S", "2024/06/relatorios", intermediate_folders=["financeiro"])
        # A função irá buscar na estrutura "EMPRESA TESTE S S/financeiro/2024/06/relatorios".
    """
    try:
        # Substituir "S/S" por "S S" e tratar '/' como separadores de diretórios
        parent_folder_name = parent_folder_name.replace("S/S", "S S")
        subfolder_name = subfolder_name.replace("S/S", "S S")
        if intermediate_folders:
            intermediate_folders = [folder.replace("S/S", "S S") for folder in intermediate_folders]

        # Função auxiliar para buscar a pasta pai
        def find_parent_folder(parent_name):
            parent_name_parts = parent_name.split("/")
            parent_folder_id = None
            for part in parent_name_parts:
                query = f"name = '{part}' and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
                if parent_folder_id:
                    query = f"'{parent_folder_id}' in parents and {query}"
                results = driver_service.files().list(q=query, pageSize=80, fields="files(id, name)").execute()
                items = results.get('files', [])
                if not items:
                    return None
                parent_folder_id = items[0]['id']
            return items[0]

        # Primeiro, encontre a pasta pai pelo nome
        parent_folder = find_parent_folder(parent_folder_name)
        if not parent_folder:
            print("Pasta pai não encontrada.")
            return None

        parent_folder_id = parent_folder['id']

        # Se houver diretórios intermediários, procure dentro deles na sequência
        if intermediate_folders:
            for intermediate_folder_name in intermediate_folders:
                query = f"'{parent_folder_id}' in parents and name = '{intermediate_folder_name}' and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
                results = driver_service.files().list(q=query, pageSize=80, fields="files(id, name)").execute()
                items = results.get('files', [])
                if not items:
                    print(f"Diretório intermediário '{intermediate_folder_name}' não encontrado.")
                    return None

                intermediate_folder = items[0]
                parent_folder_id = intermediate_folder['id']

        # Se a busca for recursiva
        if recursive:
            # Função auxiliar para busca recursiva
            def busca_recursiva(folder_id, subfolder_name):
                found_items = []
                page_token = None
                while True:
                    query = f"'{folder_id}' in parents and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
                    results = driver_service.files().list(
                        q=query, pageSize=80, fields="files(id, name, parents)", pageToken=page_token
                    ).execute()
                    items = results.get('files', [])
                    for item in items:
                        if item['name'] == subfolder_name:
                            found_items.append(item)
                        # Busca recursiva nas subpastas
                        found_items.extend(busca_recursiva(item['id'], subfolder_name))
                    
                    page_token = results.get('nextPageToken')
                    if not page_token:
                        break
                return found_items

            return busca_recursiva(parent_folder_id, subfolder_name)

        # Caso não seja recursiva, procure a subpasta diretamente
        query = f"'{parent_folder_id}' in parents and name = '{subfolder_name}' and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
        results = driver_service.files().list(q=query, pageSize=80, fields="files(id, name)").execute()
        items = results.get('files', [])
        if not items:
            print("Nenhuma subpasta encontrada.")
            return None

        print("Subpastas encontradas:")
        for item in items:
            print(f'{item["name"]} ({item["id"]})')
        return items  # Retorna todos os resultados encontrados

    except Exception as error:
        print(f"Erro ao procurar subpasta no Google Drive: {error}")
        return None

# ==================== MÉTODOS DE CADA ETAPA DO PROCESSO========================
def gera_relatorio_dentistas_norte(mes, mes_nome, ano, dir_dentistas_norte_modelo, dir_dentistas_norte_destino):
    try:
        pythoncom.CoInitialize()
        dentistas_norte = procura_clientes_por_regiao("Manaus", db_conf)
        dentistas_norte.reverse()
        linha = 3
        indice = 1
        total = 0

        if dentistas_norte:
            dir_dentistas_norte_destino.mkdir(parents=True, exist_ok=True)
            nome_arquivo = f"Dentistas_Norte_{mes}_{ano}.xlsx"
            caminho_relatorio = f"{dir_dentistas_norte_destino}\\{nome_arquivo}"
            copy(dir_dentistas_norte_modelo, dir_dentistas_norte_destino / nome_arquivo)
            try:
                workbook, sheet, style_moeda = carrega_excel(f"{dir_dentistas_norte_destino}\\{nome_arquivo}")
                sheet.title = f"Dentistas_Norte_{mes}_{ano}"
                sheet['C2'] = mes_nome
                for cliente in dentistas_norte:
                    cliente_id = cliente[0]
                    cliente_nome = cliente[1]
                    valores = procura_valores(cliente_id, db_conf, mes, ano)
                    if valores:
                        economia_formal = valores[4]
                        total = economia_formal + total
                        sheet[f'C{linha}'].style = style_moeda
                        sheet[f'A{linha}'].border = Border(top=Side(style='thin'), bottom=Side(style='thin'), left=Side(style='thin'), right=Side(style='thin'))
                        sheet[f'B{linha}'].border = Border(top=Side(style='thin'), bottom=Side(style='thin'), left=Side(style='thin'), right=Side(style='thin'))
                        sheet[f'C{linha}'].border = Border(top=Side(style='thin'), bottom=Side(style='thin'), left=Side(style='thin'), right=Side(style='thin'))
                        sheet[f'A{linha}'] = f"{indice}"
                        sheet[f'B{linha}'] = f"{cliente_nome}"
                        sheet[f'C{linha}'] = economia_formal
                        linha += 1
                        indice += 1
                sheet[f'C{linha}'].style = style_moeda
                sheet[f'A{linha}'].border = Border(top=Side(style='thin'), bottom=Side(style='thin'), left=Side(style='thin'), right=Side(style='thin'))
                sheet[f'B{linha}'].border = Border(top=Side(style='thin'), bottom=Side(style='thin'), left=Side(style='thin'), right=Side(style='thin'))
                sheet[f'C{linha}'].border = Border(top=Side(style='thin'), bottom=Side(style='thin'), left=Side(style='thin'), right=Side(style='thin'))
                sheet[f'B{linha}'] = "Valor total da Economia Pós pagamento a Human"
                sheet[f'C{linha}'] = total
                workbook.save(caminho_relatorio)
                workbook.close()
                try:
                    excel = win32.gencache.EnsureDispatch('Excel.Application')
                    excel.Visible = True
                    wb = excel.Workbooks.Open(caminho_relatorio)
                    ws = wb.Worksheets[f"Dentistas_Norte_{mes}_{ano}"]
                    sleep(3)
                    ws.ExportAsFixedFormat(0, str(dir_dentistas_norte_destino) + f"\\Dentistas_Norte_{mes}_{ano}")
                    wb.Close()
                    excel.Quit()
                    print("Relatório Gerado!")
                except Exception as error:
                    print(error)
            except Exception as error:
                print(error)
    except Exception as error:
        print(error)
    finally:
        pythoncom.CoUninitialize()

def envia_email(dir_dentistas_norte_destino):
    try:
        emails_clientes = os.getenv('EMAILS_CLIENTES').split(",")
        corpo_email = os.getenv('CORPO_EMAIL_01')
        emails_formatado = []
        anexos = []
        for email in emails_clientes:
            emails_formatado.append(email.replace("\n", "").strip())
        arquivos = listagem_arquivos(dir_dentistas_norte_destino)
        for arquivo in arquivos:
            if arquivo.__contains__(".pdf"):
                anexos.append(arquivo)
                enviar_email_com_anexos(emails_formatado, "Relatório de Redução de Custos Trabalhistas Mensal - Grupo Dentistas do Norte", corpo_email, anexos)
        if anexos == []:
            print("Relatório não foi encontrado")
    except Exception as error:
        print(error)

def relatorio_economia_geral_mensal(mes, ano, particao, lista_dir_clientes, dir_economia_mensal_modelo):
    try:
        pythoncom.CoInitialize()
        workbook_emails, sheet_emails, style_moeda_emails = carrega_excel(f"{particao}:\\Meu Drive\\Arquivos_Automacao\\emails para envio relatorio human.xlsx") # TODO: PRECISA DOS EMAILS DE CADA CLIENTE
        ceo_email = os.getenv('CEO_EMAIL')
        corpo_email = os.getenv('CORPO_EMAIL_02')
        cliente_emails = [ceo_email]
        relatorio_enviado = False
        for row in sheet_emails.iter_rows(min_row=1, max_row=12, min_col=1, max_col=2):
            cliente = procura_cliente(row[0].value, db_conf)
            if cliente:
                cliente_id = cliente[0]
                cliente_nome = row[0].value
                cliente_emails.append(row[1].value)
                valores = procura_todos_valores_ano(cliente_id, db_conf, ano)
                if valores: # TODO: Ao invés de verificar somente se existem valores, verificar também se o cliente está ativo
                    valores.reverse()
                    for valor in valores:
                        if valor[6] == int(mes) and valor[7] == int(ano) and valor[8] == 1:
                            relatorio_enviado = True
                            break
                        else:
                            relatorio_enviado = False

                    if relatorio_enviado == False:
                        pasta_cliente = procura_pasta_drive_por_nome(cliente_nome)
                        if pasta_cliente != None:
                            pasta_cliente = pasta_cliente[0]
                        
                        if pasta_cliente:
                            pasta_economia_mensal = procura_subpasta_drive_por_nome(cliente_nome, ano, ["Economia Mensal"])
                            if pasta_economia_mensal != None:
                                pasta_economia_mensal = pasta_economia_mensal[0]
                            
                            caminho_arquivo_excel = f"/tmp/Economia_Mensal_{cliente_nome}_{ano}.xlsx"
                            
                            copy(dir_economia_mensal_modelo, caminho_arquivo_excel)
                            sleep(0.5)
                            workbook_economia, sheet_economia, style_moeda_economia = carrega_excel(caminho_arquivo_excel)
                            sheet_economia[f'C1'] = f"Relatorio demonstrativo de economia previdenciaria {ano}"
                            sheet_economia[f'C2'] = cliente_nome
                            for indice, valor in enumerate(valores):
                                sheet_economia['C4'].style = style_moeda_economia
                                sheet_economia['A4'].border = Border(top=Side(style='thin'), bottom=Side(style='thin'), left=Side(style='thin'))
                                sheet_economia['B4'].border = Border(top=Side(style='thin'), bottom=Side(style='thin'))
                                sheet_economia['C4'].border = Border(top=Side(style='thin'), bottom=Side(style='thin'))
                                sheet_economia['D4'].border = Border(top=Side(style='thin'), bottom=Side(style='thin'), left=Side(style='thin'))
                                sheet_economia['E4'].border = Border(top=Side(style='thin'), bottom=Side(style='thin'), right=Side(style='thin'))
                                mes_valor = calendar.month_name[int(valor[6])].capitalize()
                                sheet_economia['A4'] = f"{mes_valor}/{ano}"
                                sheet_economia['D4'] = valor[3]
                                if not indice == len(valores) - 1:
                                    sheet_economia.insert_rows(4)
                            for row in sheet_economia.iter_rows(min_row=1, min_col=1, max_col=5):
                                if row[0].value == "Total economia/ano":
                                    sheet_economia[f'D{row[0].row}'] = f"=SUM(D4:D{row[0].row - 1})"
                            workbook_economia.save(caminho_arquivo_excel)
                            workbook_economia.close()
                            try:
                                excel = win32.gencache.EnsureDispatch('Excel.Application')
                                excel.Visible = True
                                wb = excel.Workbooks.Open(caminho_arquivo_excel)
                                ws = wb.Worksheets[f"Página1"]
                                sleep(3)
                                ws.ExportAsFixedFormat(0, f"{pasta_economia_mensal}" + f"\\Economia_Mensal_{cliente_nome}_{ano}")
                                wb.Close()
                                excel.Quit()
                                print("Relatório Gerado!")
                            except Exception as error:
                                print(error)
                            sleep(0.5)
                            caminho_arquivo_pdf = [f"{pasta_economia_mensal}\\Economia_Mensal_{cliente_nome}_{ano}.pdf"]
                            sleep(0.5)
                            enviar_email_com_anexos(cliente_emails, f"Relatório demonstrativo de economia previdenciaria {ano}", corpo_email, caminho_arquivo_pdf)
                            query_atualiza_relatorios = ler_sql("sql/atualiza_relatorios_cliente.sql")
                            values_relatorio = (cliente_id, mes, ano)
                            with mysql.connector.connect(**db_conf) as conn, conn.cursor() as cursor:
                                cursor.execute(query_atualiza_relatorios, values_relatorio)
                                conn.commit()
                        else:
                            print("Pasta do cliente não encontrada")
                    else:
                        print(f"Relatório ja foi enviado para {cliente_nome}!")
                else:
                    print("Nenhum registro de valor encontrado para o cliente")
            else:
                print("Nenhum cliente encontrado")
            cliente_emails = []
            cliente_emails = [ceo_email]
        workbook_emails.close()
    except Exception as error:
        print(error)
    finally:
        pythoncom.CoUninitialize()


def lambda_handler(event, context):
    # parse body da requisição
    body = json.loads(event['body'])
    sleep(2)
    mes = body['mes']
    ano = body['ano']
    particao = body['particao']
    rotina = body['rotina']

    # ========================PARAMETROS INICIAS==============================
    clientes_itaperuna_id = os.getenv('CLIENTES_ITAPERUNA_FOLDER_ID')
    clientes_manaus_id = os.getenv('CLIENTES_MANAUS_FOLDER_ID')
    dentista_norte_id = os.get_env('DENTISTAS_NORTE_FOLDER_ID') # TODO: TEM QUE SER CRIADO

    arquivos_itaperuna = lista_pastas_em_diretorio(clientes_itaperuna_id)
    arquivos_manaus = lista_pastas_em_diretorio(clientes_manaus_id)
    arquivos_dentistas_norte = lista_pastas_em_diretorio(dentista_norte_id)

    lista_dir_clientes = arquivos_itaperuna + arquivos_manaus

    dir_dentistas_norte_destino = Path(f"{particao}:\\Meu Drive\\Relatorio_Dentista_do_Norte\\{mes}-{ano}") # TODO: TEM QUE DESCOBRIR COMO CHEGAR NESSE CAMINHO OU CRIAR CASO NÃO EXISTA
    dir_dentistas_norte_modelo = Path(f"templates\\dentistas_norte_modelo_00_0000_python.xlsx") # TODO: TEM QUE MOVER PARA A PASTA TEMPLATES
    dir_economia_mensal_modelo = Path(f"templates\\modelo_relatorio_demonstrativo_economia_previdencia.xlsx") # TODO: TEM QUE MOVER PARA A PASTA TEMPLATES
    mes_nome = calendar.month_name[int(mes)].capitalize()
    sucesso = False

    # ========================LÓGICA DE EXECUÇÃO DO ROBÔ===========================
    if rotina == "1. Gerar Relatorio Dentista do Norte":
        gera_relatorio_dentistas_norte(mes, mes_nome, ano, dir_dentistas_norte_modelo, dir_dentistas_norte_destino)
        envia_email(dir_dentistas_norte_destino)
        relatorio_economia_geral_mensal(mes, ano, particao, lista_dir_clientes, dir_economia_mensal_modelo)
        sucesso = True
    elif rotina == "2. Enviar Email":
        envia_email(dir_dentistas_norte_destino)
        relatorio_economia_geral_mensal(mes, ano, particao, lista_dir_clientes, dir_economia_mensal_modelo)
        sucesso = True
    elif rotina == "3. Relatorio Economia Geral Mensal":
        relatorio_economia_geral_mensal(mes, ano, particao, lista_dir_clientes, dir_economia_mensal_modelo)
        sucesso = True
    else:
        print("Nenhuma rotina selecionada, encerrando o robô...")
        sucesso = False

    if sucesso:
      return {'message': 'Relatorios gerados com sucesso'}, 200
    else:
      return {'message': 'Erro ao gerar relatorios'}, 500