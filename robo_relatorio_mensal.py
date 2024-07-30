# =========================IMPORTAÇÕES DE BIBLIOTECAS E COMPONENTES========================
from components.importacao_diretorios_windows import listagem_arquivos, procura_pasta_cliente
from components.importacao_caixa_dialogo import DialogBox
from components.checar_ativacao_google_drive import checa_google_drive
from components.configuracao_db import configura_db, ler_sql
from components.importacao_automacao_excel_openpyxl import carrega_excel
from components.procura_cliente import procura_cliente, procura_clientes_por_regiao
from components.procura_valores import procura_valores, procura_todos_valores_ano
from components.enviar_emails import enviar_email_com_anexos
from components.google_drive import procura_subpasta_drive_por_nome, lista_pastas_em_diretorio, procura_pasta_drive_por_nome
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

locale.setlocale(locale.LC_ALL, 'pt_BR.utf8')
load_dotenv()

# =====================CONFIGURAÇÂO DO BANCO DE DADOS===========================
db_conf = configura_db()

# =============CHECANDO SE O GOOGLE FILE STREAM ESTÁ INICIADO NO SISTEMA========
checa_google_drive()

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
                cliente_ativo = cliente[7]
                valores = procura_todos_valores_ano(cliente_id, db_conf, ano)
                if valores and cliente_ativo:
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
    dir_dentistas_norte_modelo = Path(f"templates\\dentistas_norte_modelo_00_0000_python.xlsx")
    dir_economia_mensal_modelo = Path(f"templates\\modelo_relatorio_demonstrativo_economia_previdencia.xlsx")
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