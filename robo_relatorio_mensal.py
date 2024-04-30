# =========================IMPORTAÇÕES DE BIBLIOTECAS E COMPONENTES========================
from components.importacao_diretorios_windows import listagem_arquivos, procura_pasta_cliente
from components.importacao_caixa_dialogo import DialogBox
from components.checar_ativacao_google_drive import checa_google_drive
from components.configuracao_db import configura_db, ler_sql
from components.importacao_automacao_excel_openpyxl import carrega_excel
from components.procura_cliente import procura_cliente, procura_clientes_por_regiao
from components.procura_valores import procura_valores, procura_todos_valores_ano
from components.enviar_emails import enviar_email_com_anexos
import mysql.connector
import tkinter as tk
from pathlib import Path
from openpyxl.styles import Border, Side, NamedStyle
from shutil import copy
import win32com.client as win32
from dotenv import load_dotenv
import os
from time import sleep
import locale
import calendar
import pandas as pd

locale.setlocale(locale.LC_ALL, 'pt_BR.utf8')
load_dotenv()

# =====================CONFIGURAÇÂO DO BANCO DE DADOS======================
db_conf = configura_db()

# =============CHECANDO SE O GOOGLE FILE STREAM ESTÁ INICIADO NO SISTEMA==============
checa_google_drive()

def main():
    root = tk.Tk()
    app = DialogBox(root)
    root.mainloop()
    return app.particao, app.rotina, app.mes, app.ano


if __name__ == "__main__":
    particao, rotina, mes, ano = main()

mes_nome = calendar.month_name[int(mes)].capitalize()

# ========================ARQUIVOS INICIAS==============================
dir_clientes_itaperuna = f"{particao}:\\Meu Drive\\Cobranca_Clientes_terceirizacao\\Clientes Itaperuna"
dir_clientes_manaus = f"{particao}:\\Meu Drive\\Cobranca_Clientes_terceirizacao\\Clientes Manaus"
lista_dir_clientes = [dir_clientes_itaperuna, dir_clientes_manaus]
dir_relatorio_926 = f"{particao}:\\Meu Drive\\Relatorio_Human_9.26_Direitos_Trabalhistas\\{ano}\\Relatório {ano} Human - 9,26_ Direitos Trabalhistas.xlsx"
dir_relatorio_taxa_adm = f"{particao}:\\Meu Drive\\Relatorio_Taxa_Administracao\\{ano}\\Taxa Administração {ano} Human.xlsx"
dir_relatorio_economia_manaus = f"{particao}:\\Meu Drive\\Relatorio_Economia_Mensal_Manaus\\{ano}\\Relatorio Economia Mensal Manaus {ano}.xlsx"
dir_dentistas_norte_modelo = Path(f"{particao}:\\Meu Drive\\Arquivos_Automacao\\Dentistas_Norte_Modelo_00_0000_python.xlsx")
dir_dentistas_norte_destino = Path(f"{particao}:\\Meu Drive\\Relatorio_Dentista_do_Norte\\{mes}-{ano}")
dir_economia_mensal_modelo = Path(f"{particao}:\\Meu Drive\\Arquivos_Automacao\\modelo relatorio demonstrativo economia previdencia.xlsx")

# ==================== MÉTODOS DE AUXÍLIO====================================
def busca_excel(row, mes):
    try:
        linha = row[0].row
        coluna_atualizar = ("D" if int(mes) == 1 
                            else "E" if int(mes) == 2 
                            else "F" if int(mes) == 3 
                            else "G" if int(mes) == 4
                            else "H" if int(mes) == 5
                            else "I" if int(mes) == 6
                            else "J" if int(mes) == 7
                            else "K" if int(mes) == 8
                            else "L" if int(mes) == 9
                            else "M" if int(mes) == 10
                            else "N" if int(mes) == 11
                            else "O" if int(mes) == 12
                            else "")  
        celula_atualizar = f"{coluna_atualizar}{linha}"
        return celula_atualizar
    except Exception as error:
        print(error)

# ==================== MÉTODOS DE CADA ETAPA DO PROCESSO=======================
def relatorio_926():
    try:
        workbook, sheet, style_moeda = carrega_excel(dir_relatorio_926)
        for row in sheet.iter_rows(min_row=4, max_row=61, min_col=3, max_col=14):
            cliente = procura_cliente(row[0].value, db_conf)
            if cliente:
                cliente_id = cliente[0]
                valores = procura_valores(cliente_id, db_conf, mes, ano)
                if valores:
                    providencia_dt = round(valores[1] * 0.0926, 2)
                    celula_atualizar = busca_excel(row, mes)
                    sheet[celula_atualizar] = providencia_dt
                else:
                    print("Nenhum registro de valor encontrado para o cliente")
            else:
                print("Nenhum cliente encontrado")
        workbook.save(dir_relatorio_926)
        workbook.close()
    except Exception as error:
        print(error)

def relatorio_taxa_adm():
    try:
        workbook, sheet, style_moeda = carrega_excel(dir_relatorio_taxa_adm)

        for row in sheet.iter_rows(min_row=4, max_row=61, min_col=3, max_col=14):
            cliente = procura_cliente(row[0].value, db_conf)
            if cliente:
                cliente_id = cliente[0]
                valores = procura_valores(cliente_id, db_conf, mes, ano)
                if valores:
                    percentual_human = valores[2]
                    celula_atualizar = busca_excel(row, mes)
                    sheet[celula_atualizar] = percentual_human
                else:
                    print("Nenhum registro de valor encontrado para o cliente")
            else:
                print("Nenhum cliente encontrado")
        workbook.save(dir_relatorio_taxa_adm)
        workbook.close()
    except Exception as error:
        print(error)

def relatorio_economia_manaus():
    try:
        workbook, sheet, style_moeda = carrega_excel(dir_relatorio_economia_manaus)

        for row in sheet.iter_rows(min_row=5, max_row=44, min_col=3, max_col=14):
            print(row[0].value)
            cliente = procura_cliente(row[0].value, db_conf)
            if cliente:
                print(cliente)
                cliente_id = cliente[0]
                valores = procura_valores(cliente_id, db_conf, mes, ano)
                if valores:
                    economia_formal = round(valores[3] - valores[2], 2)
                    celula_atualizar = busca_excel(row, mes)
                    sheet[celula_atualizar] = economia_formal
                    query_economia_formal = ler_sql('sql/registra_economia_formal.sql')
                    values = (economia_formal, cliente_id)
                    with mysql.connector.connect(**db_conf) as conn, conn.cursor() as cursor:
                        cursor.execute(query_economia_formal, values)
                        conn.commit()
                else:
                    print("Nenhum registro de valor encontrado para o cliente")
            else:
                print("Nenhum cliente encontrado")
        workbook.save(dir_relatorio_economia_manaus)
        workbook.close()
    except Exception as error:
        print(error)

def gera_relatorio_dentistas_norte():
    try:
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

def envia_email():
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
                print(anexos)
                input("Pressione ENTER para enviar o e-mail")
                enviar_email_com_anexos(emails_formatado, "Relatório de Redução de Custos Trabalhistas Mensal - Grupo Dentistas do Norte", corpo_email, anexos)
        if anexos == []:
            print("Relatório não foi encontrado")
    except Exception as error:
        print(error)

def relatorio_economia_geral_mensal():
    try:
        workbook_emails, sheet_emails, style_moeda_emails = carrega_excel(f"{particao}:\\Meu Drive\\Arquivos_Automacao\\emails para envio relatorio human.xlsx")
        cliente_emails = ["victor.pena@acpcontabil.com.br"]
        corpo_email = os.getenv('CORPO_EMAIL_02')
        relatorio_enviado = False
        for row in sheet_emails.iter_rows(min_row=1, max_row=12, min_col=1, max_col=2):
            cliente = procura_cliente(row[0].value, db_conf)
            if cliente:
                cliente_id = cliente[0]
                cliente_nome = row[0].value
                cliente_emails.append(row[1].value)
                valores = procura_todos_valores_ano(cliente_id, db_conf, ano)
                if valores:
                    valores.reverse()
                    for valor in valores:
                        if valor[6] == int(mes) and valor[7] == int(ano) and valor[8] == 1:
                            relatorio_enviado = True
                            break
                        else:
                            relatorio_enviado = False
                    print(cliente_emails)
                    if relatorio_enviado == False:
                        pasta_cliente = procura_pasta_cliente(cliente_nome, lista_dir_clientes)
                        if pasta_cliente:
                            pasta_economia_mensal = Path(f"{pasta_cliente}\\Economia Mensal\\{ano}")
                            caminho_arquivo_excel = f"{pasta_economia_mensal}\\Economia_Mensal_{cliente_nome}_{ano}.xlsx"
                            copy(dir_economia_mensal_modelo, pasta_economia_mensal / f"Economia_Mensal_{cliente_nome}_{ano}.xlsx")
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
                                    print(f"Elemento atual: {valor}")
                                else:
                                    print(f"Último elemento: {valor}")
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
                            input("Pressione ENTER para enviar o e-mail")
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
            cliente_emails = ["victor.pena@acpcontabil.com.br"]
        workbook_emails.close()
    except Exception as error:
        print(error)

# ========================LÓGICA DE EXECUÇÃO DO ROBÔ===========================
if rotina == "1. Relatorio 9.26%":
    relatorio_926()
    relatorio_taxa_adm()
    relatorio_economia_manaus()
    gera_relatorio_dentistas_norte()
    envia_email()
    relatorio_economia_geral_mensal()
elif rotina == "2. Relatorio de Taxa ADM":
    relatorio_taxa_adm()
    relatorio_economia_manaus()
    gera_relatorio_dentistas_norte()
    envia_email()
    relatorio_economia_geral_mensal()
elif rotina == "3. Relatório Economia de Manaus":
    relatorio_economia_manaus()
    gera_relatorio_dentistas_norte()
    envia_email()
    relatorio_economia_geral_mensal()
elif rotina == "4. Gerar Relatório Dentista do Norte":
    gera_relatorio_dentistas_norte()
    envia_email()
    relatorio_economia_geral_mensal()
elif rotina == "5. Enviar Email":
    envia_email()
    relatorio_economia_geral_mensal()
elif rotina == "6. Relatorio Economia Geral Mensal":
    relatorio_economia_geral_mensal()
else:
    print("Nenhuma rotina selecionada, encerrando o robô...")