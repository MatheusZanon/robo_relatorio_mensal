from components.configuracao_db import ler_sql
import mysql.connector

def procura_cliente(nome_cliente, db_conf):
    try:
        query_procura_cliente = ler_sql('sql/procura_cliente.sql')
        values_procura_cliente = (nome_cliente,)
        with mysql.connector.connect(**db_conf) as conn, conn.cursor() as cursor:
            cursor.execute(query_procura_cliente, values_procura_cliente)
            cliente = cursor.fetchone()
            conn.commit()
        if cliente:
            return cliente
        else:
            cliente_mod = procura_cliente_mod(str(nome_cliente).replace("S S", "S/S"), db_conf)
            return cliente_mod
    except Exception as error:
        print(error)

def procura_cliente_mod(nome_cliente, db_conf):
    try:
        query_procura_cliente = ler_sql('sql/procura_cliente.sql')
        values_procura_cliente = (nome_cliente,)
        with mysql.connector.connect(**db_conf) as conn, conn.cursor() as cursor:
            cursor.execute(query_procura_cliente, values_procura_cliente)
            cliente = cursor.fetchone()
            conn.commit()
        if cliente:
            return cliente
    except Exception as error:
        print(error)

def procura_clientes_por_regiao(regiao, db_conf):
    try:
        query_procura_cliente = ler_sql('sql/procura_clientes_por_regiao.sql')
        values_procura_cliente = (regiao,)
        with mysql.connector.connect(**db_conf) as conn, conn.cursor() as cursor:
            cursor.execute(query_procura_cliente, values_procura_cliente)
            clientes = cursor.fetchall()
            conn.commit()
        if clientes:
            return clientes
        else:
            return None
    except Exception as error:
        print(error)
