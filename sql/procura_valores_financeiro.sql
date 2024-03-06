SELECT 
cliente_id, 
soma_salarios_provdt,
percentual_human,
economia_mensal,
economia_formal,
total_fatura
FROM clientes_financeiro_valores WHERE 
cliente_id = %s AND mes = %s AND ano = %s