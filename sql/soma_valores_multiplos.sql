SELECT 
cliente_id, 
ROUND(SUM(soma_salarios_provdt), 2) AS soma_salarios_provdt,
ROUND(SUM(percentual_human), 2) AS percentual_human,
ROUND(SUM(economia_mensal), 2) AS economia_mensal,
ROUND(SUM(economia_liquida), 2) AS economia_liquida,
ROUND(SUM(total_fatura), 2) AS total_fatura
FROM clientes_financeiro_valores WHERE cliente_id = %s AND mes = %s and ano = %s
GROUP BY cliente_id;