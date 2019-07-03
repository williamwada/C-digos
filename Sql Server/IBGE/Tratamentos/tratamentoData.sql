
-- convertendo string para date 
select CONVERT(date,replace(D3N,'janeiro ','01/01/')),D3N from teste_API.dbo.t_stage_IPCA where D3N like 'janeiro%'

select format(CONVERT(date,replace(D3N,'fevereiro ','02/01/')),'MMM/yyyy'),D3N from teste_API.dbo.t_stage_IPCA where D3N like 'fevereiro%'
select replace(D3N,'fevereiro ','02/01/'),D3N from teste_API.dbo.t_stage_IPCA where D3N like 'fevereiro%'


UPDATE tb_stage_ipca
SET D3N = CASE
WHEN D3N like 'janeiro%' THEN replace(D3N,'janeiro ','01/01/')
WHEN D3N like 'fevereiro%' THEN replace(D3N,'fevereiro ','02/01/')
WHEN D3N like 'março%' THEN replace(D3N,'março ','03/01/')
WHEN D3N like 'abril%' THEN replace(D3N,'abril ','04/01/')
WHEN D3N like 'maio%' THEN replace(D3N,'maio ','05/01/')
WHEN D3N like 'junho%' THEN replace(D3N,'junho ','06/01/')
WHEN D3N like 'julho%' THEN replace(D3N,'julho ','07/01/')
WHEN D3N like 'agosto%' THEN replace(D3N,'agosto ','08/01/')
WHEN D3N like 'setembro%' THEN replace(D3N,'setembro ','09/01/')
WHEN D3N like 'outubro%' THEN replace(D3N,'outubro ','10/01/')
WHEN D3N like 'novembro%' THEN replace(D3N,'novembro ','11/01/')
WHEN D3N like 'dezembro%' THEN replace(D3N,'dezembro ','12/01/')
END;

select * from tb_stage_ipca
order by D3N