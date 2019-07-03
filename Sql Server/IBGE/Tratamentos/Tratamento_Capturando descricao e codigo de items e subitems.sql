
--capturando os codigos de items no campo D4N e descricao
select * from teste_API.dbo.t_stage_IPCA


--descricao
select substring(D4N,CHARINDEX('.',D4N,0)+1,LEN(D4N)),D4N from IBGE.dbo.tb_stage_IPCA


--Codigo
select substring(D4N,0,CHARINDEX('.',D4N,0)),D4N from dbo.tb_stage_IPCA



--ordenacao categorias
select substring(D4N,0,CHARINDEX('.',D4N,0)),substring(D4N,CHARINDEX('.',D4N,0)+1,LEN(D4N)) 
from dbo.tb_stage_IPCA
where D4N not like '%geral%'
group by substring(D4N,0,CHARINDEX('.',D4N,0)),substring(D4N,CHARINDEX('.',D4N,0)+1,LEN(D4N))
order by substring(D4N,0,CHARINDEX('.',D4N,0))


--ordernacao por regiao
select convert(int,D1C),D1N,count(1) 
from tb_stage_ipca
group by D1C,D1N
order by convert(int,D1C)

select * from tb_stage_ipca

