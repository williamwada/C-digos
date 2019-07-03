

--contar quantos decimais tem apos o ponto
-- verificar maximo de decimais
select LEN(SUBSTRING(V,CHARINDEX('.',V,0)+1,LEN(V))),V from teste_API.dbo.t_stage_IPCA

-- Retorna decimal maximo na tabela
--maximo retorno 4
select max(LEN(SUBSTRING(V,CHARINDEX('.',V,0)+1,LEN(V)))) from teste_API.dbo.t_stage_IPCA


-- Retorna numerico maximo
-- maximo retorno 3
select max(LEN(SUBSTRING(V,0,CHARINDEX('.',V,0)))) from teste_API.dbo.t_stage_IPCA

