
--Captura do estado
select trim(SUBSTRING(D1N,CHARINDEX('-',D1N,0)+2,LEN(D1N))),D1N from t_stage_IPCA

--Captura da cidade
select trim(SUBSTRING(D1N,1,CHARINDEX('-',D1N,0)-1)),D1N from t_stage_IPCA

