-- Verificar espaco utilizado pelo database
Use teste_P1
go
sp_spaceused 

--verificar o espaco utilizado pelas tabelas de determinado database
exec sp_msforeachtable 'sp_spaceused ''?'',true'