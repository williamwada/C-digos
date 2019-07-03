
--Desligando o database para nao haver inputs ou updats de dados
ALTER DATABASE IBGE SET OFFLINE WITH ROLLBACK IMMEDIATE;



-- Transferindo o file de dados do database Imports para outro drive
USE master; --do this all from the master
ALTER DATABASE IBGE
MODIFY FILE (name='teste_P1'
             ,filename='E:\DB_Files\teste_P1.mdf'); --Filename is new location

-- Transferindo o file de logs do database Imports para outro drive
USE master; --do this all from the master
ALTER DATABASE IBGE
MODIFY FILE (name='teste_P1_log'
             ,filename='E:\DB_Files\teste_P1_log.ldf'); --Filename is new location

--Religando o database
ALTER DATABASE IBGE SET ONLINE;