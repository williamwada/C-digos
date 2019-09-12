--
--Como Migrar as bases de dados para outro local(drive)
--

--1. deixar o database offline
ALTER DATABASE TEST SET OFFLINE;  

--2. Apontar para o novo path 
ALTER DATABASE TEST MODIFY FILE ( NAME = TEST, FILENAME = 'D:\DataBases\TEST.mdf' );

--3. Apontar o arquivo Log para o novo path
ALTER DATABASE TEST MODIFY FILE ( NAME = TEST_log, FILENAME = 'D:\DataBases\TEST_log.ldf' );

--4. Colocar em modo online o Db (Caso de acesso negado a pasta dar grant full control no arquivo db e log para system
ALTER DATABASE TEST SET ONLINE;  

--5. Verificar o novo endereco
SELECT name, physical_name AS CurrentLocation, state_desc  
FROM sys.master_files  
WHERE database_id = DB_ID(N'TEST');  

