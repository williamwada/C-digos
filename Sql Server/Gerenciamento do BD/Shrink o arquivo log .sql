USE BACEN;  
GO  

SELECT (total_log_size_in_bytes - used_log_space_in_bytes)*1.0/1024/1024/1024 AS [free log space in GB]  
FROM sys.dm_db_log_space_usage;  

--Checar o nome do arquivo de log
SELECT name FROM sys.master_files WHERE type_desc = 'LOG'

-- Modificar para o recovery simple para realizar o shrink.
ALTER DATABASE BACEN
        SET RECOVERY SIMPLE
        GO

-- Realizar a compressao do arquivo log
DBCC SHRINKFILE (BACEN_log, 1)
GO

--retornar em modo FULL recovery
ALTER DATABASE BACEN
SET RECOVERY FULL
        

