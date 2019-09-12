declare @DBName as varchar(20)
set @DBName = (select DB_NAME(db_id())) --'DDS_MUNNK_PRD'

SELECT  spid,
        sp.[status],
        loginame [Login],
        hostname, 
        blocked BlkBy,
        sd.name DBName, 
        cmd Command,
        cpu CPUTime,
        physical_io DiskIO,
        last_batch LastBatch,
        [program_name] ProgramName   
FROM master.dbo.sysprocesses sp 
JOIN master.dbo.sysdatabases sd ON sp.dbid = sd.dbid
where sd.name=@DBName
ORDER BY spid ;


