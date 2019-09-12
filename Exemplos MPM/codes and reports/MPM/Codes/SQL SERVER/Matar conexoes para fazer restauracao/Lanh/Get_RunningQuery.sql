SELECT er.session_id as SPID
	,ses.STATUS AS STATUS
	,ses.login_name AS [Login] 
	,ses.host_name AS Host 
	,er.blocking_session_id AS BlkBy 
	,DB_Name(er.database_id) AS DBName 
	,er.command AS CommandType 
	,OBJECT_NAME(st.objectid) AS ObjectName
	,er.cpu_time AS CPUTime 
	,er.start_time AS StartTime 
	,CAST(GETDATE() - er.start_time AS TIME) AS TimeElapsed
	,st.text AS SQLStatement 
FROM    sys.dm_exec_requests er
    OUTER APPLY sys.dm_exec_sql_text(er.sql_handle) st
    LEFT JOIN sys.dm_exec_sessions ses
    ON ses.session_id = er.session_id
LEFT JOIN sys.dm_exec_connections con
    ON con.session_id = ses.session_id
WHERE   st.text IS NOT NULL