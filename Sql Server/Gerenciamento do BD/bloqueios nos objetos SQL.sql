
--verificar se algum objeto esta bloqueado
exec sp_Lock

--verificar quem esta fazendo o bloqueio. 
exec sp_Who2 

-- select que mostra qual Sid esta bloqueando para fazer o kill
select spid, blocked, hostname=left(hostname,20), program_name=left(program_name,20),
       WaitTime_Seg = convert(int,(waittime/1000))  ,open_tran, status
From master.dbo.sysprocesses 
where blocked > 0
order by spid

--mata a sessao liberando entao o objeto.
--kill 52

commit;

-- select que mostra qual Sid esta bloqueando para fazer o kill
SELECT conn.session_id, host_name, program_name,
    nt_domain, login_name, connect_time, last_request_end_time 
FROM sys.dm_exec_sessions AS sess
JOIN sys.dm_exec_connections AS conn
   ON sess.session_id = conn.session_id;