
declare @files table (ID int IDENTITY, FileName varchar(100))
declare @num int
insert into @files execute xp_cmdshell 'dir c:\test\ /b'
select @num = count(*)
from @files
where FileName like '%MUN%'

print @num
IF @num > 0
begin;

select 'Arquivo incorreto' as 'ERROR Message'

return;
end;

select *
from @files