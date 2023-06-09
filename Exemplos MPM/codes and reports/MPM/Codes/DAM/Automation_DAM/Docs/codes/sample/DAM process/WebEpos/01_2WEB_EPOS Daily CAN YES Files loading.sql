

/**
	Daily GAN (Cancer) WEB EPOS YES File loading.
	This program would looking for YES file in the staging area
	and load it to database then move file to historic folder.
	Copy YES file to D:\EPOS\DATA_AREA\STAGING_AREA\WEB_YES_File before run this.
	
-- =============================================
-- Author:		<William Wada>
-- Create date: <07 Dec 2015>
-- Description:	<Importing WEB EPOS YES FILE file daily>
-- =============================================

	
**/

USE EPOS
GO

declare @staging varchar(100)
declare @yes varchar(255)
declare @yesfullpath varchar(255)
declare @yes_exists int
declare @wkday int
declare @fcounter int
declare @fend int
/*
-- first of all , make a database backup
declare @dbbackup varchar(255)
set @dbbackup = 'D:\Dev_EPOS\Data base\Backup\WEB_EPOS_' + convert(varchar,getdate(),112) + '.bak';

EXEC master..xp_fileexist @dbbackup, @yes_exists output

backup database Dev_EPOS to disk = @dbbackup
*/
set @staging = 'D:\EPOS\DATA_AREA\STAGING_AREA\WEB_YES_File\'
set @wkday = DATEPART(WEEKDAY,GETDATE()) -1

set @fend = 0
set @fcounter = -24


while @fcounter < @fend
begin

	set @yes = 'YES-EPWEBCAN' + convert(varchar,DATEADD(day,@fcounter,getdate()),112) + '.txt'
	set @yesfullpath = @staging + @yes

	-- check file exists 
	EXEC master..xp_fileexist @yesfullpath, @yes_exists output

	if @yes_exists = 1
	begin
	-- file exists then load it
		declare @return_value int
		declare @loaddate date
		set @loaddate = getdate()
		exec @return_value = WEB_P1_CAN_YES_Import @yesfullpath
		if @return_value = 0
			select 'YES File Loading with '+@yes+' completed.' as 'YES Load Result'
		else
			select 'YES File Loading with '+@yes+' failed!' as 'YES Load Result'
			
		declare @cmd varchar(1024)
		set @cmd = 'COPY /Y '+ @yesfullpath + ' /B ' + @staging + 'HISTORIC\'+@yes

		-- copy file to historic folder
		EXEC master..xp_cmdshell @cmd, no_output
		-- then delete the original one
		set @cmd = 'DEL /F ' + @yesfullpath

		EXEC master..xp_cmdshell @cmd, no_output
		
	end
	
	set @fcounter = @fcounter + 1
	
end

--select * from WEB_EPOS_CAN_YES_FILE
--where YES_FILE_LOAD_DATE = cast(GETDATE() as DATE);