USE [DDS_EPOS_PRD]
GO
/****** Object:  StoredProcedure [dbo].[Daily_BackUp_EPOS_PRD]    Script Date: 05/30/2017 11:14:44 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

ALTER PROCEDURE [dbo].[Daily_BackUp_EPOS_PRD] 
AS
BEGIN

	-- first of all , make a database backup
	declare @dbbackup varchar(255)
	declare @bk_flg int
	--Added on 2014/12/05
	declare @zipapp varchar(100)
	declare @result int
	declare @cmd varchar(1000)
	set @zipapp = 'D:\EPOS_PRD\DATABASE\backup\7za920\7za.exe'
	--
	set @dbbackup = 'D:\EPOS_PRD\DATABASE\backup\DDS_EPOS_PRD_' + CONVERT(VARCHAR(8), GETDATE(), 112) + REPLACE(CONVERT(varchar, GETDATE(), 108), ':','') + '.bak';

	EXEC master..xp_fileexist @dbbackup,@bk_flg output

	if @bk_flg = 0
	begin
		backup database [DDS_EPOS_PRD] to disk = @dbbackup
		--Added by R.J. 2014-12-05
		--Zip bak file
		set @cmd = @zipapp + ' a ' + @dbbackup + '.zip ' + @dbbackup;

		EXEC @result = xp_cmdshell @cmd
		if @result = 0
		begin
			set @cmd = 'del ' + @dbbackup
			EXEC xp_cmdshell @cmd
		end
		--End
	end

END



