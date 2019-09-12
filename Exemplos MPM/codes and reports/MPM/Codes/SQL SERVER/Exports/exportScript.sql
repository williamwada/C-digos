declare @sql nvarchar(255)

set @sql = 'bcp.exe Shinkin.dbo.V_ADDRESS_DIV OUT '+
		'D:\Shinkin\DATA_AREA\OUTPUT_AREA\F2\TEST_' + convert(varchar,GETDATE(),112) +
		'.csv -c -t "|" -S  ADCF-PCKT\ADCFSQL -T'

EXECUTE xp_cmdshell @sql


