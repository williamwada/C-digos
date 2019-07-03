@ECHO OFF
 
 REM IPCA_MENSAL_VARIACAO
Set Server=WILLIAM-T\LCA_P1
Set Username=sistemaAPI
Set Password=lca12345
Set Arq_header="c:\temp\HeadersOnly.csv"
Set Arquivo="c:\temp\TableDataWithoutHeaders.csv"
Set Saida = "c:\temp\TableData.csv"
Set Log="c:\temp\tabela.log_exp"
Set Query="select * from [IBGE].[dbo].[view_IPCA_Mensal_Variacao] order by CODIGO, REFERENCIA"
Set BCP_EXPORT_DB=IBGE
Set BCP_EXPORT_TABLE=view_IPCA_Mensal_Variacao

ECHO Inicio do BCP...: %TIME%
ECHO Aguarde a carga dos dados...

BCP "DECLARE @colnames VARCHAR(max);SELECT @colnames = COALESCE(@colnames + ';', '') + column_name from %BCP_EXPORT_DB%.INFORMATION_SCHEMA.COLUMNS where TABLE_NAME='%BCP_EXPORT_TABLE%'; select @colnames;" queryout %Arq_header% -o %Log% -U%Username% -P%Password% -S%Server% -c -C "ACP" -t ; -w

BCP %Query% queryout %Arquivo% -o %Log% -U%Username% -P%Password% -S%Server% -c -C "ACP" -t ; -w 

copy /b %Arq_header%+%Arquivo% c:\temp\IPCA_Mensal_Variacao.csv

REM del %Arq_header%
REM del %Arquivo% 

 
ECHO Termino do BCP..: %TIME%
ECHO Log no arquivo %Log%
 


 
PAUSE