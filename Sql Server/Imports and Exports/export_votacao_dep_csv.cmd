@ECHO OFF
 
 REM CAMARA DEPUTADOS VOTACAO
Set Server=WWA\LCA_P1
Set Username=sistemaAPI
Set Password=Lc@12345
Set Arq_header="C:\Users\william.wada\Desktop\OneDrive - LCA Consultores\Tarefas\Camara\output\csv\HeadersOnly.csv"
Set Arquivo="C:\Users\william.wada\Desktop\OneDrive - LCA Consultores\Tarefas\Camara\output\csv\TableDataWithoutHeaders.csv"
Set Saida = "C:\Users\william.wada\Desktop\OneDrive - LCA Consultores\Tarefas\Camara\output\csv\Votacao_deputados.csv"
Set Log="C:\Users\william.wada\Desktop\OneDrive - LCA Consultores\Tarefas\Camara\output\csv\log\tabela.log_exp"
Set Query="select * from Camara_Dep.dbo.v_dados_votacao"
Set BCP_EXPORT_DB=Camara_Dep
Set BCP_EXPORT_TABLE=v_dados_votacao

ECHO Inicio do BCP...: %TIME%
ECHO Aguarde a carga dos dados...

REM BCP "DECLARE @colnames VARCHAR(max);SELECT @colnames = COALESCE(@colnames + ';', '') + column_name from %BCP_EXPORT_DB%.INFORMATION_SCHEMA.COLUMNS where TABLE_NAME='%BCP_EXPORT_TABLE%'; select @colnames;" queryout %Arq_header% -o %Log% -U%Username% -P%Password% -S%Server% -c -C "ACP" -t \t -w

BCP "DECLARE @colnames VARCHAR(max);SELECT @colnames = COALESCE(@colnames + '	', '') + column_name from %BCP_EXPORT_DB%.INFORMATION_SCHEMA.COLUMNS where TABLE_NAME='%BCP_EXPORT_TABLE%'; select @colnames;" queryout %Arq_header% -o %Log% -U%Username% -P%Password% -S%Server% -c -C "ACP" -t \t -w

BCP %Query% queryout %Arquivo% -o %Log% -U%Username% -P%Password% -S%Server% -c -C "ACP" -t \t -w 

copy /b %Arq_header%+%Arquivo% "C:\Users\william.wada\Desktop\OneDrive - LCA Consultores\Tarefas\Camara\output\csv\Votacao_deputados.csv"

del %Arq_header%
del %Arquivo% 

 
ECHO Termino do BCP..: %TIME%
ECHO Log no arquivo %Log%
 


 
PAUSE