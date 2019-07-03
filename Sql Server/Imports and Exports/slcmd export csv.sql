
!!sqlcmd -S WILLIAM-T\LCA_P1 -d CNPJ -E -Q "set nocount on; select top 10 TRIM(FORMA),TRIM(CNPJ),TRIM(NM_PAIS) from [dbo].[TB_PRINCIPAL_CNPJ]" -o "E:\outputs\cnpj_cnae.csv" -W -s";" -h-1