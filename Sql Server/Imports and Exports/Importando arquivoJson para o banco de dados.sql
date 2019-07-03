Declare @countYear int
Declare @SQL nvarchar(max)
Declare @path varchar(max)

set @countYear =2017

WHILE @countYear <2020
BEGIN
	SET @path ='C:\Users\william.wada\Desktop\OneDrive - LCA Consultores\Tarefas\Camara\raw_data\proposicoes\json\converted\proposicoes-'+convert(varchar,@countYear)+'-unicode.json'
	
	SET @SQL='Declare @JSON NVARCHAR(max) '+			
			+'SELECT @JSON = BulkColumn '+
			+'FROM OPENROWSET (BULK '''+@path+''' , SINGLE_NCLOB) as j '+
			+'insert into Camara_Dep.dbo.tb_raw_proposicoes '+
			+'SELECT * FROM OPENJSON (@JSON,''$.dados'') '+
			+'With( id NVARCHAR(10) ,	uri VARCHAR(100) ,	siglaTipo VARCHAR(10) ,	numero VARCHAR(10) ,	ano VARCHAR(10) ,	codTipo VARCHAR(10) ,	descricaoTipo NVARCHAR(500) ,	ementa nVARCHAR(MAX) ,	ementaDetalhada nVARCHAR(MAX) ,	keywords nVARCHAR(MAX) ,	dataApresentacao VARCHAR(50) ,	uriOrgaoNumerador VARCHAR(100) ,	uriPropAnterior VARCHAR(100) ,	uriPropPrincipal VARCHAR(100) ,	uriPropPosterior VARCHAR(100) ,	urlInteiroTeor VARCHAR(100) ,	urnFinal VARCHAR(100) ,	ultimoStatus_dataHora VARCHAR(25) ,	ultimoStatus_sequencia VARCHAR(5) ,	ultimoStatus_uriRelator VARCHAR(100) ,	ultimoStatus_idOrgao VARCHAR(15) ,	ultimoStatus_siglaOrgao nVARCHAR(25) ,	ultimoStatus_uriOrgao VARCHAR(100) ,	ultimoStatus_regime VARCHAR(250) ,	ultimoStatus_descricaoTramitacao nVARCHAR(250) ,	ultimoStatus_idTipoTramitacao nVARCHAR(10) ,	ultimoStatus_descricaoSituacao nVARCHAR(250) ,	ultimoStatus_idSituacao nVARCHAR(10) ,	ultimoStatus_despacho nVARCHAR(MAX) ,	ultimoStatus_url nVARCHAR(100) ) '
	 SET @countYear = @countYear +1;
	
	PRINT(@SQL)
	EXEC(@SQL)
END;

GO