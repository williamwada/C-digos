USE [Camara_Dep]
GO
/****** Object:  StoredProcedure [dbo].[csv_insert_to_table]    Script Date: 29/05/2019 16:39:32 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[json_insert_to_table]
as

Declare @countYear int
Declare @SQL nvarchar(max)
Declare @path varchar(max)

set @countYear =1945

WHILE @countYear <2020
BEGIN
	SET @path ='C:\Users\william.wada\Desktop\OneDrive - LCA Consultores\Tarefas\Camara\raw_data\proposicoes\json\converted\proposicoes-'+convert(varchar,@countYear)+'-unicode.json'
	
	SET @SQL='Declare @JSON NVARCHAR(max) '+			
			+'SELECT @JSON = BulkColumn '+
			+'FROM OPENROWSET (BULK '''+@path+''' , SINGLE_NCLOB) as j '+
			+'insert into Camara_Dep.dbo.tb_raw_proposicoes(id,uri,siglaTipo,numero,ano,codTipo,descricaoTipo,ementa,ementaDetalhada,keywords,dataApresentacao,uriOrgaoNumerador,uriPropAnterior,uriPropPrincipal,uriPropPosterior,urlInteiroTeor,ultimoStatus_dataHora,ultimoStatus_sequencia,ultimoStatus_uriRelator,ultimoStatus_idOrgao,ultimoStatus_siglaOrgao,ultimoStatus_uriOrgao,ultimoStatus_regime,ultimoStatus_descricaoTramitacao,ultimoStatus_idTipoTramitacao,ultimoStatus_descricaoSituacao,ultimoStatus_idSituacao,ultimoStatus_despacho,ultimoStatus_url) '+
			+'SELECT id,uri,siglaTipo,numero,ano,codTipo,descricaoTipo,ementa,ementaDetalhada,keywords,dataApresentacao,uriOrgaoNumerador,uriPropAnterior,uriPropPrincipal,uriPropPosterior,urlInteiroTeor, data as ultimoStatus_dataHora,  sequencia as ultimoStatus_sequencia,  uriRelator as ultimoStatus_uriRelator,  codOrgao as ultimoStatus_idOrgao,  siglaOrgao as ultimoStatus_siglaOrgao,  uriOrgao as ultimoStatus_uriOrgao,  regime as ultimoStatus_regime,  descricaoTramitacao as ultimoStatus_descricaoTramitacao,  idTipoTramitacao as ultimoStatus_idTipoTramitacao,  descricaoSituacao as ultimoStatus_descricaoSituacao,  idSituacao as ultimoStatus_idSituacao,  despacho as ultimoStatus_despacho,  url as ultimoStatus_url '+
			+'FROM OPENJSON (@JSON,''$.dados'') '+ 
			+'With( id NVARCHAR(10) ,	uri VARCHAR(100) ,	siglaTipo VARCHAR(10) ,	numero VARCHAR(10) ,	ano VARCHAR(10) ,	codTipo VARCHAR(10) ,	descricaoTipo NVARCHAR(500) ,	ementa nVARCHAR(MAX) ,	ementaDetalhada nVARCHAR(MAX) ,	keywords nVARCHAR(MAX) ,	dataApresentacao VARCHAR(50) ,	uriOrgaoNumerador VARCHAR(100) ,	uriPropAnterior VARCHAR(100) ,	uriPropPrincipal VARCHAR(100) ,	uriPropPosterior VARCHAR(100) ,	urlInteiroTeor VARCHAR(100), ultimoStatus nvarchar(max) AS JSON ) as J '+
			+'CROSS APPLY '+
			+'OPENJSON (J.[ultimoStatus]) '+
			+'with(data varchar(20),sequencia varchar(50),uriRelator varchar(50),codOrgao varchar(50),siglaOrgao varchar(10),uriOrgao varchar(100),regime varchar(100),descricaoTramitacao varchar(100),idTipoTramitacao varchar(50),descricaoSituacao varchar(50),idSituacao varchar(50),despacho varchar(100),url varchar(100)) '

	 SET @countYear = @countYear +1;
	
	PRINT(@countYear)
	PRINT(@SQL)
	EXEC(@SQL)
END;

GO




