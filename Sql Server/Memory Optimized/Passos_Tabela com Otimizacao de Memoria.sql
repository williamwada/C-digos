-- No management Studio,botao direito no database, propriedades
-- 0. criar grupo de arquivo na aba Dados com otimizacao de memoria
--1. Criar arquivo de banco de dados do tipo Dados FileStream apontando o grupo de arquivo criado no item anterior.


--2. criacao da tabela habilitando a otimizacao de memoria.
-- para tabelas q tenham a otimizacao de memoria habilitada utilizar prefixo so_*
USE [IBGE]
GO

/****** Object:  Table [dbo].[so_raw_ipca]    Script Date: 02/07/2019 18:06:33 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[so_raw_ipca]
(
	[D4N] [nchar](150) COLLATE Latin1_General_CI_AS NULL,
	[D1C] [nchar](50) COLLATE Latin1_General_CI_AS NULL,
	[D1N] [nchar](150) COLLATE Latin1_General_CI_AS NULL,
	[D2N] [nchar](150) COLLATE Latin1_General_CI_AS NULL,
	[D3N] [nchar](150) COLLATE Latin1_General_CI_AS NULL,
	[MN] [nchar](50) COLLATE Latin1_General_CI_AS NULL,
	[V] [nchar](50) COLLATE Latin1_General_CI_AS NULL,

INDEX [ix1] NONCLUSTERED 
(
	[D4N] ASC
)
)WITH ( MEMORY_OPTIMIZED = ON , DURABILITY = SCHEMA_ONLY )
GO

-- 3. Habilitar snapshot ON na base especificada
ALTER DATABASE IBGE
    SET MEMORY_OPTIMIZED_ELEVATE_TO_SNAPSHOT = ON;

--4. modificar o nome da tabela no script.