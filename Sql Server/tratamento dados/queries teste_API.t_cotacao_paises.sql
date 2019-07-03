select * from teste_API.dbo.t_cotacao_paises order by serie

select * from teste_API.dbo.t_producao 

INSERT INTO teste_API.dbo.t_cotacao_paises (serie,valor, data) VALUES(1,3.8339, '10/04/2019')
--insert
insert into teste_API.dbo.t_cotacao_paises(serie,valor,data,pais)
values (10813,10.80,'01/05/1999','Nigeria');


--Delete series 10813
begin transaction;
delete from teste_API.dbo.t_cotacao_paises where serie ='10813';
--commit;
--rollback;

--Delete all t_producao
begin transaction;
delete from teste_API.dbo.t_producao ;
--commit;
--rollback;


--Delete ALL	
begin transaction;
delete from teste_API.dbo.t_cotacao_paises;
--commit;
--rollback;

