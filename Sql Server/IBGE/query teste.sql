--query modelo1
select * 
from tb_ipca_mensal_ibge i,
     tb_regioes_ipca_ibge r,
	 tb_grupo_ipca_ibge g
where i.id_regioes_ibge =1
and year(i.data) ='2017'
and i.id_regioes_ibge = r.id_regioes_ibge
and i.id_grupo_ipca_ibge = g.id_grupo_ipca_ibge
and g.id_grupo_ipca_ibge =1
and i.id_subgrupo_ipca_ibge is null
and i.id_item_ipca_ibge is null
and i.id_subItem_ipca_ibge is null
order by indice

-- Query modelo 2
select *
from tb_ipca_mensal_ibge i,
tb_classificacao_ibge c,
tb_regioes_ipca_ibge r
where i.id_regioes_ibge =r.id_regioes_ibge
and i.id_class_ibge = c.id_class_ibge
and r.descricao ='Brasil'
and c.descricao ='Alimentação e bebidas'
and year(i.data)='2017'
order by indice


--
-- Modelo1 ficou muito rebuscado para trazer as informacoes. Precisa-se criar queries mais complicadas. As insercoes tambem foram necessarios logicas mais pesadas para se chegar no resultado final
-- Modelo2 mais simples e com queries mais intuitivas. Performance melhor que o modelo1 e no caso de outra pessoa assumir sera bem mais facil para entender todo o fluxo
--
-- O modelo 1 tenta normalizar as tabelas para poder facilitar nos updates e na integridade dos dados.
-- O modelo2 usa o encadeamento dos codigos para acessar os subconjuntos de regioes e classificacoes.
-- o motivo do modelo 1 ter ficado complexo nas consultas e nas insercoes eh porque houve a necessidade de se manter os dados agregados em conjunto com os dados granulados.
-- Como a logica de negocio obriga o salvamento desses dados a alternatica do modelo 2 com encadeamento dos codigos se torna a mais interessante.


