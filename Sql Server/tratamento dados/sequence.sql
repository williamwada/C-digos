

select NEXT VALUE FOR sequence_1 ,D1N,D1C from tb_stage_ipca


CREATE SEQUENCE sequence_1
start with 1
increment by 1
minvalue 0
maxvalue 100
cycle;

DROP SEQUENCE sequence_1;

select D1N,D1C 
from tb_stage_ipca
group by D1C,D1N


