
/*1.Alimentação e bebidas                                                                                                                               
	11.Alimentação no domicílio  
		1101.Cereais, leguminosas e oleaginosas
			1101002.Arroz                                                                                                                                                                                                                                                                                                                                                                                 
			*/
select convert(int,substring(D4N,0,CHARINDEX('.',D4N,0))),D4N from dbo.tb_raw_ipca
order by convert(int,substring(D4N,0,CHARINDEX('.',D4N,0)))

select convert(int,substring(D4N,0,CHARINDEX('.',D4N,0))),D4N,SUM(convert(float,replace(replace(V,'...','0'),'-',''))) 
from dbo.tb_raw_ipca
group by convert(int,substring(D4N,0,CHARINDEX('.',D4N,0))),D4N
order by convert(int,substring(D4N,0,CHARINDEX('.',D4N,0)))


select convert(float,replace(replace(V,'...','0'),'-','')),V
from tb_raw_ipca

select top 5282 replace(replace(V,'...','0'),'-',''),V
from tb_raw_ipca

select * from tb_raw_ipca