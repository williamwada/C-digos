--DD/MM/YYYY
DECLARE @d DATE = convert(DATE,'15/04/2019',103);  
select @d

select getdate()

select convert(varchar,GETDATE(),'yyyy/mm/dd')

select CONVERT(DATE,'15/04/2019',103)


DECLARE @d DATETIME = '10/01/2011';  

SELECT FORMAT(@d,'MMM/yyyy')

SELECT FORMAT ( @d, 'd', 'en-US' ) AS 'US English Result'  
      ,FORMAT ( @d, 'd', 'en-gb' ) AS 'Great Britain English Result'  
      ,FORMAT ( @d, 'd', 'de-de' ) AS 'German Result'  
      ,FORMAT ( @d, 'd', 'zh-cn' ) AS 'Simplified Chinese (PRC) Result';
  
SELECT FORMAT ( @d, 'D', 'en-US' ) AS 'US English Result'  
      ,FORMAT ( @d, 'D', 'en-gb' ) AS 'Great Britain English Result'  
      ,FORMAT ( @d, 'D', 'de-de' ) AS 'German Result'  
      ,FORMAT ( @d, 'D', 'zh-cn' ) AS 'Chinese (Simplified PRC) Result';  