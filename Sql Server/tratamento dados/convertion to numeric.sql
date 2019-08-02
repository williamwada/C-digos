select CAST(LEFT(ligaAluminio, CHARINDEX('.', ligaAluminio) - 1) + '.' + SUBSTRING(ligaAluminio,(CHARINDEX('.',ligaAluminio)+1),3) AS DECIMAL(13,6)) CastedNumeri
from teste_API.dbo.ligaAluminio