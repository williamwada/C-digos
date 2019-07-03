INSERT INTO BMG.dbo.[tb_raw_data]([data],[ARS_Curncy],[AUD_Curncy],[ATS_Curncy],[BOB_Curncy],[BRL_Curncy],[BGN_Curncy],[CAD_Curncy],[CLP_Curncy],[CNY_Curncy],[SGD_Curncy],[COP_Curncy],[ECS_Curncy],[SKK_Curncy],[SIT_Curncy],[ESP_Curncy],[EUR_Curncy],[PHP_Curncy],[FRF_Curncy],[HKD_Curncy],[HUF_Curncy],[INR_Curncy],[IDR_Curncy],[GBP_Curncy],[IEP_Curncy],[ILS_Curncy],[ITL_Curncy],[JPY_Curncy],[KRW_Curncy],[MYR_Curncy],[MXN_Curncy],[NOK_Curncy],[NZD_Curncy],[PYG_Curncy],[PEN_Curncy],[PLN_Curncy],[RON_Curncy],[RUB_Curncy],[SEK_Curncy],[CHF_Curncy],[THB_Curncy],[TWD_Curncy],[CZK_Curncy],[TRY_Curncy],[UAH_Curncy],[UYU_Curncy],[VND_Curncy],[ZAR_Curncy],[USTWBROA_Index],[DXY_Curncy])
values ('2019-06-13 00:00:00','43.5294','0.6915','12.2025','6.91','3.8494','1.7346','1.3327','696.11','6.9216','1.3668','3268.81','25000.0','26.7154','212.5099','147.5491','1.1276','51.887','5.8169','7.8286999999999995','285.48','69.5137','14280.0','1.2674','1.4318','3.5984','1717.06','108.38','1183.01','4.165','19.1947','8.6803','0.6568','6244.2','3.331','3.7734','4.185','64.5632','9.4831','0.994','31.201','31.474','22.6655','5.8688','26.4412','35.28','23317.0','14.8675','128.5107','97.013')
SELECT MAX(LEN(RIGHT(CAD_Curncy,LEN(CAD_Curncy)-CHARINDEX('.',CAD_Curncy)))) as qtde 
FROM tb_raw_data


select CAD_Curncy, round(CAD_Curncy,16)
from tb_raw_data