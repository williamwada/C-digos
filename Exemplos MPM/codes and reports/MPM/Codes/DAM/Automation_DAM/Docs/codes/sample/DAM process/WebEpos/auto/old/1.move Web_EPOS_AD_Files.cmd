::move files to load
COPY  F:\EPOS_AD\f13\*.xls I:\DailyReports
move F:\EPOS_AD\f13\*.xls Y:\DATA_AREA\STAGING_AREA\WEB_F13\
move F:\EPOS_AD\logfile\adcan\*.csv Y:\DATA_AREA\STAGING_AREA\WEB_log\ADCAN\
move F:\EPOS_AD\logfile\mgi\*.csv Y:\DATA_AREA\STAGING_AREA\WEB_log\ADMGI\
move F:\EPOS_AD\yesfile\adcan\*.txt Y:\DATA_AREA\STAGING_AREA\WEB_YES_File\
move F:\EPOS_AD\yesfile\mgi\*.txt Y:\DATA_AREA\STAGING_AREA\WEB_YES_File\
pause