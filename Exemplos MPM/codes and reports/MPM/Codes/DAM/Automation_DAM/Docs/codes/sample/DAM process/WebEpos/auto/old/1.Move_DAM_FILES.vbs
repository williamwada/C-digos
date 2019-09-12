Dim vArq As String

vArq = Dir("F:\EPOS_AD\f13\*.xls")

'COPY FILES TO Another USB
If vArq <> "" Then
objFSO.CopyFile "F:\EPOS_AD\f13\*.xls", "I:\DailyReports\"
End If

'MOVE EPOS_F13
vArq = Dir("F:\EPOS_AD\f13\*.xls")
If vArq <> "" Then
	objFSO.MoveFile "F:\EPOS_AD\f13\*.xls","Y:\DATA_AREA\STAGING_AREA\WEB_F13\"
End If

'MOVE EPOS_LOG_CAN
vArq = Dir("F:\EPOS_AD\logfile\adcan\*.csv")
If vArq <> "" Then
	objFSO.MoveFile "F:\EPOS_AD\logfile\adcan\*.csv","Y:\DATA_AREA\STAGING_AREA\WEB_log\ADCAN\"
End If

'MOVE EPOS_LOG_MGI
vArq = Dir("F:\EPOS_AD\logfile\mgi\*.csv")
If vArq <> "" Then
	objFSO.MoveFile "F:\EPOS_AD\logfile\mgi\*.csv","Y:\DATA_AREA\STAGING_AREA\WEB_log\ADMGI\"
End If

'MOVE EPOS_YESFILE_CAN
vArq = Dir("F:\EPOS_AD\yesfile\adcan\*.txt")
If vArq <> "" Then
	objFSO.MoveFile "F:\EPOS_AD\f13\*.xls","Y:\DATA_AREA\STAGING_AREA\WEB_YES_File\"
End If

'MOVE EPOS_YESFILE_MGI
vArq = Dir("F:\EPOS_AD\yesfile\mgi\*.txt")
If vArq <> "" Then
	objFSO.MoveFile "F:\EPOS_AD\f13\*.xls","Y:\DATA_AREA\STAGING_AREA\WEB_YES_File\"
End If



End Sub