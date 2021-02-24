
IF NOT EXIST "C:\Temp" GOTO MAKETEMP
IF NOT EXIST "C:\Temp\WOMDataExtractionTool(New)" GOTO MAKEDIR

IF EXIST "C:\Temp\WOMDataExtractionTool(New)" GOTO DELDIR

:DELDIR

DEL C:\Temp\WOMDataExtractionTool(New) /Y

:MAKETEMP
MKDIR C:\Temp

:MAKEDIR
MKDIR "C:\Temp\WOMDataExtractionTool(New)"

COPY "\\wkn-appserver\Access$\VBPrograms\WOMDataExtractionTool(New)\WOMDataExtractionTool(New).exe" "C:\Temp\WOMDataExtractionTool(New)" /Y

:Run
START C:\Temp\WOMDataExtractionTool(New)\WOMDataExtractionTool(New).exe
EXIT