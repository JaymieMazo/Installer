
IF EXIST "C:\Temp\WOMDataExtractionTool" GOTO DELDIR

:DELDIR

DEL C:\Temp\WOMDataExtractionTool /Y


IF NOT EXIST "C:\Temp" GOTO MAKETEMP
IF NOT EXIST "C:\Temp\WOMDataExtractionTool" GOTO MAKEDIR

:MAKETEMP
MKDIR C:\Temp

:MAKEDIR
MKDIR "C:\Temp\WOMDataExtractionTool"

COPY "\\wkn-appserver\Access$\VBPrograms\WOMDataExtractionTool\WOMDataExtractionTool.exe" "C:\Temp\WOMDataExtractionTool" /Y

:Run
START C:\Temp\WOMDataExtractionTool\WOMDataExtractionTool.exe
EXIT