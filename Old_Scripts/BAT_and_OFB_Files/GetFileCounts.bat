@ECHO off
CLS
SETLOCAL EnableDelayedExpansion

REM Explain what's gonna happen to the user
ECHO This script will count the number of MCD, PLX, and TXT files
ECHO in every subdirectory of the current directory.
ECHO You can use this to check if errors occurred while processing any files
ECHO and at what stage of the pipeline the error occurred.
ECHO.
PAUSE

REM Store counts of MCD, PLX, and TXT files into FileCounts.txt
ECHO Displaying counts of MCD, PLX, and TXT files in the current directory and all its subdirectories 
FOR /r /d %%d IN (*) DO (
	SET mcdCount=0
	SET plxCount=0
	SET txtCount=0
	FOR %%f IN ("%%~fd\*.mcd") DO set /a mcdCount+=1
	FOR %%f IN ("%%~fd\*.plx") DO set /a plxCount+=1
	FOR %%f IN ("%%~fd\*.txt") DO set /a txtCount+=1
	ECHO %%~fd,!mcdCount!,!plxCount!,!txtCount!> FileCounts.csv
)

ENDLOCAL
