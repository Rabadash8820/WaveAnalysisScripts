@ECHO off

CLS

SETLOCAL EnableDelayedExpansion

REM Delete the previously generated FileCounts.txt file
REM Don't output "Couldn't find file" if there wasn't one previously generated
IF EXIST "FileCounts.csv" (
	DEL "FileCounts.csv"
)

REM Store counts of MCD, PLX, and TXT files into FileCounts.txt
ECHO Displaying counts of MCD, PLX, and TXT files in the current directory and all its subdirectories 
FOR /r /d %%d IN (*) DO (
	SET mcdCount=0
	SET plxCount=0
	SET txtCount=0
	FOR %%f IN ("%%~fd\*.mcd") DO set /a mcdCount+=1
	FOR %%f IN ("%%~fd\*.plx") DO set /a plxCount+=1
	FOR %%f IN ("%%~fd\*.txt") DO set /a txtCount+=1
	ECHO %%~fd,!mcdCount!,!plxCount!,!txtCount!>> FileCounts.csv
)

REM Deallocate environment variables
SET mcdCount=
SET plxCount=
SET txtCount=

ENDLOCAL