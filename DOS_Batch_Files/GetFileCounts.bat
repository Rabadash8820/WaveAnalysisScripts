@ECHO off

CLS

SETLOCAL EnableDelayedExpansion

CHDIR "..\..\CORDs MEA data"
DEL FileCounts.txt

FOR /r /d %%d IN (*) DO (
	SET mcdCount=0
	SET plxCount=0
	SET txtCount=0
	FOR %%f IN ("%%~fd\*.mcd") DO set /a mcdCount+=1
	FOR %%f IN ("%%~fd\*-01.plx") DO set /a plxCount+=1
	FOR %%f IN ("%%~fd\*-01.txt") DO set /a txtCount+=1
	ECHO %%~fd,!mcdCount!,!plxCount!,!txtCount!>> ../Code/DOS_Batch_Files/FileCounts.txt
)

SET mcdCount=
SET plxCount=
SET txtCount=

ENDLOCAL