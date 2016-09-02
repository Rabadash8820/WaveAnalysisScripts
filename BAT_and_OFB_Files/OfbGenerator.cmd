@ECHO off
CLS
SETLOCAL EnableDelayedExpansion

:: Store the current directory
SET currPath=%CD%

:: Get a valid rootPath from the user
ECHO Enter the fully qualified path to a root directory that contains
ECHO subdirectories with MCD files ^(without quotes^).
ECHO OFB files will be generated for all subdirectories.
SET /p rootPath=">"
IF NOT EXIST "!rootPath!" (
	ECHO Directory not found
	PAUSE & EXIT
)

:: Get a valid destination path from the user
ECHO.
ECHO Enter the fully qualified name of a template OFB file on which
ECHO the generated OFB files will be based ^(again, without quotes^).
ECHO This file will have Dir statements automatically generated.
SET /p templateFile=">"
IF NOT EXIST "!templateFile!" (
	ECHO File not found
	PAUSE & EXIT
)

:: Get a valid destination path from the user
ECHO.
ECHO Finally, enter the fully qualified path to a directory where you
ECHO where you wish to place the generated OFB files ^(without quotes^).
ECHO All .ofb files in this directory will be overwritten or removed.
SET /p destPath=">"
IF NOT EXIST "!destPath!" (
	ECHO Directory not found
	PAUSE & EXIT
)

:: Delete all the old .ofb files in the destination directory
DEL "!destPath!\*.ofb" 2> NUL

:: Move to the provided path (may be on a different drive)
CHDIR /D "!rootPath!"

:: For each subdirectory in the rootPath
SET count=1
ECHO.
FOR /d /r %%D IN (*) DO (
	:: If this subdirectory contains at least one .mcd file...
	IF EXIST "%%~fD\*.mcd" (	
		:: Create an .ofb file named like "count_subDirectoryName.ofb"		
		:: Add the line that queues all .mcd files in the rootPath directory
		> "!destPath!\!count!_%%~nD.ofb" (
			ECHO // Queue all .MCD files in the directory
			ECHO Dir %%~fD\*.mcd
			ECHO.
		)
		
		:: Append lines from the template .ofb file
		TYPE "!templateFile!" >> "!destPath!\!count!_%%~nD.ofb"
		
		:: Increase the file count and show the fileName just processed on the console
		ECHO Created "!count!_%%~nD.ofb"
		SET /a count+=1
	) 2> NUL
)

ECHO.
ECHO All OFB files are now in the directory "!destPath!"

PAUSE
ENDLOCAL
@ECHO on
