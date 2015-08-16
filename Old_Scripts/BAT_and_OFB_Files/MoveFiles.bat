@ECHO off
CLS
SETLOCAL EnableDelayedExpansion

REM Store the current directory
SET currPath=%CD%

REM Get the file filter from the user
ECHO Enter the filter for the files that you want to move
ECHO (e.g., "adch_*.txt" for all text files with a name starting with adch_)
ECHO The current directory is being processed, so full paths should not be given.
SET /p filter=">"

REM Get a valid destination path from the user
ECHO.
ECHO Enter the fully qualified path to a directory
ECHO where you wish to move all files matching the filter.
SET /p destPath=">"
IF NOT EXIST "!destPath!" (
	ECHO Directory not found
	PAUSE & EXIT
)

REM Check whether the user wants to move or copy the files
ECHO.
ECHO Finally, do you want to MOVE these files or COPY them?
ECHO Please type either "m" to move or "c" to copy.
SET /p copy=">"
IF NOT "!copy!"=="m" (
	IF NOT "!copy!"=="c" (
		ECHO You must type either an "m" or a "c"
		PAUSE & EXIT
	)
)

REM Notify the user of what's about to happen
ECHO.
ECHO Ok, the xcopy program will take over from here...

REM Move all files matching the filter to the provided destination
REM /f outputs the names of each copied file
REM /j prevents buffering of files to speed things up
REM /t replicates the source's tree structure
REM /w prompts the user before starting anything
XCOPY "!currPath!\!filter!" "!destPath!" /f /j /s /w

REM If the user wants to MOVE files, then delete all files at the source location
IF !copy!==c (
	DEL /S "!currPath!\!filter!"
)

ECHO.
ECHO All files have now been moved to "!destPath!"!

PAUSE
ENDLOCAL
@ECHO on
