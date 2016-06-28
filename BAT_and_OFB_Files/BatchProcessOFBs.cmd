:: void BatchProcessOFBs(string rootPath)
::     'rootPath' is the full path to a directory containing OFB files.
::     Every OFB file in that directory will be sequentially executed by Offline Sorter.

@ECHO OFF 
SETLOCAL EnableDelayedExpansion

:: If no directory path was provided then ask for one
SET rootPath=%1
IF "%rootPath%"=="" (
    ECHO Enter the fully-qualified path to a directory containing OFB files.
    ECHO Every OFB file in that directory will be sequentially executed by Offline Sorter.
    SET /p rootPath=">"
)

:: Either way, make sure that the provided directory actually exists
IF NOT EXIST "!rootPath!" (
	ECHO Directory not found
    ECHO. && PAUSE & EXIT /B
)

:: If so, tell Offline Sorter to sequentially execute every OFB file therein
CD /D "%rootPath%"
FOR %%f in (*.ofb) DO (
    ECHO Beginning execution of %%f...
    "C:\Program Files (x86)\Plexon Inc\Offline Sorter x64 V4\OfflineSorterx64V4.exe" /b "%%~ff"
)
ECHO. && PAUSE

ENDLOCAL
@ECHO ON