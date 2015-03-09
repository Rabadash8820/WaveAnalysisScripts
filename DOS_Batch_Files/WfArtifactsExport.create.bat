@echo off
CLS
SETLOCAL EnableDelayedExpansion

REM Get the rootPath from the user
ECHO Enter the fully qualified path to the directory containing .mcd files
ECHO Note that .mcd files in all subdirectories will also be processed
SET /p rootPath=">"

REM If that path exists, continue
IF NOT EXIST "!rootPath!" (
	ECHO Directory not found
) ELSE (
	REM Move to the provided path
	CHDIR "!rootPath!"
	
	REM Get the rootPath from the user
	ECHO.
	ECHO Enter the fully qualified path to the directory where you wish to place the generated .ofb files
	SET /p destPath=">"
	
	REM If that path exists, continue
	IF NOT EXIST "!destPath!" (
		ECHO Directory not found
	) ELSE (	
		REM Delete all previous .plx files if the user wishes
		ECHO.
		ECHO Delete previous .plx files from all subdirectories also? ^(y/n^)
		SET /p deleteMCDs=">"
		IF !deleteMCDs!==y DEL /s "*.plx"

		REM Delete all previous .ofb files
		REM This makes sure it doesn't output "Couldn't find file" when there are no .ofb files left
		IF EXIST "!destPath!\*.ofb" (
			DEL "!destPath!\*.ofb"
		)

		REM For each subdirectory in the rootPath
		SET count=0
		ECHO.
		FOR /r /d %%D IN (*) DO (
			REM If it contains at least one .mcd file
			IF EXIST "%%~fD\*.mcd" (
				SET /a count+=1
				
				REM Create the .ofb file named like "count_subdirectoryName.ofb"  (the spaces after DetectSigmas and DetectDead are necessary for some reason)
				ECHO Dir %%~fD\*.mcd>> !count!_%%~nD.ofb
				ECHO Set DetectSigmas 5 >> !count!_%%~nD.ofb
				ECHO Set DetectNPW 50>> !count!_%%~nD.ofb
				ECHO Set DetectNPre 10>> !count!_%%~nD.ofb
				ECHO Set DetectDead 0 >> !count!_%%~nD.ofb
				ECHO ForEachChannel Detect>> !count!_%%~nD.ofb
				ECHO Set ArtifactWidth 10>> !count!_%%~nD.ofb
				ECHO Set ArtifactPercentage 70>> !count!_%%~nD.ofb
				ECHO ForEachFile InvalidateArtifactsAfter>> !count!_%%~nD.ofb
				ECHO Set SaveCont 0 >> !count!_%%~nD.ofb
				ECHO ForEachFile ExportToPlx>> !count!_%%~nD.ofb
				ECHO Process>> !count!_%%~nD.ofb
				
				ECHO Created !count!_%%~nD.ofb
			)
		)

		REM Move all .ofb files to the appropriate Code folder
		FOR /r %%F in ("*.ofb") DO move "%%~fF" "!destPath!"
		ECHO.
		ECHO All .ofb files are now in the directory "!destPath!"
		
	)
)

REM Deallocate environment variables
SET rootPath=
SET destPath=
SET deleteMCDs=
SET count=

PAUSE
ENDLOCAL
@echo on