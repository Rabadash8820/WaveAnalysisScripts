@ECHO off
CLS
SETLOCAL EnableDelayedExpansion

REM Get the rootPath from the user
ECHO Enter the fully qualified path to a directory that
ECHO does not contain MCD files, but whose subdirectories do.
ECHO OFB files will be generated for all subdirectories.
SET /p rootPath=">"

REM If that path exists, continue
IF NOT EXIST "!rootPath!" (
	ECHO Directory not found
) ELSE (
	REM Move to the provided path (may be on a different drive)
	CHDIR /D "!rootPath!"
	
	REM Get the rootPath from the user
	ECHO.
	ECHO Enter the fully qualified path to a directory where you wish to place generated OFB files
	SET /p destPath=">"
	
	REM If that path exists, continue
	IF NOT EXIST "!destPath!" (
		ECHO Directory not found
	) ELSE (	
		REM Delete all previous .plx files if the user wishes
		ECHO.
		ECHO Delete previous PLX files from all subdirectories also? ^(y/n^)
		SET /p deleteMCDs=">"
		IF !deleteMCDs!==y DEL /s "*.plx"

		REM Delete all previous .ofb files
		REM This makes sure it doesn't output "Couldn't find file" when there are no .ofb files left
		IF EXIST "!destPath!\*.ofb" (
			DEL "!destPath!\*.ofb"
		)

		REM For each subdirectory in the rootPath
		SET count=1
		ECHO.
		FOR /r /d %%D IN (*) DO (
			REM If it contains at least one .mcd file
			IF EXIST "%%~fD\*.mcd" (						
				REM Create the .ofb file named like "count_subdirectoryName.ofb"
				REM Spaces before >> operators are only necessary sometimes for whatever reason
				>> !count!_%%~nD.ofb (
					ECHO // Work with all .MCD files in the Data directory
					ECHO(
					ECHO // Queue all .MCD files in the Data directory
					ECHO Dir %%~fD\*.mcd
					ECHO Set DetectSigmas 5
					ECHO Set DetectNPW 50
					ECHO Set DetectNPre 10
					ECHO Set DetectDead 0
					ECHO ForEachChannel Detect
					ECHO(
					ECHO // Remove artifacts
					ECHO Set ArtifactWidth 10
					ECHO Set ArtifactPercentage 70
					ECHO ForEachFile InvalidateArtifactsAfter
					ECHO(
					ECHO // Export unsorted timestamps to .PLX files
					ECHO Set SaveCont 0
					ECHO ForEachFile ExportToPlx
					ECHO(					
					ECHO // Run T-Distribution E-M sorting on all channels, using principle components 1-3 as features
					ECHO Set FeatureX 0
					ECHO Set FeatureY 1
					ECHO Set FeatureZ 2
					ECHO ForEachChannel TDist3d
					ECHO(					
					ECHO // Export sorted timestamps to new .PLX files
					ECHO Set SaveCont 0
					ECHO ForEachFile ExportToPlx
					ECHO(					
					ECHO Process
				)
				
				REM Increase the file count and show the fileName just processed on the console
				ECHO Created !count!_%%~nD.ofb
				SET /a count+=1
			)
		)

		REM Move all .ofb files to the appropriate Code folder
		FOR /r %%F in ("*.ofb") DO (
			MOVE "%%~fF" "!destPath!"
		)
		ECHO.
		ECHO All OFB files are now in the directory "!destPath!"
		
	)
)

REM Deallocate environment variables
SET rootPath=
SET destPath=
SET deleteMCDs=
SET count=

PAUSE
ENDLOCAL
@ECHO on