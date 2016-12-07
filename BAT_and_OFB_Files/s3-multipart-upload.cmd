::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:: int main(string[] args)
::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:main

:: Set up environment
@ECHO OFF
SETLOCAL EnableDelayedExpansion

:: If no arguments were provided then show usage and exit
IF "%1"=="" (
    ECHO Error: Missing required arguments 1>&2
    ECHO.
    CALL :showUsage
    EXIT /B 0
)

:: Otherwise, parse arguments
CALL :parseArgs %*

:: If there were any parsing errors then just exit
SET errorCode=%ERRORLEVEL%
IF %errorCode%==1 EXIT /B 1
IF %errorCode%==2 EXIT /B 1

:: If help was requested then show proper usage
IF %help%==true (
    CALL :showUsage
    EXIT /B 1
)

CALL :initMultipartUpload
CALL :doMultipartUpload

EXIT /B 0

::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:: void showUsage()
::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:showUsage
>CON (
    ECHO Usage: s3-multipart-upload -b^|--bucket ^<value^> [-h^|--help] -k^|--key ^<value^>
    ECHO                            [-c^|--compress] [-m^|--metadata ^<key=value,...^>]
    ECHO                            -p^|--profile ^<value^>
    ECHO.
    ECHO    -b, --bucket     Name of the AWS S3 bucket to which you are uploading
    ECHO    -c, --compress   Compress the file before splitting into parts (you should only use this if the file is not already compressed^)
    ECHO    -h, --help       Show this help text
    ECHO    -k, --key        A unique key (name^) to identify the uploaded object in the bucket (e.g., "my-file"^)
    ECHO    -m, --metadata   A comma-delimited list of "key=value" pairs (e.g., "Name=my-file,Project=Derp"^)
    ECHO    -p, --profile    A credentials profile to pass to the AWS CLI (must have already been created with "aws configure"^)
)
EXIT /B 0

::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:: int parseArgs(string[] args)
::
:: Return value:
::    0 - parsed without error
::    1 - argument was missing subsequent arguments
::    2 - unrecognized argument
::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:parseArgs

:: Initialize arguments
SET compress=false
SET help=false

:: Parse each argument
SET result=0
:loop
IF NOT "%1"=="" (
    SET validArg=false

    :: Parse bucket name
    SET isBucket=false
    IF "%1"=="-b" SET isBucket=true
    IF "%1"=="/b" SET isBucket=true && IF "%1"=="/B" SET isBucket=true
    IF "%1"=="--bucket" SET isBucket=true
    IF !isBucket!==true (
        SET arg=%2
        IF "!arg!"=="" SET result=1
        IF "!arg:~0,1!"=="-" SET result=1
        IF !result!==0 (SHIFT) ELSE ECHO Error: --bucket requires an argument 1>&2
        SET bucket=!arg!
        SHIFT
    )
    IF !isBucket!==true SET validArg=true

    :: Parse compress flag
    SET isCompress=false
    IF "%1"=="-c" SET isCompress=true
    IF "%1"=="/c" SET isCompress=true & IF "%1"=="/C" SET isCompress=true
    IF "%1"=="--compress" SET isCompress=true
    IF !isCompress!==true (
        SET compress=true
        SHIFT
    )
    IF !isCompress!==true SET validArg=true

    :: Parse help flag
    SET isHelp=false
    IF "%1"=="-h" SET isHelp=true
    IF "%1"=="/h" SET isHelp=true & IF "%1"=="/H" SET isHelp=true
    IF "%1"=="--help" SET isHelp=true
    IF !isHelp!==true (
        SET help=true
        SHIFT
    )
    IF !isHelp!==true SET validArg=true
    
    :: Parse object key
    SET isKey=false
    IF "%1"=="-k" SET isKey=true
    IF "%1"=="/k" SET isKey=true & IF "%1"=="/K" SET isKey=true
    IF "%1"=="--key" SET isKey=true
    IF !isKey!==true (
        SET arg=%2
        IF "!arg!"=="" SET result=1
        IF "!arg:~0,1!"=="-" SET result=1
        IF !result!==0 (SHIFT) ELSE ECHO Error: --key requires an argument 1>&2
        SET key=!arg!
        SHIFT
    )
    IF !isKey!==true SET validArg=true
    
    :: Parse metadata
    SET isMetadata=false
    IF "%1"=="-m" SET isMetadata=true
    IF "%1"=="/m" SET isMetadata=true & IF "%1"=="/M" SET isMetadata=true
    IF "%1"=="--metadata" SET isMetadata=true
    IF !isMetadata!==true (
        SET arg=%2
        IF "!arg!"=="" SET result=1
        IF "!arg:~0,1!"=="-" SET result=1
        IF !result!==0 (SHIFT) ELSE ECHO Error: --metadata requires an argument 1>&2
        SET metadata=!arg!
        SHIFT
    )
    IF !isMetadata!==true SET validArg=true
    
    :: Parse AWS credentials profile
    SET isProfile=false
    IF "%1"=="-p" SET isProfile=true
    IF "%1"=="/p" SET isProfile=true & IF "%1"=="/P" SET isProfile=true
    IF "%1"=="--profile" SET isProfile=true
    IF !isProfile!==true (
        SET arg=%2
        IF "!arg!"=="" SET result=1
        IF "!arg:~0,1!"=="-" SET result=1
        IF !result!==0 (SHIFT) ELSE ECHO Error: --profile requires an argument 1>&2
        SET profile=!arg!
        SHIFT
    )
    IF !isProfile!==true SET validArg=true
    
    :: If this arg was invalid...
    IF !validArg!==false (
        ECHO Error: unrecognized argument "%1" 1>&2
        SET result=2
        SHIFT
    )
    
    :: If there were any parsing errors then just exit
    IF !result!==0 (GOTO loop) ELSE (EXIT /B !result!)
)

:: If help was requested, then just show usage and exit
IF %help%==true (EXIT /B 0)

:: Validate arguments
SET valid=true
IF NOT DEFINED bucket (SET valid=false & ECHO Error: You must provide a bucket name ^(--bucket^)!) 1>&2
IF NOT DEFINED key (SET valid=false & ECHO Error: You must provide an object key ^(--key^)!) 1>&2
IF NOT DEFINED profile (SET valid=false & ECHO Error: You must provide an AWS CLI credentials profile ^(--profile^)!) 1>&2
IF %valid%==false EXIT /B 1

:: Unset local vars before exit
SET arg=
SET validArg=
SET valid=
SET result=

SET isBucket=
SET isCompress=
SET isHelp=
SET isKey=
SET isMetadata=
SET isProfile=

EXIT /B 0

::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:: void initMultipartUpload(string bucket, string key, string metadata="", string profile, ref uploadID)
::
:: If successful, the upload-id of the new multipart upload is stored in uploadID
::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:initMultipartUpload

:: Create the AWS S3 multipart upload (can't pass null string to --metadata argument)
ECHO Creating multipart upload...
SET RESPONSE_FILE=response.txt
SET ERROR_FILE=error.txt
IF "%metadata%"=="" (
    aws s3api create-multipart-upload --bucket "%bucket%" --key "%key%" --profile "%profile%"> "%RESPONSE_FILE%" 2> "%ERROR_FILE%"
) ELSE (
    aws s3api create-multipart-upload --bucket "%bucket%" --key "%key%" --metadata "%metadata%" --profile "%profile%"> "%RESPONSE_FILE%" 2> "%ERROR_FILE%"
)

:: Rethrow any error messages from the AWS API
FOR /F %%i IN ("%ERROR_FILE%") DO SET size=%%~zi
IF %size% GTR 0 (
    ECHO Creation failed with error message: 1>&2
    TYPE "%ERROR_FILE%" 1>&2
    DEL "%ERROR_FILE%"
    EXIT /B 1
)
DEL "%ERROR_FILE%"

:: Store the new upload ID
SET ID_FILE=upload-id.txt
FINDSTR /C:"UploadId" "%RESPONSE_FILE%"> "%ID_FILE%"
DEL "%RESPONSE_FILE%"
SET /P uploadID= < "%ID_FILE%"
DEL "%ID_FILE%"
SET uploadID=%uploadID:"=%          &:: Remove double quotes
SET uploadID=%uploadID: =%          &:: Remove spaces
SET uploadID=%uploadID::=%          &:: Remove colons
SET uploadID=%uploadID:,=%          &:: Remove commas
SET uploadID=%uploadID:UploadId=%   &:: Remove other text
ECHO Creation succeeded with upload-id:
ECHO    %uploadID%

:: Unset local vars before exit
SET RESPONSE_FILE=
SET ERROR_FILE=
SET ID_FILE=
SET UPLOAD_STR=
SET size=

EXIT /B 0

::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:: void initMultipartUpload()
::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:doMultipartUpload

SETLOCAL EnableDelayedExpansion

EXIT /B 0

:: Add initial lines to parts JSON file
ECHO Uploading object parts...
SET ETAGS_FILE=upload-part-etags.json
> %ETAGS_FILE% (
    ECHO {
    ECHO    "Parts": [
)

:: Count the number of MCD file parts in this directory
SET numFiles=0
FOR %%F IN (*.mcd.*) DO SET /A numFiles+=1

:: Add a JSON block for each object part
SET counter=0
FOR %%F IN (*.mcd.*) DO (
    :: Display progress
    SET /A counter+=1
    ECHO Uploading part !counter!/%numFiles%...
    
    :: Get the ETag for this object part =%
     : Trim leading spaces and remove escpaed quotes =%
    aws s3api upload-part --bucket "%bucket%" --key "%key%" --part-number !counter! --body %%F --upload-id "%uploadID%" --profile "%profile%" | findstr ETag> tmp.txt
    ECHO     "ETag\": \"adsfPOIjad89K"> tmp.txt
    SET /P etag= < tmp.txt
    SET etag=!etag: =!
    SET etag=!etag::=: !
    SET etag=!etag:\"="!
    
    :: Export the JSON block for this part
     : Makes sure there's no comma after the last one!
    IF !counter!==%numFiles% (SET closeBrace=}) ELSE (SET closeBrace=},)
    >> %ETAGS_FILE% (
        ECHO        {
        ECHO            !etag!,
        ECHO            "PartNumber": !counter!
        ECHO        !closeBrace!
    )
)
DEL tmp.txt

:: Add final lines to parts JSON file
>> %ETAGS_FILE% (
    ECHO    ]
    ECHO }
)
