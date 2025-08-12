@echo off
setlocal enabledelayedexpansion

set "max_counter=0"
set "counter=0"
cd /D "C:\Users\uiv16084\SAM_Pictures"
set "workdir=C:\Users\uiv16084\SAM_Pictures"

echo Working directory: %workdir%

:: Loop through .tiff files, max 5
for /f "tokens=1* delims=:" %%A in ('dir /b /a-d /o-d ^| findstr /n /e ".tiff"') do (
    if !max_counter! GEQ 5 (
        goto :done
    )
    set "file.%%A=%%B"
    echo Found file: %%B
    set /a counter+=1
	set /a max_counter+=1
)

:done
echo Processed !max_counter! files.

:: Check if any files were found
set "HasImg=False"
if !counter! GTR 0 (
    set "HasImg=True"
)

if "!HasImg!"=="True" (
    call "P:\SAM_Daily_Save\DailyPicSave.bat"
) else (
    echo No images found. Exiting.
    exit /b
)
endlocal
