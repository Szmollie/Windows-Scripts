::Wirte a Batch for organizing the SAM Pictures saved folder
::-Criteria1: Run this batch from TaskScheduler every day between 5:30p.m. 6:30p.m., if missed run try tomorrow. Save predecessor days only.

::e.g. 202502051404_00000AAA9107010000@0000368203@24092703395_Scan1_Channel0
::-Criteria2: Based on file save name, determine if it is today generated. If it is, save it to a newly created folder, named with the date.
::-Error handling: 
	::-If there is a folder already with today day, increase "date_value++"
	::-There are no runs from that day create the folder e.g."date_NoTestThisDay"
	::-Double check for days before yesterday, if anything left, move to correct folder. (if, for--, array)
	::-
::-
::-

cd /D "C:\\Users\uiv16084\SAM_Pictures"

set "max_counter=0"
set path=%cd%
echo %PATH%
set var=dir /b /a-d /o-d 
::run the dir command on pre set path
%var%%path%
@echo off
setlocal enabledelayedexpansion
::Filter return message, store original filename in %msg% variable, write %msg%
for /F "tokens=1*" %%i IN (' %var% ') DO (
    set "array=!array! %%i"
)

::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
::Finish this segment: Here cut the array elemnts to date, Then create a new array that only contains said day elements
set "index=1"  REM Index for the 2nd element (starting from 0)
set "counter=0"
::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
::This esegment is for validating the array existance and element correctly stored, and can be referenced individually in the future
set "index=1"  REM Index for the 2nd element (starting from 0)
set "counter=0"
::
for %%i in (%array%) do (
    if !counter! == %index% (
        echo 2nd element: %%i
		set msg=%%i
    )
    set /a counter+=1
)
::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
::fileday: the date that the first file in the folder has as first 8 characters
		SET "_test=%msg%"
		set fileday=%_test:~0,8%
		::echo %fileday%

:: Create folder named after %fileday% if it doesn't exist
if not exist "%fileday%" (
    mkdir "%fileday%"
)

:: Move all files that start with %fileday% into the folder
for %%F in (*.*) do (
	if !max_counter! GEQ 600 (
		echo !max_counter!, more than 600 files found, stopped after 600 files moved
        goto :done
    )
	set /a max_counter+=1
    set "filename=%%~nF"
    setlocal enabledelayedexpansion
    set "prefix=!filename:~0,8!"
    if "!prefix!"=="%fileday%" (
        echo Moving file: %%F to folder: %fileday%
        move "%%F" "%fileday%\"
    )
    endlocal
)
echo !max_counter!, less than 600 files found, stopped after full run
:done
exit