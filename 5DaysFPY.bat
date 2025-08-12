::Using todaywhat.bat to define which CW we on
		@echo off
		call "P:\5DaysFPY\Scripts\TodayWhat.bat"
	::Open save location for the actual weekly report
		cd /D "P:\5DaysFPY\Scripts"
	:: Read the week lastdaydber, day, dayNo, month; from the file
		set /p todayw=<todayw.txt & set /p todayd=<todayd.txt & set /p todaydNo=<todaydNo.txt & set /p todaym=<todaym.txt
	::Read the PWs and Users
		cd /D "P:\UserInfo"
		set /p mysecurity=<mysecurity.txt & set /p myuser=<myuser.txt
::Decide if today is Monday or any other day, so it knows if it is Last Weekend or Last Day report needed
if %todaydNO% gtr 1 (
		set lastwhat="Last day"
    ) else (
        set lastwhat="Last weekend"
    )
::Reset remio Tool
setlocal
tasklist | findstr /I "MES Report from Dremio.exe" >nul
if %ERRORLEVEL%==0 (
echo MES Report from Dremio.exe found, the script closes the app, then opens a new instance.
:: Terminate the process
taskkill /IM "MES Report from Dremio.exe" /F
) else (
echo MES Report from Dremio.exe not found, the script opens the app now.
)
endlocal
::Run Dremio Tool for DQ200 & 400 linereport, and login and setup run, commented out to make debug easier
	"Z:\99_TestEngTools\11_MES Report from Dremio\MES Report from Dremio.exe" -psw %mysecurity% -timerange %lastwhat% -retest 7 -report "DQ200 Line Report" "DQ400 Line Report"	
	::This line was only created to see if I can open the file folder correctly with a dynamic Week calculation
	::start %windir%\explorer.exe "X:\97_Weekly Report\CW%todayw%\DQ200 Line Report\"
:::::DQ200::::::::::DQ200::::::::::DQ200::::::::::DQ200::::::::::DQ200::::::::::DQ200::::::::::DQ200::::::::::DQ200:::::
::If it is Monday report should look for last weeks folder, also it will set MondayRedW to make sure we only modify variables that come from monday
::Trim the CW0x to just x to be able to make the x-1 week calc, for Monday	
REM Check if todaydW is less than 9
if %todayw:~0,1% == 0 (
    REM Trim todayw to get the second character
    set /A todaywtrim=%todayw:~1,1%
) else (
    REM Trim todayw to get the first two characters
    set /A todaywtrim=%todayw:~0,2%
)

::Calculate x-1week for Mondays, IF we are after CW10
if %todayw% GEQ 10 (
 if %todaydNO% == 1 (set /A todayw=%todaywtrim%-1 & set /A MondayRedW=1
 )  else (set /A MondayRedW=0)
)
if %MondayRedW% == 1 (
		set todayw=%todayw%
	)
if %todayw% LSS 10 (
 if %todaydNO% == 1 (set /A todayw=%todaywtrim% & set todayw=0%todayw%
 )  else (set /A MondayRedW=0)
)
::Change directory to start the find of LastDay run Excel
		cd /D "X:\97_Weekly Report\CW%todayw%\DQ200 Line Report\"
::::::::::::::::::::Simple file content output to 1 array (not used here)::::::::::::::::::::
::setlocal enabledelayedexpansion
::set params=
::for /f "delims=" %%a in ('dir /s/b') do set params=!params! %%a
::Echo !params!
::pause
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

::Find all the matching files, put to array::::::::::::::::::::
		@echo off
		setlocal enableDelayedExpansion
	:: Load the file path "array"
	::Set the monthday variable, which is coming from the TodayWhat.bat file, bringing the actual date of run
		set monthday=%todaym%%todayd%
	::Findstr /e function, to look from the ending of the file name for a match, based on a triggerword
		for /f "tokens=1* delims=:" %%A in ('dir /s /b^|findstr /n /e "%monthday%.xlsx"') do (
		set "file.%%A=%%B"
		set "file.count=%%A"
		)
:: Access the values segment
	:: This line outputs all finds for the files array
		for /l %%N in (1 1 !file.count!) do set actKPIpath=!file.%%N!
	:: This line only puts out our first find, for now this is good
		::echo Path:!actKPIpath!
::The searched file is not found go to error
if not exist "!actKPIpath!" echo Your file not found in the directory! Run the report again please & goto :end
)
	::TRIMs down path and today's date from file name, leaving only the reports starting date as an output
	::while using the string based trim you cant use variable values, so for the working solution, I'm removing the first 61 chars, and taking the 4 coming chars only
		SET "_test=!actKPIpath!"
		set lastdayd=%_test:~61,4%
		::SET "_lastdayd=%_test:X:\97_Weekly Report\CW46\DQ200 Line Report\DQ200 Line Report@=%" ::Not working
		::SET lastdayd=%_lastdayd:~0,-10% :: Solution, for removing the last 10 chars, not needed
	::Here I'm just visual checking if the excel file xxyy- date in the name is what i wanted, it is
		::Echo %lastdayd%

	::This was just to see the output of the array element I'm going to use
		::echo !actKPIpath!
	::Save the actKPIpath to txt
		cd /D "P:\5DaysFPY\Scripts"
		echo/!actKPIpath!> actKPIpath.txt
		echo/%lastdayd%> lastdayd.txt
::NEXT STEP: write the segment of calculation, whether it is a weekly report or a daily/weekendly. And only process the daily report file of today
::Possible end of month turning variations:
::February, February+1: F=28, M=3, F=29 M=3; F=27 M=2, F=28 M=2; F=26 M=1, F27 M=1 
::30 day month: F=30 M=3, F=29 M=2, F=28 M=1
::31 day month: F=31 M=3,F=30 M=2, F=29 M=1

::/a is set for making sure this is an integer, adding +1 with a dec offset makes sure batch doesnt think our lastdaydber is an octal value
		::echo lastdayd: %lastdayd%
		set "lastdayd=1%lastdayd%"
		set "monthday=1%monthday%"
		set /a dayspassed=%monthday%-%lastdayd%
::echo %dayspassed%
::echo %todayd%
::echo %lastdayd%
if %dayspassed% gtr 1 (
    if %dayspassed% leq 4 (
        echo You are running a %lastwhat% report for the file: !actKPIpath!
    ) else (
        echo The automatic solution found the weekly report from Monday first, instead of the %lastwhat% report you wanted. Please replace or remove the weekly report: !actKPIpath!
    )
) else if %dayspassed%==1 (
    echo You are running a %lastwhat% report for the file: !actKPIpath!
) else (
    echo The automatic solution found a month end crossover report. Please manually check the value, to ensure you get what you wanted: !actKPIpath!
)
::If CWxx folder does not exists create it and copy today's file
if not exist "P:\5DaysFPY\RawExcel\CW%todayw%\" mkdir "P:\5DaysFPY\RawExcel\CW%todayw%\" & copy "X:\97_Weekly Report\CW%todayw%\DQ200 Line Report\DQ200 Line Report@????-??%todayd%.xlsx" "P:\5DaysFPY\RawExcel\CW%todayw%\DQ200 Line Report@????-??%todayd%.xlsx"
::If CWxx file exists but the file does not copy it
IF not EXIST "P:\5DaysFPY\RawExcel\CW%todayw%\DQ200 Line Report@????-??%todayd%.xlsx" copy "X:\97_Weekly Report\CW%todayw%\DQ200 Line Report\DQ200 Line Report@????-??%todayd%.xlsx" "P:\5DaysFPY\RawExcel\CW%todayw%\DQ200 Line Report@????-??%todayd%.xlsx"
::Here comes the local excel evaluation part
	::Removed function|Rename today's report to TodayKPIReport
	::Removed function|ren "DQ200 Line Report@????-??%todayd%.xlsx" "TodayKPIReport.xlsx"
	::Run VBS script, that will create XLSM from XLSX 
		cd /D "P:\5DaysFPY\Scripts\"
		%SystemRoot%\System32\cscript "XLSXtoXLSM_Rmacro_DQ200.vbs" "P:\5DaysFPY\RawExcel\CW%todayw%\DQ200 Line Report@????-??%todayd%.xlsx"
::taskkill /IM EXCEL.exe /F

:::::DQ200::::::::::DQ200::::::::::DQ200::::::::::DQ200::::::::::DQ200::::::::::DQ200::::::::::DQ200::::::::::DQ200::::::::::DQ200::::::::::DQ200::::::::::DQ200


:::::DQ400::::::::::DQ400::::::::::DQ400::::::::::DQ400::::::::::DQ400::::::::::DQ400::::::::::DQ400::::::::::DQ400:::::
	::Open save location for the actual weekly report
		cd /D "P:\5DaysFPY\Scripts"
	:: Read the week lastdaydber, day, dayNo, month; from the file
		set /p todayw=<todayw.txt & set /p todayd=<todayd.txt & set /p todaydNo=<todaydNo.txt & set /p todaym=<todaym.txt
	::This line was only created to see if I can open the file folder correctly with a dynamic Week calculation
	::start %windir%\explorer.exe "X:\97_Weekly Report\CW%todayw%\DQ400 Line Report\"
::Decide if today is Monday or any other day, so it knows if it is Last Weekend or Last Day report needed
if %todaydNO% gtr 1 (
		set lastwhat="Last day"
    ) else (
        set lastwhat="Last weekend"
    )
	::This line was only created to see if I can open the file folder correctly with a dynamic Week calculation
	::start %windir%\explorer.exe "X:\97_Weekly Report\CW%todayw%\DQ200 Line Report\"
:::::DQ200::::::::::DQ200::::::::::DQ200::::::::::DQ200::::::::::DQ200::::::::::DQ200::::::::::DQ200::::::::::DQ200:::::
::If it is Monday report should look for last weeks folder, also it will set MondayRedW to make sure we only modify variables that come from monday
::Trim the CW0x to just x to be able to make the x-1 week calc, for Monday	
REM Check if todaydW is less than 9
if %todayw:~0,1% == 0 (
    REM Trim todayw to get the second character
    set /A todaywtrim=%todayw:~1,1%
) else (
    REM Trim todayw to get the first two characters
    set /A todaywtrim=%todayw:~0,2%
)

::Calculate x-1week for Mondays, IF we are after CW10
if %todayw% GEQ 10 (
 if %todaydNO% == 1 (set /A todayw=%todaywtrim%-1 & set /A MondayRedW=1
 )  else (set /A MondayRedW=0)
)
if %MondayRedW% == 1 (
		set todayw=%todayw%
	)
if %todayw% LSS 10 (
 if %todaydNO% == 1 (set /A todayw=%todaywtrim% & set todayw=0%todayw%
 )  else (set /A MondayRedW=0)
)
::Change directory to start the find of LastDay run Excel
		cd /D "X:\97_Weekly Report\CW%todayw%\DQ400 Line Report\"
::::::::::::::::::::Simple file content output to 1 array (not used here)::::::::::::::::::::
::setlocal enabledelayedexpansion
::set params=
::for /f "delims=" %%a in ('dir /s/b') do set params=!params! %%a
::Echo !params!
::pause
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

::Find all the matching files, put to array::::::::::::::::::::
::Change directory to start the find of LastDay run Excel
		cd /D "X:\97_Weekly Report\CW%todayw%\DQ400 Line Report\"
::::::::::::::::::::Simple file content output to 1 array (not used here)::::::::::::::::::::
::setlocal enabledelayedexpansion
::set params=
::for /f "delims=" %%a in ('dir /s/b') do set params=!params! %%a
::Echo !params!
::pause
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

::Find all the matching files, put to array::::::::::::::::::::
		@echo off
		setlocal enableDelayedExpansion
	:: Load the file path "array"
	::Set the monthday variable, which is coming from the TodayWhat.bat file, bringing the actual date of run
		set monthday=%todaym%%todayd%
	::Findstr /e function, to look from the ending of the file name for a match, based on a triggerword
		for /f "tokens=1* delims=:" %%A in ('dir /s /b^|findstr /n /e "%monthday%.xlsx"') do (
		set "file.%%A=%%B"
		set "file.count=%%A"
		)
:: Access the values segment
	:: This line outputs all finds for the files array
		for /l %%N in (1 1 !file.count!) do set actKPIpath=!file.%%N!
	:: This line only puts out our first find, for now this is good
		::echo Path:!actKPIpath!
	::TRIMs down path and today's date from file name, leaving only the reports starting date as an output
	::while using the string based trim you cant use variable values, so for the working solution, I'm removing the first 61 chars, and taking the 4 coming chars only
		SET "_test=!actKPIpath!"
		SET lastdayd=%_test:~61,4%
		::SET "_lastdayd=%_test:X:\97_Weekly Report\CW46\DQ400 Line Report\DQ400 Line Report@=%" ::Not working
		::SET lastdayd=%_lastdayd:~0,-10% :: Solution, for removing the last 10 chars, not needed
	::Here I'm just visual checking if the excel file xxyy- date in the name is what i wanted, it is
		::Echo %lastdayd%

	::This was just to see the output of the array element I'm going to use
		::echo !actKPIpath!
	::Save the actKPIpath to txt
		cd /D "P:\5DaysFPY\Scripts"
		echo/!actKPIpath!> actKPIpath.txt

::NEXT STEP: write the segment of calculation, whether it is a weekly report or a daily/weekendly. And only process the daily report file of today
::Possible end of month turning variations:
::February, February+1: F=28, M=3, F=29 M=3; F=27 M=2, F=28 M=2; F=26 M=1, F27 M=1 
::30 day month: F=30 M=3, F=29 M=2, F=28 M=1
::31 day month: F=31 M=3,F=30 M=2, F=29 M=1

::/a is set for making sure this is an integer
		set "lastdayd=1%lastdayd%"
		set "monthday=1%monthday%"
set /a dayspassed=%monthday%-%lastdayd%
::echo %dayspassed%
::echo %todayd%
::echo %lastdayd%

if %dayspassed% gtr 1 (
    if %dayspassed% leq 4 (
        echo You are running a %lastwhat% report for the file: !actKPIpath!
    ) else (
        echo The automatic solution found the weekly report from Monday first, instead of the %lastwhat% report you wanted. Please replace or remove the weekly report: !actKPIpath!
    )
) else if %dayspassed%==1 (
    echo You are running a %lastwhat% report for the file: !actKPIpath!
) else (
    echo The automatic solution found a month end crossover report. Please manually check the value, to ensure you get what you wanted: !actKPIpath!
)
::If CWxx folder does not exists create it and copy today's file
if not exist "P:\5DaysFPY\RawExcel\CW%todayw%\" mkdir "P:\5DaysFPY\RawExcel\CW%todayw%\" & copy "X:\97_Weekly Report\CW%todayw%\DQ400 Line Report\DQ400 Line Report@????-??%todayd%.xlsx" "P:\5DaysFPY\RawExcel\CW%todayw%\DQ400 Line Report@????-??%todayd%.xlsx"


::If CWxx folder exists but the file does not copy it
IF not EXIST "P:\5DaysFPY\RawExcel\CW%todayw%\DQ400 Line Report@????-??%todayd%.xlsx" copy "X:\97_Weekly Report\CW%todayw%\DQ400 Line Report\DQ400 Line Report@????-??%todayd%.xlsx" "P:\5DaysFPY\RawExcel\CW%todayw%\DQ400 Line Report@????-??%todayd%.xlsx"

::Here comes the local excel evaluation part
	::Removed function|Rename today's report to TodayKPIReport
	::Removed function|ren "DQ400 Line Report@????-??%todayd%.xlsx" "TodayKPIReport.xlsx"
	::Run VBS script, that will create XLSM from XLSX 
		cd /D "P:\5DaysFPY\Scripts\"
		%SystemRoot%\System32\cscript "XLSXtoXLSM_Rmacro_DQ400.vbs" "P:\5DaysFPY\RawExcel\CW%todayw%\DQ400 Line Report@????-??%todayd%.xlsx"
taskkill /IM EXCEL.exe /F
:::::DQ400::::::::::DQ400::::::::::DQ400::::::::::DQ400::::::::::DQ400::::::::::DQ400::::::::::DQ400::::::::::DQ400::::::::::DQ400::::::::::DQ400::::::::::DQ400
::Run Macro for creating the 5 days chart the top FPY losses
		call "P:\5DaysFPY\Scripts\CreateChart_DQx00.bat"
		::%SystemRoot%\System32\cscript "CreateChart_DQx00.vbs" "P:\5DaysFPY\Database\DQ200_Rep_TP_DB.xlsm"
::Setup the RNG for the trigger.txt(used in the last lines of the code) ::Write Trigger file4
SET /A RNGvalue=%RANDOM% * 30000 / 50000 + 1
cd /D "P:\5DaysFPY\FlowTrigger"
echo/ %todayd% %RNGvalue% > Trigger.txt
::Sync OneDrive
powershell -Command "Start-Process -FilePath 'C:\Program Files\Microsoft OneDrive\OneDrive.exe' -ArgumentList '/background'"
::Sync OneDrive   ::OUTDATED UPDATE ONEDRIVE VERSION
::cd /D "C:\Program Files\Microsoft OneDrive"
::OneDrive.exe /sync

::Archive the result to FinalExcel folder
if not exist "P:\5DaysFPY\FinalExcel\CW%todayw%\" mkdir "P:\5DaysFPY\FinalExcel\CW%todayw%\" & copy "P:\5DaysFPY\Database\DQ200_Rep_TP_DB.xlsm" "P:\5DaysFPY\FinalExcel\CW%todayw%\DQ200_%todaym%_%todayd%.xlsm"
if exist "P:\5DaysFPY\FinalExcel\CW%todayw%\" copy "P:\5DaysFPY\Database\DQ200_Rep_TP_DB.xlsm" "P:\5DaysFPY\FinalExcel\CW%todayw%\DQ200_%todaym%_%todayd%.xlsm"
::Archive the result to FinalExcel folder
if not exist "P:\5DaysFPY\FinalExcel\CW%todayw%\" mkdir "P:\5DaysFPY\FinalExcel\CW%todayw%\" & copy "P:\5DaysFPY\Database\DQ400_Rep_TP_DB.xlsm" "P:\5DaysFPY\FinalExcel\CW%todayw%\DQ400_%todaym%_%todayd%.xlsm"
if exist "P:\5DaysFPY\FinalExcel\CW%todayw%\" copy "P:\5DaysFPY\Database\DQ400_Rep_TP_DB.xlsm" "P:\5DaysFPY\FinalExcel\CW%todayw%\DQ400_%todaym%_%todayd%.xlsm"