::=======================================================This BAT runs a script in vbs, that runs macros on demand=============================================================
cd /D "P:\Daily KPI Report\Scripts"
set /p todayw=<todayw.txt & set /p todayd=<todayd.txt
::echo %todayd%
::echo %todayw%
cscript XLSXtoXLSM_Rmacro_ALLtestKPI.vbs "P:\FF2 and FADQ Test_KPI\RawExcel\CW%todayw%\DQ Test Stations Report@????-??%todayd%.xlsx
::echo P:\Daily KPI Report\RawExcel\CW%todayw%\DQ Test Stations Report@????-??%todayd%.xlsx
::=============================================================================================================================================================================