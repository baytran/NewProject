@echo off

set LIB_PATH=%~dp0
cd %LIB_PATH%

copy /Y "..\Recovery Scenario\RecoveryScenarioHandler.vbs" "C:\RecoveryScenarioHandler.vbs"
copy /Y "..\Recovery Scenario\RS_LogAndExitEncompass.qrs" "C:\RS_LogAndExitEncompass.qrs"

if exist "*.qfl" (
	del /f /q *.qfl
)

set BIZ_LIB=Business.qfl
set FRM_LIB=Framework.qfl
set UTL_LIB=Util.qfl

for /r "%LIB_PATH%\Business" %%a in ("*.qfl") do (
	if exist "%%a" (
		echo append file "%%a" into %BIZ_LIB% 
		echo ''''''''''''''''''''''''''''''''''''''''''''''''''''>> %BIZ_LIB%
		echo ' File "%%a" Start >> %BIZ_LIB%
		echo ''''''''''''''''''''''''''''''''''''''''''''''''''''>> %BIZ_LIB%
		type "%%a" >> %BIZ_LIB%
		echo ''''''''''''''''''''''''''''''''''''''''''''''''''''>> %BIZ_LIB%
		echo ' File "%%a" End >> %BIZ_LIB%
		echo ''''''''''''''''''''''''''''''''''''''''''''''''''''>> %BIZ_LIB%
	)
)

for /r "%LIB_PATH%\Framework" %%a in ("*.qfl") do (
	if exist "%%a" (
		echo append file "%%a" into %FRM_LIB% 
		echo ''''''''''''''''''''''''''''''''''''''''''''''''''''>> %FRM_LIB%
		echo ' File "%%a" Start >> %FRM_LIB%
		echo ''''''''''''''''''''''''''''''''''''''''''''''''''''>> %FRM_LIB%
		type "%%a" >> %FRM_LIB%
		echo ''''''''''''''''''''''''''''''''''''''''''''''''''''>> %FRM_LIB%
		echo ' File "%%a" End >> %FRM_LIB%
		echo ''''''''''''''''''''''''''''''''''''''''''''''''''''>> %FRM_LIB%
	)
)

for /r "%LIB_PATH%\GUI" %%a in ("*.qfl") do (
	if exist "%%a" (
		echo append file "%%a" into %FRM_LIB% 
		echo ''''''''''''''''''''''''''''''''''''''''''''''''''''>> %FRM_LIB%
		echo ' File "%%a" Start >> %FRM_LIB%
		echo ''''''''''''''''''''''''''''''''''''''''''''''''''''>> %FRM_LIB%
		type "%%a" >> %FRM_LIB%
		echo ''''''''''''''''''''''''''''''''''''''''''''''''''''>> %FRM_LIB%
		echo ' File "%%a" End >> %FRM_LIB%
		echo ''''''''''''''''''''''''''''''''''''''''''''''''''''>> %FRM_LIB%
	)
)

for /r "%LIB_PATH%\Util" %%a in ("*.qfl") do (
	if exist "%%a" (
		echo append file "%%a" into %UTL_LIB% 
		echo ''''''''''''''''''''''''''''''''''''''''''''''''''''>> %UTL_LIB%
		echo ' File "%%a" Start >> %UTL_LIB%
		echo ''''''''''''''''''''''''''''''''''''''''''''''''''''>> %UTL_LIB%
		type "%%a" >> %UTL_LIB%
		echo ''''''''''''''''''''''''''''''''''''''''''''''''''''>> %UTL_LIB%
		echo ' File "%%a" End >> %UTL_LIB%
		echo ''''''''''''''''''''''''''''''''''''''''''''''''''''>> %UTL_LIB%
	)
)

goto:eof




