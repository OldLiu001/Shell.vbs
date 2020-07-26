@echo off
if not "%VBSH_Home%"=="" (
	echo You have installed the program. Do you want to reinstalled the program?
	if %errorlevel%==1 (
		goto Install
	)
	else
	(
		exit
	)
)

:Install
echo Setting up environmental varibles...
setx VBSH_Home %cd%>nul
setx VBSH_FLib %cd%\lib\vbs>nul
setx VBSH_DLL %cd%\lib\dll>nul

echo Creating shortcup of VBS Shell...
copy VBSH* %windir%\system32>nul