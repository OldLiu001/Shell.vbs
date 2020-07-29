@echo off
if not "%VBSH_Home%"=="" (
	echo You have installed the program. Do you want to reinstalled the program?
	echo Installtion will be started after 10 seconds.
	TimeOut /T 10 /NoBreak
)

:Install
echo Setting up environmental varibles...
echo Creating VBSH_Home
setx VBSH_Home %cd%>nul
echo Creating VBSH_FLib
setx VBSH_FLib %cd%\librarys\VBScript>nul
echo Creating VBSH_DLL
setx VBSH_DLL %cd%\librarys\ActiveX>nul
pause