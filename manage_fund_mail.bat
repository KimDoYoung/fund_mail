@echo off
cd /d "%~dp0"

set VENV_PATH=.venv
set PYTHON_SCRIPT=src/service_wrapper.py
set SERVICE_NAME=FundMailFetchService

echo ====================================
echo Fund Email Fetch Service Management
echo ====================================

:MENU
echo.
echo 1. Install Service
echo 2. Start Service
echo 3. Stop Service
echo 4. Remove Service
echo 5. Service Status
echo 6. View Logs
echo 7. Exit
echo.
set /p choice=Select option (1-7): 

if "%choice%"=="1" goto INSTALL
if "%choice%"=="2" goto START
if "%choice%"=="3" goto STOP
if "%choice%"=="4" goto REMOVE
if "%choice%"=="5" goto STATUS
if "%choice%"=="6" goto LOGS
if "%choice%"=="7" goto EXIT

echo Invalid choice. Please try again.
goto MENU

:INSTALL
echo Installing service...
call %VENV_PATH%\Scripts\activate.bat
python %PYTHON_SCRIPT% install
if %ERRORLEVEL% EQU 0 (
    echo Service installed successfully!
) else (
    echo Failed to install service!
)
pause
goto MENU

:START
echo Starting service...
net start %SERVICE_NAME%
if %ERRORLEVEL% EQU 0 (
    echo Service started successfully!
) else (
    echo Failed to start service!
)
pause
goto MENU

:STOP
echo Stopping service...
net stop %SERVICE_NAME%
if %ERRORLEVEL% EQU 0 (
    echo Service stopped successfully!
) else (
    echo Failed to stop service!
)
pause
goto MENU

:REMOVE
echo Removing service...
net stop %SERVICE_NAME% 2>nul
call %VENV_PATH%\Scripts\activate.bat
python %PYTHON_SCRIPT% remove
if %ERRORLEVEL% EQU 0 (
    echo Service removed successfully!
) else (
    echo Failed to remove service!
)
pause
goto MENU

:STATUS
echo Checking service status...
sc query %SERVICE_NAME%
pause
goto MENU

:LOGS
echo Opening log directory...
start C:\fund_mail\logs
goto MENU

:EXIT
echo Goodbye!
pause