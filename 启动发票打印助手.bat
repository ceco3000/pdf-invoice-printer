@echo off
setlocal EnableExtensions
cd /d "%~dp0"

set "CAND_FILE=%TEMP%\invoice_printer_python_candidates.txt"
set "PY="
set "PYW="
set "PY_READY="

call :reset_and_pick
if not defined PY call :install_python
if not defined PY goto :python_failed

if not defined PYW (
    for %%I in ("%PY%") do (
        if /i "%%~nxI"=="python.exe" (
            if exist "%%~dpIpythonw.exe" set "PYW=%%~dpIpythonw.exe"
        )
    )
)

echo Using Python: %PY%

if not defined PY_READY (
    echo Installing required packages...
    "%PY%" -m pip --version >nul 2>&1
    if errorlevel 1 (
        echo Preparing pip...
        "%PY%" -m ensurepip --upgrade >nul 2>&1
    )
    "%PY%" -m pip install --disable-pip-version-check -r requirements.txt
    if errorlevel 1 goto :deps_failed
    "%PY%" -c "import pymupdf, PIL, win32print, tkinterdnd2" >nul 2>&1
    if errorlevel 1 goto :deps_failed
)

if exist "%CAND_FILE%" del "%CAND_FILE%" >nul 2>&1

if defined PYW (
    start "" "%PYW%" "invoice_printer.py"
) else (
    "%PY%" "invoice_printer.py"
)
exit /b 0

:python_failed
if exist "%CAND_FILE%" del "%CAND_FILE%" >nul 2>&1
echo.
echo Python 3.12 could not be installed automatically.
echo Install it manually from https://www.python.org/downloads/windows/
pause
exit /b 1

:deps_failed
if exist "%CAND_FILE%" del "%CAND_FILE%" >nul 2>&1
echo.
echo Failed to install the required Python packages.
echo Run manually:
echo "%PY%" -m pip install -r requirements.txt
pause
exit /b 1

:install_python
echo.
echo Python was not found. Installing Python 3.12...

where winget >nul 2>&1
if not errorlevel 1 (
    echo Trying winget...
    winget install --id Python.Python.3.12 -e --source winget --scope user --silent --accept-package-agreements --accept-source-agreements --disable-interactivity
    call :reset_and_pick
    if defined PY exit /b 0
    echo winget did not complete successfully. Trying the official installer...
)

echo Downloading Python 3.12.10 from python.org...
powershell -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference='Stop'; $url='https://www.python.org/ftp/python/3.12.10/python-3.12.10-amd64.exe'; $out=Join-Path $env:TEMP 'python-3.12.10-amd64.exe'; try { [Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -Uri $url -OutFile $out -UseBasicParsing; $sig=Get-AuthenticodeSignature -FilePath $out; if ($sig.Status -ne 'Valid' -or $sig.SignerCertificate.Subject -notlike '*Python Software Foundation*') { throw 'Installer signature validation failed.' }; $p=Start-Process -FilePath $out -ArgumentList '/quiet InstallAllUsers=0 PrependPath=1 Include_pip=1 Include_launcher=1 Include_test=0' -Wait -PassThru; Remove-Item $out -Force -ErrorAction SilentlyContinue; if ($p.ExitCode -ne 0 -and $p.ExitCode -ne 3010) { throw ('Installer exit code: ' + $p.ExitCode) } } catch { Write-Host $_.Exception.Message; exit 1 }"
call :reset_and_pick
exit /b 0

:reset_and_pick
set "PY="
set "PYW="
set "PY_READY="
call :collect_python_candidates
call :pick_python
exit /b 0

:collect_python_candidates
type nul > "%CAND_FILE%"
for /f "delims=" %%I in ('python -c "import sys;print(sys.executable)" 2^>nul') do >>"%CAND_FILE%" echo %%I
for /f "delims=" %%I in ('py -3.12 -c "import sys;print(sys.executable)" 2^>nul') do >>"%CAND_FILE%" echo %%I
for /f "delims=" %%I in ('py -3 -c "import sys;print(sys.executable)" 2^>nul') do >>"%CAND_FILE%" echo %%I
>>"%CAND_FILE%" echo %LocalAppData%\Programs\Python\Python312\python.exe
>>"%CAND_FILE%" echo %LocalAppData%\Programs\Python\Python313\python.exe
>>"%CAND_FILE%" echo %LocalAppData%\Programs\Python\Python311\python.exe
>>"%CAND_FILE%" echo %LocalAppData%\Programs\Python\Python310\python.exe
>>"%CAND_FILE%" echo %ProgramFiles%\Python312\python.exe
>>"%CAND_FILE%" echo %ProgramFiles%\Python313\python.exe
>>"%CAND_FILE%" echo %ProgramFiles%\Python311\python.exe
>>"%CAND_FILE%" echo %ProgramFiles%\Python310\python.exe
>>"%CAND_FILE%" echo %ProgramFiles(x86)%\Python312\python.exe
>>"%CAND_FILE%" echo %ProgramFiles(x86)%\Python311\python.exe
>>"%CAND_FILE%" echo %ProgramFiles(x86)%\Python310\python.exe
call :collect_registry
exit /b 0

:collect_registry
call :append_registry "HKCU\Software\Python\PythonCore\3.12\InstallPath"
call :append_registry "HKLM\Software\Python\PythonCore\3.12\InstallPath"
call :append_registry "HKCU\Software\Python\PythonCore\3.13\InstallPath"
call :append_registry "HKLM\Software\Python\PythonCore\3.13\InstallPath"
call :append_registry "HKCU\Software\Python\PythonCore\3.11\InstallPath"
call :append_registry "HKLM\Software\Python\PythonCore\3.11\InstallPath"
call :append_registry "HKCU\Software\Python\PythonCore\3.10\InstallPath"
call :append_registry "HKLM\Software\Python\PythonCore\3.10\InstallPath"
exit /b 0

:append_registry
set "PYDIR="
for /f "tokens=2,*" %%A in ('reg query "%~1" /ve 2^>nul ^| findstr /i "REG_SZ"') do set "PYDIR=%%B"
if defined PYDIR if exist "%PYDIR%python.exe" >>"%CAND_FILE%" echo %PYDIR%python.exe
exit /b 0

:pick_python
if not exist "%CAND_FILE%" exit /b 0
for /f "usebackq delims=" %%I in ("%CAND_FILE%") do (
    if not defined PY_READY call :consider "%%I"
)
exit /b 0

:consider
if defined PY_READY exit /b 0
if not exist "%~1" exit /b 0
"%~1" -c "import sys; raise SystemExit(0 if sys.version_info >= (3,10) else 1)" >nul 2>&1
if errorlevel 1 exit /b 0
if not defined PY set "PY=%~1"
"%~1" -c "import pymupdf, PIL, win32print, tkinterdnd2" >nul 2>&1
if not errorlevel 1 (
    set "PY=%~1"
    set "PY_READY=1"
)
exit /b 0
