@echo off
rem slidesmith installer for Windows 11 - per-user, no admin rights.
rem Finds Python 3.10+ and runs install.py. Options are passed through, e.g.
rem   install.cmd --workspace C:\work\my-kiro-project
rem   install.cmd --index-url https://pypi.mycorp.example/simple
rem   install.cmd --uninstall
setlocal
cd /d "%~dp0"

set "PY="
where py >nul 2>nul && (
  py -3 -c "import sys; sys.exit(0 if sys.version_info >= (3, 10) else 1)" >nul 2>nul && set "PY=py -3"
)
if not defined PY (
  where python >nul 2>nul && (
    python -c "import sys; sys.exit(0 if sys.version_info >= (3, 10) else 1)" >nul 2>nul && set "PY=python"
  )
)
if not defined PY (
  echo.
  echo Python 3.10 or newer was not found.
  echo Install it for your user only ^(no admin needed^), then run install.cmd again:
  echo     winget install -e --id Python.Python.3.12 --scope user
  echo or get it from Company Portal / Software Center, or https://www.python.org/downloads/windows/
  echo ^(tick "Add python.exe to PATH"^).
  set "RC=1"
  goto :end
)

%PY% "%~dp0install.py" %*
set "RC=%errorlevel%"

:end
rem keep the window open when started by double-click from Explorer
echo %cmdcmdline% | findstr /i /c:"%~nx0" >nul && (echo. & pause)
exit /b %RC%
