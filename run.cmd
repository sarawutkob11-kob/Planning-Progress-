@echo off
rem Planning Progress - same as run.ps1, for when PowerShell scripts are blocked.
rem   run              open the app (syncs with the real Excel file)
rem   run -NoSync      open the app, read-only towards Excel
rem   run sync status  any other uat_tracker command
setlocal
set "HERE=%~dp0"
set "PY=%USERPROFILE%\.venvs\uat-tracker\Scripts\python.exe"
if not exist "%PY%" (
  echo Python environment not found: %PY%
  exit /b 1
)
pushd "%HERE%"
if "%~1"=="" (
  "%PY%" -m uat_tracker --config "%HERE%config.toml" ui
) else if /i "%~1"=="-NoSync" (
  "%PY%" -m uat_tracker --config "%HERE%config.toml" ui --no-sync
) else (
  "%PY%" -m uat_tracker --config "%HERE%config.toml" %*
)
set RC=%ERRORLEVEL%
popd
exit /b %RC%
