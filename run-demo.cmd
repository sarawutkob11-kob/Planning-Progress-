@echo off
rem Planning Progress - DEMO. Uses config.demo.toml: a copy of the workbook in %USERPROFILE%\PlanningProgressDemo
rem and its own data folder %LOCALAPPDATA%\PlanningProgress_DEMO. The real workbook and real data are never touched.
rem   run-demo              open the demo app
rem   run-demo -NoSync      open the demo app, read-only towards the demo Excel copy
rem   run-demo <command>    any other uat_tracker command, on the demo data
setlocal
set "HERE=%~dp0"
set "PY=%USERPROFILE%\.venvs\uat-tracker\Scripts\python.exe"
set "CFG=%HERE%config.demo.toml"
if not exist "%PY%" (
  echo Python environment not found: %PY%
  exit /b 1
)
if not exist "%CFG%" (
  echo Demo config not found: %CFG%
  exit /b 1
)
title Planning Progress - DEMO
pushd "%HERE%"
if "%~1"=="" (
  "%PY%" -m uat_tracker --config "%CFG%" ui
) else if /i "%~1"=="-NoSync" (
  "%PY%" -m uat_tracker --config "%CFG%" ui --no-sync
) else (
  "%PY%" -m uat_tracker --config "%CFG%" %*
)
set RC=%ERRORLEVEL%
popd
exit /b %RC%
