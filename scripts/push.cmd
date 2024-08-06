@echo off

set scriptsdir=%~dp0
set root=%scriptsdir%\..
set destination=%root%\deploy
set version=%1

if "%version%"=="" (
  echo Please remember to specify which version to push as an argument.
  goto exit_fail
)

if not exist "%destination%" (
  echo Could not find %destination%
  echo.
  echo Did you remember to build the packages before running this script?
)

set nuget=%root%\tools\NuGet.exe

if not exist "%nuget%" (
  echo Could not find NuGet here:
  echo.
  echo    "%nuget%"
  echo.
  goto exit_fail
)

"%nuget%" push "%destination%\*.%version%.nupkg" -Source https://nuget.org
if %ERRORLEVEL% neq 0 (
  echo NuGet push failed.
  goto exit_fail
)

goto exit_success
:exit_fail
exit /b 1
:exit_success