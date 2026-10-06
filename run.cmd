@echo off
REM One-click launcher for the Nature Index -> PNU Scholar agent.
REM Passes any extra args through to run.ps1 (e.g. run.cmd -FreshCache).
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0run.ps1" %*
