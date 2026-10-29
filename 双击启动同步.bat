@echo off
setlocal
title InfiniteGithubUpstreamSync
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0InfiniteGithubUpstreamSync.ps1" -Menu
if errorlevel 1 pause
