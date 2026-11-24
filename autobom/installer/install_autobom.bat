@echo off
chcp 65001 >nul
title AutoBOM 一键安装
cd /d "%~dp0"
echo.
echo   ==========================================
echo     AutoBOM 一键安装（会自动申请管理员权限）
echo   ==========================================
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0install_autobom.ps1" %*
if errorlevel 1 pause