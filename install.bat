@echo off
chcp 65001 >nul
title EXCLE助手 · 一键安装
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0install.ps1"
echo.
if %errorlevel% neq 0 (
  echo 安装失败，请查看上方错误信息。
) else (
  echo 安装完成！请重启 Excel 或 WPS 表格生效。
)
pause
