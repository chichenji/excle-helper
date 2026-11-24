@echo off
chcp 65001 >nul
title EXCLE助手 · 一键卸载
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0uninstall.ps1"
echo.
if %errorlevel% neq 0 (
  echo 卸载失败，请查看上方错误信息。
) else (
  echo 卸载完成！请重启 Excel 或 WPS 表格生效。
)
pause
