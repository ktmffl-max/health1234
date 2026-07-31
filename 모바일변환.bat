@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo.
echo  운동프로그램 워크북 -^> 모바일 HTML
echo  ----------------------------------
where python >nul 2>nul
if %errorlevel%==0 (python build_mobile.py) else (py -3 build_mobile.py)
echo.
pause
