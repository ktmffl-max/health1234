@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo.
echo  [1/2] 워크북 -^> 모바일 HTML
echo  --------------------------------
where python >nul 2>nul
if %errorlevel%==0 (python build_mobile.py) else (py -3 build_mobile.py)
if errorlevel 1 goto :fail

echo.
echo  [2/2] Netlify 배포
echo  --------------------------------
if not exist "docs\index.html" (echo  docs\index.html 이 없습니다. 변환이 실패한 것 같습니다. & goto :fail)
call netlify deploy --prod --dir=docs
if errorlevel 1 goto :fail

echo.
echo  완료. 폰에서 새로고침하면 반영됩니다.
goto :end

:fail
echo.
echo  실패했습니다. 위 메시지를 확인하세요.

:end
echo.
pause
