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
echo  [2/2] GitHub 푸시
echo  --------------------------------
git rev-parse --is-inside-work-tree >nul 2>nul
if errorlevel 1 (echo  이 폴더는 아직 git 저장소가 아닙니다. 최초 1회 설정이 필요합니다. & goto :fail)

git add -A
if errorlevel 1 goto :fail

git diff --cached --quiet
if not errorlevel 1 (echo  바뀐 내용이 없습니다. & goto :end)

git commit -m "운동프로그램 갱신 %date% %time%"
if errorlevel 1 goto :fail

git push
if errorlevel 1 goto :fail

echo.
echo  완료. 30초쯤 뒤 폰에서 새로고침하면 반영됩니다.
goto :end

:fail
echo.
echo  실패했습니다. 위 메시지를 확인하세요.

:end
echo.
pause
