@echo off
setlocal
cd /d "%~dp0"
echo ===================================================
echo   🚀 중앙지사 바탕화면 바로가기 자동 설치
echo ===================================================

copy /y "중앙지사_실행.bat" "%USERPROFILE%\Desktop\중앙지사_자재관리.bat"
copy /y "중앙지사_모바일서버.bat" "%USERPROFILE%\Desktop\중앙지사_모바일서버.bat"
copy /y "중앙지사_구글동기화.bat" "%USERPROFILE%\Desktop\중앙지사_구글동기화.bat"
copy /y "중앙지사_배포.bat" "%USERPROFILE%\Desktop\중앙지사_배포.bat"

echo.
echo [완료] 바탕화면에 중앙지사 바로가기 4종이 성공적으로 생성되었습니다!
echo   1. 중앙지사_자재관리.bat
echo   2. 중앙지사_모바일서버.bat
echo   3. 중앙지사_구글동기화.bat
echo   4. 중앙지사_배포.bat
echo.
pause
endlocal

