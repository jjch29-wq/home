@echo off
cd /d "c:\Users\jjch2\Desktop\PMI"
"c:\Users\jjch2\Desktop\PMI\.venv\Scripts\python.exe" "c:\Users\jjch2\Desktop\PMI\home\src\비파괴검사보고서.py" > "c:\Users\jjch2\Desktop\PMI\crash.log" 2>&1
echo 종료코드: %ERRORLEVEL% >> "c:\Users\jjch2\Desktop\PMI\crash.log"
