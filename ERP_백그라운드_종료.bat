@echo off
echo ERP 서버(Streamlit)를 종료합니다...
taskkill /F /IM python.exe /FI "WINDOWTITLE eq streamlit*" >nul 2>&1
taskkill /F /IM streamlit.exe >nul 2>&1
echo 종료 완료!
pause
