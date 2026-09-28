@echo off
chcp 65001 > nul
:: 관리자 권한 확인
net session >nul 2>&1
if %errorLevel% == 0 (
    echo 방화벽 규칙을 추가합니다 (포트 8501 허용)...
    powershell -Command "New-NetFirewallRule -DisplayName 'Streamlit ERP' -Direction Inbound -LocalPort 8501 -Protocol TCP -Action Allow -ErrorAction SilentlyContinue"
    echo.
    echo 방화벽 설정이 완료되었습니다! 이제 스마트폰 등에서 접속 가능합니다.
    pause
) else (
    echo 관리자 권한이 필요하여 권한 요청 창을 띄웁니다...
    powershell -Command "Start-Process '%~dpnx0' -Verb RunAs"
)
