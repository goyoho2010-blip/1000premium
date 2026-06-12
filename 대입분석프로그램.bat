@echo off
chcp 65001 >nul
title 1000Premium 입시 분석 시스템 자동 실행기

echo ====================================================
echo   천명의선택 대입 분석 및 진로교과 시스템을 구동합니다.
echo   최초 실행 시 1~2분 정도 소요될 수 있습니다...
echo ====================================================

cd /d "C:\Users\김태영\Documents\1000Premium"

echo [1/2] 필수 프로그램(엔진)을 확인하고 설치 중입니다...
python -m pip install --upgrade pip >nul 2>&1
python -m pip install streamlit pandas numpy openpyxl >nul 2>&1

echo.
echo [2/2] 시스템을 엽니다! 인터넷 창이 뜰 때까지 기다려주세요...
:: [핵심 해결책] streamlit 명령어가 안 먹히는 윈도우를 위해 파이썬(python) 엔진을 강제로 거쳐서 엽니다.
python -m streamlit run 1000app.py

pause