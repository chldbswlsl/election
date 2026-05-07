@echo off
chcp 65001 > nul
title 여론조사 데이터 업데이트
cd /d "%~dp0"

echo.
echo ====================================================
echo   여론조사 데이터 업데이트 시작
echo ====================================================
echo.

REM 1) polls.csv 가 변경되었는지 확인
git diff --quiet polls.csv
if %ERRORLEVEL%==0 (
    echo   ⚠ polls.csv 가 변경되지 않았습니다.
    echo     먼저 polls.csv 에 새 여론조사 한 줄을 추가해주세요.
    echo.
    pause
    exit /b 1
)

REM 2) CSV → JSON 변환
echo   [1/4] polls.csv → polls.json 변환 중...
python csv_to_json.py
if %ERRORLEVEL% neq 0 (
    echo.
    echo   ❌ 변환 실패. polls.csv 형식 확인하세요.
    pause
    exit /b 1
)

REM 3) git 커밋
echo   [2/4] git 변경사항 등록...
git add polls.csv polls.json

set /p MSG="  [3/4] 커밋 메시지 (엔터=날짜만): "
if "%MSG%"=="" set MSG=data: 여론조사 추가 (%DATE%)

git commit -m "%MSG%"

REM 4) push
echo   [4/4] 깃허브에 push 중...
git push
if %ERRORLEVEL% neq 0 (
    echo.
    echo   ⚠ push 실패. 인터넷 연결 또는 인증 확인.
    pause
    exit /b 1
)

echo.
echo ====================================================
echo   ✓ 완료! 1~2분 후 사이트에 자동 반영됩니다.
echo   사이트: https://chldbswlsl.github.io/election/
echo ====================================================
echo.
pause
