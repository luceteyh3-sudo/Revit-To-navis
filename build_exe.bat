@echo off
chcp 65001 > nul
echo ============================================
echo  Revit to Navisworks Converter - EXE 빌드
echo ============================================
echo.

:: PyInstaller 설치 확인
pip show pyinstaller > nul 2>&1
if %errorlevel% neq 0 (
    echo [INFO] PyInstaller 설치 중...
    pip install pyinstaller
    if %errorlevel% neq 0 (
        echo [ERROR] PyInstaller 설치 실패
        pause
        exit /b 1
    )
)

echo [INFO] EXE 빌드 시작...
echo.

pyinstaller ^
    --onefile ^
    --windowed ^
    --name "RevitToNavisConverter" ^
    main.py

if %errorlevel% neq 0 (
    echo.
    echo [ERROR] 빌드 실패! 위 오류 메시지를 확인해주세요.
    pause
    exit /b 1
)

echo.
echo ============================================
echo  빌드 완료!
echo  실행 파일 위치: dist\RevitToNavisConverter.exe
echo ============================================
pause
