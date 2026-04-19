@echo off
chcp 65001 > nul
echo [빌드 시작] HWP_PDF변환기
echo.

pyinstaller hwp_to_pdf_oneclick.spec --clean --noconfirm

echo.
if exist "dist\HWP_PDF변환기.exe" (
    echo [완료] dist\HWP_PDF변환기.exe 생성됨
) else (
    echo [실패] exe 파일이 생성되지 않았습니다.
    exit /b 1
)
