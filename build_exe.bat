@echo off
chcp 65001 > nul
echo ========================================================
echo     BUILD CLAIM HELPER - APPLE UX/UI EXECUTABLE (.EXE)
echo ========================================================
echo.

echo [1/3] Đang kiểm tra thư viện...
python -m pip install -r requirements.txt
echo.

echo [2/3] Đang đóng gói ClaimHelper.exe (Apple Design)...
pyinstaller --noconfirm --onedir --windowed --name "ClaimHelper" ^
  --collect-all customtkinter ^
  --collect-all tkinterdnd2 ^
  --clean claim_helper.py

echo.
echo ========================================================
if exist "dist\ClaimHelper\ClaimHelper.exe" (
    echo [3/3] HOÀN TẤT THÀNH CÔNG!
    echo File phần mềm đã tạo tại: dist\ClaimHelper\ClaimHelper.exe
    echo Bạn có thể nén thư mục 'dist\ClaimHelper' lại và gửi cho đồng nghiệp sử dụng ngay.
) else (
    echo [LỖI] Đóng gói thất bại. Vui lòng kiểm tra lại log bên trên.
)
echo ========================================================
echo.
pause
