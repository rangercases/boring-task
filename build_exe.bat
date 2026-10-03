@echo off
chcp 65001 > nul
echo ========================================================
echo     BUILD CLAIM HELPER - 1 FILE EXECUTABLE (.EXE)
echo ========================================================
echo.

echo [1/3] Đang kiểm tra thư viện...
python -m pip install -r requirements.txt
echo.

echo [2/3] Đang đóng gói 1 file ClaimHelper.exe duy nhất (Apple Design)...
python -m PyInstaller --noconfirm --onefile --windowed --name "ClaimHelper" ^
  --collect-all customtkinter ^
  --collect-all tkinterdnd2 ^
  --clean claim_helper.pyw

echo.
echo ========================================================
if exist "dist\ClaimHelper.exe" (
    copy /y "dist\ClaimHelper.exe" "ClaimHelper.exe" > nul
    echo [3/3] HOÀN TẤT THÀNH CÔNG!
    echo File phần mềm duy nhất đã tạo tại: ClaimHelper.exe
    echo Bạn chỉ cần click đúp vào ClaimHelper.exe để mở ứng dụng ngay lập tức (không có CMD).
) else (
    echo [LỖI] Đóng gói thất bại. Vui lòng kiểm tra lại log bên trên.
)
echo ========================================================
echo.
pause
