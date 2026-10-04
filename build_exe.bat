@echo off
chcp 65001 > nul
echo ========================================================
echo     BUILD BORING TASK - 1 FILE EXECUTABLE (.EXE)
echo ========================================================
echo.

echo [1/3] Đang kiểm tra thư viện...
python -m pip install -r requirements.txt pyinstaller>=6.0.0
echo.

echo [2/3] Đang đóng gói 1 file BoringTask.exe duy nhất...
python -m PyInstaller --noconfirm --onefile --windowed --name "BoringTask" ^
  --icon "assets\app_icon.ico" ^
  --add-data "assets;assets" ^
  --collect-all customtkinter ^
  --collect-all tkinterdnd2 ^
  --clean boring_task.pyw

echo.
echo ========================================================
if exist "dist\BoringTask.exe" (
    copy /y "dist\BoringTask.exe" "BoringTask.exe" > nul
    echo [3/3] HOÀN TẤT THÀNH CÔNG!
    echo File phần mềm duy nhất đã tạo tại: BoringTask.exe
    echo Bạn chỉ cần click đúp vào BoringTask.exe để mở ứng dụng ngay lập tức (không có CMD).
) else (
    echo [LỖI] Đóng gói thất bại. Vui lòng kiểm tra lại log bên trên.
)
echo ========================================================
echo.
pause
