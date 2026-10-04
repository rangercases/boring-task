@echo off
chcp 65001 > nul
title Cài đặt Boring Task

echo ========================================================
echo          CÀI ĐẶT PHẦN MỀM BORING TASK
echo ========================================================
echo.

set "SCRIPT_DIR=%~dp0"
set "TARGET_PYW=%SCRIPT_DIR%boring_task.pyw"
set "ICON_PATH=%SCRIPT_DIR%assets\app_icon.ico"

where python >nul 2>nul
if %errorlevel% neq 0 (
    echo [LỖI] Máy tính chưa cài đặt Python!
    echo Vui lòng cài đặt Python (tích chọn 'Add Python to PATH') trước khi chạy.
    echo.
    pause
    exit /b 1
)

echo [1/3] Đang kiểm tra và cài đặt thư viện cần thiết...
python -m pip install -r "%SCRIPT_DIR%requirements.txt" --quiet

echo.
echo [2/3] Đang tạo biểu tượng Boring Task ngoài Desktop...

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$WshShell = New-Object -ComObject WScript.Shell; " ^
  "$DesktopPath = [System.Environment]::GetFolderPath('Desktop'); " ^
  "$Shortcut = $WshShell.CreateShortcut((Join-Path $DesktopPath 'Boring Task.lnk')); " ^
  "$Shortcut.TargetPath = 'pythonw.exe'; " ^
  "$Shortcut.Arguments = '\"%TARGET_PYW%\"'; " ^
  "$Shortcut.WorkingDirectory = '%SCRIPT_DIR%'; " ^
  "$Shortcut.IconLocation = '%ICON_PATH%,0'; " ^
  "$Shortcut.Description = 'Boring Task - Hệ thống tự động hóa tác vụ'; " ^
  "$Shortcut.Save(); " ^
  "Write-Host '  ✓ Đã tạo thành công biểu tượng Boring Task trên Desktop.'"

echo.
echo [3/3] Đang khởi động phần mềm...
start "" pythonw.exe "%TARGET_PYW%"

echo.
echo ========================================================
echo  HOÀN TẤT! Từ bây giờ bạn chỉ cần bấm vào biểu tượng
echo  \"Boring Task\" ngoài màn hình Desktop để mở ứng dụng.
echo ========================================================
echo.
timeout /t 3 > nul
exit
