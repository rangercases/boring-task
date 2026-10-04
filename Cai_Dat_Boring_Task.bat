@echo off
chcp 65001 > nul
title Cài đặt Boring Task

echo ========================================================
echo          CÀI ĐẶT PHẦN MỀM BORING TASK
echo ========================================================
echo.

set "SCRIPT_DIR=%~dp0"
set "TARGET_PYW=%SCRIPT_DIR%boring_task.pyw"
set "LOCAL_ICON=%SCRIPT_DIR%assets\app_icon.ico"

:: 1. Kiểm tra Python thực sự (tránh lỗi shortcut Microsoft Store)
python -c "import sys" >nul 2>nul
if %errorlevel% equ 0 (
    set "PY_CMD=python"
) else (
    py -3 -c "import sys" >nul 2>nul
    if %errorlevel% equ 0 (
        set "PY_CMD=py -3"
    ) else (
        echo [LỖI] Máy tính chưa cài đặt Python!
        echo.
        echo Vui lòng cài đặt Python (tích chọn ô 'Add Python to PATH'):
        echo Link tải: https://www.python.org/downloads/
        echo.
        pause
        exit /b 1
    )
)

:: 2. Cài đặt thư viện với cờ --user (chống lỗi thiếu quyền Admin)
echo [1/3] Đang kiểm tra và cài đặt thư viện cần thiết...
%PY_CMD% -m pip install --user -r "%SCRIPT_DIR%requirements.txt" --disable-pip-version-check --quiet

:: Xác thực lại các thư viện cốt lõi
%PY_CMD% -c "import openpyxl, customtkinter, PIL" >nul 2>nul
if %errorlevel% neq 0 (
    echo.
    echo [CẢNH BÁO] Chưa thể tải đầy đủ thư viện (có thể do mất mạng hoặc bị chặn).
    echo Đang thử lại với thông báo chi tiết...
    %PY_CMD% -m pip install --user -r "%SCRIPT_DIR%requirements.txt" --disable-pip-version-check
    %PY_CMD% -c "import openpyxl, customtkinter, PIL" >nul 2>nul
    if %errorlevel% neq 0 (
        echo.
        echo [LỖI] Cài đặt thư viện không thành công. Vui lòng kiểm tra lại kết nối mạng Internet.
        echo.
        pause
        exit /b 1
    )
)

:: 3. Sao lưu Logo vào thư mục an toàn của hệ thống (chống mất icon khi di chuyển thư mục)
echo.
echo [2/3] Đang cấu hình Logo và tạo lối tắt Desktop...
set "SYS_ICON_DIR=%LOCALAPPDATA%\BoringTask\assets"
if not exist "%SYS_ICON_DIR%" mkdir "%SYS_ICON_DIR%" >nul 2>nul
if exist "%LOCAL_ICON%" (
    copy /y "%LOCAL_ICON%" "%SYS_ICON_DIR%\app_icon.ico" >nul 2>nul
    set "FINAL_ICON=%SYS_ICON_DIR%\app_icon.ico"
) else (
    set "FINAL_ICON=%LOCAL_ICON%"
)

:: 4. Tạo Shortcut ngoài Desktop với đường dẫn pythonw.exe tuyệt đối và Icon hệ thống
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$WshShell = New-Object -ComObject WScript.Shell; " ^
  "$DesktopPath = [System.Environment]::GetFolderPath('Desktop'); " ^
  "$Shortcut = $WshShell.CreateShortcut((Join-Path $DesktopPath 'Boring Task.lnk')); " ^
  "$pyw = (Get-Command -Name 'pythonw.exe' -ErrorAction SilentlyContinue).Source; " ^
  "if (-not $pyw) { $p = (Get-Command -Name 'python.exe' -ErrorAction SilentlyContinue).Source; if ($p) { $c = Join-Path (Split-Path $p) 'pythonw.exe'; if (Test-Path $c) { $pyw = $c } } }; " ^
  "if (-not $pyw) { $pyw = 'pythonw.exe' }; " ^
  "$Shortcut.TargetPath = $pyw; " ^
  "$Shortcut.Arguments = '\"%TARGET_PYW%\"'; " ^
  "$Shortcut.WorkingDirectory = '%SCRIPT_DIR%'; " ^
  "$Shortcut.IconLocation = '%FINAL_ICON%,0'; " ^
  "$Shortcut.Description = 'Boring Task - Hệ thống tự động hóa tác vụ'; " ^
  "$Shortcut.Save(); " ^
  "Write-Host '  ✓ Đã tạo thành công biểu tượng Boring Task trên Desktop.'"

:: 5. Khởi động phần mềm ngầm
echo.
echo [3/3] Đang khởi động phần mềm...
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$pyw = (Get-Command -Name 'pythonw.exe' -ErrorAction SilentlyContinue).Source; " ^
  "if (-not $pyw) { $p = (Get-Command -Name 'python.exe' -ErrorAction SilentlyContinue).Source; if ($p) { $c = Join-Path (Split-Path $p) 'pythonw.exe'; if (Test-Path $c) { $pyw = $c } } }; " ^
  "if (-not $pyw) { $pyw = 'pythonw.exe' }; " ^
  "Start-Process -FilePath $pyw -ArgumentList '\"%TARGET_PYW%\"' -WorkingDirectory '%SCRIPT_DIR%'"

echo.
echo ========================================================
echo  HOÀN TẤT! Từ bây giờ bạn chỉ cần bấm vào biểu tượng
echo  \"Boring Task\" ngoài màn hình Desktop để mở ứng dụng.
echo ========================================================
echo.
timeout /t 3 > nul
exit
