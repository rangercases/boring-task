@echo off
chcp 65001 > nul
title Cài đặt Boring Task

echo ========================================================
echo          CÀI ĐẶT PHẦN MỀM BORING TASK
echo ========================================================
echo.

set "SCRIPT_DIR=%~dp0"
set "INSTALL_DIR=%LOCALAPPDATA%\BoringTask"
set "GITHUB_RAW=https://raw.githubusercontent.com/rangercases/claim-helper/main"

:: 1. Kiểm tra môi trường Python thực sự (tránh lỗi shortcut Microsoft Store)
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

:: 2. Khởi tạo thư mục cài đặt chuẩn Windows
if not exist "%INSTALL_DIR%" mkdir "%INSTALL_DIR%" >nul 2>nul
if not exist "%INSTALL_DIR%\assets" mkdir "%INSTALL_DIR%\assets" >nul 2>nul

:: 3. Đồng bộ mã nguồn và assets vào thư mục cài đặt
if exist "%SCRIPT_DIR%boring_task.pyw" (
    echo [1/4] Đang sao chép mã nguồn vào thư mục ứng dụng...
    copy /y "%SCRIPT_DIR%boring_task.pyw" "%INSTALL_DIR%\boring_task.pyw" >nul 2>nul
    copy /y "%SCRIPT_DIR%requirements.txt" "%INSTALL_DIR%\requirements.txt" >nul 2>nul
    if exist "%SCRIPT_DIR%assets" (
        xcopy /y /e "%SCRIPT_DIR%assets" "%INSTALL_DIR%\assets\" >nul 2>nul
    )
) else (
    echo [1/4] Đang tải mã nguồn mới nhất từ GitHub...
    powershell -NoProfile -ExecutionPolicy Bypass -Command ^
      "$installDir = [System.Environment]::GetFolderPath('LocalApplicationData') + '\BoringTask'; " ^
      "$baseUrl = '%GITHUB_RAW%'; " ^
      "$wc = New-Object System.Net.WebClient; " ^
      "$wc.Headers.Add('User-Agent', 'BoringTask-Installer'); " ^
      "$files = @('boring_task.pyw', 'requirements.txt', 'assets/app_icon.ico', 'assets/boring_task_banner.png', 'assets/logo.png'); " ^
      "foreach ($f in $files) { " ^
      "  $target = Join-Path $installDir $f; " ^
      "  $url = $baseUrl + '/' + $f; " ^
      "  try { $wc.DownloadFile($url, $target) } catch { Write-Host '  Lỗi tải:' $f } " ^
      "}; " ^
      "Write-Host '  ✓ Đã tải xong mã nguồn và tài nguyên.'"
)

:: 4. Cấu hình phân hệ làm việc (chỉ hỏi 1 lần duy nhất lúc cài đặt)
if not exist "%INSTALL_DIR%\config.json" (
    echo.
    echo ========================================================
    echo  CHỌN PHÂN HỆ LÀM VIỆC:
    echo  [1] Purchase Cost Auto-Filled (Chị A)
    echo  [2] Image Inserter - Place in Cell (Chị B)
    echo  [3] Bản Đầy Đủ (Admin)
    echo ========================================================
    set "USER_CHOICE=1"
    set /p "USER_CHOICE=Nhập số (1, 2 hoặc 3) [Mặc định: 1]: "
    if "%USER_CHOICE%"=="2" (
        echo {"mode": "images"} > "%INSTALL_DIR%\config.json"
        echo   - Đã thiết lập phân hệ: Image Inserter.
    ) else if "%USER_CHOICE%"=="3" (
        echo {"mode": "all"} > "%INSTALL_DIR%\config.json"
        echo   - Đã thiết lập phân hệ: Bản Đầy Đủ.
    ) else (
        echo {"mode": "cost"} > "%INSTALL_DIR%\config.json"
        echo   - Đã thiết lập phân hệ: Purchase Cost Auto-Filled.
    )
    echo.
)

:: 5. Cài đặt thư viện phụ thuộc an toàn (chống lỗi thiếu quyền Admin)
echo [2/4] Đang cài đặt thư viện cần thiết...
%PY_CMD% -m pip install --user -r "%INSTALL_DIR%\requirements.txt" --disable-pip-version-check --quiet

%PY_CMD% -c "import openpyxl, customtkinter, PIL" >nul 2>nul
if %errorlevel% neq 0 (
    echo.
    echo [CẢNH BÁO] Chưa thể tải đầy đủ thư viện (có thể do mất mạng).
    echo Đang thử lại với thông báo chi tiết...
    %PY_CMD% -m pip install --user -r "%INSTALL_DIR%\requirements.txt" --disable-pip-version-check
    %PY_CMD% -c "import openpyxl, customtkinter, PIL" >nul 2>nul
    if %errorlevel% neq 0 (
        echo.
        echo [LỖI] Cài đặt thư viện không thành công. Vui lòng kiểm tra lại kết nối Internet.
        echo.
        pause
        exit /b 1
    )
)

:: 6. Tạo biểu tượng ngoài Desktop trỏ vào thư mục an toàn
echo.
echo [3/4] Đang tạo biểu tượng Boring Task trên Desktop...
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$WshShell = New-Object -ComObject WScript.Shell; " ^
  "$DesktopPath = [System.Environment]::GetFolderPath('Desktop'); " ^
  "$Shortcut = $WshShell.CreateShortcut((Join-Path $DesktopPath 'Boring Task.lnk')); " ^
  "$installDir = [System.Environment]::GetFolderPath('LocalApplicationData') + '\BoringTask'; " ^
  "$targetPyw = Join-Path $installDir 'boring_task.pyw'; " ^
  "$iconPath = Join-Path $installDir 'assets\app_icon.ico'; " ^
  "$pyw = (Get-Command -Name 'pythonw.exe' -ErrorAction SilentlyContinue).Source; " ^
  "if (-not $pyw) { $p = (Get-Command -Name 'python.exe' -ErrorAction SilentlyContinue).Source; if ($p) { $c = Join-Path (Split-Path $p) 'pythonw.exe'; if (Test-Path $c) { $pyw = $c } } }; " ^
  "if (-not $pyw) { $pyw = 'pythonw.exe' }; " ^
  "$Shortcut.TargetPath = $pyw; " ^
  "$Shortcut.Arguments = '\"' + $targetPyw + '\"'; " ^
  "$Shortcut.WorkingDirectory = $installDir; " ^
  "$Shortcut.IconLocation = $iconPath + ',0'; " ^
  "$Shortcut.Description = 'Boring Task - Hệ thống tự động hóa tác vụ'; " ^
  "$Shortcut.Save(); " ^
  "Write-Host '  ✓ Đã tạo thành công biểu tượng Boring Task trên Desktop.'"

:: 7. Khởi động phần mềm ngầm
echo.
echo [4/4] Đang khởi động phần mềm...
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$installDir = [System.Environment]::GetFolderPath('LocalApplicationData') + '\BoringTask'; " ^
  "$targetPyw = Join-Path $installDir 'boring_task.pyw'; " ^
  "$pyw = (Get-Command -Name 'pythonw.exe' -ErrorAction SilentlyContinue).Source; " ^
  "if (-not $pyw) { $p = (Get-Command -Name 'python.exe' -ErrorAction SilentlyContinue).Source; if ($p) { $c = Join-Path (Split-Path $p) 'pythonw.exe'; if (Test-Path $c) { $pyw = $c } } }; " ^
  "if (-not $pyw) { $pyw = 'pythonw.exe' }; " ^
  "Start-Process -FilePath $pyw -ArgumentList ('\"' + $targetPyw + '\"') -WorkingDirectory $installDir"

echo.
echo ========================================================
echo  HOÀN TẤT CÀI ĐẶT THÀNH CÔNG!
echo.
echo  • Phần mềm đã được cài vào:
echo    %%LOCALAPPDATA%%\BoringTask\
echo  • Biểu tượng \"Boring Task\" đã xuất hiện trên Desktop.
echo  • Bạn có thể xóa file Cai_Dat_Boring_Task.bat này.
echo ========================================================
echo.
timeout /t 5 > nul
exit
