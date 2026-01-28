@echo off
REM Nuitka build script for TaskManager - Fast startup!
echo ========================================
echo Building TaskManager with Nuitka
echo This may take 5-10 minutes on first build
echo ========================================
echo.

REM Clean previous builds
echo Cleaning previous builds...
if exist "TaskManager.dist" rmdir /s /q "TaskManager.dist"
if exist "TaskManager.build" rmdir /s /q "TaskManager.build"
if exist "TaskManager.exe" del "TaskManager.exe"
echo.

REM Build with Nuitka
echo Building executable (this will take a while)...
python -m nuitka ^
    --standalone ^
    --onefile ^
    --windows-console-mode=disable ^
    --enable-plugin=tk-inter ^
    --include-data-file=TaskManager_main.png=TaskManager_main.png ^
    --company-name="Your Company" ^
    --product-name="Task Manager" ^
    --file-version=0.23.0 ^
    --product-version=0.23.0 ^
    --file-description="Task Management Application" ^
    --output-filename=TaskManager.exe ^
    TaskManager_0.17.py

echo.
if exist "TaskManager.exe" (
    echo ========================================
    echo Build successful!
    echo Executable: TaskManager.exe
    echo ========================================
    echo.
    echo Startup time: 2-5 seconds (vs PyInstaller's 30-120 sec)
    echo File size: Usually smaller than PyInstaller
    echo.
    echo You can now test or share TaskManager.exe
    pause
) else (
    echo ========================================
    echo Build failed! Check errors above.
    echo ========================================
    pause
)

