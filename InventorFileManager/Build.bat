@echo off
echo Building Inventor File Manager Plugin...

REM Check if .NET Framework SDK or Visual Studio is available
where msbuild >nul 2>nul
if %ERRORLEVEL% NEQ 0 (
    echo Error: MSBuild not found in PATH
    echo Please install Visual Studio or .NET Framework SDK
    echo Or run this from Visual Studio Developer Command Prompt
    pause
    exit /b 1
)

REM Clean previous build
if exist "bin" rmdir /s /q "bin"
if exist "obj" rmdir /s /q "obj"

REM Build the project
echo Building project...
msbuild InventorFileManager.csproj /p:Configuration=Release /p:Platform="Any CPU" /verbosity:minimal

if %ERRORLEVEL% NEQ 0 (
    echo Build failed!
    pause
    exit /b 1
)

REM Create distribution folder
if not exist "Distribution" mkdir "Distribution"

REM Copy built files
echo Copying files for distribution...
copy "bin\Release\InventorFileManager.dll" "Distribution\" >nul
copy "InventorFileManager.addin" "Distribution\" >nul

REM Copy dependencies
if exist "bin\Release\EPPlus.dll" (
    copy "bin\Release\EPPlus.dll" "Distribution\" >nul
)

REM Copy installer
copy "Setup\Install.bat" "Distribution\" >nul

REM Create README for distribution
(
echo Inventor File Manager Plugin
echo ===========================
echo.
echo Installation Instructions:
echo 1. Run Install.bat as Administrator
echo 2. Restart Autodesk Inventor
echo 3. Look for "File Name Export" and "File Rename" buttons in the Add-Ins tab
echo.
echo Features:
echo - Export file names ^(.iam, .ipt, .idw^) from folders to Excel
echo - Rename files based on Excel mapping with old/new file names
echo - Supports subfolders and maintains file relationships
echo.
echo Requirements:
echo - Autodesk Inventor 2022 or newer
echo - Windows 10 or newer
echo - .NET Framework 4.8
echo.
echo For support, contact your system administrator.
) > "Distribution\README.txt"

echo.
echo Build completed successfully!
echo Distribution files are in the 'Distribution' folder.
echo.
echo To install the plugin:
echo 1. Copy the Distribution folder to the target machine
echo 2. Run Install.bat as Administrator
echo 3. Restart Inventor
echo.
pause
