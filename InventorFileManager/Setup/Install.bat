@echo off
echo Installing Inventor File Manager Plugin...

REM Get the current directory
set PLUGIN_DIR=%~dp0

REM Check if Inventor is installed
if not exist "C:\Program Files\Autodesk\Inventor 2022" (
    if not exist "C:\Program Files\Autodesk\Inventor 2023" (
        if not exist "C:\Program Files\Autodesk\Inventor 2024" (
            echo Error: Inventor 2022 or newer not found!
            echo Please install Autodesk Inventor 2022 or newer before installing this plugin.
            pause
            exit /b 1
        )
    )
)

REM Create plugin directory in ProgramData
set INSTALL_DIR=%ProgramData%\Autodesk\Inventor Addins\InventorFileManager

if not exist "%INSTALL_DIR%" (
    mkdir "%INSTALL_DIR%"
)

REM Copy plugin files
echo Copying plugin files...
copy "%PLUGIN_DIR%..\InventorFileManager.dll" "%INSTALL_DIR%\" >nul
copy "%PLUGIN_DIR%..\InventorFileManager.addin" "%INSTALL_DIR%\" >nul

REM Copy dependencies
if exist "%PLUGIN_DIR%..\EPPlus.dll" (
    copy "%PLUGIN_DIR%..\EPPlus.dll" "%INSTALL_DIR%\" >nul
)

REM Register the plugin by copying the .addin file to the Inventor addins folder
set INVENTOR_ADDINS=%ProgramData%\Autodesk\Inventor Addins

if not exist "%INVENTOR_ADDINS%" (
    mkdir "%INVENTOR_ADDINS%"
)

copy "%INSTALL_DIR%\InventorFileManager.addin" "%INVENTOR_ADDINS%\" >nul

REM Create uninstaller
echo Creating uninstaller...
(
echo @echo off
echo echo Uninstalling Inventor File Manager Plugin...
echo.
echo REM Remove plugin files
echo if exist "%INSTALL_DIR%" ^(
echo     rmdir /s /q "%INSTALL_DIR%"
echo ^)
echo.
echo REM Remove addin registration
echo if exist "%INVENTOR_ADDINS%\InventorFileManager.addin" ^(
echo     del "%INVENTOR_ADDINS%\InventorFileManager.addin"
echo ^)
echo.
echo echo Plugin uninstalled successfully!
echo echo Please restart Inventor to complete the uninstallation.
echo pause
) > "%INSTALL_DIR%\Uninstall.bat"

REM Create desktop shortcut for uninstaller
echo Creating uninstaller shortcut...
powershell -Command "$WshShell = New-Object -comObject WScript.Shell; $Shortcut = $WshShell.CreateShortcut('%USERPROFILE%\Desktop\Uninstall Inventor File Manager.lnk'); $Shortcut.TargetPath = '%INSTALL_DIR%\Uninstall.bat'; $Shortcut.WorkingDirectory = '%INSTALL_DIR%'; $Shortcut.IconLocation = 'shell32.dll,31'; $Shortcut.Description = 'Uninstall Inventor File Manager Plugin'; $Shortcut.Save()"

echo.
echo Installation completed successfully!
echo.
echo The Inventor File Manager plugin has been installed.
echo Please restart Inventor to load the plugin.
echo.
echo You will find two new buttons in the Add-Ins tab:
echo - File Name Export: Export file names to Excel
echo - File Rename: Rename files based on Excel mapping
echo.
echo To uninstall, use the shortcut created on your desktop or run:
echo "%INSTALL_DIR%\Uninstall.bat"
echo.
pause
