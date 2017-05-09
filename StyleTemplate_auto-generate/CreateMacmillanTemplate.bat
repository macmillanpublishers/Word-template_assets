@echo off
SET mypath=%~dp0
echo.

tasklist /FI "IMAGENAME eq EXCEL.exe" 2>NUL | find /I /N "EXCEL.exe">NUL
if "%ERRORLEVEL%"=="0" (
echo "Excel is running - please quit Excel and launch this again!"
timeout /t 1 /nobreak > NUL
echo. & echo Press any key to exit this window
pause >nul
EXIT /B
)
tasklist /FI "IMAGENAME eq WINWORD.exe" 2>NUL | find /I /N "WINWORD.exe">NUL
if "%ERRORLEVEL%"=="0" (
echo "Word is running - please quit Word and launch this again!"
timeout /t 1 /nobreak > NUL
echo. & echo Press any key to exit this window
pause >nul
EXIT /B
)

echo "Exporting styles to Styles.json ... Should be 10-20 seconds"
PowerShell -NoProfile -ExecutionPolicy Bypass -Command "%mypath%runMacro_Wd-or-Xls_NoSave.ps1 excel WordTemplateStyles.xlsm autorun_StylesToJSON"
echo "Done exporting to Styles.json!" & echo.

echo "Exporting bookmaker HTML-mappings to style_config.json ... Should be <10 seconds"
PowerShell -NoProfile -ExecutionPolicy Bypass -Command "%mypath%runMacro_Wd-or-Xls_NoSave.ps1 excel WordTemplateStyles.xlsm autorun_HTMLmappingsToJSON"
echo "Done exporting to style_config.json!" & echo.

echo "Exporting VBA style-mappings to vba_style_config.json ... Should be <10 seconds"
PowerShell -NoProfile -ExecutionPolicy Bypass -Command "%mypath%runMacro_Wd-or-Xls_NoSave.ps1 excel WordTemplateStyles.xlsm autorun_VBAmappingsToJSON"
echo "Done exporting to vba_style_config.json!" & echo.

echo "(Re)creating macmillan.dotm ... this should take 60-90 seconds"
timeout /t 1 /nobreak > NUL & echo.
echo "Keep your eye on this window for the all clear!"
PowerShell -NoProfile -ExecutionPolicy Bypass -Command "%mypath%runMacro_Wd-or-Xls_NoSave.ps1 word StyleTemplateCreator.docm WriteTemplatefromJson"
echo "OK, finished writing macmillan.dotm!" & echo.

echo. & echo "(Re)creating macmillan_NoColor.dotm ... this should take  60-90 seconds again."
timeout /t 1 /nobreak > NUL
echo. & echo "Keep your eye on this window for the all clear (again)!"
PowerShell -NoProfile -ExecutionPolicy Bypass -Command "%mypath%runMacro_Wd-or-Xls_NoSave.ps1 word StyleTemplateCreator.docm WriteNoColorTemplatefromJson"
echo "Yay, finished writing macmillan_NoColor.dotm!"

echo. & echo.
echo "All done!  This is so awesome!" & echo.
echo (Press any key to exit this window)
pause >nul
