@echo off
echo.
echo. Manually unregistering custom ArcGIS HLU Tool extension v4 in ArcMap Desktop ...
echo.
setlocal
set PATH=%PATH%;C:\Program Files (x86)\Common Files\ArcGIS\bin;C:\Program Files\Common Files\ArcGIS\bin

ESRIRegAsm.exe /p:Desktop "C:\Program Files (x86)\HLU\HLU Tool v4\HluArcMapExtensionv4.dll" /u

endlocal

echo. Unregistration complete.
echo.