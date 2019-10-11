@echo off
echo.
echo. Manually registering custom ArcGIS HLU Tool extension in ArcMap Desktop ...
echo.
setlocal
set PATH=%PATH%;C:\Program Files (x86)\Common Files\ArcGIS\bin;C:\Program Files\Common Files\ArcGIS\bin

ESRIRegAsm.exe /p:Desktop "C:\Program Files (x86)\HLU\HLU GIS Tool\HluArcMapExtension.dll"

endlocal

echo. Registration complete.
echo.