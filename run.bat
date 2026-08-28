@echo off
setlocal
cd /d "%~dp0"

dotnet run --project "src\eneBridge.Wpf\eneBridge.Wpf.csproj"

if errorlevel 1 (
    echo.
    echo Build or run failed - see errors above.
    pause
)

endlocal
