@echo off
setlocal
cd /d "%~dp0"
call .\mvnw javafx:run
endlocal