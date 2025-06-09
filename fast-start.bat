@echo off
setlocal
cd /d "%~dp0"
call .\mvnw clean package
call .\mvnw javafx:run
endlocal