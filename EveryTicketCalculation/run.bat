@echo off
echo Building and running Ticket Analyzer...
echo.

echo Step 1: Compiling the project...
mvn compile
if errorlevel 1 (
    echo Failed to compile. Please check for errors above.
    pause
    exit /b 1
)

echo.
echo Step 2: Running the analysis...
mvn exec:java
if errorlevel 1 (
    echo Failed to run. Please check for errors above.
    pause
    exit /b 1
)

echo.
echo Analysis completed successfully!
pause