@echo off
setlocal EnableDelayedExpansion

echo.
echo ================================================
echo   BUILDING BIOMETRIC ATTENDANCE PARSER
echo ================================================
echo.

rem --- 1. Maven: Build Fat JAR ---
echo [1/4] Building fat JAR with Maven...
call mvn clean package -DskipTests
if not %errorlevel% == 0 (
    echo *** ERROR: Maven failed! ***
    pause
    exit /b 1
)

rem --- 2. jlink: Create Custom JRE ---
echo.
echo [2/4] Creating slim JRE with jlink...
rmdir /s /q custom-jre 2>nul

if not defined JAVA_HOME (
    echo *** ERROR: JAVA_HOME not set! ***
    pause
    exit /b 1
)

if not exist "%JAVA_HOME%\jmods" (
    echo *** ERROR: jmods not found! JAVA_HOME = %JAVA_HOME% ***
    pause
    exit /b 1
)

echo Running jlink command...
jlink ^
    --module-path "%JAVA_HOME%\jmods" ^
    --add-modules java.base,java.desktop,java.logging,java.xml,java.naming,java.sql,java.management,java.instrument,java.security.jgss,java.security.sasl,java.scripting,java.compiler,java.net.http,java.prefs,java.datatransfer ^
    --strip-debug --no-header-files --no-man-pages --compress=2 ^
    --output custom-jre

if not %errorlevel% == 0 (
    echo.
    echo *** ERROR: jlink failed with error code %errorlevel%! ***
    echo.
    echo Common jlink issues:
    echo - Module names might be incorrect
    echo - Java version compatibility issues
    echo - Missing modules
    echo.
    echo Let's try a simpler module set...
    
    rem Try with minimal modules
    jlink --module-path "%JAVA_HOME%\jmods" --add-modules java.base,java.desktop,java.logging,java.xml --strip-debug --no-header-files --no-man-pages --output custom-jre
    
    if not %errorlevel% == 0 (
        echo.
        echo *** ERROR: jlink failed even with minimal modules! ***
        pause
        exit /b 1
    )
)

echo jlink completed successfully!

rem --- 3. jpackage: Create portable app (no WiX required) ---
echo.
echo [3/4] Creating portable application...
rmdir /s /q output 2>nul

rem Check if icon file exists
set ICON_OPTION=
if exist "src\main\resources\app.ico" (
    set ICON_OPTION=--icon src\main\resources\app.ico
    echo Using icon: src\main\resources\app.ico
) else (
    echo Warning: Icon file not found, proceeding without icon
)

echo Running jpackage command...
jpackage --type app-image --input target --name "Biometric Attendance Parser" --main-jar Biometric-Attendance-Parser-1.0-SNAPSHOT.jar --main-class org.bioparse.cleaning.AttendanceAppGUI --runtime-image custom-jre %ICON_OPTION% --dest output --win-console --java-options "--add-opens=java.desktop/javax.swing=ALL-UNNAMED"

if not %errorlevel% == 0 (
    echo.
    echo *** ERROR: jpackage failed! ***
    pause
    exit /b 1
)

rem --- 4. SUCCESS ---
echo.
echo [4/4] SUCCESS!
echo.
echo ================================================
echo   PORTABLE APPLICATION CREATED!
echo ================================================
echo.
echo Application location: output\"Biometric Attendance Parser"
echo.
echo To run the application:
echo   output\"Biometric Attendance Parser\Biometric Attendance Parser.exe"
echo.
echo You can copy the entire folder to any Windows PC!
echo.
pause