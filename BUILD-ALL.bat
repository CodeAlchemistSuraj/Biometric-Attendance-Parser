@echo off
echo ========================================
echo 🚀 BUILDING ATTENDANCE APP - FULL AUTOMATION
echo ========================================

REM Ensure we're in project root
cd /d "C:\Users\MOHD GULZAR KHAN\Desktop\Suraj Working\java.git.projects\Biometric-Attendance-Parser"

echo [1/6] Building JAR...
mvn clean package

echo [2/6] Creating dist folder...
if not exist "dist" mkdir dist

echo [3/6] Downloading JRE (50MB)...
powershell -Command ^
"try { ^
    Invoke-WebRequest -Uri 'https://github.com/adoptium/temurin21-binaries/releases/download/jdk-21.0.4%2B7/OpenJDK21U-jre_x64_windows_hotspot_21.0.4_7.zip' -OutFile 'dist\jre.zip' -UseBasicParsing; ^
    Write-Host 'JRE Download SUCCESS!' ^
} catch { ^
    Write-Host 'Download failed, trying alternative...'; ^
    Invoke-WebRequest -Uri 'https://download.java.net/java/GA/jdk21.0.4/2dac2138f8304deea6a0b5c5700d9a7c/13/GPL/openjdk-21.0.4_windows-x64_bin.zip' -OutFile 'dist\jre.zip' -UseBasicParsing ^
}"

echo [4/6] Extracting JRE...
powershell -Command "Expand-Archive -Path 'dist\jre.zip' -DestinationPath 'dist' -Force"
powershell -Command "ren 'dist\jdk-21.0.4' 'jre'"
powershell -Command "rd /s /q 'dist\jre.zip'"

echo [5/6] Creating LICENSE...
echo Attendance Analytics Dashboard v1.0 > LICENSE.txt
echo Copyright (c) 2025 BioParse Solutions > LICENSE.txt

echo [6/6] Building EXE...
mkdir output
copy "target\Biometric-Attendance-Parser-1.0-SNAPSHOT.jar" "output\AttendanceApp.jar"

REM Simple Launch4J (download if needed)
if not exist "launch4j\bin\launch4j.exe" (
    echo Downloading Launch4J...
    powershell -Command "Invoke-WebRequest -Uri 'https://sourceforge.net/projects/launch4j/files/launch4j-3/3.5/launch4j-3.5-win32.zip/download' -OutFile 'launch4j.zip' -UseBasicParsing"
    powershell -Command "Expand-Archive -Path 'launch4j.zip' -DestinationPath '.' -Force"
    rmdir /s /q launch4j.zip
)

REM Launch4J Config
(
echo ^<?xml version="1.0" encoding="UTF-8"?^>
echo ^<launch4jConfig^>
echo   ^<headerType^>gui^</headerType^>
echo   ^<outfile^>output\AttendanceApp.exe^</outfile^>
echo   ^<jar^>output\AttendanceApp.jar^</jar^>
echo   ^<jre^>
echo     ^<path^>jre^</path^>
echo     ^<minVersion^>21.0^</minVersion^>
echo   ^</jre^>
echo ^</launch4jConfig^>
) > launch4j.xml

launch4j\bin\launch4j.exe launch4j.xml

echo ========================================
echo 🎉 SUCCESS! 
echo ========================================
echo 📁 JAR:        target\Biometric-Attendance-Parser-1.0-SNAPSHOT.jar
echo 🖥️  EXE:        output\AttendanceApp.exe  
echo 📦 JRE:        dist\jre\
echo ========================================
echo NEXT: Install Inno Setup, then run installer.iss
pause