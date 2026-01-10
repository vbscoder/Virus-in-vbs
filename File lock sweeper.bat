@echo off
title Instant Permanent Delete

echo WARNING! Files will be PERMANENTLY deleted!
set /p "file=File name to NUKE: "

echo.
echo Deleting everything called: %file%
echo.

del /f /q "%file%" 2>nul
del /f /q /a "%file%" 2>nul

for %%d in (C D E F G H I J K L M N O P Q R S T U V W X Y Z) do (
    if exist "%%d:\$Recycle.Bin" (
        del /f /q "%%d:\$Recycle.Bin\*\*%file%*" 2>nul
    )
)

echo Done. If file is still here → it's probably in use.
pause
