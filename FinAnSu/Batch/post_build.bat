@echo off
echo ------------------
echo FinAnSu Post-Build
echo ------------------
echo.

echo %PATH%
cd ..\bin\Release\

echo Recreating x64 and x86 directories...
md x86
md x64
del x86 /s /q
del x64 /s /q
echo.

echo Copying DNA, DLL and XLL to bin\Release\x86...
copy FinAnSu.dll x86
copy ..\..\Resources\ExcelDna.dna x86\FinAnSu.dna /y
copy ..\..\Resources\ExcelDna.xll x86\FinAnSu.xll /y
echo.

echo Copying DNA, DLL and XLL to bin\Release\x64...
copy FinAnSu.dll x64 /y
copy ..\..\Resources\ExcelDna64.dna x64\FinAnSu.dna /y
copy ..\..\Resources\ExcelDna64.xll x64\FinAnSu.xll /y
echo.

echo Packing into single XLL . . .
ExcelDnaPack x86\FinAnSu.dna /y
if errorlevel 9009 goto ExcelDnaNotFound
move /y x86\FinAnSu-packed.xll FinAnSu.xll
ExcelDnaPack x64\FinAnSu.dna /y
move /y x64\FinAnSu-packed.xll FinAnSu_x64.xll
echo.

echo Getting rid of extraneous config files...
del FinAnSu.dll.config

echo Cleaning up archive...
7za d FinAnSu.zip
if errorlevel 9009 goto SevenZipNotFound
echo.

echo Packing into archive...
7za a -tzip FinAnSu.zip *.*
echo.
goto Exit

:ExcelDnaNotFound
echo You must have the following ExcelDna files in your PATH:
echo   - ExcelDnaPack.exe
echo   - ExcelDna.Integration.dll
echo   - ExcelDna.xll
echo   - ExcelDna64.xll
exit /b 1

:SevenZipNotFound
echo You must have 7za.exe somewhere in your PATH
exit /b 2

:Exit
exit /b 0
