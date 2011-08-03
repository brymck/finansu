@echo off
echo ------------------
echo FinAnSu Post-Build
echo ------------------
echo.

echo Moving DNA, DLL and XLL to bin\Release\unpacked . . .
cd ..\bin\Release\
xcopy FinAnSu.* unpacked\FinAnSu.* /i /y
del FinAnSu.*
echo.

echo Packing into single XLL . . .
..\..\Resources\ExcelDnaPack unpacked\FinAnSu.dna /y
move unpacked\FinAnSu-packed.xll FinAnSu.xll
echo.

echo Packing into ZIP archive . . .
..\..\Resources\7za a -tzip FinAnSu.zip *.*
echo.

echo Copying to %AppData%\Microsoft\AddIns . . .
copy FinAnSu.xll "%AppData%\Microsoft\AddIns\FinAnSu.xll"
pause