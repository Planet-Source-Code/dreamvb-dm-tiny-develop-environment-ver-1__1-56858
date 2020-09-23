@echo off
CLS

echo this will compress exehead using UPX this process
echo will only need to preformed once.
echo.
pause


if exist "upx.exe" goto CompressExe

goto ende


:CompressExe
upx --best --crp-ms=100000 exehead.exe 2>log.txt
goto Finsihed

:Finsihed
CLS
echo. exehead.exe has now been compressed.
echo.
pause
goto Close

:ende
CLS
echo.
echo. UPX was not found on your system.
echo. Please make sure the application was not deleted by mistake.
echo.
pause

:Close
