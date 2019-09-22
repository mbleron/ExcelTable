@echo off
pushd %cd%
cd %~dp0\lib
echo Please enter target database information...
setlocal
set sid=%ORACLE_SID%
set /p sid="SID [%sid%]: "
set /p user="User: "
call loadjava -u %user%@%sid% -r -v -jarsasdbobjects -fileout ..\install_jdk6.log exceldbtools-1.6.jar
pause
endlocal
popd
