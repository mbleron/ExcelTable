@echo off
pushd %cd%
cd %~dp0\lib
echo Please enter target database information...
setlocal
set sid=%ORACLE_SID%
set /p sid="SID [%sid%]: "
set /p user="User: "
call loadjava -u %user%@%sid% -r -v -jarsasdbobjects -fileout ..\install_jdk5.log stax-api-1.0-2.jar sjsxp-1.0.2.jar exceldbtools-1.5.jar 
pause
endlocal
popd
