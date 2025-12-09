CHCP 1258 >nul 2>&1
CHCP 65001 >nul 2>&1
@echo off
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
if '%errorlevel%' NEQ '0' (
    echo  Chay CMD voi quyen Quan tri vien...
    goto goUAC 
) else (
 goto goADMIN )

:goUAC
    echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\getadmin.vbs"
    set params = %*:"=""
    echo UAC.ShellExecute "cmd.exe", "/c %~s0 %params%", "", "runas", 1 >> "%temp%\getadmin.vbs"
    "%temp%\getadmin.vbs"
    del "%temp%\getadmin.vbs"
    exit /B

:goADMIN
    pushd "%CD%"
    CD /D "%~dp0"
    
::Cài đặt Project-Visio
::Version: 2.0
::Developer: Thanos

title Cài đặt Project-Visio cho máy tính!
cls
color f0
mode con: cols=60 lines=27

:MainMenu
del /s /f /q Configuration.xml
cls
set zz=
set pp=
set tt=
set mm=
set ee=
set cc=
set nn=
set vv=
set gg=
set ff=
set xx=
set yy=

:===========================================================================================================
:batdau
>> "Configuration.xml" echo ^<Configuration^>

echo.
echo.          Chọn phiên bản 32bit hoặc 64bit
echo.
echo.
echo.      (A): 32bit        ;         (B): 64bit
echo.
Choice /N /C AB /M "* Nhập Lựa Chọn Của Bạn [A hoặc B] :
if ERRORLEVEL 2 set xx=64
if ERRORLEVEL 1 set xx=32
>> "Configuration.xml" echo  ^<Add OfficeClientEdition="%xx%" ^>
cls
echo.
echo.
echo.
:project
echo. 1. Bạn có muốn cài Project Pro?
Choice /N /C YN /M "* Y:Có , N:Không - [Y hoặc N] :
if ERRORLEVEL 2 echo. == Không ==&goto:visio
if ERRORLEVEL 1 echo. == Có ==&set zz=ProjectPro&set vv=Project Professional

echo.
echo.   (1): phiên bản 2016
echo.   (2): phiên bản 2019
echo.   (3): phiên bản 2021
echo.
echo.
choice /c:123 /n /m "Nhập số của phiên bản muốn cài đặt [1,2,3] : "
if %errorlevel% EQU 3 set pp=2021
if %errorlevel% EQU 2 set pp=2019
if %errorlevel% EQU 1 set pp=&set gg=2016_

::retail-volume
set tt=Retail



:display
>> "Configuration.xml" echo  ^<Product ID="%zz%%pp%%tt%"^>
>> "Configuration.xml" echo  ^<Language ID="en-us" /^>
>> "Configuration.xml" echo  ^</Product^>




:visio
echo.
echo. 2. Bạn có muốn cài Visio Pro?
Choice /N /C YN /M "* Y:Có , N:Không - [Y hoặc N] :
if ERRORLEVEL 2 echo. == Không ==&goto:end_all
if ERRORLEVEL 1 echo. == Có ==&set mm=VisioPro&set nn=Visio Professional

echo.
echo.   (1): phiên bản 2016
echo.   (2): phiên bản 2019
echo.   (3): phiên bản 2021
echo.
echo.
choice /c:123 /n /m "Nhập số của phiên bản muốn cài đặt [1,2,3] : "
if %errorlevel% EQU 3 set cc=2021
if %errorlevel% EQU 2 set cc=2019
if %errorlevel% EQU 1 set cc=&set ff=2016_

::retail-volume
set ee=Retail


:display
>> "Configuration.xml" echo  ^<Product ID="%mm%%cc%%ee%"^>
>> "Configuration.xml" echo  ^<Language ID="en-us" /^>
>> "Configuration.xml" echo  ^</Product^>




:===========================================================================================================
:end_all
>> "Configuration.xml" echo  ^</Add^>
>> "Configuration.xml" echo  ^<Display AcceptEULA="True" /^>
>> "Configuration.xml" echo  ^<Extend CreateShortcuts="true" /^>
>> "Configuration.xml" echo  ^</Configuration^>

cls
echo.
echo.
echo.
::xet dieu kien
if [%zz%] EQU [ProjectPro] goto:chuyentiep
if [%zz%] NEQ [ProjectPro] goto:chuyentiep2

 :chuyentiep
if [%mm%] EQU [VisioPro] set chicopro=no&goto:co_pro_ne
if [%mm%] NEQ [VisioPro] set chicopro=yes&goto:co_pro_ne

:chuyentiep2
if [%mm%] NEQ [VisioPro] goto:MainMenu
if [%mm%] EQU [VisioPro] goto:co_vi_ne



::DISPLAY
:co_pro_ne
echo.      === %vv% %gg%%pp%_%tt%_%xx%bit ===
echo.
echo.
if [%chicopro%] EQU [yes] goto:endgame

:co_vi_ne
echo.      === %nn% %ff%%cc%_%ee%_%xx%bit ===
echo.
echo.






:endgame
echo.
echo.               === BẮT ĐẦU CÀI ĐẶT? ===
echo.
echo.             (Y): Có     ;      (N): Không
echo.
Choice /N /C YN /M "* Nhập Lựa Chọn Của Bạn [Y hoặc N] :
if ERRORLEVEL 2 del /s /f /q Configuration.xml&cls&goto:MainMenu
if ERRORLEVEL 1 cls

mode con: cols=50 lines=15
echo.
echo. Đang bắt đầu quá trình cài đặt Project/Visio....
echo.
setup.exe /configure Configuration.xml
exit






