CHCP 1258 >nul 2>&1
CHCP 65001 >nul 2>&1
@echo off
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
if '%errorlevel%' NEQ '0' (
    echo  Run CMD as Administrator...
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

		

::Version: 2.0
::Developer: Thanos
::OS support [32+64bit]: Windows 7/8/8.1 (chi cai duoc Office 2010, 2013, 2016 Volume), Windows 10 (cai duoc moi ban), Windows 11 (cai duoc moi ban)

:========================================================================================================
:MainMenu
title Cai dat Word,Excel,Powerpoint... cho may tinh!
color f0
mode con: cols=58 lines=27

:startok
del /s /f /q Configuration.xml
cls
set aa=
set bb=
set xx=
set yy=
set off365=
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
set quickmode=
>> "Configuration.xml" echo ^<Configuration^>
echo.
echo. ===  Lua chon phien ban Office ban muon cai dat ===
echo.
echo.
ECHO              1. Office 2007 Pro Plus
ECHO              -----------------------
ECHO              2. Office 2010 Pro Plus
ECHO              -----------------------
ECHO              3. Office 2013 Pro Plus
ECHO              -----------------------
ECHO              4. Office 2016 Pro Plus
ECHO              -----------------------
ECHO              5. Office 2019 Pro Plus 
ECHO              -----------------------
ECHO              6. Office 2021 Pro Plus 
ECHO              -----------------------
ECHO              7. Office 2024 Pro Plus 
ECHO              -----------------------
ECHO              8. Office 365 Pro Plus 
echo.
echo.
echo. -----------------------
choice /c:12345678 /n /m "Chon phien ban muon cai dat [1,2,3,4,5,6,7,8] : "
if %errorlevel% EQU 1 set aa=2007&set yy=Office Professional Plus 2007&set quickinstall=0&goto:1
if %errorlevel% EQU 2 set aa=2010&set yy=Office Professional Plus 2010&set quickinstall=0&goto:1
if %errorlevel% EQU 3 set aa=2013&set yy=Office Professional Plus 2013&set quickinstall=0&goto:1
if %errorlevel% EQU 4 set aa=ProPlus&set yy=Office Professional Plus 2016&set quickinstall=0&goto:1
if %errorlevel% EQU 5 set aa=ProPlus2019&set yy=Office Professional Plus 2019&goto:quickselect
if %errorlevel% EQU 6 set aa=ProPlus2021&set yy=Office Professional Plus 2021&goto:quickselect
if %errorlevel% EQU 7 set aa=ProPlus2024&set yy=Office Professional Plus 2024&goto:quickselect
if %errorlevel% EQU 8 set aa=O365ProPlus&set yy=Office 365&set off365=ok&goto:quickselect
if %errorlevel% NEQ 1 goto:startok
if %errorlevel% NEQ 2 goto:startok
if %errorlevel% NEQ 3 goto:startok
if %errorlevel% NEQ 4 goto:startok
if %errorlevel% NEQ 5 goto:startok
if %errorlevel% NEQ 6 goto:startok
if %errorlevel% NEQ 7 goto:startok
if %errorlevel% NEQ 8 goto:startok

:quickselect
cls
echo.
echo. ===  Lua chon che do cai dat ===
echo.
echo.      (A): Cai dat Nhanh (Word, Excel, PowerPoint,...)
echo.      (B): Cai dat Tuy chinh (Chon tung ung dung)
echo.
Choice /N /C AB /M "* Nhap lua chon cua ban [A hoac B] : "
if errorlevel 2 goto:custominstall
if errorlevel 1 goto:quickinstall

:custominstall
set quickmode=0
goto:1

:quickinstall
set quickmode=1
goto:1

:1
echo.
echo.      (A): 32bit     ;      (B): 64bit
echo.
choice /c AB /n /m "Nhap lua chon cua ban [A hoac B] : "
if errorlevel 2 set xx=64&set bb=Volume
if errorlevel 1 set xx=32&set bb=Volume
>> "Configuration.xml" echo  ^<Add OfficeClientEdition="%xx%" ^>

::retail-volume
if [%off365%] EQU [ok] set bb=Retail&goto:tiepdi
set bb=Volume&goto:tiepdi

:tiepdi
::display
if [%aa%] EQU [2007] goto:download
if [%aa%] EQU [2010] goto:download
if [%aa%] EQU [2013] goto:download
if [%aa%] EQU [ProPlus] set aa=2016&goto:download
if [%aa%] EQU [ProPlus2024] goto:display
goto:display

:display
>> "Configuration.xml" echo  ^<Product ID="%aa%%bb%"^>
cls
goto:part1

::Option_App
:part1
>> "Configuration.xml" echo  ^<Language ID="en-us" /^>

::Check if quickinstall mode - auto select apps
if not "%quickmode%"=="1" goto:custommode
echo.
echo. (Cai dat nhanh: Word, Excel, PowerPoint, Outlook, Teams)
echo.
set a=Word
set b= + Excel
set c= + PowerPoint
set f= + Outlook
set i= + Teams
>> "Configuration.xml" echo  ^<ExcludeApp ID="Access" /^>
>> "Configuration.xml" echo  ^<ExcludeApp ID="Publisher" /^>
>> "Configuration.xml" echo  ^<ExcludeApp ID="OneNote" /^>
>> "Configuration.xml" echo  ^<ExcludeApp ID="OneDrive" /^>
goto:endoffice

:custommode

echo.
echo.
echo.    ___Lua chon phan mem ban muon "cai/khong cai"____
echo.
echo.
echo.
echo. 1/ Ban co muon cai Word?
Choice /N /C YN /M "* Y:Co , N:Khong - [Y hoac N] :
if ERRORLEVEL 2 echo. == Khong ==&goto:lem1
if ERRORLEVEL 1 echo. == Co ==&set a=Word&goto:part2
:lem1
>> "Configuration.xml" echo  ^<ExcludeApp ID="Word" /^> 

:part2
echo.   
echo. 2/ Ban co muon cai Excel?
Choice /N /C YN /M "* Y:Co , N:Khong - [Y hoac N] :
if ERRORLEVEL 2 echo. == Khong ==&goto:lem2
if ERRORLEVEL 1 echo. == Co ==&set b= + Excel&goto:part3
:lem2
>> "Configuration.xml" echo  ^<ExcludeApp ID="Excel" /^> 

:part3
echo.
echo. 3/ Ban co muon cai PowerPoint?
Choice /N /C YN /M "* Y:Co , N:Khong - [Y hoac N] :
if ERRORLEVEL 2 echo. == Khong ==&goto:lem3
if ERRORLEVEL 1 echo. == Co ==&set c= + PowerPoint&goto:part4
:lem3
>> "Configuration.xml" echo  ^<ExcludeApp ID="PowerPoint" /^> 

:part4
echo.
echo. 4/ Ban co muon cai Access?
Choice /N /C YN /M "* Y:Co , N:Khong - [Y hoac N] :
if ERRORLEVEL 2 echo. == Khong ==&goto:lem4
if ERRORLEVEL 1 echo. == Co ==&set d= + Access&goto:part5
:lem4
>> "Configuration.xml" echo  ^<ExcludeApp ID="Access" /^> 


::Check if Office 2024 - skip Publisher question
if [%aa%] EQU [ProPlus2024] goto:skip_publisher

:part5
echo.
echo. 5/ Ban co muon cai Publisher?
Choice /N /C YN /M "* Y:Co , N:Khong - [Y hoac N] :
if ERRORLEVEL 2 echo. == Khong ==&goto:lem5
if ERRORLEVEL 1 echo. == Co ==&set e= + Publisher&goto:part6
:lem5
>> "Configuration.xml" echo  ^<ExcludeApp ID="Publisher" /^> 
goto:part6

:skip_publisher
echo.
echo. (Publisher khong co trong Office 2024 Pro Plus)
>> "Configuration.xml" echo  ^<ExcludeApp ID="Publisher" /^>


:part6
echo.
echo. 6/ Ban co muon cai Outlook?
Choice /N /C YN /M "* Y:Co , N:Khong - [Y hoac N] :
if ERRORLEVEL 2 echo. == Khong ==&goto:lem6
if ERRORLEVEL 1 echo. == Co ==&set f= + Outlook&goto:part7
:lem6
>> "Configuration.xml" echo  ^<ExcludeApp ID="Outlook" /^> 


:part7
echo.
echo. 7/ Ban co muon cai OneNote?
Choice /N /C YN /M "* Y:Co , B:Khong - [Y hoac N] :
if ERRORLEVEL 2 echo. == Khong ==&goto:lem7
if ERRORLEVEL 1 echo. == Co ==&set g= + OneNote&goto:part8
:lem7
>> "Configuration.xml" echo  ^<ExcludeApp ID="OneNote" /^> 



:part8
echo.
echo. 8/ Ban co muon cai OneDrive?
Choice /N /C YN /M "* Y:Co , N:Khong - [Y hoac N] :
if ERRORLEVEL 2 echo. == Khong ==&goto:lem8
if ERRORLEVEL 1 echo. == Co ==&set h= + OneDrive&goto:part9
:lem8
>> "Configuration.xml" echo  ^<ExcludeApp ID="OneDrive" /^> 


:part9
if [%off365%] EQU [ok] goto:tieptuc
if [%off365%] NEQ [ok] goto:endoffice
:tieptuc
echo.
echo. 9/ Ban co muon cai Microsoft Teams?
Choice /N /C YN /M "* Y:Co , N:Khong - [Y hoac N] :
if ERRORLEVEL 2 echo. == Khong ==&goto:lem9
if ERRORLEVEL 1 echo. == Co ==&set i= + Teams&goto:endoffice
:lem9
>> "Configuration.xml" echo  ^<ExcludeApp ID="Teams" /^> 


:endoffice
>> "Configuration.xml" echo  ^<ExcludeApp ID="Groove" /^> 
>> "Configuration.xml" echo  ^<ExcludeApp ID="Lync" /^> 
>> "Configuration.xml" echo  ^</Product^>

::Check if Office 2019, 2021 or 365 - skip Project/Visio
if [%aa%] EQU [ProPlus2019] goto:end_all
if [%aa%] EQU [ProPlus2021] goto:end_all
if [%aa%] EQU [ProPlus2024] goto:end_all
if [%aa%] EQU [O365ProPlus] goto:end_all

:===========================================================================================================
:part12
echo.
echo.
echo.      ==================================
echo.               Project - Visio
echo.      ==================================
echo.
echo.
echo. 10/ Ban co muon cai Project Pro?
Choice /N /C AB /M "* A:Co , B:Khong - [A hoac B] :
if ERRORLEVEL 2 echo. == Khong ==&goto:part13
if ERRORLEVEL 1 echo. == Co ==&set zz=ProjectPro&set vv=Project Professional

echo.
echo.   (1): phien ban 2016
echo.   (2): phien ban 2019
echo.   (3): phien ban 2021
echo.
echo.
choice /c:123 /n /m "Nhap number cua phien ban muon cai dat [1,2,3] : "
if %errorlevel% EQU 3 set pp=2021
if %errorlevel% EQU 2 set pp=2019
if %errorlevel% EQU 1 set pp=&set gg=2016_

::retail-volume
set tt=Retail



:display
>> "Configuration.xml" echo  ^<Product ID="%zz%%pp%%tt%"^>
>> "Configuration.xml" echo  ^<Language ID="en-us" /^>
>> "Configuration.xml" echo  ^</Product^>




:part13
echo.
echo. 11/ Ban co muon cai Visio Pro?
Choice /N /C AB /M "* A:Co , B:Khong - [A hoac B] :
if ERRORLEVEL 2 echo. == Khong ==&goto:end_all
if ERRORLEVEL 1 echo. == Co ==&set mm=VisioPro&set nn=Visio Professional

echo.
echo.   (1): phien ban 2016
echo.   (2): phien ban 2019
echo.   (3): phien ban 2021
echo.
echo.
choice /c:123 /n /m "Nhap number cua phien ban muon cai dat [1,2,3] : "
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
echo.       ===========================================
echo.        %yy%_%bb%_%xx%bit
echo.       ===========================================
echo. Bao gom: %a%%b%%c%%d%%e%%f%%g%%h%%i%
echo.

::Notify if Office 2024
if [%aa%] EQU [ProPlus2024] (
echo. Luu y: Publisher khong co trong Office 2024 Pro Plus
echo.
)
echo.
echo.
if [%zz%] NEQ [ProjectPro] goto:buocnhay
echo.      === %vv% %gg%%pp%_%tt%_%xx%bit ===
echo.
echo.
:buocnhay
if [%mm%] NEQ [VisioPro] goto:buocnhay2
echo.      === %nn% %ff%%cc%_%ee%_%xx%bit ===
echo.
echo.
echo.
echo.
:buocnhay2
echo.
echo.               === BAT DAU CAI DAT? ===
echo.
echo.             (Y): Yes     ;      (N): No
echo.
Choice /N /C YN /M "* Nhap Lua Chon Cua Ban - [Y hoac N] :
if ERRORLEVEL 2 del /s /f /q Configuration.xml&cls&goto:startok
if ERRORLEVEL 1 cls

mode con: cols=50 lines=15
echo.
echo. Dang bat dau qua trinh cai dat Office....
echo.
setup.exe /configure Configuration.xml
exit

:download
mode con: cols=62 lines=20
if [%aa%] EQU [2007] goto:2007
if [%aa%] EQU [2010] goto:2010 
if [%aa%] EQU [2013] goto:2013
if [%aa%] EQU [2016] goto:2016ne

:2007
if [%xx%] NEQ [32] goto:64bit2007
if [%xx%] EQU [32] cls
if [%bb%] EQU [Volume] start https://drive.massgrave.dev/en_office_professional_plus_2007_x86_x15-74074.exe
goto:tieptheo
:64bit2007
if [%bb%] EQU [Volume] start https://drive.massgrave.dev/en_office_professional_plus_2007_x64_x15-74137.exe
goto:tieptheo

:2010
if [%xx%] NEQ [32] goto:64bitne
if [%xx%] EQU [32] cls
if [%bb%] EQU [Volume] start https://drive.massgrave.dev/SW_DVD5_Office_Professional_Plus_2010w_SP1_W32_English_CORE_MLF_X17-76748.ISO
goto:tieptheo
:64bitne
if [%bb%] EQU [Volume] start https://drive.massgrave.dev/SW_DVD5_Office_Professional_Plus_2010w_SP1_64Bit_English_CORE_MLF_X17-76756.ISO
goto:tieptheo

:2013
if [%xx%] NEQ [32] goto:64bitnha
if [%xx%] EQU [32] cls
if [%bb%] EQU [Volume] start https://drive.massgrave.dev/SW_DVD5_Office_Professional_Plus_2013w_SP1_W32_English_MLF_X19-35978.ISO
goto:tieptheo
:64bitnha
if [%bb%] EQU [Volume] start https://drive.massgrave.dev/SW_DVD5_Office_Professional_Plus_2013w_SP1_64Bit_English_MLF_X19-35906.ISO
goto:tieptheo


:2016ne
if [%xx%] EQU [32] start https://drive.massgrave.dev/SW_DVD5_Office_Professional_Plus_2016_W32_English_MLF_X20-42426.ISO
if [%xx%] EQU [64] start https://drive.massgrave.dev/SW_DVD5_Office_Professional_Plus_2016_64Bit_English_MLF_X20-42432.ISO
goto:tieptheo


:tieptheo
cls
echo.
echo.
echo.
echo.        === %yy%_%bb%_%xx%bit ===
echo.
echo. Nhan phim bat ky de quay lai MENU...
pause >nul
start Office.cmd
exit