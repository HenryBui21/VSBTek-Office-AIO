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
	
::Version: 3.0
::Developer: Thanos
::OS support [32+64bit]: Windows 7/8/8.1 (chi cai duoc Office 2010, 2013, 2016 Volume), Windows 10 (cai duoc moi ban), Windows 11 (cai duoc moi ban)

====================================================================
title Ho tro cac van de ve Office cho may tinh!
color f0


:MainMenu
mode con: cols=68 lines=25
echo. 
cls
echo.
echo.                          == MENU ==
echo.      
echo.      [  1. Cai dat Office (Word, Excel...)    : Nhan so 1  ] 
echo.
echo.      [  2. Cai dat Project - Visio            : Nhan so 2  ]
echo.	  
echo.      [  3. Xuat bieu tuong ra Desktop         : Nhan so 3  ]
echo.
echo.      [  4. Go cai Office tan goc              : Nhan so 4  ]
echo.
echo.      [  5. Download file ISO Office           : Nhan so 5  ]
echo.	  
echo.        
echo.
@echo ===========================
Choice /N /C 12345 /M "* Nhap lua chon cua ban: 

if ERRORLEVEL 5 goto:downloadISO
if ERRORLEVEL 4 goto:uninstalloffice      
if ERRORLEVEL 3 goto:in_shortcut_office
if ERRORLEVEL 2 goto:installproject_visio
if ERRORLEVEL 1 goto:installoffice




:============================================================================================================
:installoffice
start office.cmd
goto:MainMenu


:============================================================================================================
:installproject_visio
start project_visio.cmd
goto:MainMenu






:============================================================================================================
:in_shortcut_office
cls
echo.
echo. ========================================
echo.   Dang xuat bieu tuong Office ra Desktop...
echo. ========================================
echo.
timeout /t 1 >nul

REM Copy shortcuts from main Programs folder
COPY /Y "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Word*.lnk" "%AllUsersProfile%\Desktop" 2>nul
COPY /Y "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Excel*.lnk" "%AllUsersProfile%\Desktop" 2>nul
COPY /Y "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\PowerPoint*.lnk" "%AllUsersProfile%\Desktop" 2>nul
COPY /Y "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Outlook*.lnk" "%AllUsersProfile%\Desktop" 2>nul
COPY /Y "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\OneNote*.lnk" "%AllUsersProfile%\Desktop" 2>nul
COPY /Y "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Access*.lnk" "%AllUsersProfile%\Desktop" 2>nul
COPY /Y "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Publisher*.lnk" "%AllUsersProfile%\Desktop" 2>nul
COPY /Y "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Visio*.lnk" "%AllUsersProfile%\Desktop" 2>nul
COPY /Y "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Project*.lnk" "%AllUsersProfile%\Desktop" 2>nul

REM Copy shortcuts from Microsoft Office folder (Office 2016+)
COPY /Y "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office\Word*.lnk" "%AllUsersProfile%\Desktop" 2>nul
COPY /Y "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office\Excel*.lnk" "%AllUsersProfile%\Desktop" 2>nul
COPY /Y "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office\PowerPoint*.lnk" "%AllUsersProfile%\Desktop" 2>nul
COPY /Y "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office\Outlook*.lnk" "%AllUsersProfile%\Desktop" 2>nul
COPY /Y "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office\OneNote*.lnk" "%AllUsersProfile%\Desktop" 2>nul
COPY /Y "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office\Access*.lnk" "%AllUsersProfile%\Desktop" 2>nul
COPY /Y "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office\Publisher*.lnk" "%AllUsersProfile%\Desktop" 2>nul
COPY /Y "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office\Visio*.lnk" "%AllUsersProfile%\Desktop" 2>nul
COPY /Y "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office\Project*.lnk" "%AllUsersProfile%\Desktop" 2>nul

REM Copy shortcuts from Office Tools folder
COPY /Y "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office Tools\*.lnk" "%AllUsersProfile%\Desktop" 2>nul

echo.
echo. ========================================
echo.   Da xuat bieu tuong Office ra Desktop!
echo. ========================================
echo.
echo. Nhan phim bat ky de quay lai menu...
pause >nul
goto:MainMenu


:=========================================================================================
:uninstalloffice
cls
mode con: cols=70 lines=25
color f0
cls
echo. 
echo.                             == MENU ==
echo.      
echo.         [  1. OfficeScrubber (Recommended)     : Nhan so 1  ]
echo.         
echo.         [  2. Revo Uninstaller (Portable)      : Nhan so 2  ]
echo.
echo.        ---------------------------------------------------
echo.
echo.                  [  3. Quay lai   : Nhan so 3  ]
echo.
echo.
@echo ===========================
Choice /N /C 123 /M "* Nhap lua chon cua ban : 
if ERRORLEVEL 3 goto:MainMenu
if ERRORLEVEL 2 goto:off_revo
if ERRORLEVEL 1 goto:off_scrubber

=========================
:off_scrubber
cls
echo.
echo.
echo. OfficeScrubber dang khoi dong (che do tuong tac)...
echo.
start remove_office_tan_goc\OfficeScrubber\OfficeScrubber.cmd
timeout 10
goto:uninstalloffice

============================
:off_revo
cls
echo.
echo. Dang tai Revo Uninstaller Portable...
echo.
start https://www.revouninstaller.com/start-freeware-download-portable/
timeout 5
goto:uninstalloffice







:======================================================================================================================================================
:convertok
mode con: cols=70 lines=27
echo. 
cls
set id1=O365ProPlusRetail
set id2=ProPlus2019Retail
set id3=ProPlus2019Volume
set id4=ProPlusRetail
set id5=Office16.PROPLUS
set id6=Office15.PROPLUSR
set id7=Office14.PROPLUSR
set id8=Officel4.PROPLUS
set id9=ProPlus2021Retail
set id10=ProPlus2021Volume

for %%b in (%id1%,%id2%,%id3%,%id4%,%id5%,%id6%,%id7%,%id8%,%id9%,%id10%) do (for /f "delims=" %%A in ('reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\%%b" /v "Displayname"') do cls&set a1=%%A&set b1=%%b)
for %%b in (%id1%,%id2%,%id3%,%id4%,%id5%,%id6%,%id7%,%id8%,%id9%,%id10%) do (for /f "delims=" %%A in ('reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\%%b - en-us" /v "Displayname"') do cls&set a1=%%A&set b1=%%b)
for %%c in (%id1%,%id2%,%id3%,%id4%,%id5%,%id6%,%id7%,%id8%,%id9%,%id10%) do (for /f "delims=" %%A in ('reg query "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\%%c" /v "Displayname"') do cls&set a1=%%A&set b1=%%c)
for %%c in (%id1%,%id2%,%id3%,%id4%,%id5%,%id6%,%id7%,%id8%,%id9%,%id10%) do (for /f "delims=" %%A in ('reg query "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\%%c - en-us" /v "Displayname"') do cls&set a1=%%A&set b1=%%c)

:menu_convert
cls
echo.
echo.
if [%b1%] EQU [O365ProPlusRetail] echo.  Dang su dung:       Microsoft 365 Apps for enterprise&set ak47=365
if [%b1%] EQU [ProPlus2019Retail] echo.  Dang su dung:     Office Professional Plus 2019 (Retail)&set ak47=19
if [%b1%] EQU [ProPlus2019Volume] echo.  Dang su dung:     Office Professional Plus 2019 (Volume)&set ak47=19
if [%b1%] EQU [ProPlusRetail]     echo.  Dang su dung:     Office Professional Plus 2016 (Retail)&set ak47=16
if [%b1%] EQU [Office16.PROPLUS]  echo.  Dang su dung:     Office Professional Plus 2016 (Volume)&set ak47=16
if [%b1%] EQU [Office15.PROPLUSR] echo.  Dang su dung:     Office Professional Plus 2013 (Retail)
if [%b1%] EQU [Office14.PROPLUSR] echo.  Dang su dung:     Office Professional Plus 2010 (Retail)
if [%b1%] EQU [Office14.PROPLUS]  echo.  Dang su dung:     Office Professional Plus 2010 (Volume)
if [%b1%] EQU [ProPlus2021Retail] echo.  Dang su dung:     Office Professional Plus 2021 (Retail)&set ak47=21
if [%b1%] EQU [ProPlus2021Volume] echo.  Dang su dung:     Office LTSC Professional Plus 2021 (Volume)&set ak47=21




set off21=""
echo.
echo.
echo.      
echo.      [  1. Retail sang Volume                      : Nhan so 1  ] 
echo.
echo.      [  2. Volume sang Retail                      : Nhan so 2  ]
echo.
echo.      [  3. ProPlus sang "Home and Student"         : Nhan so 3  ]
echo.
echo.      [  4. ProPlus sang "Home and Bussiness"       : Nhan so 4  ]
echo.
echo.      [  5. "Student" or "Bussiness" sang Pro Plus  : Nhan so 5  ]
echo.
echo.           -------------------------------------------
echo.
echo.               [  6.Quay lai: Nhan so 6  ]
echo.	  
echo.	  
echo.        
echo.
@echo ===========================
Choice /N /C 123456 /M "* Nhap lua chon cua ban: 

if ERRORLEVEL 6 goto:orther    
if ERRORLEVEL 5 set option=retail&goto:khoidau      
if ERRORLEVEL 4 set option=buss&set off21=no&goto:khoidau      
if ERRORLEVEL 3 set option=student&set off21=no&goto:khoidau      
if ERRORLEVEL 2 set option=retail&goto:khoidau    
if ERRORLEVEL 1 set option=vl&goto:khoidau




:khoidau
if [%ak47%] EQU [16] goto:office2016
if [%ak47%] EQU [19] goto:office2019
if [%ak47%] EQU [21] goto:office2021
if [%ak47%] EQU [365] echo.&echo. Khong Duoc Phep Convert!!!&timeout 5&goto:menu_convert


:office2019
if [%option%] EQU [vl] goto:vl19
if [%option%] EQU [retail] goto:retail19
if [%option%] EQU [student] goto:student19
if [%option%] EQU [buss] goto:buss19

:retail19
set Key=NW69C-TFYXD-YBDBD-KBMHD-MDYCT
set Description=ProPlus2019MSDNR_Retail
goto:ketthuc
:student19
set Key=FMTDN-KMP4R-H983J-6J87C-YKPY7
set Description=HomeStudent2019R_OEM_Perp
goto:ketthuc
:buss19
set Key=BNH3K-7JG4D-C7XF8-6FM6B-8FJFG
set Description=HomeBusiness2019R_Retail
goto:ketthuc
:vl19
set Key=W9HYN-C8J79-2YGTT-JVQW8-K2GT3
set Description=ProPlus2019VL_MAK_AE
goto:ketthuc



:office2021
if [%option%] EQU [vl] goto:vl21
if [%option%] EQU [retail] goto:retail21
if [%off21%] EQU [no] goto:echo.&echo. Khong Co Data de Convert!!!&timeout 5&goto:menu_convert
if [%off21%] EQU [no] goto:echo.&echo. Khong Co Data de Convert!!!&timeout 5&goto:menu_convert

:retail21
set Key=9D9QG-NKMFF-GGGR8-TYF9Y-3GRY9
set Description=ProPlus2021MSDNR_Retail
goto:ketthuc
:vl21
set Key=XN3GG-8V6WR-GXYKF-3MX6H-JQP89
set Description=ProPlus2021VL_MAK_AE
goto:ketthuc



:offfice2016
if [%option%] EQU [vl] goto:vl16
if [%option%] EQU [retail] goto:retail16
if [%option%] EQU [student] goto:student16
if [%option%] EQU [buss] goto:buss16

:retail16
set Key=BQ8H7-N8MWT-9FD39-24MDC-TMVHC
set Description=Office16_ProPlusMSDNR_Retail
goto:ketthuc
:student16
set Key=QN97V-HRX9H-QQVB6-8Q9BM-YWRFK
set Description=Office16_HomeStudentR_Retail
goto:ketthuc
:buss16
set Key=WBXCT-N9PQ9-WJG7Q-CYH4D-BG9VX
set Description=Office16_HomeBusinessR_Retail
goto:ketthuc
:vl16
set Key=KC7N8-WXH8P-8M8R7-QWYKK-JXCQY
set Description=Office16_ProPlusVL_MAK
goto:ketthuc




:ketthuc
::Nguon code convert Office : VN- ZOOM (user: Nhanchu)
mode con: cols=80 lines=25
reg Delete HKLM\Software\Wow6432Node\Microsoft\Office\16.0\Common\OEM /f
reg Delete HKLM\Software\Microsoft\Office\16.0\Common\OEM /f
for %%a in (4,5,6) do (if exist "%ProgramFiles%\Microsoft Office\Office1%%a\ospp.vbs" (cd /d "%ProgramFiles%\Microsoft Office\Office1%%a")
If exist "%ProgramFiles% (x86)\Microsoft Office\Office1%%a\ospp.vbs" (cd /d "%ProgramFiles% (x86)\Microsoft Office\Office1%%a"))&cls
for /f "tokens= 8" %%b in ('cscript //nologo OSPP.VBS /dstatus ^| findstr /b /c:"Last 5"') do (cscript //nologo ospp.vbs /unpkey:%%b)
for /f %%i in ('dir /b ..\root\Licenses16\%Description%*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%i"
cscript ospp.vbs /remhst
cscript ospp.vbs /ckms-domain
cscript ospp.vbs /inpkey:%key%
cscript //nologo ospp.vbs /act
echo.
echo.=========================
echo. Da convert thanh cong!
echo.=========================
echo.
echo.
timeout 7
goto:menu_convert
::Nguon code convert Office : VN- ZOOM (user: Nhanchu)








:======================================================================================================================================================
:downloadISO
cls
echo.
echo.  ===  Lua chon phien ban Office ban muon tai xuong ===
echo.
echo.
ECHO              1. Office 2019 Pro Plus 
ECHO              -----------------------
ECHO              2. Office 2021 Pro Plus 
ECHO              -----------------------
ECHO              3. Office 2024 Pro Plus 
ECHO              -----------------------
ECHO              4. Office 365 Pro Plus 
ECHO              -----------------------
ECHO              5. Quay lai 
echo.
echo.
echo. -----------------------
choice /c:12345 /n /m "Chon phien ban muon tai xuong [1,2,3,4,5] : "
if %errorlevel% EQU 1 goto:2019_retail_3264bit
if %errorlevel% EQU 2 goto:2021_retail_3264bit
if %errorlevel% EQU 3 goto:2024_retail_3264bit
if %errorlevel% EQU 4 goto:365_retail_3264bit
if %errorlevel% EQU 5 goto:MainMenu
if %errorlevel% NEQ 1 goto:downloadISO
if %errorlevel% NEQ 2 goto:downloadISO
if %errorlevel% NEQ 3 goto:downloadISO
if %errorlevel% NEQ 4 goto:downloadISO
if %errorlevel% NEQ 5 goto:downloadISO



:oknha
:Retail
:2019_retail_3264bit
cls
echo.
echo. ==================================================
echo.   Dang mo trang tai Office 2019...
echo. ==================================================
echo.
echo. Huong dan: 
echo. 1. Tai file ISO (format .img hoac .iso)
echo. 2. Extract file ISO neu can thiet
echo. 3. Chay setup.exe de cai dat Office
echo.
start https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/ProPlus2019Retail.img
timeout 3
goto:downloadISO

:2021_retail_3264bit
cls
echo.
echo. ==================================================
echo.   Dang mo trang tai Office 2021...
echo. ==================================================
echo.
echo. Huong dan: 
echo. 1. Tai file ISO (format .img hoac .iso)
echo. 2. Extract file ISO neu can thiet
echo. 3. Chay setup.exe de cai dat Office
echo.
start https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/ProPlus2021Retail.img
timeout 3
goto:downloadISO

:2024_retail_3264bit
cls
echo.
echo. ==================================================
echo.   Dang mo trang tai Office 2024...
echo. ==================================================
echo.
echo. Huong dan: 
echo. 1. Tai file ISO (format .img hoac .iso)
echo. 2. Extract file ISO neu can thiet
echo. 3. Chay setup.exe de cai dat Office
echo.
start https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/ProPlus2024Retail.img
timeout 3
goto:downloadISO

:365_retail_3264bit
cls
echo.
echo. ==================================================
echo.   Dang mo trang tai Office 365...
echo. ==================================================
echo.
echo. Huong dan: 
echo. 1. Tai file ISO (format .img hoac .iso)
echo. 2. Extract file ISO neu can thiet
echo. 3. Chay setup.exe de cai dat Office
echo.
start https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/O365ProPlusRetail.img
timeout 3
goto:downloadISO






:======================================================================================================================================================
:Exit
echo. Good Bye!
timeout 3
exit
































::cach in ky tu ra file notepad thanh cong!
::Tất cả các ký hiệu Greater-Than (>), Less-Than (<), Pipe (|), Ampersand (&) và Caret (^) đều cần phải được thoát bằng dấu mũ (^) trừ khi chúng được chứa trong “dấu ngoặc kép”
::vd: echo ^)                                                                                                  ::
::vd2: echo ^> ^< ^| ^& ^^