::  Office Aktivasyon Aracç
::  Written by @BoomBookTR
::  https://github.com/BoomBookTR/Office-Legal-Activation-Script
::  Lisans Anahtarç Telegram Kanalçmçz: https://t.me/windows_office_etkinlestir


@echo off
setlocal EnableDelayedExpansion

::net file to test privileges, 1>NUL redirects output, 2>NUL redirects errors
NET FILE 1>NUL 2>NUL
if '%errorlevel%' == '0' ( goto START ) else ( goto getPrivileges ) 

:getPrivileges
if '%1'=='ELEV' ( goto START )

set "batchPath=%~f0"
set "batchArgs=ELEV"

::Add quotes to the batch path, if needed
set "script=%0"
set script=%script:"=%
IF '%0'=='!script!' ( GOTO PathQuotesDone )
    set "batchPath=""%batchPath%"""
:PathQuotesDone

::Add quotes to the arguments, if needed.
:ArgLoop
IF '%1'=='' ( GOTO EndArgLoop ) else ( GOTO AddArg )
    :AddArg
    set "arg=%1"
    set arg=%arg:"=%
    IF '%1'=='!arg!' ( GOTO NoQuotes )
        set "batchArgs=%batchArgs% "%1""
        GOTO QuotesDone
        :NoQuotes
        set "batchArgs=%batchArgs% %1"
    :QuotesDone
    shift
    GOTO ArgLoop
:EndArgLoop

::Create and run the vb script to elevate the batch file
echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\OEgetPrivileges.vbs"
echo UAC.ShellExecute "cmd", "/c ""!batchPath! !batchArgs!""", "", "runas", 1 >> "%temp%\OEgetPrivileges.vbs"
"%temp%\OEgetPrivileges.vbs" 
exit /B

:START
::Remove the elevation tag and set the correct working directory
IF '%1'=='ELEV' ( shift /1 )
cd /d %~dp0

:: .... your code start ....


@echo off

setlocal
call :setESC

cls

:: color help <<<<<<<<<<<<<<<<<<bÅtÅn renk bilgileri iáin cmd ekrançna yaz
:: color [arkaplanrengi][yazçrengi]
::color 1F
mode con lines=50 cols=150

:: Global options for ospp.vbs
:: https://docs.microsoft.com/en-us/deployoffice/vlactivation/tools-to-manage-volume-activation-of-office
::Product ID
::https://docs.microsoft.com/en-us/microsoft-365/troubleshoot/installation/product-ids-supported-office-deployment-click-to-run


title Office 2013-2016-2019-2021 Etkinleütirme Scripti (MenÅlÅ)
echo ============================================================================&
echo ============================================================================&
echo %ESC%[101;93m #Proje: Sadece Lisans Kodunu girerek otomatik aktivasyon iülemi saßlançr. %ESC%[0m&
echo %ESC%[101;93mNot: Bu script sadece var olan lisans anahtarçnçzç sisteme cmd Åzerinden iülenmesini ve sistemin aktivasyonunu saßlar.%ESC%[0m
echo %ESC%[101;93mKçsaca KMS vb. lisans anahtarç girmez vb. iülemler yÅrÅtmez.%ESC%[0m
echo.&
echo.&
echo Written by %ESC%[101;93m@BoomBookTR%ESC%[0m
echo.&
echo https://github.com/BoomBookTR/Office-Legal-Activation-Script
echo.&
echo Telegram Kanalç: https://t.me/windows_office_etkinlestir
echo Telegram Grubu: https://t.me/windows_office_etkinlestirme
echo.&
echo.&
echo ****************************************************************************&
echo ****************************************************************************&

echo %ESC%[104m #Desteklenen ÅrÅnler: %ESC%[0m& 
echo ============================================================================&
echo %ESC%[93m Office 2013 %ESC%[0m& 
echo %ESC%[93m Office 2016 %ESC%[0m&
echo %ESC%[93m Office 2019 %ESC%[0m& 
echo %ESC%[93m Office 2021 %ESC%[0m& 
echo.&
echo.& 
echo ============================================================================&


if exist "C:\Program Files\Microsoft Office\Office16\ospp.vbs" cd /d "C:\Program Files\Microsoft Office\Office16"
if exist "C:\Program Files (x86)\Microsoft Office\Office16\ospp.vbs" cd /d "C:\Program Files (x86)\Microsoft Office\Office16"
if exist "C:\Program Files\Microsoft Office\Office15\ospp.vbs" cd /d "C:\Program Files\Microsoft Office\Office15"
if exist "C:\Program Files (x86)\Microsoft Office\Office15\ospp.vbs" cd /d "C:\Program Files (x86)\Microsoft Office\Office15"
if exist "C:\Program Files\Microsoft Office\Office14\ospp.vbs" cd /d "C:\Program Files\Microsoft Office\Office14"
if exist "C:\Program Files (x86)\Microsoft Office\Office14\ospp.vbs" cd /d "C:\Program Files (x86)\Microsoft Office\Office14"
::for %%a in (4,5,6) do (if exist "%ProgramFiles%\Microsoft Office\Office1%%a\ospp.vbs" (cd "%ProgramFiles%\Microsoft Office\Office1%%a")
::if exist "%ProgramFiles(x86)%\Microsoft Office\Office1%%a\ospp.vbs" (cd "%ProgramFiles(x86)%\Microsoft Office\Office1%%a"))

set officedir=%cd%
::echo %officedir%



:baslangic
echo.&
echo %ESC%[101;93mNot:%ESC%[0m Bunu kullanmadan înce kullançm videosunu izleyiniz ve hemen altçndaki kçsa notlarç okuyunuz.
echo %ESC%[101;93mVideo ve detaylar:%ESC%[0m https://github.com/BoomBookTR/Office-Legal-Activation-Script
echo.&
echo Retail anahtarlar; Retail ISO sÅrÅmÅnÅ,
echo Volume anahtarlar; Office Deployment Tool ile Volume sÅrÅmÅnÅ etkinleütirir.
echo Volume anahtarlar ile Retail ISO sÅrÅmÅnÅ etkinleütirmek iáin Retail2Volume iülemi yapçnçz.
echo.&
echo ============================================================================&
echo.&
echo %ESC%[101;93mGenel anlamda doßru aktivasyon iáni sçralama üu üekildedir;%ESC%[0m
echo Mevcut lisans anahtarlarçnçn tÅmÅnÅ sil;Volume anahtar(MAK) girilecekse Retail2Volume yap;Anahtar gir;Etkinleütir
echo.&
ECHO ----------------------------------------------------------------------------&
ECHO %ESC%[92m1.%ESC%[0m Convert òülemi (Retail2Volume/Volume2Retail)
ECHO %ESC%[92m2.%ESC%[0m Lisans Anahtarlarçnçn TÅmÅnÅ Sil
ECHO %ESC%[92m3.%ESC%[0m Lisans Anahtarç Gir
ECHO %ESC%[92m4.%ESC%[0m Etkinleütirme Adçmçna Geá
ECHO %ESC%[92m5.%ESC%[0m Lisans Durumuna Bak
ECHO %ESC%[92m6.%ESC%[0m YÅklenmiü Lisansç Etkinleütirmeyi Dene
ECHO %ESC%[92m7.%ESC%[0m Lisans Yedekle
ECHO ----------------------------------------------------------------------------&
echo.&
echo.&
set choice=
set /p choice=%ESC%[101;93mYapçlacak iüleme ait numarayç yazçnçz! = %ESC%[0m
if not '%choice%'=='' set choice=%choice:~0,1%
if '%choice%'=='1' goto convert
if '%choice%'=='2' goto secim
if '%choice%'=='3' goto keygir
if '%choice%'=='4' goto onoff
if '%choice%'=='5' goto lisansdurum
if '%choice%'=='6' goto lisansetkinlestir
if '%choice%'=='7' goto yedekleme
ECHO "%choice%" geáersiz numara girdiniz.
ECHO.
goto baslangic


:convert
set /P j=%ESC%[7mRetail to VL%ESC%[0m (%ESC%[92mV%ESC%[0m) //// %ESC%[7mVL to Retail%ESC%[0m (%ESC%[92mR%ESC%[0m) //// %ESC%[7mAtla%ESC%[0m (%ESC%[92mA%ESC%[0m) -----------%ESC%[101;93mSEÄ%ESC%[0m---------[%ESC%[92mV%ESC%[0m/%ESC%[92mR%ESC%[0m/%ESC%[92mA%ESC%[0m]?
if /I "%j%" EQU "V" goto :Retail2VL
if /I "%j%" EQU "R" goto :VL2Retail
if /I "%j%" EQU "A" goto :baslangic
goto :convert


echo ============================================================================&

:Retail2VL
if exist "C:\Program Files\Microsoft Office\Office16\ospp.vbs" cd /d "C:\Program Files\Microsoft Office\Office16"
if exist "C:\Program Files (x86)\Microsoft Office\Office16\ospp.vbs" cd /d "C:\Program Files (x86)\Microsoft Office\Office16"
if exist "C:\Program Files\Microsoft Office\Office15\ospp.vbs" cd /d "C:\Program Files\Microsoft Office\Office15"
if exist "C:\Program Files (x86)\Microsoft Office\Office15\ospp.vbs" cd /d "C:\Program Files (x86)\Microsoft Office\Office15"
if exist "C:\Program Files\Microsoft Office\Office14\ospp.vbs" cd /d "C:\Program Files\Microsoft Office\Office14"
if exist "C:\Program Files (x86)\Microsoft Office\Office14\ospp.vbs" cd /d "C:\Program Files (x86)\Microsoft Office\Office14"
::for %%a in (4,5,6) do (if exist "%ProgramFiles%\Microsoft Office\Office1%%a\ospp.vbs" (cd "%ProgramFiles%\Microsoft Office\Office1%%a")
::if exist "%ProgramFiles(x86)%\Microsoft Office\Office1%%a\ospp.vbs" (cd "%ProgramFiles(x86)%\Microsoft Office\Office1%%a"))
echo ============================================================================&
echo ============================================================================&
echo ============================================================================&
echo ============================================================================&
echo ============================================================================&
echo ============================================================================&

set /P d=%ESC%[93mRetail to Volume (VL) iülemi iáin sÅrÅm seáiniz:%ESC%[0m %ESC%[7m2013%ESC%[0m (%ESC%[92m1%ESC%[0m) //// %ESC%[7m2016%ESC%[0m (%ESC%[92m2%ESC%[0m) //// %ESC%[7m2019%ESC%[0m (%ESC%[92m3%ESC%[0m) //// %ESC%[7m2021%ESC%[0m (%ESC%[92m4%ESC%[0m) //// %ESC%[7mATLA%ESC%[0m (%ESC%[92mA%ESC%[0m) iáin sçrasçyla [%ESC%[92m1%ESC%[0m/%ESC%[92m2%ESC%[0m/%ESC%[92m3%ESC%[0m/%ESC%[92m4%ESC%[0m/%ESC%[92mA%ESC%[0m] tçkla?
if /I "%d%" EQU "1" goto :retailtovolume2013
if /I "%d%" EQU "2" goto :retailtovolume2016
if /I "%d%" EQU "3" goto :retailtovolume2019
if /I "%d%" EQU "4" goto :retailtovolume2021
if /I "%d%" EQU "A" goto :baslangic
goto :Retail2VL

:retailtovolume2013
for /f %%x in ('dir /b "..\..\Microsoft Office 15\root\Licenses\"ProPlusVL*.xrm-ms') do cscript ospp.vbs /inslic:"..\..\Microsoft Office 15\root\Licenses\%%x"
goto :baslangic

:retailtovolume2016
for /f %%x in ('dir /b "..\root\Licenses16\"ProPlusVL*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x"
goto :baslangic

:retailtovolume2019
for /f %%x in ('dir /b ..\root\Licenses16\ProPlus2019VL*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x"
goto :baslangic

:retailtovolume2021
for /f %%x in ('dir /b ..\root\Licenses16\ProPlus2021VL*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x"
goto :baslangic

echo ============================================================================&

:VL2Retail
if exist "C:\Program Files\Microsoft Office\Office16\ospp.vbs" cd /d "C:\Program Files\Microsoft Office\Office16"
if exist "C:\Program Files (x86)\Microsoft Office\Office16\ospp.vbs" cd /d "C:\Program Files (x86)\Microsoft Office\Office16"
if exist "C:\Program Files\Microsoft Office\Office15\ospp.vbs" cd /d "C:\Program Files\Microsoft Office\Office15"
if exist "C:\Program Files (x86)\Microsoft Office\Office15\ospp.vbs" cd /d "C:\Program Files (x86)\Microsoft Office\Office15"
if exist "C:\Program Files\Microsoft Office\Office14\ospp.vbs" cd /d "C:\Program Files\Microsoft Office\Office14"
if exist "C:\Program Files (x86)\Microsoft Office\Office14\ospp.vbs" cd /d "C:\Program Files (x86)\Microsoft Office\Office14"
::for %%a in (4,5,6) do (if exist "%ProgramFiles%\Microsoft Office\Office1%%a\ospp.vbs" (cd "%ProgramFiles%\Microsoft Office\Office1%%a")
::if exist "%ProgramFiles(x86)%\Microsoft Office\Office1%%a\ospp.vbs" (cd "%ProgramFiles(x86)%\Microsoft Office\Office1%%a"))

set /P d= %ESC%[93mVolume to Retail (R) iülemi iáin sÅrÅm seáiniz:%ESC%[0m %ESC%[7m2013%ESC%[0m (%ESC%[92m1%ESC%[0m) //// %ESC%[7m2016%ESC%[0m (%ESC%[92m2%ESC%[0m) //// %ESC%[7m2019%ESC%[0m (%ESC%[92m3%ESC%[0m) //// %ESC%[7m2021%ESC%[0m (%ESC%[92m4%ESC%[0m) //// %ESC%[7mATLA%ESC%[0m (%ESC%[92mA%ESC%[0m) iáin sçrasçyla [%ESC%[92m1%ESC%[0m/%ESC%[92m2%ESC%[0m/%ESC%[92m3%ESC%[0m/%ESC%[92m4%ESC%[0m/%ESC%[92mA%ESC%[0m] tçkla?
if /I "%d%" EQU "1" goto :volumetoretail2013
if /I "%d%" EQU "2" goto :volumetoretail2016
if /I "%d%" EQU "3" goto :volumetoretail2019
if /I "%d%" EQU "4" goto :volumetoretail2021
if /I "%d%" EQU "A" goto :baslangic
goto :VL2Retail

:volumetoretail2013
for /f %%x in ('dir /b "..\..\Microsoft Office 15\root\Licenses\"ProPlus*R_Retail*.xrm-ms') do cscript ospp.vbs /inslic:"..\..\Microsoft Office 15\root\Licenses\%%x"
goto :baslangic

:volumetoretail2016
for /f %%x in ('dir /b "..\root\Licenses16\"ProPlus*R_Retail*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x"
goto :baslangic

:volumetoretail2019
for /f %%x in ('dir /b ..\root\Licenses16\ProPlus2019*R_Retail*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x"
goto :baslangic

:volumetoretail2021
for /f %%x in ('dir /b ..\root\Licenses16\ProPlus2021*R_Retail*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x"
goto :baslangic

echo ============================================================================&



:devamet
if exist "C:\Program Files\Microsoft Office\Office16\ospp.vbs" cd /d "C:\Program Files\Microsoft Office\Office16"
if exist "C:\Program Files (x86)\Microsoft Office\Office16\ospp.vbs" cd /d "C:\Program Files (x86)\Microsoft Office\Office16"
if exist "C:\Program Files\Microsoft Office\Office15\ospp.vbs" cd /d "C:\Program Files\Microsoft Office\Office15"
if exist "C:\Program Files (x86)\Microsoft Office\Office15\ospp.vbs" cd /d "C:\Program Files (x86)\Microsoft Office\Office15"
if exist "C:\Program Files\Microsoft Office\Office14\ospp.vbs" cd /d "C:\Program Files\Microsoft Office\Office14"
if exist "C:\Program Files (x86)\Microsoft Office\Office14\ospp.vbs" cd /d "C:\Program Files (x86)\Microsoft Office\Office14"
::for %%a in (4,5,6) do (if exist "%ProgramFiles%\Microsoft Office\Office1%%a\ospp.vbs" (cd "%ProgramFiles%\Microsoft Office\Office1%%a")
::if exist "%ProgramFiles(x86)%\Microsoft Office\Office1%%a\ospp.vbs" (cd "%ProgramFiles(x86)%\Microsoft Office\Office1%%a"))

set officedir=%cd%
::echo %officedir%



echo.&
echo ============================================================================&



:secim

if exist "C:\Program Files\Microsoft Office\Office16\ospp.vbs" cd /d "C:\Program Files\Microsoft Office\Office16"
if exist "C:\Program Files (x86)\Microsoft Office\Office16\ospp.vbs" cd /d "C:\Program Files (x86)\Microsoft Office\Office16"
if exist "C:\Program Files\Microsoft Office\Office15\ospp.vbs" cd /d "C:\Program Files\Microsoft Office\Office15"
if exist "C:\Program Files (x86)\Microsoft Office\Office15\ospp.vbs" cd /d "C:\Program Files (x86)\Microsoft Office\Office15"
if exist "C:\Program Files\Microsoft Office\Office14\ospp.vbs" cd /d "C:\Program Files\Microsoft Office\Office14"
if exist "C:\Program Files (x86)\Microsoft Office\Office14\ospp.vbs" cd /d "C:\Program Files (x86)\Microsoft Office\Office14"
::for %%a in (4,5,6) do (if exist "%ProgramFiles%\Microsoft Office\Office1%%a\ospp.vbs" (cd "%ProgramFiles%\Microsoft Office\Office1%%a")
::if exist "%ProgramFiles(x86)%\Microsoft Office\Office1%%a\ospp.vbs" (cd "%ProgramFiles(x86)%\Microsoft Office\Office1%%a"))

set officedir=%cd%
::echo %officedir%


set /P e=%ESC%[93mYÅklenmiü tÅm lisans anahtarlarç silinecektir. Silinsin mi? ------%ESC%[0m(%ESC%[101;93mSòLòNMESò ôNERòLòR%ESC%[0m)%ESC%[93m------%ESC%[0m [%ESC%[92mE%ESC%[0m/%ESC%[92mH%ESC%[0m]?
if /I "%e%" EQU "E" goto :lisanssil
if /I "%e%" EQU "H" goto :baslangic
goto :secim


:lisanssil


if exist "C:\Program Files\Microsoft Office\Office16\ospp.vbs" cd /d "C:\Program Files\Microsoft Office\Office16"
if exist "C:\Program Files (x86)\Microsoft Office\Office16\ospp.vbs" cd /d "C:\Program Files (x86)\Microsoft Office\Office16"
if exist "C:\Program Files\Microsoft Office\Office15\ospp.vbs" cd /d "C:\Program Files\Microsoft Office\Office15"
if exist "C:\Program Files (x86)\Microsoft Office\Office15\ospp.vbs" cd /d "C:\Program Files (x86)\Microsoft Office\Office15"
if exist "C:\Program Files\Microsoft Office\Office14\ospp.vbs" cd /d "C:\Program Files\Microsoft Office\Office14"
if exist "C:\Program Files (x86)\Microsoft Office\Office14\ospp.vbs" cd /d "C:\Program Files (x86)\Microsoft Office\Office14"
::for %%a in (4,5,6) do (if exist "%ProgramFiles%\Microsoft Office\Office1%%a\ospp.vbs" (cd "%ProgramFiles%\Microsoft Office\Office1%%a")
::if exist "%ProgramFiles(x86)%\Microsoft Office\Office1%%a\ospp.vbs" (cd "%ProgramFiles(x86)%\Microsoft Office\Office1%%a"))

set officedir=%cd%
::echo %officedir%


echo Deneme SÅrÅmÅ veya YÅklenmiü TÅm Lisans Anahtarlarç Siliniyor (KMS Anahtarç ve Sunucusu da dahil)...&
::Office Lisanslarçnç Sil
::for /f "tokens=8" %b in ('cscript ospp.vbs /dstatus ^| findstr /b /c:"Last 5"') do (cscript ospp.vbs /unpkey:%b)
for /f "tokens=8" %%b in ('cscript ospp.vbs /dstatus ^| findstr /b /c:"Last 5"') do (cscript ospp.vbs /unpkey:%%b)
::@For /F "Tokens=1* Delims=:" %%G In ('^""%__AppDir__%cscript.exe" "%ProgramFiles%\Microsoft Office\Office16\OSPP.VBS" /DStatus 2^> NUL ^| "%__AppDir__%find.exe" "Last 5"^"') Do @For %%I In (%%H) Do @If /I Not "XXXXX" == "%%I" "%__AppDir__%cscript.exe" "%ProgramFiles%\Microsoft Office\Office16\OSPP.VBS" /UnPKey:%%I

::KMS HOST SòL
cscript //nologo ospp.vbs /remhst
cscript //nologo %SystemRoot%\system32\slmgr.vbs /ckms

::cscript C:\Windows\System32\slmgr.vbs -ckms
::cscript %SystemRoot%\system32\slmgr.vbs -ckms
cscript //nologo %SystemRoot%\system32\slmgr.vbs /cpky
cscript //nologo %SystemRoot%\system32\slmgr.vbs /ckms

rem/||(
cscript ospp.vbs /unpkey:6MWKP
cscript ospp.vbs /unpkey:BTDRB
cscript ospp.vbs /unpkey:DRTFM
cscript ospp.vbs /unpkey:WFG99
cscript ospp.vbs /unpkey:27GXM
)
goto :baslangic


:keygir

if exist "C:\Program Files\Microsoft Office\Office16\ospp.vbs" cd /d "C:\Program Files\Microsoft Office\Office16"
if exist "C:\Program Files (x86)\Microsoft Office\Office16\ospp.vbs" cd /d "C:\Program Files (x86)\Microsoft Office\Office16"
if exist "C:\Program Files\Microsoft Office\Office15\ospp.vbs" cd /d "C:\Program Files\Microsoft Office\Office15"
if exist "C:\Program Files (x86)\Microsoft Office\Office15\ospp.vbs" cd /d "C:\Program Files (x86)\Microsoft Office\Office15"
if exist "C:\Program Files\Microsoft Office\Office14\ospp.vbs" cd /d "C:\Program Files\Microsoft Office\Office14"
if exist "C:\Program Files (x86)\Microsoft Office\Office14\ospp.vbs" cd /d "C:\Program Files (x86)\Microsoft Office\Office14"
::for %%a in (4,5,6) do (if exist "%ProgramFiles%\Microsoft Office\Office1%%a\ospp.vbs" (cd "%ProgramFiles%\Microsoft Office\Office1%%a")
::if exist "%ProgramFiles(x86)%\Microsoft Office\Office1%%a\ospp.vbs" (cd "%ProgramFiles(x86)%\Microsoft Office\Office1%%a"))

set officedir=%cd%
::echo %officedir%

set /p LicenseKey=%ESC%[101;93mLisans Anahtarç Gir:%ESC%[0m
cscript //nologo ospp.vbs /inpkey:%LicenseKey%



echo ************************************ &
echo.&
echo.&
echo ============================================================================&
echo Office Etkinleütirilecektir...&

echo ============================================================================&

:onoff
if exist "C:\Program Files\Microsoft Office\Office16\ospp.vbs" cd /d "C:\Program Files\Microsoft Office\Office16"
if exist "C:\Program Files (x86)\Microsoft Office\Office16\ospp.vbs" cd /d "C:\Program Files (x86)\Microsoft Office\Office16"
if exist "C:\Program Files\Microsoft Office\Office15\ospp.vbs" cd /d "C:\Program Files\Microsoft Office\Office15"
if exist "C:\Program Files (x86)\Microsoft Office\Office15\ospp.vbs" cd /d "C:\Program Files (x86)\Microsoft Office\Office15"
if exist "C:\Program Files\Microsoft Office\Office14\ospp.vbs" cd /d "C:\Program Files\Microsoft Office\Office14"
if exist "C:\Program Files (x86)\Microsoft Office\Office14\ospp.vbs" cd /d "C:\Program Files (x86)\Microsoft Office\Office14"
::for %%a in (4,5,6) do (if exist "%ProgramFiles%\Microsoft Office\Office1%%a\ospp.vbs" (cd "%ProgramFiles%\Microsoft Office\Office1%%a")
::if exist "%ProgramFiles(x86)%\Microsoft Office\Office1%%a\ospp.vbs" (cd "%ProgramFiles(x86)%\Microsoft Office\Office1%%a"))

set officedir=%cd%
::echo %officedir%


set /P f=%ESC%[101;93mOffice áevrimiái mi etkinleütirilsin?%ESC%[0m [%ESC%[92mE%ESC%[0m/%ESC%[92mH%ESC%[0m]?
if /I "%f%" EQU "E" goto :online
if /I "%f%" EQU "H" goto :offline
goto :onoff

:online
::Online Activation
::KMS veya MAK anahtarçyla aktivasyon geráekleütirir.
cscript //nologo ospp.vbs /act | find /i "Product activation successful" && (echo.&echo ************************************************* &echo.&choice /n /c HE /m "Aktivasyon baüarçlç...Kapatmak istiyor musunuz? (E/H)" & if errorlevel 2 goto yenidendene) || (echo Aktivasyon Baüarçsçz...! Yeniden baülançyor...) &
cscript //nologo ospp.vbs /dstatus

:tekrardene
set /P g=%ESC%[7;31mTekrar denemek ister misiniz?%ESC%[0m [%ESC%[92mE%ESC%[0m/%ESC%[92mH%ESC%[0m]
if /I "%g%" EQU "E" goto :online
if /I "%g%" EQU "H" goto :yenidendene
goto :tekrardene


echo ============================================================================&

:yenidendene
set /P h=%ESC%[101;93mAktivasyon iülemine en baütan baülansçn mç? %ESC%[0m [%ESC%[92mE%ESC%[0m/%ESC%[92mH%ESC%[0m]?
if /I "%h%" EQU "E" goto :baslangic
if /I "%h%" EQU "H" goto :yedekleme
goto :yenidendene

echo ============================================================================&

:offline
::Offline Activation
::KMS veya MAK anahtarçyla aktivasyon geráekleütirir.
::YÅkleme ID gîsterir
echo ============================================================================&
echo ============================================================================&
cscript ospp.vbs /dinstid > "%~dp0"\IID_office.txt

pushd %~dp0
start IID_office.txt
echo %ESC%[93mIID_office.txt dosyasç aáçlmçü olmalç.%ESC%[0m
echo ============================================================================&
echo %ESC%[93mInstallation ID kçsmçndan ID numarasçnç kopyalayçn. Kaza ile kapatçrsançz dosya yolu aüaßçda belirtilmiü. %ESC%[0m
echo ============================================================================&
echo %ESC%[91mIID_office.txt Yolu:%ESC%[0m %ESC%[94m %~dp0IID_office.txt %ESC%[0m
echo ============================================================================&
echo ============================================================================&

echo %ESC%[93mGîsterilen áevrimdçüç etkinleütirme iáin Kurulum Kimlißini (Installation ID) kopyalayçn.%ESC%[0m
echo %ESC%[93mOnay Kimlißi (Confirmation ID) alçp bu ekrana dînÅn.%ESC%[0m
echo ============================================================================&
echo ============================================================================&



:cidgir
echo %ESC%[101;93mNOT:%ESC%[0m %ESC%[91mCID kodu 363624231932455202567656237413441780894815599191 üu formatta olmalç. Aralarda - varsa silin.%ESC%[0m
echo ============================================================================&
echo ============================================================================&
echo ============================================================================&
echo ============================================================================&
echo ============================================================================&


pushd %officedir%

set /p CIDKey=%ESC%[101;93mCID Kodunu Gir:%ESC%[0m
cscript //nologo ospp.vbs /actcid:%CIDKey%

:lisansdurum
echo ============================================================================&

if exist "C:\Program Files\Microsoft Office\Office16\ospp.vbs" cd /d "C:\Program Files\Microsoft Office\Office16"
if exist "C:\Program Files (x86)\Microsoft Office\Office16\ospp.vbs" cd /d "C:\Program Files (x86)\Microsoft Office\Office16"
if exist "C:\Program Files\Microsoft Office\Office15\ospp.vbs" cd /d "C:\Program Files\Microsoft Office\Office15"
if exist "C:\Program Files (x86)\Microsoft Office\Office15\ospp.vbs" cd /d "C:\Program Files (x86)\Microsoft Office\Office15"
if exist "C:\Program Files\Microsoft Office\Office14\ospp.vbs" cd /d "C:\Program Files\Microsoft Office\Office14"
if exist "C:\Program Files (x86)\Microsoft Office\Office14\ospp.vbs" cd /d "C:\Program Files (x86)\Microsoft Office\Office14"
::for %%a in (4,5,6) do (if exist "%ProgramFiles%\Microsoft Office\Office1%%a\ospp.vbs" (cd "%ProgramFiles%\Microsoft Office\Office1%%a")
::if exist "%ProgramFiles(x86)%\Microsoft Office\Office1%%a\ospp.vbs" (cd "%ProgramFiles(x86)%\Microsoft Office\Office1%%a"))

set officedir=%cd%
::echo %officedir%


echo.&
echo ============================================================================&
cscript //nologo ospp.vbs /dstatus

goto :baslangic


:lisansetkinlestir
echo ============================================================================&

if exist "C:\Program Files\Microsoft Office\Office16\ospp.vbs" cd /d "C:\Program Files\Microsoft Office\Office16"
if exist "C:\Program Files (x86)\Microsoft Office\Office16\ospp.vbs" cd /d "C:\Program Files (x86)\Microsoft Office\Office16"
if exist "C:\Program Files\Microsoft Office\Office15\ospp.vbs" cd /d "C:\Program Files\Microsoft Office\Office15"
if exist "C:\Program Files (x86)\Microsoft Office\Office15\ospp.vbs" cd /d "C:\Program Files (x86)\Microsoft Office\Office15"
if exist "C:\Program Files\Microsoft Office\Office14\ospp.vbs" cd /d "C:\Program Files\Microsoft Office\Office14"
if exist "C:\Program Files (x86)\Microsoft Office\Office14\ospp.vbs" cd /d "C:\Program Files (x86)\Microsoft Office\Office14"
::for %%a in (4,5,6) do (if exist "%ProgramFiles%\Microsoft Office\Office1%%a\ospp.vbs" (cd "%ProgramFiles%\Microsoft Office\Office1%%a")
::if exist "%ProgramFiles(x86)%\Microsoft Office\Office1%%a\ospp.vbs" (cd "%ProgramFiles(x86)%\Microsoft Office\Office1%%a"))

set officedir=%cd%
::echo %officedir%


echo.&
echo ============================================================================&
cscript //nologo ospp.vbs /act | find /i "---LICENSED---" && (echo.&echo ************************************************* &echo.&choice /n /c HE /m "Aktivasyon baüarçlç...Devam edelim mi? [E/H]" & if errorlevel 2 goto yenidendenensinmi) || (echo Aktivasyon Baüarçsçz...! Yeniden baülançyor...) &
::cscript //nologo ospp.vbs /act | find /i "product activation successful" && (echo.&echo ************************************************* &echo.&choice /n /c HE /m "Aktivasyon baüarçlç...Kapatmak istiyor musunuz? (E/H)" & if errorlevel 2 exit) || (echo Aktivasyon Baüarçsçz...! Yeniden baülançyor...) &
::cscript //nologo ospp.vbs /dstatus | find /i "---LICENSED---" && (echo.&echo ************************************************* &echo.&choice /n /c HE /m "Aktivasyon baüarçlç...Kapatmak istiyor musunuz? (E/H)" & if errorlevel 2 exit) || (echo Aktivasyon Baüarçsçz...! Yeniden baülançyor...) &
cscript //nologo ospp.vbs /dstatus

:tekrardenensinmi
set /P g=%ESC%[7;31mTekrar denemek ister misiniz?%ESC%[0m [%ESC%[92mE%ESC%[0m/%ESC%[92mH%ESC%[0m]
if /I "%g%" EQU "E" goto :baslangic
if /I "%g%" EQU "H" goto :yenidendenensinmi
goto :tekrardenensinmi


echo ============================================================================&

:yenidendenensinmi
set /P i=%ESC%[93mAktivasyon iülemine en baütan baülansçn mç?%ESC%[0m [%ESC%[92mE%ESC%[0m/%ESC%[92mH%ESC%[0m]
if /I "%i%" EQU "E" goto :baslangic
if /I "%i%" EQU "H" goto :yedekleme
goto :yenidendenensinmi

:yedekleme
set /P k=%ESC%[93mYedek Alçnsçn mç?%ESC%[0m [%ESC%[92mE%ESC%[0m/%ESC%[92mH%ESC%[0m]?
if /I "%k%" EQU "E" goto :yedekleniyor
if /I "%k%" EQU "H" goto :baslangic
goto :yedekleme

:yedekleniyor
if exist "c:\ofis_yedek" rd /s /q "c:\ofis_yedek"
xcopy /i /e "C:\Windows\System32\spp" "c:\ofis_yedek"
echo %ESC%[93mOffice baüarçyla yedeklendi. Yedekleme dosyalarçnç gÅvenilir bir konuma taüçyçnçz.%ESC%[0m
echo %ESC%[93mYedekleme Konumu: "c:\ofis_yedek"%ESC%[0m


goto :baslangic

:bitis


:devametme
Echo %ESC%[101;93mÄçkçü Yapçldç...%ESC%[0m
timeout 5
exit

echo.&
echo ============================================================================&




:setESC
for /F "tokens=1,2 delims=#" %%a in ('"prompt #$H#$E# & echo on & for %%b in (1) do rem"') do (
  set ESC=%%b
  exit /B 0
)
exit /B 0




::BURADA KOD BòTTò
::BURADA KOD BòTTò