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
setlocal EnableDelayedExpansion



::for %%a in (4,5,6) do (if exist "%ProgramFiles%\Microsoft Office\Office1%%a\ospp.vbs" (cd /d "%ProgramFiles%\Microsoft Office\Office1%%a")
::if exist "%ProgramFiles(x86)%\Microsoft Office\Office1%%a\ospp.vbs" (cd /d "%ProgramFiles(x86)%\Microsoft Office\Office1%%a")else (echo YÅklÅ Office yazçlçmç yoktur ya da hatalç kurulmuütur. & goto :end_loop)) & cls

::echo %cd%
::pause 


::for %%a in (4,5,6) do (if not exist "%ProgramFiles%\Microsoft Office\Office1%%a\ospp.vbs" (echo YÅklÅ Office yazçlçmç yoktur ya da hatalç kurulmuütur. & goto :end_point)
::if not exist "%ProgramFiles(x86)%\Microsoft Office\Office1%%a\ospp.vbs" (echo YÅklÅ Office yazçlçmç yoktur ya da hatalç kurulmuütur. & goto :end_point)) & cls


::for %%b in (4,5,6) do (if exist "%ProgramFiles%\Microsoft Office\Office1%%b\ospp.vbs" (cd /d "%ProgramFiles%\Microsoft Office\Office1%%b")
::if exist "%ProgramFiles(x86)%\Microsoft Office\Office1%%b\ospp.vbs" (cd /d "%ProgramFiles(x86)%\Microsoft Office\Office1%%b")) & cls


if exist "C:\Program Files\Microsoft Office\Office16\ospp.vbs" (cd /d "C:\Program Files\Microsoft Office\Office16" & goto :IIDstart)
if exist "C:\Program Files (x86)\Microsoft Office\Office16\ospp.vbs" (cd /d "C:\Program Files (x86)\Microsoft Office\Office16" & goto :IIDstart)
if exist "C:\Program Files\Microsoft Office\Office15\ospp.vbs" (cd /d "C:\Program Files\Microsoft Office\Office15" & goto :IIDstart)
if exist "C:\Program Files (x86)\Microsoft Office\Office15\ospp.vbs" (cd /d "C:\Program Files (x86)\Microsoft Office\Office15" & goto :IIDstart)
if exist "C:\Program Files\Microsoft Office\Office14\ospp.vbs" (cd /d "C:\Program Files\Microsoft Office\Office14" & goto :IIDstart)
if exist "C:\Program Files (x86)\Microsoft Office\Office14\ospp.vbs" (cd /d "C:\Program Files (x86)\Microsoft Office\Office14" & goto :IIDstart)
echo YÅklÅ Office yazçlçmç yoktur ya da hatalç kurulmuütur. & goto :cid_point

set officedir=%cd%
::echo %officedir%

:IIDstart
cscript ospp.vbs /dinstid > "%~dp0IID_office.txt"
::cscript ospp.vbs /dinstid > "%~dp0"\IID_office.txt

rem pushd "%~dp0"

cd /d %~dp0

:: IID_office.txt dosyasçndaki Installation ID for: satçrlarçnç listeleyin ve menÅ oluüturun


:: IID_office.txt dosyasçnda "Installation" satçrç var mç?
find "Installation" "IID_office.txt" >nul || (echo YÅklÅ anahtar bulunamadç. & goto :cid_point)


echo YÅkleme Kimlikleri aüaßçda listelenmiütir.
echo.
set "i=0"
for /f "usebackq tokens=1* delims=: " %%a in ("IID_office.txt") do (
if "%%a"=="Installation" (
set /a i+=1
echo !i! - %%b
)
)
echo.
:: Kullançcçdan seáim yapmasçnç isteyin
set /p "secim=LÅtfen seáim yapçnçz: "

:: MenÅdeki seáimi IID deßiükenine atayçn
set "i=0"
for /f "usebackq tokens=1* delims=: " %%a in ("IID_office.txt") do (
if "%%a"=="Installation" (
set /a i+=1
if "!i!"=="%secim%" set "IID=%%b"
)
)

:: IID kodunun baüçndaki metinleri siler
::set "IID=%IID:*for: =%"
::set "IID=%IID:, =%"
::set "IID=%IID:*edition: =%"

set "IID=%IID:*edition: =%"


if defined IID (
echo IID kodu: %IID%
echo IID kodu: %IID% > IID_office.txt
goto :cid_point
) else (
echo YÅklÅ anahtar bulunamadç.
goto :cid_point
)

:cid_point

::curl -L "https://getconfirmationid.com/ajax/cidms_api?iids=%IID%&username=trogiup24h&password=PHO" > CID_windows.txt



rem URL ve giriü bilgilerini tançmlayçn
set "url=https://getconfirmationid.com/ajax/cidms_api?iids=%IID%&username=trogiup24h&password=PHO"
set "username=trogiup24h"
set "password=PHO"

rem JSON verisini indirin ve dosyaya kaydedin
curl -u "%username%:%password%" -o response.json "%url%"


for /f "tokens=5 delims=:," %%j in ('find "result" response.json') do set "CIDDurum=%%j"
set "CIDDurum=%CIDDurum:"=%"
::echo %CIDDurum%
set "CID=null"

::echo %CID%
if "%CIDDurum%" == "Successfully" (
  for /f "tokens=23 delims=:," %%c in ('find "confirmation_id_no_dash" response.json') do set "CID=%%c"

  
) else (
  echo CID kodu alçnamadç. Girilen YÅkleme Kimlißine ait Lisans anahtarç geáersiz veya hatalç olabilir.
)
rem CID kodundan " iüaretini siler
set "CID=%CID:"=%"

rem "CID" deßiükeni ekrana yazdçrçlçr
echo CID kodu: %CID%
echo CID kodu: %CID% > CID_office.txt


::Onaylama ID (CID)yÅkler ve aktif eder. 
::CID anahtarçnç ekle
::set /p CIDkey=CID Gir:
::set "CIDkey=%CID%"
:CIDGir
if "%CID%" == "" (
  set /p CIDkey=CID Gir:
) else if /i "%CID%" == "null" (
  set /p CIDkey=CID Gir:
) else (
  set "CIDkey=%CID%"
)


pushd %officedir%

set /p CIDKey=%ESC%[101;93mCID Kodunu Gir:%ESC%[0m
cscript //nologo ospp.vbs /actcid:%CIDKey%

::Offline Activation
::KMS veya MAK anahtarçyla aktivasyon geráekleütirir.
cscript //nologo ospp.vbs /act | find /i "---LICENSED---" && (echo.&echo ************************************************* &echo.&choice /n /c HE /m "Aktivasyon baüarçlç...Devam edelim mi? [E/H]" & if errorlevel 2 goto yenidendenensinmi) || (echo Aktivasyon Baüarçsçz...!) &
cscript //nologo ospp.vbs /dstatus
::cscript //nologo ospp.vbs /act | find /i "product activation successful" && (echo.&echo ************************************************* &echo.&choice /n /c HE /m "Aktivasyon baüarçlç...Kapatmak istiyor musunuz? (E/H)" & if errorlevel 2 exit) || (echo Aktivasyon Baüarçsçz...! Yeniden baülançyor...) &
::cscript //nologo ospp.vbs /dstatus | find /i "---LICENSED---" && (echo.&echo ************************************************* &echo.&choice /n /c HE /m "Aktivasyon baüarçlç...Kapatmak istiyor musunuz? (E/H)" & if errorlevel 2 exit) || (echo Aktivasyon Baüarçsçz...! Yeniden baülançyor...) &

:secim-tb
set /P t=Tekrar dene/Baüa Dîn [T/B]?
if /I "%t%" EQU "T" goto :CIDGir
if /I "%t%" EQU "B" goto :devametme
goto :secim-tb

:devametme
pause
