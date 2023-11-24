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
call :setESC

:: color help <<<<<<<<<<<<<<<<<<b�t�n renk bilgileri i�in cmd ekran�na yaz
:: color [arkaplanrengi][yaz�rengi]
::color 1F
mode con lines=20 cols=150




::for %%a in (4,5,6) do (if exist "%ProgramFiles%\Microsoft Office\Office1%%a\ospp.vbs" (cd /d "%ProgramFiles%\Microsoft Office\Office1%%a")
::if exist "%ProgramFiles(x86)%\Microsoft Office\Office1%%a\ospp.vbs" (cd /d "%ProgramFiles(x86)%\Microsoft Office\Office1%%a")else (echo Y�kl� Office yaz�l�m� yoktur ya da hatal� kurulmu�tur. & goto :end_loop)) & cls

::echo %cd%
::pause 


::for %%a in (4,5,6) do (if not exist "%ProgramFiles%\Microsoft Office\Office1%%a\ospp.vbs" (echo Y�kl� Office yaz�l�m� yoktur ya da hatal� kurulmu�tur. & goto :end_point)
::if not exist "%ProgramFiles(x86)%\Microsoft Office\Office1%%a\ospp.vbs" (echo Y�kl� Office yaz�l�m� yoktur ya da hatal� kurulmu�tur. & goto :end_point)) & cls


::for %%b in (4,5,6) do (if exist "%ProgramFiles%\Microsoft Office\Office1%%b\ospp.vbs" (cd /d "%ProgramFiles%\Microsoft Office\Office1%%b")
::if exist "%ProgramFiles(x86)%\Microsoft Office\Office1%%b\ospp.vbs" (cd /d "%ProgramFiles(x86)%\Microsoft Office\Office1%%b")) & cls


if exist "C:\Program Files\Microsoft Office\Office16\ospp.vbs" (cd /d "C:\Program Files\Microsoft Office\Office16" & goto :IIDstart)
if exist "C:\Program Files (x86)\Microsoft Office\Office16\ospp.vbs" (cd /d "C:\Program Files (x86)\Microsoft Office\Office16" & goto :IIDstart)
if exist "C:\Program Files\Microsoft Office\Office15\ospp.vbs" (cd /d "C:\Program Files\Microsoft Office\Office15" & goto :IIDstart)
if exist "C:\Program Files (x86)\Microsoft Office\Office15\ospp.vbs" (cd /d "C:\Program Files (x86)\Microsoft Office\Office15" & goto :IIDstart)
if exist "C:\Program Files\Microsoft Office\Office14\ospp.vbs" (cd /d "C:\Program Files\Microsoft Office\Office14" & goto :IIDstart)
if exist "C:\Program Files (x86)\Microsoft Office\Office14\ospp.vbs" (cd /d "C:\Program Files (x86)\Microsoft Office\Office14" & goto :IIDstart)
echo Y�kl� Office yaz�l�m� yoktur ya da hatal� kurulmu�tur. & goto :cid_point

set officedir=%cd%
::echo %officedir%

:IIDstart
cscript ospp.vbs /dinstid > "%~dp0IID_office.txt"
::cscript ospp.vbs /dinstid > "%~dp0"\IID_office.txt

rem pushd "%~dp0"

cd /d %~dp0

:: IID_office.txt dosyas�ndaki Installation ID for: sat�rlar�n� listeleyin ve men� olu�turun


:: IID_office.txt dosyas�nda "Installation" sat�r� var m�?
find "Installation" "IID_office.txt" >nul || (echo Y�kl� anahtar bulunamad�. & goto :cid_point)


echo Y�kleme Kimlikleri a�a��da listelenmi�tir.
echo.
set "i=0"
for /f "usebackq tokens=1* delims=: " %%a in ("IID_office.txt") do (
if "%%a"=="Installation" (
set /a i+=1
echo !i! - %%b
)
)
echo.
:: Kullan�c�dan se�im yapmas�n� isteyin
set /p "secim=L�tfen se�im yap�n�z: "

:: Men�deki se�imi IID de�i�kenine atay�n
set "i=0"
for /f "usebackq tokens=1* delims=: " %%a in ("IID_office.txt") do (
if "%%a"=="Installation" (
set /a i+=1
if "!i!"=="%secim%" set "IID=%%b"
)
)

:: IID kodunun ba��ndaki metinleri siler
::set "IID=%IID:*for: =%"
::set "IID=%IID:, =%"
::set "IID=%IID:*edition: =%"

set "IID=%IID:*edition: =%"


if defined IID (
echo IID kodu: %IID%
echo IID kodu: %IID% > IID_office.txt
goto :cid_point
) else (
echo Y�kl� anahtar bulunamad�.
goto :cid_point
)

:cid_point

::curl -L "https://getconfirmationid.com/ajax/cidms_api?iids=%IID%&username=trogiup24h&password=PHO" > CID_windows.txt



rem URL ve giri� bilgilerini tan�mlay�n
set "url=https://pidkey.com/ajax/cidms_api?iids=%IID%&username=trogiup24h&password=PHO"
::set "url=https://getconfirmationid.com/ajax/cidms_api?iids=%IID%&username=trogiup24h&password=PHO"
set "username=trogiup24h"
set "password=PHO"
::Alternatif Yollar.
::https://pidkey.com/apis
::https://pidkey.com/ajax/cidms_api?iids=%IID%&username=trogiup24h&password=PHO
::Sayfa cevab�: {"short_result":"Confirmation ID (CID):\r\n033815-357101-860630-218090-894495-243386-132221-970210","result":"Successfully","typeiid":1,"is_chat_ms":0,"confirmationid":"033815-357101-860630-218090-894495-243386-132221-970210","confirmationid_chat_ms":"","have_cid":1,"ultimate_cid":null,"ultimate_have_cid":-1,"confirmation_id_with_dash":"033815-357101-860630-218090-894495-243386-132221-970210","confirmation_id_no_dash":"033815357101860630218090894495243386132221970210","error_executing":null,"had_occurred":0}

::https://khoatoantin.com/apis
::https://khoatoantin.com/ajax/cidms_api?iids=%IID%&username=trogiup24h&password=PHO
::Sayfa cevab�: {"short_result":"Confirmation ID (CID):\r\n033815-357101-860630-218090-894495-243386-132221-970210","result":"Successfully","typeiid":1,"is_chat_ms":0,"confirmationid":"033815-357101-860630-218090-894495-243386-132221-970210","confirmationid_chat_ms":"","have_cid":1,"ultimate_cid":null,"ultimate_have_cid":-1,"confirmation_id_with_dash":"033815-357101-860630-218090-894495-243386-132221-970210","confirmation_id_no_dash":"033815357101860630218090894495243386132221970210","error_executing":null,"had_occurred":0}


::BUNDA FARKLI KOD KULLANILMALI ALTTAK� ��E YARAMAZ. response.json yok yani.
::https://pidkey.vip/GetCID.aspx?id=GetCID&pass=GetCID&iid=%IID%
::Sayfa Kayna��: 033815-357101-860630-218090-894495-243386-132221-970210

::https://kichhoat24h.com/apis
::https://kichhoat24h.com/user-api/get-cid?iid=%IID%&price=0.5&token=[token_id]&send_to_email=[send_to_email]&callback_url=[callback_url]
::https://kichhoat24h.com/user-api/get-cid?iid=%IID%&price=0.5&token=[token_id]


::BUNDA DA FARKLI KOD KULLANILMALI ALTTAK� ��E YARAMAZ. c_cid k�sm�n�n alttaki kodda confirmation_id_no_dash ile de�i�tirilmesi gerekiyor.(Belki daha fazlas�.)
::https://getcid.xyz/api
::https://bs.getcid.xyz/webapi/get-cid/?autocall=1&iid=%IID%
::iid (mandatory): Your installation ID. You can send it with or without hyphens.
::autocall (optional, default 1): You can set autocall to 0 if you just want to get information without making a call.
::onlycid (optional, default 0): You can set onlycid to 1 if you want the user information not to be included.
::Valid examples:
::https://bs.getcid.xyz/webapi/get-cid/?autocall=0&iid=%IID%
::https://bs.getcid.xyz/webapi/get-cid/?onlycid=1&iid=%IID%
::https://bs.getcid.xyz/webapi/get-cid/?token=************************************************&iid=%IID%
::      {
::        "cid": "435504-697485-138242-433182-497552-083534-176493-978266",
::        "c_cid": "435504697485138242433182497552083534176493978266",
::        "user_data": {
::          "credit": 100.0,
::          "autocall_price": 0.5
::        },
::      }



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
  echo CID kodu al�namad�. Girilen Y�kleme Kimli�ine ait Lisans anahtar� ge�ersiz veya hatal� olabilir.
)
rem CID kodundan " i�aretini siler
set "CID=%CID:"=%"

rem "CID" de�i�keni ekrana yazd�r�l�r
echo CID kodu: %CID%
echo CID kodu: %CID% > CID_office.txt


::Onaylama ID (CID)y�kler ve aktif eder. 
::CID anahtar�n� ekle
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


if exist "C:\Program Files\Microsoft Office\Office16\ospp.vbs" cd /d "C:\Program Files\Microsoft Office\Office16"
if exist "C:\Program Files (x86)\Microsoft Office\Office16\ospp.vbs" cd /d "C:\Program Files (x86)\Microsoft Office\Office16"
if exist "C:\Program Files\Microsoft Office\Office15\ospp.vbs" cd /d "C:\Program Files\Microsoft Office\Office15"
if exist "C:\Program Files (x86)\Microsoft Office\Office15\ospp.vbs" cd /d "C:\Program Files (x86)\Microsoft Office\Office15"
if exist "C:\Program Files\Microsoft Office\Office14\ospp.vbs" cd /d "C:\Program Files\Microsoft Office\Office14"
if exist "C:\Program Files (x86)\Microsoft Office\Office14\ospp.vbs" cd /d "C:\Program Files (x86)\Microsoft Office\Office14"


echo.
echo.
echo.

::set /p CIDKey=%ESC%[101;93mCID Kodunu Gir:%ESC%[0m
cscript //nologo ospp.vbs /actcid:%CID%

::Offline Activation
::KMS veya MAK anahtar�yla aktivasyon ger�ekle�tirir.
cscript //nologo ospp.vbs /act | find /i "---LICENSED---" && (echo.&echo ************************************************* &echo.&choice /n /c HE /m "Aktivasyon ba�ar�l�...Devam edelim mi? [E/H]" & if errorlevel 2 goto yenidendenensinmi) || (echo Aktivasyon Ba�ar�s�z...!) &
cscript //nologo ospp.vbs /dstatus
::cscript //nologo ospp.vbs /act | find /i "product activation successful" && (echo.&echo ************************************************* &echo.&choice /n /c HE /m "Aktivasyon ba�ar�l�...Kapatmak istiyor musunuz? (E/H)" & if errorlevel 2 exit) || (echo Aktivasyon Ba�ar�s�z...! Yeniden ba�lan�yor...) &
::cscript //nologo ospp.vbs /dstatus | find /i "---LICENSED---" && (echo.&echo ************************************************* &echo.&choice /n /c HE /m "Aktivasyon ba�ar�l�...Kapatmak istiyor musunuz? (E/H)" & if errorlevel 2 exit) || (echo Aktivasyon Ba�ar�s�z...! Yeniden ba�lan�yor...) &

:secim-tb
set /P t=Tekrar dene/Ba�a D�n [T/B]?
if /I "%t%" EQU "T" goto :CIDGir
if /I "%t%" EQU "B" goto :devametme
goto :secim-tb

:devametme
pause


:setESC
for /F "tokens=1,2 delims=#" %%a in ('"prompt #$H#$E# & echo on & for %%b in (1) do rem"') do (
  set ESC=%%b
  exit /B 0
)
exit /B 0
