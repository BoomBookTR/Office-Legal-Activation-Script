![image](https://user-images.githubusercontent.com/10184695/183362497-925c9410-fdcd-420e-a170-3bb6d86e0d9c.png)


# Nedir?

Bu Batch Script, Office aktivasyonunun kod ile kolayca yapılmasını sağlar. 
Office 2013, 2016, 2019, 2021 tüm sürümlerini destekler. Office 365 sürümü öğretmen ve öğrencilere (uygun edu.tr ve k12.tr gibi mail adresi olanlara) ücretsiz. Böyle bir imkanınız varsa bu konuyla ilgilenmeseniz de olur ama genel bilgi için faydalı olacaktır.



# ÖZELLİKLER

*Retail to Volume (Retail ISO dosyasıyla yapılan kurulumlarda Volume:MAK anahtarların etkinleştirilmesi için gereklidir)

*Volume to Retail

*Mevcut lisans anahtarlarını silme

*Lisans anahtarı ekleme

*Çevrimiçi aktivasyon

*Çevrimdışı aktivasyon

*Installation ID (IID) - Yükleme Kimliği alır.

*Confirmation ID (CID) - Onay Kimliği alır.

*Lisans yedekleme (MAK lisanslar için birebir ki sonradan aynı lisansı kullanmak mümkün sorunsuz bir şekilde.)

# Detaylar
Bizim burada kullanacağımız etkinleştirme yöntemi çevrimiçi ve çevrimdışı aktivasyon yönteminin ikisinin de kod ile yapılabilmesi üzerine olacaktır.  
Bununla ilgili Microsoft Windows aktivasyon dökümanı şuradan detaylı bilgi edinilebilir ve kullanılabilir. Link: https://docs.microsoft.com/en-us/deployoffice/vlactivation/tools-to-manage-volume-activation-of-office


# Temel Kodlar

Script üzerinde kullanılan 5 temel kod var.
Bunlar hakkında bilgiye üstteki linkten ulaşabilirsiniz. Kodlar şunlar;

`for /f "tokens=8" %b in ('cscript ospp.vbs /dstatus ^| findstr /b /c:"Last 5"') do (cscript ospp.vbs /unpkey:%b)` kodu `cscript //nologo ospp.vbs /unpkey:(Lisans kodunun son 5 hanesi)` şeklinde çalışan lisansları topluca silme kodu.

`cscript //nologo ospp.vbs /inpkey:(ABCDE-ABCDE-ABCDE-ABCDE-ABCDE şeklinde lisans anahtarı)`

`cscript //nologo ospp.vbs /dinstid`

`cscript //nologo ospp.vbs /actcid:(000000000000000000000000000000000000000000000000 şekline Onay Kimliği kodu)`

`cscript //nologo ospp.vbs /dstatus`


# KULLANIM
Bat komut dosyasını çalıştırın ve yönergeleri takip edin.

Kullanım Videosu: (Önce çevrimiçi aktivasyonu deneyebilirsiniz ki çevrimiçi aktivasyonda CID istemez. Olmazsa çevrimdışı aktivasyonu deneyiniz.)



--------------------------------------------------------------------------------------------------------
[![Legal Office Telefon Aktivasyonu](https://yt-embed.herokuapp.com/embed?v=m05XuXU58yw)](https://www.youtube.com/watch?v=m05XuXU58yw "Legal Office Telefon Aktivasyonu")
--------------------------------------------------------------------------------------------------------


# Genel Bilgilendirme
Altta arayüz üzerinden nasıl aktivasyon yapılabileceği ile ilgili video mevcuttur.  
MAK gibi çevrimiçi etkinleştirmeye müsait olan lisans anahtarları ile aktivasyon kolayca yapılabilir. Retail ISO kurulumunda Retail anahtarlar kullanılabilir, MAK anahtarlar çalışmayacaktır ki bu hatayı aşmak için Retail2Volume işlemi yapmak gerekiyor. 


[![Legal Office Arayüz Aktivasyonu](https://yt-embed.live/embed?v=Ni3vSDHdd2I)](https://youtu.be/Ni3vSDHdd2I?t=162 "Legal Office Arayüz Aktivasyonu")

Diğer bir aktivasyon yöntemi ise çevrimdışı (internet bağlantısı olmama durumunda vb. kullanılabilen bir aktivasyon yöntemi) aktivasyon yöntemidir.
RETAIL anahtarlar bu şekilde kolayca aktif edilebilir. Ücretsiz Microsoft Etkinleştirme Telefon Hattı: 0(800) 211 3939

[![Legal Office Telefon Aktivasyonu](https://yt-embed.live/embed?v=H7cJOp2L5FU)](https://www.youtube.com/watch?v=H7cJOp2L5FU "Legal Office Telefon Aktivasyonu")



# ÖNEMLİ NOTLAR

![image](https://user-images.githubusercontent.com/10184695/183368431-07979414-1b67-491d-ac30-34baceca9c1e.png)


![image](https://user-images.githubusercontent.com/10184695/183439783-d679cc81-4424-4bab-a1fe-52085df680b7.png)


Telefon etkinleştirmesi ile Microsoft aranır ve ekranda görülen Yükleme Kimliği (IID) kodu telefondan girilir ve Onay Kimliği (CID) kodu alınır ve ekrana girildiğinde aktivasyon gerçekleşir.

Arayüz üzerinden telefon etkinleştirmesine ulaşım ve kullanımı; https://support.microsoft.com/tr-tr/office/etkinle%C5%9Ftirme-sihirbaz%C4%B1n%C4%B1-kullanarak-office-i-etkinle%C5%9Ftirme-1144e0de-e849-496e-8e33-ed6fb1b34202#bkmk_phone

Aşağıda bir satıcının bu yöntem ile aktivasyonun nasıl yapıldığını anlattığı videosunu görüyorsunuz.

[![Legal Office Arayüz Aktivasyonu](https://yt-embed.herokuapp.com/embed?v=HASfbIpboxQ)](https://www.youtube.com/watch?v=HASfbIpboxQ "Legal Office Arayüz Aktivasyonu")

Bu da başka bir satıcının Telefon etkinleştirmesi yönlendirmesi.

![image](https://user-images.githubusercontent.com/10184695/183441832-7ca9c86e-956b-4d8c-9097-d2e95a42bac1.png)

Bu da başkası.

[![Legal Office Arayüz Aktivasyonu](https://yt-embed.herokuapp.com/embed?v=zdF5HO7xy8g)](https://www.youtube.com/watch?v=zdF5HO7xy8g "Legal Office Arayüz Aktivasyonu")




# Lisans anahtarı bulmak
Lisans anahtarınız yoksa bir tane edinmeniz gerekiyor. İnternet siteleri ucuza satıyorlar ve bence buna gerek de yok.
Telegram grupları (https://t.me/windows_office_etkinlestir), bazı internet siteleri üzerinden paylaşılan anahtarlar her gün yayılıyor ve emin olun Microsoft'un çok pahalı sattığı anahtarları satan bu satıcılardan aldıktan sonra Google üzerinde arama yaparsanız nette olduğunu göreceksiniz. Düşmediyse bile 1-2 güne düşer. Neyse satın aldınız veya belirttiğim şekilde buldunuz ki bu kodların çalışıp çalışmadığını öğrenmeniz gerekiyor. Bunu kontrol etmek için ise aşağıdaki siteleri ve programları kullanabilirsiniz.

# 🔥Lisans Durumunu Nasıl Öğrenirim?🔥

Lisans kodunu kullanmadan önce durumunu kendiniz kontrol edebilirsiniz. Sorgulamadaki hata kodları (bknz: /hatakodlari) lisans anahtarı durumunu gösterir. 

🚩PID (Product ID) Anahtarı Durumu Sorgulama Siteleri

🔗https://khoatoantin.com/pidms ( Username: trogiup24h Password: PHO veya Username: HQCNTH - Password: MIGOI )

🔗https://khoatoantin.com/office365 (Office365 Hesap durumu sorgulamak için) ( Username: trogiup24h Password: PHO veya Username: HQCNTH - Password: MIGOI )

🔗https://pidkey.top

🔗https://doonoi.top

🔗https://webact.185.hk/mskey.php (WeChat isteyebilir)

🔗https://dbmer.com/checkkey


🚩PID Sorgulama Programları

🔗http://khoatoantin.com/products/cidms.zip ( Username: trogiup24h Password: PHO veya Username: HQCNTH - Password: MIGOI )

🔗https://github.com/laomms/PidKeyTool

🔗https://github.com/Ja7ad/PIDChecker

Elinizdeki lisans anahtarının durumunu program veya siteye girip sorgulatınız. Size şu gibi bir çıktı verecek ve burada önemli olan ErrorCode:0xC004C008 ve MAK anahtarlarda Remaining kısmı.

MAK keyleriyle üstte belirttiğim gibi arayüz üzerinden kolayca etkinleştirebilirsiniz ancak diğer lisans anahtarları arayüzde hata verecektir. Bu sebeple telefon ile aktivasyon gerçekleştirilir.

`ProductKey:9KHK2-PN74D-HYAG4-P4JM8-CNTWK`

`Description:Office21_ProPlus2021MSDNR_Retail`

`ErrorCode:0xC004C008`

`CheckedTime:2022-08-07 10:10:04 AM(UTC+03:00)`

-----------------------------------------------------
-----------------------------------------------------


`Key: CV2D9-NFMX6-FTACV-Y5B9D-T6C3D`

`Description: Office19_ProPlus2019VL_MAK_AE`

`Remaining: 19862`

`Time: 8/8/2022 2:37:59 PM (GMT+7)`

-----------------------------------------------------
-----------------------------------------------------


# 🔥BİLİNMESİ GEREKEN HATA KODLARI🔥

`0xC004C008` ve `0xC004C020`
CID kodu alınarak telefon etkinleştirmesi ile aktivasyon yapmak mümkündür.

`0xC004C004`
Sahte ürün anahtarıdır ve kullanılamaz.

`0xC004C060` ve `0xC004C003`
Geçerli ürün anahtarı değildir ve kullanılamaz.



# 🔥CID kodunu nasıl alabilirim?🔥

Microsoft'u ücretsiz arayarak alabilirsiniz. Ücretsiz Microsoft hattı: 0(800) 211 3939

![image](https://user-images.githubusercontent.com/10184695/183365865-c4138831-fe17-4538-8cfb-c688855f60fa.png)

CID kodunu siteler aracılığıyla da almak mümkün. Aşağıda belirtilen siteler IID kodunu giriğinizde size CID kodunu verecektir.

🚩CID alma siteleri

🔗https://0xc004c008.com ( Username: trogiup24h Password: PHO )

🔗https://khoatoantin.com/cidms ( Username: trogiup24h Password: PHO )

🔗https://doonoi.top/GenCID.aspx

🔗https://pidkey.top/GenCID.aspx

🔗https://pintipin.com/aktivasyon

🔗https://pintiaktivasyon.com

🔗https://microsoft.gointeract.io/interact/index?interaction=1461173234028-3884f8602eccbe259104553afa8415434b4581-05d1&accountId=microsoft&loadFrom=CDN&appkey=196de13c-e946-4531-98f6-2719ec8405ce&Language=English&name=pana&CountryCode=en-US&Click%20To%20Call%20Caller%20Id=+16265860337&startedFromSmsToken=e9dmQhQ&dnis=1&token=UEuUbt

🔗https://getcid.info

🔗https://webact.185.hk


# 🔥Office Seçmeli Kurulum🔥
 
 Office Deployment Tool ile İstenilen Office Sürümün Kurulması


🚩Office Seçmeli Kurulum için gerekli olan aracı indirin.

Office Deployment Tool indirme linki:

🔗 https://www.microsoft.com/en-us/download/details.aspx?id=49117

Office 2013 Deployment Tool:

🔗 https://www.microsoft.com/en-us/download/details.aspx?id=36778


Kurulum şu şekilde:

✔️Önce exe içerisindeki dosyalar çıkarılır.

✔️İçerisindeki config dosyasının adını config.xml olarak değiştirin (değiştirmeseniz de olur ancak alt satırdaki kodda config adını doğru girmelisiniz) ve kendi isteğimize göre düzenlenir. https://config.office.com/deploymentsettings adresinden online olarak da config.xml dosyanızı oluşturabilirsiniz.

✔️Bulunduğunuz dizinde cmd komut satırı açıp setup.exe /configure config.xml kodunu çalıştırın ve kurulumun bitmesini bekleyin.

🚩Önemli Bilgi ve Yardım Sayfaları

🔗 https://docs.microsoft.com/en-us/microsoft-365/troubleshoot/installation/product-ids-supported-office-deployment-click-to-run

🔗 https://docs.microsoft.com/en-us/office365/troubleshoot/installation/product-ids-supported-office-deployment-click-to-run

🔗 https://www.heidoc.net/joomla/technology-science/microsoft/79-create-an-office-2013,-2016-and-365-offline-installer-with-the-office-deployment-tool

🔗 https://docs.microsoft.com/tr-tr/deployoffice/overview-deploying-languages-microsoft-365-apps#install-the-same-languages-as-the-operating-system



# 🔥 Office ProPlus İndirme Linkleri 🔥


Not: RETAIL olduğu için MAK lisans aktivasyonu yapmak için Retail to Volume işlemi yapmak gerekiyor.

Office 365 ProPlus:
🔗 http://officecdn.microsoft.com/pr/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/tr-tr/O365ProPlusRetail.img

Office 2013 ProPlus:
🔗 https://officecdn.microsoft.com/pr/39168D7E-077B-48E7-872C-B232C3E72675/media/tr-TR/ProfessionalRetail.img

Office 2016 ProPlus:
🔗 https://officecdn.microsoft.com/pr/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/tr-TR/ProPlusRetail.img

Office 2019 ProPlus:
🔗 https://officecdn.microsoft.com/pr/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/tr-tr/ProPlus2019Retail.img

Office 2021 ProPlus:
🔗 https://officecdn.microsoft.com/pr/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/tr-tr/ProPlus2021Retail.img



# 🔥Lisans Yedeği Nasıl Alınır?🔥

Scriptin son aşamasında yedek alınması ile ilgili soru soracaktır. İstenilirse alttaki yönergeye göre manuel olarak da yedek alınabilir.

Etkinleştirme bilgileri yerel diskinizde depolanır. Ancak, sistemin yeniden yüklenmesi veya diğer bazı eylemler etkinleştirme bilgilerini silecektir. 
Özellikle MAK aktivasyonu yapıldığı durumlarda format sonrası yeniden etkinleştirmek için aynı anahtarı kullandığınızda kullanım hakkı kalmadığı için yeniden aktivasyon gerçekleşmez. 

Sürekli lisans anahtarı aramamak için lisans yedeğinin alınması önerilir. Bu işlemi ücretsiz yazılımlarla yapabileceğiniz gibi kendiniz de yapabilirsiniz.


🚩Lisans yedeği şu şekilde alınır:

✔️`C:\Windows\System32\spp` klasörünü güvenli bir alana yedekleyin.


🚩Lisans yedeği şu şekilde geri yüklenir:

✔️Geri yükleme yapılırken Komut Satırı (cmd) ekranı yönetici olarak açılır. 

✔️`net stop sppsvc` komutunu göndererek Yazılım Koruması (Software Protection Platform) servisini kapatın.

✔️Yeniden kurulum sonrası yedeklenmiş store klasörünü `C:\Windows\System32\spp` alanına yapıştırıp üzerine kopyalayın.

✔️`net start sppsvc` komutunu göndererek Yazılım Koruması (Software Protection Platform) servisini açın.

✔️Bu işlem sonrası yazılım çevrimdışı olarak kendini etkinleştirecektir.
