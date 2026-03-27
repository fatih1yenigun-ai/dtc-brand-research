"""
yuvacim_v3_batch1.py
Post-2018 ecommerce-native DTC brands — HOME / BEDDING / ORGANIZER niche
6 categories, 200+ brands total
All descriptions in Turkish (brand names & URLs in English)
"""

BATCH_DATA = {

    # ─────────────────────────────────────────────────────────────────
    # 1. YATAK ÇARŞAFI & NEVRESİM  (Sheets & Bedding Sets)
    # ─────────────────────────────────────────────────────────────────
    "Yatak Çarşafı & Nevresim": [
        ("Miracle Brand", "miraclebrand.com", "Antimikrobiyal Çarşaf",
         "Gümüş iyonlu antimikrobiyal kumaş; 3x daha az yıkama gerektirir; NASA ilhamlı teknoloji; self-cleaning iddiası",
         "2018'de kuruldu; Shark Tank'ta yatırım aldı; DTC olarak büyüdü; sosyal medya reklamlarıyla viral; aylık $2M+ gelir"),

        ("Sheets & Giggles", "sheetsandgiggles.com", "Ökaliptüs Çarşaf",
         "100% ökaliptüs liyosel; organik ve sürdürülebilir; pamuktan 10x daha az su kullanımı; termal düzenleme",
         "2018'de Denver'da kuruldu; Indiegogo'da $280K topladı; çevre-dostu mesajla DTC büyüme; komik marka sesi; online-only"),

        ("Ettitude", "ettitude.com", "Bambu Çarşaf",
         "CleanBamboo kumaş teknolojisi; OEKO-TEX sertifikalı; su tasarruflu üretim; ipeksi yumuşak doku",
         "2019'da ABD pazarına girdi; Avustralya kökenli; bambu bazlı çarşaf DTC; sosyal medya ve influencer marketing; B Corp sertifikalı"),

        ("Comma Home", "comma-home.com", "Minimalist Nevresim",
         "Scandinavian estetik; GOTS organik pamuk; aşırı minimalist tasarım; nötr renk paleti",
         "2019'da kuruldu; Instagram-first marka; DTC Shopify; influencer ve UGC odaklı büyüme; premium segment"),

        ("Sunday Citizen", "sundaycitizen.co", "Bambu ve Kristal Çarşaf",
         "Snug bambu kumaş; kristal-infüzyon (ametist, kuvars) serisi; sürdürülebilir; ultra-yumuşak",
         "2019'da kuruldu; wellness x yatak tekstili kavramı; DTC ve Amazon; sosyal medya pazarlama; hızlı büyüme"),

        ("Sijo", "sijohome.com", "TENCEL Lüks Çarşaf",
         "TENCEL Lyocell ve French Linen serileri; AirWeight teknolojisi; termal düzenleme; minimalist ambalaj",
         "2019'da kuruldu; Instagram ve Facebook Ads odaklı DTC; sürdürülebilirlik mesajı; premium fiyat segmenti"),

        ("Pizuna Linens", "pizunalinens.com", "Uzun Elyaf Pamuk Çarşaf",
         "1000 thread count; %100 uzun elyaf pamuk; otel kalitesinde; OEKO-TEX sertifikalı",
         "2019'da kuruldu; Amazon-native marka; DTC genişleme; uygun fiyatlı lüks segment; online-only büyüme"),

        ("Bedsure", "bedsurehome.com", "Bütçe Dostu Çarşaf",
         "Amazon'un en çok satan çarşaf markası; saten, bambu, polar serileri; erişilebilir fiyat",
         "2018'de kuruldu; Amazon-first strateji; TikTok ve influencer pazarlama; $500M+ yıllık gelir; global DTC"),

        ("Quince", "onequince.com", "Lüks Çarşaf (Uygun Fiyat)",
         "Doğrudan fabrikadan tüketiciye; lüks kalite %50-80 daha ucuz; organik pamuk ve keten serileri",
         "2018'de kuruldu; aracısız tedarik zinciri; DTC only; hızlı büyüyen e-ticaret; çoklu kategori genişleme"),

        ("Cozy Earth", "cozyearth.com", "Bambu Çarşaf Premium",
         "Oprah's Favorite Things listesinde; %100 bambu viskon; 10 yıl garanti; termoregülasyon",
         "2018'de kuruldu; Oprah desteğiyle patlama; premium DTC; Shopify tabanlı; aylık $5M+ gelir"),

        ("Luxome", "luxome.com", "Ağırlıklı Battaniye + Çarşaf",
         "Bambu çarşaf ve ağırlıklı battaniye birleşimi; TENCEL Lyocell; çıkarılabilir kılıflar",
         "2019'da kuruldu; niş DTC marka; Shopify ve Amazon; uyku aksesuarı eko-sistemi; online-only"),

        ("Cariloha", "cariloha.com", "Bambu Resort Çarşaf",
         "Resort bambu kumaş; termal düzenleme; OdorFree teknolojisi; otel tarzı deneyim",
         "2018 sonrası DTC'ye geçiş; daha önce resort satış; e-ticaret odaklı büyüme; Shopify ve Amazon"),

        ("Nolah", "nolahsleep.com", "AirFoam Çarşaf & Yatak",
         "Yatak markası olarak başladı; çarşaf serisini ekledi; AirFoam uyku teknolojisi; DTC sleep brand",
         "2018'de kuruldu; yatak markası olarak DTC büyüme; çarşaf hattı genişleme; online-only"),

        ("Authenticity50", "authenticity50.com", "Made in USA Çarşaf",
         "%100 ABD üretimi; Supima pamuk; şeffaf tedarik zinciri; Buy American konsepti",
         "2018'de kuruldu; ABD üretimi odaklı DTC; Shopify; küçük ama sadık müşteri kitlesi"),

        ("Harvest Green Mattress", "harvestgreenmattress.com", "Organik Pamuk Çarşaf",
         "GOTS organik sertifikalı; çevre dostu üretim; yatak ve çarşaf seti; doğal malzemeler",
         "2018'de kuruldu; organik yatak markası; çarşaf hattı ekledi; DTC; çevre-bilinçli segment"),

        ("Beddy's", "beddys.com", "Fermuarlı Çarşaf Sistemi",
         "Patentli fermuarlı yatak çarşafı sistemi; çocuk ve ranza yatakları için ideal; kolay yatak yapma",
         "2018 sonrası viral büyüme; Instagram ve TikTok; DTC Shopify; anne blog'ları ile pazarlama"),

        ("CloudTen", "cloudtenbedding.com", "Bambu-Eucalyptus Çarşaf",
         "Bambu ve ökaliptüs karışımı; ekstra soğutma; derin cepli; OEKO-TEX",
         "2020'de kuruldu; Amazon-native; DTC genişleme; TikTok pazarlama; bütçe-dostu premium"),

        ("QuickZip", "quickzip.com", "Modüler Çarşaf Sistemi",
         "Patentli zip-on çarşaf; lastik kenar derdi yok; hızlı değiştirme; modüler tasarım",
         "2019'da DTC lansmanı; Shark Tank sunumu; Shopify; Amazon; yenilikçi yatak tekstili"),

        ("Olive + Crate", "oliveandcrate.com", "Bambu Çarşaf",
         "Premium bambu viskon; termoregülasyon; hipoalerjenik; ipek hissi",
         "2020'de kuruldu; Amazon-first; DTC büyüme; sosyal medya pazarlama; sleep niche"),

        ("CGK Unlimited", "cgkunlimited.com", "Mikrolif Çarşaf",
         "Amazon best-seller; çeşitli desen ve renk; ultra yumuşak mikrofiber; bütçe dostu",
         "2018'de kuruldu; Amazon-native marka; yüksek hacim-düşük fiyat stratejisi; e-ticaret odaklı"),

        ("California Design Den", "californiadesignden.com", "Organik Pamuk Çarşaf",
         "%100 organik pamuk; 400TC sateen; GOTS sertifikalı; uygun fiyatlı premium",
         "2018'de kuruldu; Amazon-native marka; DTC web sitesi genişleme; sürdürülebilir mesaj"),

        ("Danjor Linens", "danjorlinens.com", "6 Parça Çarşaf Seti",
         "Amazon'da en çok satan bütçe çarşaf; 1800 serisi mikrofiber; yumuşak ve dayanıklı",
         "2018'de kuruldu; Amazon-only marka; düşük fiyat yüksek hacim stratejisi; e-ticaret DTC"),

        ("Buffy", "buffy.co", "Ökaliptüs Çarşaf",
         "Geri dönüştürülmüş malzemeler; ökaliptüs fiber; çevre dostu ambalaj; yumuşak doku",
         "2018'de kuruldu; ağırlıklı battaniye ve çarşaf hattı; DTC; sosyal medya viral; sürdürülebilirlik odaklı"),

        ("Lavish Home", "lavishhome.co", "Mikrofiber Çarşaf",
         "Ultra yumuşak mikrofiber; geniş renk yelpazesi; bütçe dostu; Amazon best-seller",
         "2019'da kuruldu; Amazon-native; e-ticaret odaklı; erişilebilir fiyat segmenti"),

        ("PeachSkinSheets", "peachskinsheets.com", "SMART Fabric Çarşaf",
         "Atletik kumaş teknolojisi; nem çekici; 1500 thread count hissi; sıcak uyuyanlar için",
         "2018 sonrası DTC büyüme; Shopify; Facebook Ads; niş termal düzenleme odaklı çarşaf"),

        ("Mellanni", "mellanni.com", "Mikrofiber Çarşaf Seti",
         "Amazon'un en çok değerlendirilen çarşaf markası; 500K+ yorum; fırçalanmış mikrofiber",
         "2018 sonrası büyüme; Amazon-native; DTC genişleme; sosyal medya pazarlama; bütçe premium"),

        ("Habitat", "habitat-home.co", "Keten Çarşaf",
         "Fransız keten; stonewashed; doğal görünüm; nefes alabilir; zamansız estetik",
         "2019'da kuruldu; Instagram-first; DTC Shopify; minimalist yaşam tarzı markası"),

        ("Piglet in Bed", "pigletinbed.com", "Keten Nevresim",
         "100% doğal keten; stonewashed yumuşaklık; renk çeşitliliği; İngiliz tasarımı; sürdürülebilir",
         "2018'de kuruldu; Instagram-first DTC; İngiltere'den global genişleme; influencer marketing; premium keten segment"),

        ("Delaney Home", "delaneyhome.com", "Percale Çarşaf",
         "Mısır pamuğu percale; çıtır ve serin doku; butik otel hissi; minimalist ambalaj",
         "2020'de kuruldu; DTC Shopify; Instagram ve Pinterest pazarlama; niş premium çarşaf"),

        ("Bed Threads", "bedthreads.com", "100% Keten Çarşaf",
         "Fransız Flax keten; 20+ renk; mix-and-match; Avustralya tasarımı",
         "2018'de kuruldu; Avustralya'dan global DTC; Instagram viral; influencer kampanyaları; $50M+ gelir"),

        ("Cultiver", "cultiver.com", "Lüks Keten Nevresim",
         "Premium Avrupa keteni; el yıkama bitişi; minimalist Avustralya tasarımı; butik lüks",
         "2019'da kuruldu; Avustralya'dan global DTC; Instagram ve tasarım blogları; premium segment"),

        ("Meela", "meela.co", "Bambu İpek Çarşaf",
         "Bambu ipek karışımı; ultra yumuşak; hipoalerjenik; termal düzenleme; modern tasarım",
         "2021'de kuruldu; DTC Shopify; TikTok ve Instagram pazarlama; yeni nesil çarşaf markası"),

        ("Silk & Snow", "silkandsnow.com", "Organik Çarşaf",
         "Kanada'nın DTC uyku markası; organik pamuk çarşaf; GOTS sertifikalı; uyku ekosistemi",
         "2018'de Kanada'da kuruldu; online-only DTC; yatak + çarşaf; sosyal medya büyüme"),

        ("Wooflinen", "wooflinen.com", "Organik Keten Çarşaf",
         "GOTS organik keten; Avrupa üretimi; sürdürülebilir ambalaj; premium doğal çarşaf",
         "2020'de kuruldu; DTC niş marka; Shopify; sürdürülebilirlik odaklı; organik keten uzmanı"),

        ("Linen House", "linenhouse.com.au", "Avustralya Nevresim",
         "Avustralya tasarımı; cesur desenler; premium pamuk ve keten; geniş koleksiyon",
         "2018 sonrası DTC genişleme; Avustralya'dan global e-ticaret; Shopify; Instagram; tasarım odaklı"),

        ("Snowe Home", "snowehome.com", "Percale & Sateen Çarşaf",
         "İtalya üretimi kumaş; percale ve sateen seçenekler; minimalist DTC; otel deneyimi",
         "2018'de kuruldu; New York DTC; doğrudan fabrikadan; Shopify; premium erişilebilir lüks"),
    ],

    # ─────────────────────────────────────────────────────────────────
    # 2. YORGAN & DOLGULU ÜRÜNLER  (Comforters & Duvets)
    # ─────────────────────────────────────────────────────────────────
    "Yorgan & Dolgulu Ürünler": [
        ("Buffy", "buffy.co", "Bulut Yorgan",
         "Geri dönüştürülmüş PET şişelerden dolgu; ökaliptüs kumaş; hipoalerjenik; çevre dostu",
         "2018'de kuruldu; Cloud Comforter ile viral; Instagram ve Facebook DTC; ücretsiz deneme modeli; $35M+ gelir"),

        ("Sunday Citizen", "sundaycitizen.co", "Bambu Yorgan",
         "Bambu dolgu ve kılıf; kristal-infüzyon opsiyonu; snug comfort; wellness odaklı",
         "2019'da kuruldu; DTC wellness yatak markası; kristal enerji konsepti; Shopify; Amazon"),

        ("Cloverlane", "cloverlane.com", "Mevsimlik Yorgan",
         "Modüler snap-on yorgan sistemi; mevsime göre kat ekleme/çıkarma; akıllı tasarım",
         "2020'de kuruldu; yenilikçi modüler yorgan; DTC Shopify; sosyal medya viral; Kickstarter kökenli"),

        ("Nolah", "nolahsleep.com", "Soğutma Yorganı",
         "AirFiber dolgu; serinletici teknoloji; hipoalerjenik; DTC uyku markası",
         "2018'de kuruldu; yatak markası olarak başladı; yorgan hattı genişleme; online-only DTC"),

        ("Rest", "rest.co", "Everloft Yorgan",
         "Everloft dolgu teknolojisi; kuş tüyü alternatifi; topaklanmaz; yıkanabilir",
         "2019'da kuruldu; DTC uyku aksesuarı markası; Shopify; sosyal medya pazarlama; yenilikçi dolgu"),

        ("Sijo", "sijohome.com", "TENCEL Yorgan",
         "TENCEL Lyocell kılıf; AirWeight dolgu; ultra hafif; termoregülasyon; all-season",
         "2019'da kuruldu; çarşaf markasından yorgan genişleme; DTC; minimalist estetik"),

        ("Cozy Earth", "cozyearth.com", "Bambu Yorgan",
         "100% bambu dolgu ve kılıf; Oprah's Favorite; termal düzenleme; 10 yıl garanti",
         "2018'de kuruldu; Oprah seçimi; premium DTC; Shopify; aylık $5M+ gelir"),

        ("Eucalypso", "eucalypso.com", "Ökaliptüs Yorgan",
         "100% TENCEL ökaliptüs; vegan; hipoalerjenik; serin uyku; sürdürülebilir",
         "2019'da kuruldu; sürdürülebilir yatak tekstili DTC; Shopify; Instagram marketing"),

        ("Quince", "onequince.com", "Kaz Tüyü Yorgan",
         "Avrupa kaz tüyü; fabrikadan direkt fiyat; lüks kalite %50 daha ucuz; RDS sertifikalı",
         "2018'de kuruldu; DTC aracısız model; uygun fiyatlı lüks yorgan; hızlı büyüme"),

        ("Puredown", "puredown.com", "Doğal Tüy Yorgan",
         "Kaz ve ördek tüyü; mevsimlik seçenekler; hipoalerjenik işlem; uygun fiyat",
         "2018 sonrası DTC büyüme; Amazon-native; global tedarik; e-ticaret odaklı"),

        ("Luna Weighted Blanket", "lunablanket.com", "Ağırlıklı Yorgan/Battaniye",
         "8-25 lb seçenekler; %100 pamuk; cam boncuk dolgu; terapötik uyku",
         "2019'da kuruldu; Amazon-first; DTC genişleme; uyku sağlığı odaklı; Shopify"),

        ("Sleep & Beyond", "sleepandbeyond.com", "Organik Yün Yorgan",
         "Organik merinos yünü dolgu; nefes alabilir; termal düzenleme; doğal malzeme",
         "2018 sonrası DTC genişleme; organik uyku markası; Shopify; Amazon; sürdürülebilir"),

        ("FluffCo", "fluffco.com", "Otel Tarzı Yorgan",
         "5 yıldızlı otel deneyimi; down-alternative dolgu; lüks hissiyat; erişilebilir fiyat",
         "2020'de kuruldu; otel lüksünü eve taşıma konsepti; DTC Shopify; Instagram reklamları"),

        ("Authenticity50", "authenticity50.com", "Made in USA Yorgan",
         "ABD üretimi; Supima pamuk kılıf; PrimaLoft dolgu; şeffaf üretim",
         "2018'de kuruldu; ABD üretimi odaklı DTC; Shopify; premium yerli üretim segment"),

        ("Luxome", "luxome.com", "Ağırlıklı Yorgan",
         "Yorgan ve ağırlıklı battaniye birleşimi; çıkarılabilir kılıf; bambu kumaş",
         "2019'da kuruldu; hibrit ürün konsepti; DTC Shopify; Amazon; niş uyku segmenti"),

        ("Silk & Snow", "silkandsnow.com", "Organik Yorgan",
         "Organik pamuk kılıf; geri dönüştürülmüş PET dolgu; Kanada DTC; vegan",
         "2018'de kuruldu; Kanada'nın DTC uyku markası; online-only; sürdürülebilir"),

        ("Brooklinen", "brooklinen.com", "Down Yorgan",
         "Kanada kaz tüyü; all-season ve lightweight seçenekler; lüks DTC",
         "2018 sonrası yorgan hattı genişleme; köklü DTC uyku markası; $200M+ gelir; Shopify + perakende"),

        ("Parachute Home", "parachutehome.com", "Down Yorgan",
         "Avrupa beyaz kaz tüyü; Oeko-Tex; premium DTC; minimalist estetik",
         "2018 sonrası yorgan genişleme; LA tabanlı DTC ev markası; Instagram-driven; perakende genişleme"),

        ("Panda London", "pandalondon.com", "Bambu Yorgan",
         "Bambu liyosel kılıf ve dolgu; hipoalerjenik; termal düzenleme; İngiliz DTC",
         "2018'de kuruldu; İngiltere'den DTC; sürdürülebilir bambu uyku markası; Amazon UK; Shopify"),

        ("Bedfolk", "bedfolk.com", "Keten Nevresim Yorgan Kılıfı",
         "100% Avrupa keteni; stonewashed; zamansız renk paleti; minimalist İngiliz marka",
         "2019'da kuruldu; İngiliz DTC keten markası; Instagram-first; premium doğal kumaş"),

        ("Hush Blankets", "hushblankets.com", "İced Soğutma Yorgan",
         "Soğutma teknolojili yorgan; sıcak uyuyanlar için; iki taraflı tasarım",
         "2018'de Kanada'da kuruldu; DTC; ağırlıklı battaniye ile başladı; yorgan hattı genişleme"),

        ("Nest Bedding", "nestbedding.com", "Easy Breather Yorgan",
         "Tencel kılıf; soğutma jel fiber dolgu; all-season; DTC uyku ekosistemi",
         "2018 sonrası DTC büyüme; uyku markası genişleme; Shopify; online-first"),

        ("Nemu", "nemuhome.com", "Japon Tarzı Yorgan",
         "Japon futon ilhamı; hafif pamuk dolgu; minimalist Japon estetik; all-season",
         "2021'de kuruldu; Japon uyku kültürü DTC; Shopify; Instagram; niş premium"),

        ("Duvetier", "duvetier.com", "Snap-On Yorgan Sistemi",
         "Patentli snap-on nevresim bağlama sistemi; kolay nevresim takma; pratik tasarım",
         "2020'de kuruldu; Kickstarter; DTC Shopify; yenilikçi yorgan sistemi; TikTok viral"),

        ("EverSnug", "eversnug.com", "Lüks Down-Alternative Yorgan",
         "Premium down-alternative; ultra kabarık; yıkanabilir; hipoalerjenik; bütçe lüks",
         "2020'de kuruldu; Amazon-first DTC; sosyal medya reklamları; hızlı büyüme"),

        ("Baloo Living", "balooliving.com", "Ağırlıklı Yorgan",
         "Doğal cam boncuk dolgu; %100 pamuk; hipoalerjenik; terapötik ağırlık",
         "2019'da kuruldu; DTC ağırlıklı uyku ürünleri; Shopify; wellness segment"),

        ("Nocturne", "nocturnebedding.com", "Premium Yorgan",
         "700+ fill power Macar kaz tüyü; el yapımı; süper premium DTC; özel sipariş",
         "2020'de kuruldu; ultra-lüks DTC yorgan; Shopify; Instagram; niş premium segment"),

        ("SCOOMS", "scooms.com", "Macar Kaz Tüyü Yorgan",
         "Macar kaz tüyü; İngiliz DTC; 3 sıcaklık seviyesi; OEKO-TEX; doğal",
         "2018'de kuruldu; İngiliz DTC yorgan markası; Shopify; sürdürülebilir lüks"),

        ("Woolroom", "thewoolroom.com", "Yün Yorgan",
         "İngiliz yünü dolgu; doğal termal düzenleme; hipoalerjenik; kimyasalsız",
         "2018 sonrası DTC genişleme; İngiliz yün uzmanı; Shopify; Amazon UK; doğal uyku"),

        ("Under the Canopy", "underthecanopy.com", "Organik Yorgan",
         "GOTS organik sertifikalı; geri dönüştürülmüş dolgu; Fair Trade; çevre dostu",
         "2019'da DTC lansmanı; sürdürülebilir ev tekstili; Shopify; Amazon; organik segment"),

        ("Coop Home Goods", "coophomegoods.com", "Ayarlanabilir Yorgan",
         "Uyku ürünleri ekosistemi; ayarlanabilir dolgu; bambu kılıf; DTC",
         "2018 sonrası büyüme; Amazon-first; DTC genişleme; yastık markasından yorgan hattı"),

        ("LANE LINEN", "lanelinen.com", "Goose Down Yorgan",
         "Premium kaz tüyü; baffle-box tasarım; all-season; otel kalitesi; Amazon best-seller",
         "2019'da kuruldu; Amazon-native DTC; hızlı büyüme; premium bütçe dostu yorgan"),

        ("Lifease", "lifease.com", "Japon Tarzı Yorgan",
         "Japon minimalizm; doğal pamuk dolgu; hafif ve nefes alabilir; basit estetik",
         "2019'da kuruldu; Asya-ilhamlı DTC; Shopify; minimalist yaşam segment; online-only"),

        ("Snug Comforter", "snugcomforter.com", "Down-Alternative Yorgan",
         "Hypoallergenic dolgu; ultra kabarık; yıkanabilir; all-season; bütçe dostu",
         "2020'de kuruldu; Amazon-native; DTC; sosyal medya pazarlama; hızlı büyüme"),

        ("Panda London", "pandalondon.com", "Bambu Cloud Yorgan",
         "Bambu liyosel dolgu ve kılıf; Cloud yorgan teknolojisi; hipoalerjenik; termoregülasyon",
         "2018'de kuruldu; İngiliz DTC; sürdürülebilir bambu uyku markası; Amazon UK; Shopify"),
    ],

    # ─────────────────────────────────────────────────────────────────
    # 3. YASTIK & UYKU AKSESUARI  (Pillows & Sleep Accessories)
    # ─────────────────────────────────────────────────────────────────
    "Yastık & Uyku Aksesuarı": [
        ("Coop Home Goods", "coophomegoods.com", "Ayarlanabilir Yastık",
         "Parçalanmış memory foam; dolgu miktarı ayarlanabilir; GreenGuard Gold sertifikalı; bambu kılıf",
         "2018 sonrası patlama; Amazon #1 yastık; 70K+ yorum; DTC ve Amazon; aylık $3M+ gelir"),

        ("Purple", "purple.com", "Hyper-Elastic Yastık",
         "Grid teknolojisi; basınç dağılımı; hiperelas polimer; hava sirkülasyonu",
         "2018 sonrası yastık hattı genişleme; yatak markasından; DTC; viral video pazarlama; $700M+ gelir"),

        ("Pillow Cube", "pillowcube.com", "Küp Yastık (Yan Uyuyanlar)",
         "Küp şeklinde; yan uyuyanlar için ideal; boşluk dolduran tasarım; memory foam",
         "2019'da kuruldu; TikTok viral; Kickstarter $1M+; DTC Shopify; eğlenceli pazarlama"),

        ("Hullo", "hullopillow.com", "Karabuğday Yastık",
         "Organik karabuğday dolgu; doğal hava akışı; ayarlanabilir yükseklik; Made in USA",
         "2018'de kuruldu; niş doğal yastık DTC; Shopify; Amazon; organik yaşam segment"),

        ("Lagoon Sleep", "lagoonsleep.com", "Kişiye Özel Yastık",
         "Quiz ile kişiye özel yastık önerisi; 6 farklı model; uyku pozisyonuna göre eşleşme",
         "2020'de kuruldu; kişiselleştirme odaklı DTC; Shopify; Instagram reklamları; yenilikçi konsept"),

        ("Pluto Pillow", "plutopillow.com", "Kişiselleştirilmiş Yastık",
         "25+ soruyla tamamen kişiye özel yastık; boy, kilo, uyku pozisyonu analizi",
         "2019'da kuruldu; dünyada ilk tam kişiselleştirilmiş yastık; DTC; Shopify; viral pazarlama"),

        ("Eli & Elm", "eliandelm.com", "Yan Uyuyan Yastığı",
         "Patentli U-şekilli tasarım; yan uyuyanlar için ergonomik; pamuk/lateks karışım dolgu",
         "2019'da kuruldu; niş ergonomik yastık DTC; Shopify; Amazon; uyku sağlığı odaklı"),

        ("Beckham Hotel Collection", "beckhamhotel.com", "Otel Tarzı Yastık",
         "Amazon'un #1 yastığı; jel fiber dolgu; otel deneyimi; 250K+ yorum",
         "2018'de kuruldu; Amazon-native; otel kalitesi yastık; DTC e-ticaret; mega best-seller"),

        ("Sleepgram", "sleepgram.com", "3'ü 1 Arada Yastık",
         "İç içe 3 yastık; yükseklik ayarı; microfiber dolgu; all-sleeping-positions",
         "2018'de kuruldu; Shark Tank yatırımı; DTC; Amazon; yenilikçi modüler tasarım"),

        ("Honeydew Sleep", "honeydewaleep.com", "Scrumptious Yastık",
         "Copper-infused CertiPUR memory foam; omurga hizalama; yan/sırt uyuyan tasarımı",
         "2019'da kuruldu; California DTC; Shopify; premium ergonomik yastık; niş büyüme"),

        ("Derila", "derila.com", "Servikal Memory Foam Yastık",
         "Ergonomik kontur tasarım; boyun desteği; memory foam; taşınabilir",
         "2020'de kuruldu; Facebook/Instagram Ads ile viral; global DTC; Shopify; hızlı büyüme"),

        ("Cozy Earth", "cozyearth.com", "Bambu Yastık",
         "Bambu viskon kılıf; parçalanmış memory foam dolgu; Oprah seçimi; premium",
         "2018'de kuruldu; uyku markasından yastık genişleme; Oprah's Favorite; DTC Shopify"),

        ("Nolah", "nolahsleep.com", "AirFiber Yastık",
         "AirFiber teknolojisi; soğutma jel; ayarlanabilir yükseklik; DTC uyku markası",
         "2018'de kuruldu; yatak markasından genişleme; yastık hattı; online-only DTC"),

        ("Tuft & Needle", "tuftandneedle.com", "Adaptif Foam Yastık",
         "T&N Adaptive foam; grafit ve jel infüzyon; soğutma; DTC uyku markası",
         "2018 sonrası yastık genişleme; DTC yatak devinden; Shopify; Amazon; $150M+ gelir"),

        ("Bear", "bearmattress.com", "Bear Pillow",
         "Celliant kılıf teknolojisi; Double Ice Fabric; atletler için tasarım; performans uyku",
         "2018 sonrası genişleme; atlet odaklı DTC uyku markası; Shopify; sporcu pazarlama"),

        ("Casper", "casper.com", "Original Pillow",
         "Yastık-içinde-yastık tasarımı; iki katman; yumuşak dış sert iç; DTC ikonu",
         "2018 sonrası yastık genişleme; DTC yatak devinden; IPO deneyimi; online + perakende"),

        ("Sijo", "sijohome.com", "AirWeight Yastık",
         "TENCEL Lyocell kılıf; hafif teknoloji; nefes alabilir; minimalist",
         "2019'da kuruldu; çarşaf markasından genişleme; DTC; premium uyku aksesuarı"),

        ("Ostrichpillow", "ostrichpillow.com", "Taşınabilir Uyku Yastığı",
         "İkonik taşınabilir yastık tasarımı; seyahat ve power nap; yenilikçi endüstriyel tasarım",
         "2018 sonrası genişleme; Kickstarter viral; global DTC; seyahat ve ofis uyku nişi"),

        ("Mellow Sleep", "mellowsleep.com", "Soğutma Jel Yastık",
         "Phase-change soğutma materyali; memory foam çekirdek; ortopedik destek; serin uyku",
         "2019'da kuruldu; soğutma uyku teknolojisi; DTC; Amazon; TikTok pazarlama"),

        ("Pancake Pillow", "pancakepillow.com", "Katmanlı Yastık",
         "6 ince katman; istediğin yüksekliği ayarla; her katman çıkarılabilir; tam kontrol",
         "2019'da kuruldu; Kickstarter; DTC Shopify; yenilikçi modüler yastık; niş büyüme"),

        ("SpineAlign", "spinealign.com", "Omurga Hizalama Yastığı",
         "Chiropraktör tasarımı; omurga hizalama teknolojisi; 3 bölge desteği; ergonomik",
         "2019'da kuruldu; sağlık profesyoneli onaylı; DTC; Shopify; wellness segment"),

        ("WonderSleep", "wondersleep.com", "Premium Yastık",
         "Ayarlanabilir loft; shredded memory foam; bambu kılıf; hipoalerjenik",
         "2019'da kuruldu; Amazon-first; DTC genişleme; premium bütçe dostu yastık"),

        ("Zoey Sleep", "zoeysleep.com", "Yan Uyuyan Yastığı",
         "Yan uyuyanlar için tasarlanmış; omuz bölgesi kesimi; memory foam; ergonomik",
         "2020'de kuruldu; niş uyku pozisyonu odaklı; DTC Shopify; Instagram pazarlama"),

        ("Cushion Lab", "cushionlab.com", "Ergonomik Yastık",
         "Ekstra yoğun memory foam; ergonomik tasarım; boyun ve sırt desteği; ofis + uyku",
         "2019'da kuruldu; ergonomik uyku ve oturma; DTC; Amazon; TikTok viral; $10M+ gelir"),

        ("Blissy", "blissy.com", "İpek Yastık Kılıfı",
         "22 momme %100 dut ipeği yastık kılıfı; cilt ve saç sağlığı; anti-aging; lüks uyku",
         "2018'de kuruldu; TikTok ve Instagram viral; DTC; $100M+ gelir; beauty-meets-sleep"),

        ("Drowsy", "drowsy.com", "İpek Uyku Maskesi",
         "Premium dut ipeği; %100 karartma; ayarlanabilir; lüks ambalaj; hediye segmenti",
         "2019'da kuruldu; İngiliz DTC; Instagram-first; influencer marketing; lüks uyku aksesuarı"),

        ("Slip", "slip.com", "İpek Yastık Kılıfı",
         "Slipsilk teknolojisi; dermatolog onaylı; anti-aging; saç koruma; Hollywood favorisi",
         "2018 sonrası DTC genişleme; influencer ve ünlü pazarlama; Instagram; Shopify; premium ipek"),

        ("Savvy Sleeper", "savvysleeper.com", "Satin Yastık Kılıfı",
         "Satin yastık kılıfı; saç ve cilt koruma; fermuar kapatma; erişilebilir fiyat",
         "2019'da kuruldu; Amazon-first DTC; influencer marketing; beauty/sleep crossover"),

        ("Night", "night.co", "Trisilk Yastık Kılıfı",
         "Triple ipek karışımı; anti-aging; nem dengeleme; İngiliz lüks uyku markası",
         "2019'da kuruldu; İngiliz DTC; premium ipek uyku aksesuarı; Shopify; Instagram"),

        ("Moonbow", "moonbow.com", "Kişiselleştirilmiş Yastık",
         "AI ile kişiye özel yastık tasarımı; uyku verisi analizi; 3D baskı foam; yenilikçi",
         "2021'de kuruldu; teknoloji odaklı kişiselleştirilmiş yastık; DTC; Shopify; yenilikçi startup"),

        ("Dosaze", "dosaze.com", "Terapötik Yastık",
         "Ortopedik tasarım; boyun ağrısı çözümü; memory foam; ayarlanabilir yükseklik",
         "2020'de kuruldu; sağlık ve wellness odaklı; DTC Shopify; Facebook Ads; niş pazar"),

        ("Lincove", "lincove.com", "Canadian Down Yastık",
         "Kanada kaz tüyü; organik pamuk kılıf; doğal lüks; el yapımı kalite",
         "2019'da kuruldu; DTC premium doğal yastık; Shopify; Amazon; lüks segment"),

        ("Snuggle-Pedic", "snugglepedic.com", "Shredded Foam Yastık",
         "Parçalanmış memory foam ve micro-gel; hava akışı; CertiPUR; ayarlanabilir",
         "2018 sonrası DTC büyüme; Amazon best-seller; ReliaCool teknolojisi; online-only"),

        ("Silvon", "silvon.co", "Antimikrobiyal Yastık Kılıfı",
         "Gümüş iyonlu antimikrobiyal; akne önleme; bakteri azaltma; temiz uyku",
         "2019'da kuruldu; akne çözümü olarak pozisyonlama; DTC Shopify; TikTok viral; niş beauty/sleep"),
    ],

    # ─────────────────────────────────────────────────────────────────
    # 4. AĞIRLIKLI BATTANİYE  (Weighted Blankets)
    # ─────────────────────────────────────────────────────────────────
    "Ağırlıklı Battaniye": [
        ("Bearaby", "bearaby.com", "Örgü Ağırlıklı Battaniye",
         "Dolgu malzeme YOK; ağırlık kumaşın kendisinden; Tree Napper organik TENCEL; estetik tasarım",
         "2018'de Kickstarter ile kuruldu; dolgu-free örgü ağırlıklı battaniye icat etti; Inc. 5000'de 82. sıra; 4999% gelir artışı; aylık $1.9M gelir; online-only DTC"),

        ("Nuzzie", "nuzzie.com", "Örgü Ağırlıklı Battaniye",
         "Bearaby alternatifi; chunky örgü; dolgu-free; daha uygun fiyat; el yapımı görünüm",
         "2020'de kuruldu; TikTok viral; Instagram influencer pazarlama; DTC Shopify; hızlı büyüme"),

        ("Baloo Living", "balooliving.com", "Cam Boncuk Ağırlıklı Battaniye",
         "%100 pamuk; doğal cam boncuk dolgu; hipoalerjenik; yıkanabilir; 15-25 lb",
         "2019'da kuruldu; doğal malzeme odaklı; DTC Shopify; Amazon; wellness segment; premium"),

        ("Luna Weighted Blanket", "lunablanket.com", "Klasik Ağırlıklı Battaniye",
         "Cam boncuk dolgu; %100 pamuk; 8-25 lb seçenekler; grid tasarım; erişilebilir fiyat",
         "2019'da kuruldu; Amazon best-seller; DTC genişleme; bütçe dostu ağırlıklı battaniye"),

        ("Hush Blankets", "hushblankets.com", "İced Ağırlıklı Battaniye",
         "Soğutma teknolojili kılıf; cam boncuk dolgu; sıcak uyuyanlar için; Kanada DTC",
         "2018'de Kanada'da kuruldu; DTC; soğutma özellikli ağırlıklı battaniye öncüsü; $30M+ gelir"),

        ("Zonli", "zonlihome.com", "Çok Amaçlı Ağırlıklı Battaniye",
         "Farklı kumaş seçenekleri; kişiye özel ağırlık; ev ve seyahat; erişilebilir",
         "2019'da kuruldu; Amazon-native; TikTok pazarlama; DTC genişleme; bütçe dostu segment"),

        ("Gravity Blanket", "gravityblankets.com", "Premium Ağırlıklı Battaniye",
         "Ağırlıklı battaniye kategorisinin popülerleştiricisi; cam mikro boncuk; %100 pamuk kılıf",
         "2018'de Kickstarter ile kuruldu; $4.7M kampanya; DTC; Target perakende genişleme; viral"),

        ("Layla Sleep", "laylasleep.com", "Ağırlıklı Battaniye",
         "Fleeced + pamuk iki taraflı kılıf; 300TC pamuk; hex dikiş; cam boncuk dolgu",
         "2018 sonrası genişleme; yatak markasından ağırlıklı battaniye hattı; DTC; Amazon"),

        ("Luxome", "luxome.com", "Bambu Ağırlıklı Battaniye",
         "Bambu-türevi kumaş; çıkarılabilir kılıf; soğutma; cam boncuk dolgu; premium",
         "2019'da kuruldu; bambu lüks segment; DTC Shopify; Amazon; niş premium"),

        ("Weighted Evolution", "weightedevolution.com", "Terapötik Ağırlıklı Battaniye",
         "Terapist tasarımı; anksiyete ve uyku bozukluğu odaklı; cam boncuk; %100 pamuk",
         "2019'da kuruldu; terapötik odaklı DTC; Shopify; wellness pazarlama; niş sağlık"),

        ("Quility", "quility.com", "Bütçe Ağırlıklı Battaniye",
         "Amazon'da en çok satan; cam boncuk; 5-30 lb seçenekler; çıkarılabilir kılıf; uygun fiyat",
         "2018'de kuruldu; Amazon-native mega best-seller; 100K+ yorum; DTC e-ticaret"),

        ("YnM", "ynmhome.com", "Cam Boncuk Ağırlıklı Battaniye",
         "7 katmanlı tasarım; cam boncuk + polyester dolgu; Amazon #1 best-seller; çok seçenek",
         "2018'de kuruldu; Amazon-native; ağırlıklı battaniye pazarının en büyük satıcılarından; DTC"),

        ("Degrees of Comfort", "degreesofcomfort.com", "Çift Taraflı Ağırlıklı Battaniye",
         "Sıcak/soğuk iki taraflı tasarım; nano seramik boncuk; mevsime göre kullanım",
         "2019'da kuruldu; Amazon-native; yenilikçi çift taraflı konsept; DTC; hızlı büyüme"),

        ("Comma Home", "comma-home.com", "Tasarım Ağırlıklı Battaniye",
         "Scandinavian estetik; minimalist tasarım; doğal malzemeler; premium ağırlıklı battaniye",
         "2019'da kuruldu; Instagram-first; tasarım odaklı ağırlıklı battaniye; DTC Shopify"),

        ("SensaCalm", "sensacalm.com", "Kişiselleştirilmiş Ağırlıklı Battaniye",
         "Özel sipariş; kumaş ve ağırlık seçimi; otizm ve duyusal işleme odaklı; terapötik",
         "2018 sonrası DTC büyüme; duyusal terapi uzmanı; Shopify; niş terapötik segment"),

        ("Mosaic Weighted Blankets", "mosaicweightedblankets.com", "Terapötik Ağırlıklı Battaniye",
         "ABD üretimi; duyusal terapi; çocuk ve yetişkin; özel tasarım seçenekleri",
         "2018 sonrası genişleme; ABD üretimi terapötik battaniye; DTC; niş pazar"),

        ("Helix Sleep", "helixsleep.com", "Ağırlıklı Battaniye",
         "Uyku markasından genişleme; cam mikro boncuk; %100 pamuk; ayarlanabilir kılıf",
         "2018 sonrası genişleme; DTC yatak markasından ağırlıklı battaniye; Shopify"),

        ("Kudd.ly", "kuddly.co", "Sherpa Ağırlıklı Battaniye",
         "Sherpa teddy kumaş; cam boncuk dolgu; kışlık sıcak tasarım; İngiliz DTC",
         "2020'de kuruldu; İngiltere'den DTC; TikTok viral; Instagram pazarlama; hızlı büyüme"),

        ("Calming Blankets", "calmingblankets.com.au", "Avustralya Ağırlıklı Battaniye",
         "Avustralya'nın DTC ağırlıklı battaniye markası; bambu ve pamuk; soğutma; terapötik",
         "2018'de kuruldu; Avustralya DTC; Shopify; Facebook Ads; wellness segment"),

        ("Weighted Idea", "weightedidea.com", "Bütçe Ağırlıklı Battaniye",
         "Amazon best-seller; çeşitli boyut ve ağırlık; cam boncuk; uygun fiyat; yüksek hacim",
         "2018'de kuruldu; Amazon-native; Çin üretimi DTC; düşük fiyat stratejisi"),

        ("Mela Weighted Blanket", "melacomfort.com", "Premium Ağırlıklı Battaniye",
         "İngiliz DTC; OEKO-TEX; %100 pamuk; cam boncuk; premium tasarım; Sunday Times tavsiyesi",
         "2019'da kuruldu; İngiliz DTC; Shopify; Instagram pazarlama; premium segment"),

        ("Buzio", "buzio.com", "Çeşitli Ağırlıklı Battaniye",
         "20+ farklı desen ve renk; sherpa, pamuk, minky; cam boncuk; Amazon best-seller",
         "2019'da kuruldu; Amazon-native; çeşitlilik odaklı; TikTok ve Instagram; DTC genişleme"),

        ("WONAP", "wonap.com", "Bambu Ağırlıklı Battaniye",
         "Bambu viskon kumaş; cam boncuk dolgu; soğutma; premium doğal; OEKO-TEX",
         "2019'da kuruldu; Amazon-first; bambu premium segment; DTC; doğal malzeme odaklı"),

        ("Relixiy", "relixiy.com", "Soğutma Ağırlıklı Battaniye",
         "Soğutma kumaşı; cam boncuk dolgu; yaz kullanımı; sıcak iklimler için",
         "2020'de kuruldu; Amazon-native; soğutma niş segment; DTC; TikTok pazarlama"),

        ("Sweetzer & Orange", "sweetzerandorange.com", "Çocuk Ağırlıklı Battaniye",
         "Çocuk odaklı tasarımlar; eğlenceli desenler; güvenli ağırlık oranları; yumuşak minky",
         "2019'da kuruldu; Amazon-native; çocuk ağırlıklı battaniye niş; DTC; anne blog pazarlama"),

        ("Amy Garden", "amygarden-home.com", "Bütçe Ağırlıklı Battaniye",
         "Geniş boyut ve ağırlık yelpazesi; cam boncuk; pamuk; uygun fiyat; Amazon popüler",
         "2018'de kuruldu; Amazon-native; düşük fiyat yüksek hacim; DTC e-ticaret"),

        ("Nest Bedding", "nestbedding.com", "Ağırlıklı Battaniye",
         "DTC uyku markasından genişleme; cam mikro boncuk; tencel kılıf; premium",
         "2018 sonrası genişleme; DTC uyku ekosistemi; Shopify; online-first"),

        ("Honeybird", "honeybirdofficial.com", "Şık Ağırlıklı Battaniye",
         "Estetik tasarım odaklı; ev dekoruna uyumlu; pamuk ve minky seçenekler; cam boncuk",
         "2020'de kuruldu; tasarım odaklı DTC; Instagram ve Pinterest pazarlama; Shopify"),

        ("ZonLi Kids", "zonlihome.com", "Çocuk Ağırlıklı Battaniye",
         "Çocuklar için güvenli ağırlık; eğlenceli desenler; minky yumuşak; cam boncuk",
         "2019'da kuruldu; Zonli'nin çocuk hattı; Amazon-native; DTC; çocuk wellness"),

        ("Slumber Cloud", "slumbercloud.com", "Outlast Ağırlıklı Battaniye",
         "NASA Outlast teknolojisi; termal düzenleme; sıcaklık dengeleme; premium teknoloji",
         "2018'de kuruldu; NASA teknolojisi lisansı; DTC Shopify; premium uyku teknolojisi"),

        ("Brooklyn Bedding", "brooklynbedding.com", "Ağırlıklı Battaniye",
         "DTC yatak markasından genişleme; cam mikro boncuk; pamuk kılıf; premium",
         "2018 sonrası genişleme; köklü DTC yatak markası; Shopify; online + showroom"),
    ],

    # ─────────────────────────────────────────────────────────────────
    # 5. BATTANİYE & THROW  (Blankets & Throws)
    # ─────────────────────────────────────────────────────────────────
    "Battaniye & Throw": [
        ("Bearaby", "bearaby.com", "Chunky Örgü Throw",
         "Velvet Napper; Tree Napper; Cotton Napper; dolgu-free örgü; dekoratif ve fonksiyonel",
         "2018'de kuruldu; ağırlıklı battaniye markasından throw genişleme; DTC; estetik ev aksesuarı"),

        ("Sunday Citizen", "sundaycitizen.co", "Snug Throw Battaniye",
         "Snug Bamboo kumaş; kristal-infüzyon serileri; wellness battaniye; ultra yumuşak",
         "2019'da kuruldu; wellness x ev tekstili; DTC; Shopify; Instagram; kristal enerji konsepti"),

        ("Rumpl", "rumpl.com", "Teknik Outdoor Battaniye",
         "Geri dönüştürülmüş malzeme; su geçirmez; packable; outdoor ve ev kullanımı; kamp battaniyesi",
         "2018 sonrası büyüme; outdoor-lifestyle DTC; Shopify; sanatçı koleksiyonları; $20M+ gelir"),

        ("Minky Couture", "minkycouture.com", "Lüks Minky Battaniye",
         "Ultra yumuşak minky kumaş; ABD üretimi; 100+ desen; premium hediye battaniye",
         "2018 sonrası DTC genişleme; Utah tabanlı; Shopify; Instagram; influencer pazarlama; premium"),

        ("Big Blanket Co", "bigblanketco.com", "10x10 Dev Battaniye",
         "Dünyanın en büyük battaniyesi; 10x10 feet; aile boyu; orijinal konsept",
         "2019'da kuruldu; TikTok ve Instagram viral; DTC Shopify; eğlenceli pazarlama; $15M+ gelir"),

        ("Saranoni", "saranoni.com", "Lush Lüks Battaniye",
         "Ultra yumuşak lush kumaş; bebek ve yetişkin; hediye ambalaj; premium hissiyat",
         "2018'de kuruldu; Instagram-first; influencer pazarlama; DTC Shopify; premium battaniye"),

        ("Kashwere", "kashwere.com", "Süper Yumuşak Throw",
         "Polyester micro chenille; otel lüksü; ultra yumuşak; geniş renk yelpazesi",
         "2018 sonrası DTC genişleme; premium throw segment; Shopify; influencer; hediye pazarı"),

        ("Tkano", "tkano.com", "Scandinavian Throw",
         "İskandinav tasarım; doğal kumaşlar; pamuk ve yün; minimalist estetik",
         "2019'da kuruldu; European DTC; minimalist ev aksesuarı; Shopify; Instagram"),

        ("ChappyWrap", "chappywrap.com", "Jacquard Throw Battaniye",
         "ABD üretimi jacquard; orijinal tasarımlar; washer-dryer uyumlu; premium kalite",
         "2018'de kuruldu; Maine USA; Instagram-first DTC; Shopify; influencer; $10M+ gelir"),

        ("Thula Tula", "thulatula.com", "Bambu Müslin Battaniye",
         "Bambu müslin; bebek ve yetişkin; Güney Afrika ilhamı; sosyal etki odaklı",
         "2019'da kuruldu; sosyal girişim; DTC Shopify; Instagram; çocuk battaniye pazarı"),

        ("Koolaburra by UGG", "koolaburra.com", "Sherpa Throw",
         "Faux-fur ve sherpa; UGG lüks hissi; ev kullanımı; premium throw",
         "2019'da DTC genişleme; UGG alt markası; online-first; Shopify; Amazon; lüks casual"),

        ("Pendleton Woolen Mills", "pendleton-usa.com", "Heritage Yün Battaniye",
         "Amerikan heritage yün battaniye; ikonik desenleri DTC'ye taşıdı; premium yün",
         "2018 sonrası DTC büyüme; e-ticaret odaklı genişleme; Shopify; geleneksel marka modernizasyonu"),

        ("Nomadix", "nomadix.co", "Geri Dönüştürülmüş Battaniye",
         "%100 geri dönüştürülmüş malzeme; hızlı kuruyan; kum tutmayan; plaj ve outdoor",
         "2018'de kuruldu; sürdürülebilir outdoor DTC; Shopify; Instagram; çevre dostu yaşam tarzı"),

        ("Bohemian Trading", "bohemiantrading.com", "Boho Throw Battaniye",
         "El dokuması; boho estetik; Türk pamuklu throw; artisan üretim",
         "2019'da kuruldu; boho lifestyle DTC; Shopify; Instagram; Etsy; artisan segment"),

        ("Pura Vida", "pfrievda.com", "Meksika Tarzı Throw",
         "Meksika ilhamlı desenler; geri dönüştürülmüş pamuk; plaj ve ev; boho yaşam tarzı",
         "2019'da battaniye genişleme; bileklik markasından; DTC; sosyal medya; lifestyle"),

        ("Half Moon Blanket", "halfmoonblanket.com", "Premium Yün Throw",
         "Merinos yünü; İskandinav tasarım; doğal boyama; sürdürülebilir üretim",
         "2020'de kuruldu; premium doğal throw; DTC Shopify; Instagram; niş tasarım segment"),

        ("Gravel", "graveltravel.com", "Seyahat Battaniyesi",
         "Packable seyahat battaniyesi; hafif; taşıma çantalı; uçak ve ofis",
         "2019'da kuruldu; seyahat aksesuarı DTC; Shopify; Instagram; dijital göçebe segment"),

        ("Kyte Baby", "kytebaby.com", "Bambu Rayon Battaniye",
         "Bambu rayon; ultra yumuşak; bebek ve yetişkin; nefes alabilir; hipoalerjenik",
         "2018'de kuruldu; bebek markasından genişleme; DTC Shopify; Instagram; TikTok; $100M+ gelir"),

        ("Solly Baby", "sollybaby.com", "Organik Müslin Battaniye",
         "GOTS organik pamuk; müslin; bebek ve throw; minimalist tasarım",
         "2018 sonrası genişleme; bebek taşıma markasından battaniye; DTC; Instagram; anne segment"),

        ("Coyuchi", "coyuchi.com", "Organik Pamuk Throw",
         "GOTS organik sertifikalı; sürdürülebilir; doğal renkler; premium organik ev tekstili",
         "2018 sonrası DTC genişleme; organik ev tekstili öncüsü; Shopify; sürdürülebilir lüks"),

        ("Viso Project", "visoproject.com", "Tasarım Throw",
         "Sanatçı kolaborasyonları; limited edition; merinos yünü; koleksiyon parçası",
         "2019'da kuruldu; sanat x ev tekstili; DTC; tasarım odaklı premium; niş koleksiyon"),

        ("Ecoloom", "ecoloom.com", "Sürdürülebilir Throw",
         "Geri dönüştürülmüş pamuk ve plastik; sürdürülebilir üretim; boho desenler",
         "2020'de kuruldu; sürdürülebilir DTC; Shopify; çevre odaklı pazarlama; artisan üretim"),

        ("Sobel Westex", "sobelathome.com", "Otel Tarzı Throw",
         "5 yıldızlı otel tedarikçisi; DTC eve taşıma; premium pamuk ve mikrolif",
         "2018 sonrası DTC büyüme; otel sektöründen tüketici DTC; Shopify; premium otel deneyimi"),

        ("Lusso", "lussoblanket.com", "İtalyan Tarzı Throw",
         "Premium Italian-style kumaş; chenille ve faux-fur; lüks ev aksesuarı",
         "2020'de kuruldu; DTC lüks throw; Shopify; Instagram; premium hediye segment"),

        ("Maker & Son", "makerandson.com", "Artisan Yün Throw",
         "İngiliz artisan üretim; doğal yün; sürdürülebilir; el işçiliği; premium",
         "2019'da kuruldu; İngiliz DTC; mobilya markasından genişleme; Shopify; premium artisan"),

        ("Happie Habitat", "happiehabitat.com", "Geri Dönüştürülmüş Pamuk Throw",
         "%100 geri dönüştürülmüş pamuk; eğlenceli pop-culture desenleri; ABD üretimi",
         "2019'da kuruldu; sürdürülebilir pop-art DTC; Shopify; Instagram; niş tasarım"),

        ("Sunday Throw", "sundaythrow.com", "Waffle Örgü Throw",
         "Waffle örgü pamuk; hafif ve nefes alabilir; İskandinav minimalizm; all-season",
         "2021'de kuruldu; minimalist DTC; Shopify; Instagram; butik ev tekstili"),

        ("Polartec", "polartec.com", "Teknik Fleece Battaniye",
         "Geri dönüştürülmüş polyester fleece; teknik kumaş; dayanıklı; outdoor + ev",
         "2018 sonrası DTC genişleme; teknik kumaş markası; Shopify; outdoor segment"),

        ("Boll & Branch", "bollandbranch.com", "Organik Waffle Throw",
         "Fair Trade organik pamuk; waffle örgü; premium DTC ev tekstili; B Corp",
         "2018 sonrası throw genişleme; premium organik DTC; Shopify; $100M+ gelir; etik üretim"),

        ("Lands Downunder", "landsdownunder.com", "Alpaka Throw",
         "Bebek alpaka yünü; İskoç tasarım; ultra yumuşak; hipoalerjenik; premium",
         "2018 sonrası DTC genişleme; premium alpaka uzmanı; Shopify; niş lüks doğal fiber"),

        ("Morrow Soft Goods", "morrowsoftgoods.com", "French Linen Throw",
         "Fransız keten; stonewashed; doğal renkler; Kaliforniya tasarımı; premium",
         "2019'da kuruldu; Kaliforniya DTC; Shopify; Instagram; premium keten ev tekstili"),

        ("Magic Linen", "magiclinen.com", "Litvanya Keten Throw",
         "Litvanya keten üretimi; OEKO-TEX; handcrafted; doğal renkler; Avrupa DTC",
         "2018'de kuruldu; Litvanya'dan DTC; Shopify; Etsy; Instagram; premium keten battaniye"),
    ],

    # ─────────────────────────────────────────────────────────────────
    # 6. GİYİLEBİLİR BATTANİYE  (Wearable Blankets)
    # ─────────────────────────────────────────────────────────────────
    "Giyilebilir Battaniye": [
        ("The Oodie", "theoodie.com", "Oversize Hoodie Battaniye",
         "Giyilebilir battaniye kategorisini patlattı; TikTok'ta 1B+ görüntülenme; 300M+ gelir",
         "2018'de Avustralya'da Davie Fogarty tarafından kuruldu; Facebook Ads ile viral; TikTok'ta patlama; $300M+ gelir; online-only; global genişleme"),

        ("The Comfy", "thecomfy.com", "Orijinal Giyilebilir Battaniye",
         "Shark Tank'ta yatırım aldı; oversize sherpa-lined hoodie; ilk orijinal giyilebilir battaniye markası",
         "2018'de kuruldu; Shark Tank Season 10; viral büyüme; DTC ve perakende; $200M+ satış; Amazon best-seller"),

        ("Blanket Hoodie", "blankethoodie.com", "Premium Hoodie Battaniye",
         "Ultra yumuşak sherpa; oversize tasarım; unisex; geniş renk yelpazesi; bütçe dostu",
         "2019'da kuruldu; DTC Shopify; TikTok ve Instagram pazarlama; Oodie alternatifi; hızlı büyüme"),

        ("Snuggy", "getsnuggy.com", "İngiliz Giyilebilir Battaniye",
         "İngiliz DTC; oversize hoodie; sherpa astarlı; eğlenceli desenler; erişilebilir fiyat",
         "2019'da kuruldu; İngiltere DTC; TikTok ve Instagram viral; Shopify; hızlı büyüme"),

        ("Kudd.ly", "kuddly.co", "Premium Giyilebilir Battaniye",
         "İngiliz DTC; ultra yumuşak teddy sherpa; oversize; ağırlıklı battaniye versiyonu da var",
         "2020'de kuruldu; İngiltere'den DTC; TikTok patlama; Instagram; hızlı büyüme; premium"),

        ("Sienna", "sienna.co.uk", "Hoodie Battaniye",
         "Ultra yumuşak fleece; sherpa astar; erişilebilir fiyat; İngiliz DTC; eğlenceli desenler",
         "2019'da kuruldu; İngiliz Amazon ve DTC; bütçe dostu giyilebilir battaniye; hızlı büyüme"),

        ("Bedsure", "bedsurehome.com", "Giyilebilir Battaniye",
         "Amazon best-seller; sherpa hoodie; bütçe dostu; geniş boyut yelpazesi; popüler",
         "2018'de kuruldu; Amazon-first; ev tekstilinden giyilebilir battaniye genişleme; global DTC"),

        ("Cozy Bliss", "cozybliss.com", "Sherpa Giyilebilir Battaniye",
         "Ultra kalın sherpa; oversized hoodie; cep tasarımı; kış sezonu best-seller",
         "2020'de kuruldu; Amazon-native; TikTok pazarlama; DTC; kış sezonluk hızlı büyüme"),

        ("Catalonia", "catalonialiving.com", "Giyilebilir Battaniye (Kollu)",
         "Kollu battaniye tasarımı; sherpa fleece; TV battaniyesi; Amazon popüler",
         "2018'de kuruldu; Amazon-first; giyilebilir battaniye çeşitliliği; DTC genişleme"),

        ("Winthome", "winthome.com", "Hoodie Battaniye",
         "Sherpa ve flannel; oversize; cep detayı; erişilebilir fiyat; Amazon best-seller",
         "2019'da kuruldu; Amazon-native; global e-ticaret; bütçe dostu giyilebilir battaniye"),

        ("FunnyFuzzy", "funnyfuzzy.com", "Evcil Hayvan + İnsan Giyilebilir Battaniye",
         "Evcil hayvan ve sahip eşleşen tasarımlar; TikTok viral; eğlenceli konsept; sevimli desenler",
         "2020'de kuruldu; TikTok viral; evcil hayvan x giyilebilir battaniye niş; DTC Shopify; Instagram; $50M+ gelir"),

        ("Zophia Creations", "zophiacreations.com", "Tasarım Giyilebilir Battaniye",
         "Artisan tasarım; özel desenler; premium kumaş; hediye segment; butik DTC",
         "2020'de kuruldu; butik DTC; Shopify ve Etsy; Instagram pazarlama; niş tasarım"),

        ("Oversized Hoodie Co", "oversizedhoodieco.com", "Mega Oversize Hoodie",
         "Ekstra büyük tasarım; çift kişilik boyut; sherpa astar; TikTok popüler",
         "2020'de kuruldu; TikTok-native; DTC Shopify; viral pazarlama; genç kitle"),

        ("PajamaSutra", "pajamasutra.com", "Loungewear Giyilebilir Battaniye",
         "Loungewear x battaniye; bamboo rayon; rahat ev kıyafeti konsepti; sürdürülebilir",
         "2021'de kuruldu; wellness loungewear DTC; Shopify; Instagram; sürdürülebilir kumaş"),

        ("Lushforest", "lushforest.com", "Bambu Giyilebilir Battaniye",
         "Bambu fiber; hafif; termal düzenleme; dört mevsim; doğal malzeme",
         "2020'de kuruldu; Amazon-native; bambu giyilebilir battaniye niş; DTC; doğal segment"),

        ("Hoodiblankie", "hoodiblankie.com", "Çocuk Giyilebilir Battaniye",
         "Çocuk odaklı tasarımlar; hayvan karakterleri; ultra yumuşak; eğlenceli renkler",
         "2020'de kuruldu; çocuk niş; DTC Shopify; Instagram; anne blog pazarlama; hediye segmenti"),

        ("Onesie Plus", "onesieplus.com", "Adult Onesie Battaniye",
         "Tam vücut onesie battaniye; fermuar kapanış; fleece; yetişkin pijama battaniye",
         "2019'da kuruldu; DTC; Amazon; yetişkin onesie x battaniye; kış sezonluk; eğlenceli"),

        ("Nap Queen", "napqueen.co", "Satin Giyilebilir Battaniye",
         "Satin iç astar; saç ve cilt koruma; gece kullanımı; beauty x comfort konsepti",
         "2021'de kuruldu; beauty x uyku DTC; TikTok; Instagram; niş kadın segment"),

        ("Navarre", "navarreblankets.com", "Lüks Giyilebilir Battaniye",
         "Premium merinos yünü; lüks tasarım; minimal estetik; ev ve outdoor",
         "2020'de kuruldu; premium DTC; Shopify; Instagram; lüks giyilebilir battaniye niş"),

        ("Dutch Decor", "dutchdecor.nl", "Teddi Giyilebilir Battaniye",
         "Hollanda DTC; teddy sherpa; oversize; Avrupa teslimat; erişilebilir fiyat",
         "2019'da DTC genişleme; Hollanda'dan Avrupa DTC; Shopify; sosyal medya pazarlama"),

        ("Utterly Snuggly", "utterlysnuggly.com", "Lüks Sherpa Hoodie",
         "Ultra premium sherpa; ekstra uzun tasarım; cep ve kapüşon; İngiliz butik DTC",
         "2020'de kuruldu; İngiliz butik DTC; Instagram; Shopify; premium giyilebilir battaniye"),

        ("Oversize Blanket Co", "oversizeblanketco.com", "Giant Hoodie Battaniye",
         "Ekstra oversize; giant hood; sherpa astar; geniş renk seçeneği; uygun fiyat",
         "2020'de kuruldu; Avustralya DTC; TikTok; Instagram; Oodie alternatifi; hızlı büyüme"),

        ("Mon Snugg", "monsnugg.com", "Fransız Giyilebilir Battaniye",
         "Fransız tasarımı; premium kumaş; minimalist renkler; Avrupa DTC; hediye odaklı",
         "2021'de kuruldu; Fransa'dan Avrupa DTC; Shopify; Instagram; premium Avrupa segment"),

        ("Koozy", "koozy.co", "Hafif Giyilebilir Battaniye",
         "Hafif kumaş; dört mevsim; nefes alabilir; oversize ama taşınabilir; aktif ev kullanımı",
         "2021'de kuruldu; DTC Shopify; TikTok; yeni nesil giyilebilir battaniye; lightweight niş"),

        ("Cozy Comfort", "cozycomfortstore.com", "Fleece Giyilebilir Battaniye",
         "Ultra kalın fleece; giant hood; kanguru cep; unisex; kış essential",
         "2020'de kuruldu; Amazon-native; DTC genişleme; kış sezonluk best-seller"),

        ("Roore", "roore.com", "Kollu TV Battaniyesi",
         "Kollu tasarım; fleece; uzaktan kumanda cebi; TV izleme odaklı; erişilebilir",
         "2019'da kuruldu; Amazon-native; kollu battaniye niş; DTC; bütçe dostu"),

        ("Slanket", "slanket.com", "Orijinal Kollu Battaniye",
         "Kollu battaniye konseptinin modernize edilmesi; fleece; geniş boyut; DTC revival",
         "2018 sonrası DTC revival; Shopify; Amazon; kollu battaniye nostalji + yenilik"),

        ("Lazydays", "lazydaysblanket.com", "Premium Lounging Battaniye",
         "Lounging odaklı; ultra soft; oversized; ev ve bahçe kullanımı; premium kumaş",
         "2021'de kuruldu; DTC Shopify; Instagram; premium lounging segment; İngiliz marka"),

        ("Wearable Blanket Co", "wearableblanketco.com", "Zip-Up Giyilebilir Battaniye",
         "Fermuar kapanışlı; tam vücut kaplama; sherpa astar; hareket özgürlüğü",
         "2020'de kuruldu; DTC Shopify; TikTok pazarlama; yenilikçi fermuar tasarımı"),

        ("Sofa Snuggler", "sofasnuggler.com", "Kanepe Battaniyesi Giyilebilir",
         "Kanepe kullanımı odaklı; kollu ve cepli; fleece sherpa; remote cebi",
         "2020'de kuruldu; niş kanepe kullanım DTC; Shopify; Instagram; hediye segment"),

        ("Huggable Hoodie", "huggablehoodie.com", "Aile Giyilebilir Battaniye",
         "Aile boyutları; ebeveyn-çocuk eşleşme; ultra yumuşak; eğlenceli desenler",
         "2021'de kuruldu; aile odaklı DTC; Shopify; Instagram; anne segment; hediye pazarı"),

        ("Starry Night", "starrynightblanket.com", "Işıldayan Giyilebilir Battaniye",
         "Karanlıkta parlayan yıldız deseni; çocuk ve yetişkin; eğlenceli; hediye",
         "2021'de kuruldu; TikTok viral; DTC Shopify; eğlenceli konsept; çocuk + yetişkin segment"),
    ],
}
