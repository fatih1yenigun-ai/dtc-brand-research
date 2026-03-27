"""
yuvacim_v3_batch6.py
Post-2018 ecommerce-native DTC brands — HOME / BEDDING / ORGANIZER niche
8 categories, 220+ brands total
All descriptions in Turkish (brand names & URLs in English)
"""

BATCH_DATA = {

    # ─────────────────────────────────────────────────────────────────
    # 1. GİYİLEBİLİR BATTANİYE  (Wearable Blankets / Hoodie Blankets / Heated Blankets)
    # ─────────────────────────────────────────────────────────────────
    "Giyilebilir Battaniye": [
        ("The Comfy", "thecomfy.com", "Giyilebilir Battaniye",
         "Oversize sherpa hoodie battaniye; tek beden; cep detaylı; unisex tasarım",
         "2018'de Shark Tank'ta yatırım aldı; TikTok ve Instagram viral büyüme; Amazon #1 bestseller; DTC Shopify mağazası"),

        ("Snuggie", "mysnuggie.com", "Kollu Battaniye",
         "Kollu polar battaniye; TV battaniyesi konsepti; hafif ve sıcak tutar",
         "Klasik konsept ancak 2019 sonrası DTC yeniden lansmanı; sosyal medya kampanyaları; Amazon DTC odaklı satış"),

        ("Oodie", "theoodie.com", "Oversize Hoodie Battaniye",
         "Ultra yumuşak sherpa astar; eğlenceli desenler; oversize kapüşonlu tasarım; unisex",
         "2018'de Avustralya'da kuruldu; TikTok viral fenomeni; global DTC genişleme; yıllık $200M+ gelir"),

        ("Wearable X Blanket", "wearablexblanket.com", "Giyilebilir Isıtmalı Battaniye",
         "USB ile şarj edilebilir ısıtma; 3 sıcaklık ayarı; polar kumaş; portatif",
         "2020'de kuruldu; Amazon FBA ve Shopify DTC; kış sezonunda viral TikTok kampanyaları"),

        ("Kally Blanket", "kallysleep.com", "Ağırlıklı Giyilebilir Battaniye",
         "Ağırlıklı boncuklu yapı; anksiyete azaltıcı; giyilebilir tasarım; yumuşak bambu kumaş",
         "2019'da İngiltere'de kuruldu; uyku sağlığı odaklı DTC; Instagram ve influencer pazarlama"),

        ("Blanket Hoodie Co", "blankethoodie.com", "Battaniye Hoodie",
         "Premium flannel; kanguru cep; geniş kapüşon; çift taraflı kumaş",
         "2019'da DTC olarak kuruldu; Facebook Ads ve Instagram Reels odaklı büyüme; Shopify native"),

        ("Huggle", "hugglehoodie.com", "Isıtmalı Hoodie Battaniye",
         "Dahili ısıtma paneli; şarj edilebilir batarya; ultra yumuşak peluş; 3 ısı seviyesi",
         "2020'de lansmanlandı; As Seen on TV + DTC hibrit model; Amazon bestseller; kış viral kampanyaları"),

        ("Catalonia Blanket", "cataloniablanket.com", "Giyilebilir TV Battaniyesi",
         "Ayak cebi; ön düğme kapama; sherpa astar; farklı beden seçenekleri",
         "2018 sonrası Amazon DTC markası; sosyal medya reklam odaklı; uygun fiyat segmenti; hızlı büyüme"),

        ("CozyChic Wearable", "cozychicwear.com", "Lüks Giyilebilir Battaniye",
         "Kaşmir karışım kumaş; minimalist tasarım; premium ambalaj; hediye odaklı pazarlama",
         "2021'de kuruldu; Instagram influencer kampanyaları; DTC Shopify; premium fiyat segmenti"),

        ("Nap Blanket", "napblanket.com", "Taşınabilir Battaniye",
         "Kompakt katlanır tasarım; seyahat dostu; snap düğmeli dönüşüm; hafif polar",
         "2020'de kuruldu; Kickstarter başarılı kampanya; DTC ve Amazon; dijital göçebe hedef kitlesi"),

        ("Comfywear", "comfywear.co", "Oversize Battaniye Sweatshirt",
         "360 derece sherpa; ekstra uzun tasarım; derin cep; elastik manşet",
         "2019'da kuruldu; TikTok Shop entegrasyonu; viral unboxing videoları; hızlı global kargo"),

        ("Bedsure Wearable", "bedsuredesign.com", "Çok Fonksiyonlu Battaniye",
         "Düğmeli ve fermuarlı; battaniye-panço dönüşümü; polar fleece; makine yıkanabilir",
         "2018'de Çin'den global DTC; Amazon dominant; sosyal medya ve influencer odaklı; uygun fiyat lider"),

        ("Snug Rug", "snugrug.co.uk", "Lüks Polar Battaniye",
         "Uzun peluş kürk; TV battaniyesi; kapüşonlu versiyon; çocuk ve yetişkin bedenleri",
         "2019'da İngiltere DTC olarak büyüdü; Amazon UK bestseller; Facebook ve Instagram Ads"),

        ("Kuddly", "kuddly.co", "Kapüşonlu Battaniye",
         "Oversized fit; teddy bear kumaş; pastel renkler; evcil hayvan dostu tasarım",
         "2020'de İrlanda'da kuruldu; TikTok viral büyüme; DTC Shopify; hızlı Avrupa genişleme"),

        ("ZonLi Wearable", "zonlihome.com", "Ağırlıklı Giyilebilir Battaniye",
         "Cam boncuk ağırlık; giyilebilir tasarım; pamuk dış yüzey; nefes alabilir",
         "2019'da ağırlıklı battaniye odaklı DTC; Amazon FBA; TikTok uyku sağlığı içerikleri"),

        ("Degrees of Comfort", "degreesofcomfort.com", "Isıtmalı Battaniye",
         "Dual-zone ısıtma; mikrofiber; otomatik kapanma; 6 ısı ayarı; yıkanabilir",
         "2019'da kuruldu; Amazon DTC; akıllı ev ürünleri segmenti; yüksek yıldız puanı"),

        ("Navaris Blanket", "navaris.de", "Giyilebilir Battaniye Pelerin",
         "Pelerin tasarım; düğmeli kapama; hafif polar; ev ve açık hava kullanımı",
         "2019'da Almanya DTC; Amazon Europe genişleme; minimalist Avrupa tasarımı; uygun fiyat"),

        ("Moonbow Blanket", "moonbowblanket.com", "Ay Işığı Battaniye",
         "Glow-in-the-dark tasarım; çocuklara özel; yumuşak mikrofiber; eğlenceli desenler",
         "2021'de kuruldu; TikTok çocuk segmentinde viral; DTC Shopify; hediye pazarı odaklı"),

        ("Sienna Blanket", "siennaliving.co.uk", "Elektrikli Giyilebilir Battaniye",
         "Elektrikli ısıtma; fleece kumaş; uzaktan kumanda; zamanlayıcı; enerji tasarruflu",
         "2020'de İngiltere DTC; Amazon UK; enerji krizi döneminde viral büyüme; uygun fiyat segmenti"),

        ("Rumpl", "rumpl.com", "Outdoor Giyilebilir Battaniye",
         "Geri dönüştürülmüş malzeme; su geçirmez dış yüzey; packable tasarım; kamp ve ev kullanımı",
         "2018 sonrası DTC büyüme; B Corp sertifikalı; outdoor influencer kampanyaları; Shopify DTC"),

        ("Cozee Home", "cozeehome.com", "Elektrikli Battaniye",
         "Hızlı ısınma teknolojisi; peluş kumaş; çift kişilik seçenek; güvenlik sertifikalı",
         "2019'da İngiltere'de DTC; QVC ve online satış; sosyal medya influencer odaklı"),

        ("Big Blanket Co", "bigblanket.co", "Dev Boy Battaniye",
         "10x10 feet büyüklük; sıcak tutma teknolojisi; premium kumaş; aile boyu",
         "2018'de kuruldu; Shark Tank sonrası viral; TikTok ve Instagram; DTC Shopify; benzersiz niş"),

        ("Graiden Blanket", "graidenblanket.com", "Merinos Giyilebilir Battaniye",
         "Merinos yünü; termoregülasyon; antibakteriyel; doğal kumaş; lüks his",
         "2021'de kuruldu; sürdürülebilir moda odaklı DTC; Instagram ve Pinterest pazarlama"),

        ("Minky Couture", "minkycouture.com", "Lüks Peluş Battaniye",
         "Ultra lüks minky kumaş; çift taraflı; özel desenler; hediye kutusu ambalaj",
         "2018'de Utah'ta kuruldu; Instagram ve influencer odaklı DTC; premium hediye segmenti"),

        ("Sunday Citizen", "sundaycitizen.com", "Kristal Ağırlıklı Battaniye",
         "Ametist kristal boncuklar; bambu kumaş; wellness odaklı; farklı ağırlık seçenekleri",
         "2019'da kuruldu; wellness ve uyku DTC; Instagram influencer; Nordstrom işbirliği; premium"),

        ("Hug Sleep", "hugsleep.com", "Giyilebilir Sıkıştırmalı Battaniye",
         "Hafif basınç uygulayan kumaş; uyku tulumu tasarım; anksiyete azaltıcı; nefes alabilir",
         "2021'de Shark Tank'ta lansmanlandı; TikTok viral; uyku sağlığı DTC; Amazon bestseller"),

        ("Barefoot Dreams", "barefootdreams.com", "Ultra Yumuşak Battaniye",
         "CozyChic örgü; lüks his; ünlü tercihi; hediye favorisi; makine yıkanabilir",
         "2018 sonrası DTC pivot; önceden perakende ağırlıklıydı; Instagram ve TikTok ile DTC büyüme"),

        ("Slanket", "slanket.com", "Kollu Battaniye 2.0",
         "Yenilenmiş kollu battaniye; ekstra uzun; sherpa astar; ayak cebi; USB ısıtma opsiyonu",
         "2019'da DTC olarak yeniden lansmanlandı; Shopify native; sosyal medya odaklı yeni nesil marka"),

        ("Woolly Mammoth Merino", "woollymammothmerino.com", "Merinos Battaniye",
         "Yeni Zelanda merinos yünü; outdoor ve ev kullanımı; antibakteriyel; hiper dayanıklı",
         "2019'da kuruldu; Amazon DTC; outdoor ve wellness segmentinde influencer pazarlama"),

        ("Hush Blankets", "hushblankets.com", "Ağırlıklı Giyilebilir Battaniye",
         "Cam boncuk ağırlık; pamuklu dış; soğutma ve ısıtma versiyonları; terapi odaklı",
         "2018'de Kanada'da kuruldu; DTC Shopify; uyku sağlığı odaklı; Instagram ve podcast sponsorlukları"),
    ],

    # ─────────────────────────────────────────────────────────────────
    # 2. HAVLU & BORNOZ & PEŞTEMAL  (Towels / Robes / Turkish Towels)
    # ─────────────────────────────────────────────────────────────────
    "Havlu & Bornoz & Pestemal": [
        ("Weezie", "weezietowels.com", "Monogramlu Havlu",
         "Kişiselleştirilmiş monogram; piké dokuma; uzun lif pamuk; otel kalitesi",
         "2018'de kuruldu; Instagram DTC viral; kişiselleştirme trendi öncüsü; Shopify native; hediye segmenti"),

        ("Onsen", "onsentowel.com", "Waffle Dokuma Havlu",
         "Japon waffle dokuma; hızlı kuruyan; hafif; antibakteriyel; minimal tasarım",
         "2019'da Kickstarter ile lansmanlandı; DTC Shopify; minimalist yaşam influencer'ları; premium segment"),

        ("Riley Home", "rileyhome.com", "Premium Banyo Tekstili",
         "Otel kalitesi havlu; Supima pamuk; spa bornoz; lüks his; dayanıklı",
         "2019'da kuruldu; doğrudan fabrikadan tüketiciye model; DTC Shopify; kalite odaklı pazarlama"),

        ("Boll & Branch", "bollandbranch.com", "Organik Pamuk Havlu",
         "Fair Trade sertifikalı; organik pamuk; çevre dostu üretim; lüks his",
         "2018 sonrası havlu hattı genişlemesi; DTC Shopify; sürdürülebilirlik odaklı; Instagram kampanyaları"),

        ("Coyuchi", "coyuchi.com", "Organik Banyo Tekstili",
         "GOTS sertifikalı organik; doğal boyama; sürdürülebilir üretim; zaman aşımı garantisi",
         "2018 sonrası DTC pivot; B Corp; çevre bilinçli tüketici hedefi; Instagram ve blog odaklı"),

        ("House No. 23", "houseno23.com", "Peştemal & Hamam Havlusu",
         "El dokuması Türk peştemalı; doğal pamuk; çok amaçlı kullanım; plaj ve banyo",
         "2019'da kuruldu; Türk peştemal geleneği DTC; Instagram estetik; Shopify; global kargo"),

        ("Dock & Bay", "dockandbay.com", "Hızlı Kuruyan Havlu",
         "Mikrofiber; ultra kompakt; hızlı kuruyan; canlı desenler; seyahat dostu",
         "2018'de lansmanlandı; Instagram ve TikTok travel influencer'lar; DTC global; Amazon bestseller"),

        ("Brooklinen", "brooklinen.com", "Lüks Banyo Koleksiyonu",
         "Super-plush havlu; waffle bornoz; Türk pamuk; ağırlıklı his; 1500 GSM",
         "2018 sonrası banyo hattı genişlemesi; DTC öncüsü; Instagram ve podcast reklamları; NYC kökenli"),

        ("Parachute", "parachutehome.com", "Premium Havlu & Bornoz",
         "Uzun lif Türk pamuğu; waffle ve klasik dokuma; spa kalitesi bornoz",
         "2018 sonrası genişleme; DTC ev tekstili öncüsü; Instagram estetik; Los Angeles kökenli"),

        ("Linen House", "linenhouse.com.au", "Keten Havlu",
         "Keten-pamuk karışım; vintage yıkama; Avrupa tarzı; nefes alabilir",
         "2019'da DTC genişleme; Avustralya kökenli; Instagram ve Pinterest odaklı pazarlama"),

        ("Kassatex", "kassatex.com", "Lüks Türk Havlu",
         "Aegean pamuk; otel kalitesi; SPA serisi; waffle ve terry seçenekleri",
         "2018 sonrası DTC odaklı pivot; New York tasarım; Instagram influencer; premium segment"),

        ("Laguna Beach Textile", "lbtowels.com", "Plaj Havlusu DTC",
         "Oversize plaj havlusu; cabana çizgili; hızlı kuruyan; canlı renkler",
         "2019'da kuruldu; Instagram lifestyle; DTC Shopify; yaz sezonu viral kampanyaları"),

        ("Mizu Towel", "mizutowel.com", "Gümüş İyonlu Havlu",
         "Gümüş iyon antibakteriyel; koku önleyici; uzun ömürlü temizlik; eco-friendly",
         "2020'de Kickstarter ile lansmanlandı; DTC; hijyen odaklı pazarlama; pandemi döneminde viral"),

        ("Cariloha", "cariloha.com", "Bambu Havlu & Bornoz",
         "Bambu viskoz kumaş; ultra yumuşak; antibakteriyel; sürdürülebilir",
         "2018 sonrası DTC genişleme; bambu tekstil öncüsü; Instagram ve Facebook Ads"),

        ("Kontex Towel", "kontex-shop.com", "Japon İmabari Havlu",
         "İmabari sertifikalı; organik pamuk; minimal tasarım; premium kalite",
         "2019'da global DTC açılım; Japon zanaat geleneği; Instagram minimal estetik"),

        ("Marici Home", "maricihome.com", "El Dokuması Peştemal",
         "Geleneksel Türk dokuması; doğal pamuk; el yapımı saçak; çok amaçlı",
         "2020'de kuruldu; Türk peştemal DTC; Etsy ve Shopify; Instagram influencer"),

        ("Loom Towels", "loomtowels.com", "Sürdürülebilir Havlu",
         "Geri dönüştürülmüş pamuk; düşük su kullanımı; karbon nötr üretim; yumuşak his",
         "2021'de kuruldu; sürdürülebilirlik DTC; Instagram ve TikTok çevre içerikleri"),

        ("Snowe Home", "snowehome.com", "Otel Kalitesi Banyo Seti",
         "Portekiz üretimi; uzun lif pamuk; 700 GSM; klasik beyaz; monogram opsiyonu",
         "2018'de kuruldu; DTC lüks ev; Instagram ve dijital reklam odaklı"),

        ("Slowtide", "slowtide.com", "Sanatçı İşbirlikli Havlu",
         "Sanatçı desenleri; sürdürülebilir üretim; plaj ve yoga; koleksiyon bazlı",
         "2019 sonrası DTC büyüme; surf ve outdoor kültürü; Instagram ve TikTok"),

        ("Towel Supercenter", "towelsupercenter.com", "Toptan DTC Havlu",
         "Salon kalitesi; toplu alım indirimi; çeşitli ağırlık seçenekleri; dayanıklı",
         "2019'da DTC modeline geçiş; B2B ve B2C hibrit; dijital pazarlama odaklı"),

        ("Hammam Linen", "hammamlinen.com", "Türk Hamam Havlusu",
         "Geleneksel Türk hamam havlusu; %100 pamuk; hızlı kuruyan; hafif",
         "2018'de kuruldu; Amazon DTC; Türk tekstil geleneği; uygun fiyat segmenti"),

        ("Nutrl Towel", "nutriltowel.com", "Antimikrobiyal Havlu",
         "Gümüş infüzyon teknolojisi; 3x daha az yıkama; koku önleyici; eko-dostu",
         "2020'de kuruldu; Indiegogo başarılı kampanya; DTC; hijyen teknolojisi odaklı"),

        ("Baina", "bainaessentials.com", "Organik Lüks Havlu",
         "GOTS organik sertifikalı; Avustralya tasarımı; earth tonları; premium ambalaj",
         "2019'da Avustralya'da kuruldu; Instagram estetik; lüks DTC; butik otel işbirlikleri"),

        ("Peshtemal City", "peshtemalcity.com", "Peştemal DTC",
         "El dokuması Türk peştemalı; doğal boyama; 300+ desen; kişiselleştirme",
         "2019'da kuruldu; Türk üreticiden direkt DTC; Etsy ve Shopify; global kargo"),

        ("Tekla", "teklafabrics.com", "Scandinavian Banyo Tekstili",
         "Portekiz üretimi organik pamuk; İskandinav minimalizm; waffle dokuma; nötr tonlar",
         "2019'da Kopenhag'da kuruldu; Instagram estetik; DTC premium; tasarım odaklı"),

        ("Nubi Towels", "nubitowels.com", "Hızlı Kuruyan Spor Havlu",
         "Nano-fiber teknolojisi; ultra hafif; kompakt; spor ve seyahat; antibakteriyel",
         "2020'de kuruldu; fitness influencer pazarlama; Amazon ve DTC; aktif yaşam segmenti"),

        ("Mayde Towel", "maydetowel.com", "Avustralya Plaj Havlusu",
         "Türk pamuk; oversize; çift taraflı desen; plaj ve havuz; tassel detay",
         "2018'de Avustralya'da kuruldu; Instagram lifestyle; surf kültürü DTC; yaz viral"),

        ("Sol Linen", "sol-linen.com", "Akdeniz Keten Havlu",
         "Belçika keteni; Akdeniz esintisi; el boyama; doğal renk paleti; spa kalitesi",
         "2021'de kuruldu; sürdürülebilir lüks DTC; Instagram ve Pinterest; niş premium segment"),
    ],

    # ─────────────────────────────────────────────────────────────────
    # 3. MAKYAJ & KOZMETİK ORGANİZERİ  (Beauty Organizers / Vanity Storage)
    # ─────────────────────────────────────────────────────────────────
    "Makyaj & Kozmetik Organizeri": [
        ("Spacejoy", "spacejoy.com", "Akrilik Makyaj Organizeri",
         "Modüler akrilik kutu sistemi; döner tepsi; çekmeceli; şeffaf tasarım",
         "2019'da kuruldu; Instagram beauty influencer kampanyaları; Amazon DTC; Shopify"),

        ("Luvo Store", "luvostore.com", "LED Işıklı Organizer",
         "Dahili LED aydınlatma; ayna entegreli; USB şarj; akrilik bölmeler",
         "2020'de kuruldu; TikTok beauty viral; DTC Shopify; estetik ve fonksiyon birleşimi"),

        ("Boxy Girl", "boxygirl.com", "Şeffaf Akrilik Organizer",
         "Üst üste istiflenebilir; modüler çekmece sistemi; şeffaf akrilik; premium his",
         "2019'da kuruldu; Instagram beauty guru pazarlama; DTC; özelleştirilebilir konfigürasyon"),

        ("Dollup Beauty", "dollupbeauty.com", "Manyetik Makyaj Paleti",
         "Manyetik yapışma sistemi; taşınabilir palet; duvar montaj; space-saving",
         "2019'da kuruldu; Shark Tank görünümü; DTC Shopify; küçük alan çözümleri"),

        ("Etoile Collective", "etoilecollective.com", "Lüks Makyaj Çantası",
         "Vegan deri; altın fermuar; iç bölmeli; seyahat dostu; monogram opsiyonu",
         "2018'de kuruldu; Instagram luxury beauty; DTC; influencer unboxing; hediye segmenti"),

        ("Fancii", "fancii.com", "Işıklı Makyaj Aynası & Organizer",
         "LED aydınlatmalı ayna; 10x büyütme; organizer tabanı; USB şarj; dokunmatik kontrol",
         "2018'de kuruldu; Amazon DTC; Instagram ve TikTok beauty; teknoloji + güzellik kesişimi"),

        ("Plouie", "plouie.com", "Döner Makyaj Organizeri",
         "360° döner platform; ayarlanabilir raflar; büyük kapasite; şık tasarım",
         "2021'de kuruldu; TikTok beauty organizer trendi; Amazon ve DTC; uygun fiyat"),

        ("Tidy Ups", "tidyups.co", "Bambu Makyaj Organizeri",
         "Bambu malzeme; doğal görünüm; çekmeceli; eko-dostu; minimalist",
         "2020'de kuruldu; sürdürülebilir güzellik trendi; Instagram; Shopify DTC"),

        ("Beauty Box Co", "beautyboxco.com", "Abonelik Organizer Kutusu",
         "Aylık organizer ürünleri; küratörlü seçim; sürpriz kutu; premium ambalaj",
         "2019'da kuruldu; abonelik model DTC; Instagram unboxing; beauty community"),

        ("Sorbus Organizer", "sorbushome.com", "Akrilik Çekmece Seti",
         "Büyük kapasiteli akrilik çekmeceler; istiflenebilir; kozmetik ve mücevher; şeffaf",
         "2018'de kuruldu; Amazon bestseller; TikTok organizasyon videoları; uygun fiyat DTC"),

        ("Glam Caddy", "glamcaddy.com", "Döner Kozmetik Rafı",
         "360° dönen raf; 3 katlı; makyaj fırçası tutucu; kompakt tasarım",
         "2020'de kuruldu; TikTok viral organizer; DTC Shopify; beauty influencer"),

        ("The Makeup Curator", "themakeupcurator.com", "Özel Tasarım Organizer",
         "Kişiselleştirilmiş bölmeler; velvet iç astar; premium ahşap; isim baskısı",
         "2021'de kuruldu; Etsy'den DTC'ye geçiş; Instagram lüks güzellik; hediye pazarı"),

        ("Jouer Cosmetics Org", "jouercosmetics.com", "Profesyonel Organizer",
         "Profesyonel makyaj sanatçısı tasarımı; modüler sistem; taşınabilir valiz tipi",
         "2019'da organizer hattı lansmanı; mevcut kozmetik markası DTC genişleme; Instagram"),

        ("Wander Case", "wandercase.co", "Seyahat Makyaj Organizeri",
         "TSA uyumlu; sızdırmaz bölmeler; kompakt; silikon tutucular; hafif",
         "2020'de kuruldu; seyahat influencer pazarlama; TikTok viral; DTC Shopify"),

        ("Cien Beauty Box", "cienbeauty.com", "Çok Katmanlı Organizer",
         "4 katlı tasarım; takı ve makyaj birleşimi; ayna kapaklı; vegan deri",
         "2021'de kuruldu; Instagram estetik; DTC; premium hediye segmenti"),

        ("Pretty Display", "prettydisplay.com", "Parfüm & Kozmetik Vitrin",
         "Akrilik vitrin; LED arka aydınlatma; parfüm ve cilt bakımı; galeri tarzı",
         "2019'da kuruldu; Instagram shelfie trendi; DTC; lüks banyo dekorasyonu"),

        ("Relavel", "relavel.com", "Büyük Makyaj Çantası",
         "Profesyonel boy; ayarlanabilir bölücüler; su geçirmez; seyahat kilidi",
         "2019'da kuruldu; Amazon DTC bestseller; TikTok MUA videoları; uygun fiyat pro segmenti"),

        ("Adore Beauty Org", "adorebeauty.com.au", "Modüler Vanity Sistemi",
         "Avustralya tasarımı; modüler bağlantı; pastel renkler; genişletilebilir",
         "2019 sonrası organizer hattı lansmanı; DTC Avustralya; Instagram; beauty podcast"),

        ("STORi Clear", "storiclear.com", "Şeffaf Plastik Organizer Seti",
         "Stackable tasarım; çekmeceli; bölmeli; banyo ve vanity; kristal şeffaf",
         "2019'da kuruldu; Amazon DTC; TikTok organizasyon ASMR; uygun fiyat; yüksek hacim"),

        ("Impressions Vanity", "impressionsvanity.com", "Hollywood Vanity Ayna & Organizer",
         "Hollywood tarzı LED ayna; entegre organizer; şarj portu; profesyonel aydınlatma",
         "2018 sonrası DTC büyüme; YouTube beauty guru favorisi; Instagram; premium segment"),

        ("Caboodles", "caboodles.com", "Retro Makyaj Kutusu",
         "90'lar retro tasarım; modern renk paleti; katlanır bölmeler; nostaljik DTC",
         "2019'da DTC yeniden lansmanı; TikTok nostalji trendi; Gen Z hedef kitlesi; Amazon"),

        ("Lingerie de Peau Org", "lorgbeauty.com", "İpek Astarlı Organizer",
         "İpek iç astar; toz geçirmez kapak; premium vizon rengi; mücevher bölmesi",
         "2021'de kuruldu; lüks DTC; Instagram luxury beauty; sınırlı üretim"),

        ("Makeup Eraser Caddy", "makeuperaser.com", "Makyaj Temizleme Organizeri",
         "Makyaj temizleme bezi + organizer set; su ile temizleme; yeniden kullanılabilir",
         "2018'de organizer seti lansmanı; TikTok viral; sürdürülebilir güzellik DTC"),

        ("Vanity Collections AU", "vanitycollections.com.au", "Premium Akrilik Organizer",
         "Avustralya tasarımı; lüks akrilik; altın detay; modüler; byob (build your own box)",
         "2018'de kuruldu; Instagram beauty organizasyon; DTC Shopify; Avustralya ve global"),

        ("Cosmocube", "cosmocube.co", "Küp Formlu Organizer",
         "Küp şekilli modüler; üst üste istiflenebilir; şeffaf; minimalist; çekmeceli",
         "2019'da kuruldu; Instagram minimalist estetik; DTC; beauty influencer kampanyaları"),

        ("Beautylish Box", "beautylish.com", "Lüks Güzellik Organizeri",
         "Küratörlü lüks güzellik kutusu; organizer dahil; premium markalar; hediye",
         "2019'da organizer hattı genişlemesi; mevcut beauty DTC; Instagram; premium hediye"),

        ("CLEARorganizer", "clearorganizer.com", "Minimal Akrilik Sistem",
         "Ultra şeffaf akrilik; modüler genişleme; manyetik kapak; stackable",
         "2020'de kuruldu; TikTok #organizewithme trendi; DTC; Marie Kondo etkisi"),

        ("Zoe Ayla", "zoeayla.com", "Kristal Organizer & Güzellik Aracı",
         "Rose quartz detay; kristal roller tutucu; estetik organizer; wellness",
         "2019'da kuruldu; wellness + güzellik DTC; Instagram kristal trendi; TikTok"),
    ],

    # ─────────────────────────────────────────────────────────────────
    # 4. EV KOKUSU & MUM  (Candles / Diffusers / Room Sprays)
    # ─────────────────────────────────────────────────────────────────
    "Ev Kokusu & Mum": [
        ("Boy Smells", "boysmells.com", "Cinsiyet Nötr Mum",
         "Hindistan cevizi-balmumu karışım; gender-neutral kokular; cesur ambalaj; çift fitil",
         "2018'de Los Angeles'ta kuruldu; Instagram estetik; queer-owned DTC; ünlü işbirlikleri"),

        ("Otherland", "otherland.com", "Sanat Esinli Mum",
         "Soya mumu; sanatçı tasarımı kutular; mevsimsel koleksiyon; vegan; el dökümü",
         "2019'da kuruldu; Instagram-first marka; DTC Shopify; koleksiyon bazlı lansmanlar"),

        ("Homesick Candles", "homesick.com", "Nostalji Kokulu Mum",
         "Şehir ve ülke temalı kokular; kişiselleştirme; soya karışım; hediye odaklı",
         "2018 sonrası DTC büyüme; viral hediye konsepti; Instagram ve Facebook Ads; Amazon"),

        ("Apotheke", "apothekeco.com", "Brooklyn Artisan Mum",
         "El dökümü soya mumu; premium koku yağları; minimalist cam kavanoz; lüks",
         "2018 sonrası DTC genişleme; Instagram ve influencer; Brooklyn kökenli; butik otel"),

        ("Voluspa", "voluspa.com", "Dekoratif Cam Mum",
         "Hindistan cevizi mumu; dekoratif cam kavanozlar; uzun yanma süresi; hediye kutusu",
         "2019 sonrası DTC pivot; premium perakendeden online geçiş; Instagram estetik"),

        ("P.F. Candle Co", "pfcandleco.com", "El Yapımı Soya Mum",
         "Soya mumu; amber kavanoz; incense ve oda spreyi hattı; doğal malzeme",
         "2018 sonrası DTC ağırlıklı büyüme; Los Angeles; Instagram; indie marka estetiği"),

        ("Vitruvi", "vitruvi.com", "Taş Difüzör",
         "Porselen difüzör; minimal tasarım; doğal esansiyel yağlar; akıllı zamanlayıcı",
         "2018'de kuruldu; wellness DTC; Instagram minimal estetik; Shopify; premium"),

        ("Candle Shack Studio", "candleshack.co.uk", "DIY Mum Kiti",
         "Evde mum yapım kiti; premium koku yağları; doğal soya mumu; hediye seti",
         "2019'da DTC genişleme; İngiltere; pandemi döneminde viral; Instagram DIY trendi"),

        ("Brooklyn Candle Studio", "brooklyncandlestudio.com", "Minimalist Mum",
         "Soya mumu; minimalist etiket; botanik kokular; sürdürülebilir ambalaj",
         "2018 sonrası DTC büyüme; Brooklyn artisan; Instagram estetik; Shopify native"),

        ("Floral Street", "floralstreet.com", "Çiçek Kokulu Mum & Difüzör",
         "Vegan; geri dönüşümlü ambalaj; çiçek bazlı kokular; eco-conscious",
         "2018'de Londra'da kuruldu; B Corp; Instagram; sürdürülebilir lüks DTC"),

        ("Capri Blue", "capri-blue.com", "Signature Koku Mum",
         "İkonik Volcano kokusu; mavi kavanoz; uzun yanma; hediye favorisi",
         "2019 sonrası DTC odaklı büyüme; TikTok viral; Anthropologie işbirliği; Instagram"),

        ("Keap Candles", "keapbk.com", "Doğal Balmumu Mum",
         "%100 balmumu; pamuk fitil; doğal koku; sürdürülebilir üretim; refill programı",
         "2019'da kuruldu; sürdürülebilir DTC; abonelik modeli; Instagram eco-conscious"),

        ("Maison Louis Marie", "maisonlouismarie.com", "Botanik Mum & Difüzör",
         "Botanik bahçe ilhamlı; clean koku; soya karışım; şık ambalaj",
         "2018 sonrası DTC genişleme; Instagram luxury; clean beauty etkisi; premium segment"),

        ("Yield Design", "yielddesign.co", "Minimal Mum & Difüzör",
         "Cam ve beton kavanoz; esansiyel yağ difüzörü; minimal tasarım; ABD üretimi",
         "2018'de kuruldu; Instagram minimal ev dekor; DTC Shopify; tasarım odaklı"),

        ("Paddywax", "paddywax.com", "Vintage Kavanoz Mum",
         "Geri dönüştürülmüş cam; soya mumu; vintage etiket; library koleksiyonu",
         "2018 sonrası DTC büyüme; hediye segmenti; Instagram estetik; Amazon bestseller"),

        ("Loewe Home Scents", "loewe.com", "Lüks Ev Kokusu",
         "Botanik mumlar; el yapımı seramik kavanoz; lüks doğal kokular; sanatsal tasarım",
         "2020'de ev kokusu hattı lansmanı; Instagram luxury; DTC ve butik perakende"),

        ("Nest New York", "nestnewwork.com", "Premium Ev Kokusu",
         "Reed difüzör; oda spreyi; mum üçlüsü; ikonik Bamboo kokusu; hediye seti",
         "2018 sonrası DTC ağırlıklı büyüme; Instagram ve TikTok; premium hediye segmenti"),

        ("Byredo", "byredo.com", "Lüks İsveç Mum",
         "Minimalist İsveç tasarımı; balmumu mum; eşsiz koku kombinasyonları; koleksiyon",
         "2019 sonrası DTC online genişleme; Instagram luxury; hype kültürü; premium"),

        ("Earl of East", "earlofeast.com", "Londra Artisan Mum",
         "El dökümü soya; seramik kavanoz; doğal fitil; atölye deneyimi; refill",
         "2018'de Londra'da kuruldu; DTC ve atölye hybrid; Instagram; sürdürülebilir"),

        ("Skandinavisk", "skandinavisk.com", "İskandinav Koku",
         "İskandinav doğa ilhamlı; geri dönüşümlü cam; vegan soya; hygge konsepti",
         "2018 sonrası DTC genişleme; İskandinav yaşam tarzı; Instagram; B Corp sertifikalı"),

        ("Lulu Candles", "lulucandles.com", "Kişiselleştirilebilir Mum",
         "Özel etiket baskısı; soya mumu; 50+ koku; hediye ve düğün; kişiselleştirme",
         "2019'da kuruldu; Amazon DTC; kişiselleştirme trendi; hediye segmenti; TikTok"),

        ("Nette", "nette.nyc", "Clean Lüks Mum",
         "Temiz içerik; hindistan cevizi ve soya karışım; minimalist cam; sosyal fayda",
         "2019'da New York'ta kuruldu; clean luxury DTC; Instagram; bağış odaklı marka"),

        ("Canopy", "getcanopy.co", "Akıllı Difüzör",
         "Filtreli nemlendirici-difüzör; anti-küf teknoloji; aromaterapi; modern tasarım",
         "2020'de kuruldu; TikTok wellness viral; DTC Shopify; teknoloji + ev kokusu"),

        ("Overose", "overose.com", "Pembe Neon Mum",
         "Pembe balmumu; bold kokular; neon ambalaj; cesur estetik; limited edition",
         "2019'da Fransa'da kuruldu; Instagram estetik; DTC; neon trend öncüsü"),

        ("Dilo", "shopdilo.com", "Slow Living Mum",
         "Hindistan cevizi mumu; pamuk fitil; botanik kokular; el etiketli; küçük parti",
         "2021'de kuruldu; slow living DTC; Instagram; sürdürülebilir yaşam topluluğu"),

        ("Aesop Aromatics", "aesop.com", "Premium Oda Kokusu",
         "Aromatik oda spreyi; botanik formül; minimalist ambalaj; ikonik tasarım",
         "2019 sonrası ev kokusu hattı genişlemesi; DTC online odaklı; Instagram; premium"),

        ("Pura", "pura.com", "Akıllı Ev Difüzörü",
         "App kontrollü difüzör; zamanlayıcı; parfümeri markaları ile işbirliği; refill sistemi",
         "2019'da kuruldu; TikTok viral; DTC subscription; akıllı ev entegrasyonu; hızlı büyüme"),

        ("Lake & Skye", "lakeandskye.com", "Clean Koku Mum",
         "11:11 ikonik koku; clean formül; soya mumu; kristal enerji; wellness",
         "2018'de kuruldu; TikTok viral koku; DTC; wellness ve spirituel pazarlama"),
    ],

    # ─────────────────────────────────────────────────────────────────
    # 5. HALI & KİLİM  (Washable Rugs / Area Rugs / Runners)
    # ─────────────────────────────────────────────────────────────────
    "Halı & Kilim": [
        ("Ruggable", "ruggable.com", "Yıkanabilir Halı",
         "Makinede yıkanabilir; iki parçalı sistem (üst halı + alt ped); 200+ desen; dayanıklı",
         "2018 sonrası hızlı DTC büyüme; Instagram ve TikTok viral; Shopify; ABD pazar lideri"),

        ("Tumble", "tumblehome.com", "Yıkanabilir Çocuk Halısı",
         "Makinede yıkanabilir; leke tutmaz; çocuk ve evcil hayvan dostu; modern desenler",
         "2021'de kuruldu; TikTok anne influencer'lar; DTC Shopify; aile odaklı pazarlama"),

        ("Rugs USA", "rugsusa.com", "Online Halı DTC",
         "Geniş ürün yelpazesi; uygun fiyat; modern ve vintage desenler; hızlı kargo",
         "2018 sonrası DTC odaklı pivot; Instagram ev dekor; dijital pazarlama; hızlı büyüme"),

        ("Boutique Rugs", "boutiquerugs.com", "Uygun Fiyat Dekoratif Halı",
         "Fabrikadan doğrudan; %80'e varan indirim; modern boho desenler; hızlı teslimat",
         "2018'de kuruldu; Instagram influencer kampanyaları; DTC; uygun fiyat lüks görünüm"),

        ("Loloi Rugs", "loloirugs.com", "Tasarımcı İşbirlikli Halı",
         "Chris Loves Julia, Amber Lewis koleksiyonları; vintage yıkama; modern klasik",
         "2018 sonrası DTC genişleme; influencer koleksiyonları; Instagram ev dekor; Amazon"),

        ("Revival Rugs", "revivalrugs.com", "Vintage El Dokuması Halı",
         "Otantik vintage Türk halıları; el dokuması; tek parça; restore edilmiş",
         "2019'da kuruldu; Türk vintage halı DTC; Instagram ev dekor trendi; Shopify"),

        ("House of Perna", "houseofperna.com", "Sanatsal Halı",
         "Sanatçı işbirlikli desenler; el tuftingi; bold renkler; statement parça",
         "2021'de kuruldu; Instagram sanat + ev dekor; DTC; sınırlı üretim koleksiyonlar"),

        ("Tulu Textiles", "tulutextiles.com", "Türk Tulu Halı",
         "Geleneksel tulu tekniği; el dokuması; doğal yün; modern renkler",
         "2019'da kuruldu; Türk dokuma geleneği DTC; Instagram; artisan üretim"),

        ("Hook & Loom", "hookandloom.com", "Eko-Dostu El Dokuması Halı",
         "Geri dönüştürülmüş malzeme; el dokuması; organik pamuk; düz dokuma",
         "2018 sonrası DTC genişleme; sürdürülebilir ev dekor; Instagram; eko-bilinçli"),

        ("Miss Amara", "missamara.com.au", "AI Destekli Halı Seçimi",
         "Yapay zeka ile oda eşleştirme; AR önizleme; geniş koleksiyon; kolay iade",
         "2018'de Avustralya'da kuruldu; teknoloji + ev dekor DTC; Instagram; AI özelliği"),

        ("RugPad USA", "rugpadusa.com", "Premium Halı Altlığı",
         "Kaymaz altlık; halı koruma; çevre dostu malzeme; özel kesim; niş DTC",
         "2019'da kuruldu; niş ürün DTC; Amazon bestseller; ev dekor influencer"),

        ("Well Woven", "wellwoven.com", "Modern Uygun Fiyat Halı",
         "Makine dokuması; modern geometrik desenler; dayanıklı; uygun fiyat; çeşitli boyut",
         "2018 sonrası DTC odaklı büyüme; Amazon bestseller; Instagram ev dekor; hızlı kargo"),

        ("Jill Zarin Rugs", "jillzarinrugs.com", "Ünlü İmzalı Halı",
         "Jill Zarin tasarımı; modern klasik; power-loomed; dayanıklı; outdoor seçenekleri",
         "2019'da kuruldu; ünlü DTC markası; Instagram; TV personality branding"),

        ("Tapis Rouge", "tapisrougestyle.com", "Lüks Fas Halısı",
         "El dokuması Fas Beni Ourain; doğal yün; boho-chic; vintage",
         "2019'da kuruldu; Fas zanaatkar DTC; Instagram boho trendi; etik üretim"),

        ("Burrow Rugs", "burrow.com", "Modüler Yıkanabilir Halı",
         "Modüler kare sistem; yıkanabilir; değiştirilebilir parçalar; leke direnci",
         "2021'de halı hattı lansmanı; mevcut DTC mobilya markası genişleme; Instagram"),

        ("Kalim", "kalim.co", "Türk Kilim DTC",
         "El dokuması Anadolu kilimi; vintage restore; modern renkler; düz dokuma",
         "2020'de kuruldu; Türk kilim DTC global; Instagram ev dekor; artisan işbirliği"),

        ("Sixhands", "sixhands.co", "Vintage Halı Marketplace",
         "Küratörlü vintage halı koleksiyonu; Türk, Fars, Fas; tek parça; premium",
         "2019'da kuruldu; vintage halı DTC; Instagram ev dekor; collectors segment"),

        ("Dash & Albert", "dashandalbert.com", "Yıkanabilir Indoor-Outdoor Halı",
         "İç ve dış mekan; makinede yıkanabilir; dayanıklı; preppy desenler; kolay bakım",
         "2018 sonrası DTC genişleme; Instagram ev dekor; Annie Selke markası; online pivot"),

        ("Rifle Paper Co Rugs", "riflepaperco.com", "Çiçek Desenli Halı",
         "El yapımı görünümlü çiçek desenleri; Loloi işbirliği; renkli; statement",
         "2019'da halı koleksiyonu lansmanı; Instagram viral; sanatsal DTC; hediye segmenti"),

        ("Nani Marquina", "nanimarquina.com", "İspanyol Tasarım Halısı",
         "El dokuması; tasarımcı işbirliği; sürdürülebilir üretim; çağdaş sanat",
         "2018 sonrası DTC online genişleme; İspanya tasarım; Instagram; premium segment"),

        ("Coral & Tusk Rugs", "coralandtusk.com", "Nakış İşlemeli Halı",
         "El nakışı detay; hayvan motifleri; sanatsal; sınırlı üretim; koleksiyon",
         "2019'da halı hattı lansmanı; Brooklyn artisan DTC; Instagram; niş premium"),

        ("Cali Rugs", "calirugs.com", "Yıkanabilir Outdoor Halı",
         "UV dayanıklı; su geçirmez; makinede yıkanabilir; teras ve balkon; modern",
         "2020'de kuruldu; outdoor living trendi; DTC Shopify; Instagram ev dekor"),

        ("The Citizenry Rugs", "the-citizenry.com", "Fair Trade Halı",
         "El dokuması; fair trade; küresel zanaatkarlar; doğal malzeme; minimal",
         "2018 sonrası halı hattı genişlemesi; etik DTC; Instagram; hikaye odaklı pazarlama"),

        ("Sunday Citizen Rugs", "sundaycitizen.com", "Bambu İplik Halı",
         "Bambu viskoz; ultra yumuşak; hipoalerjenik; modern desenler; makine yıkanabilir",
         "2020'de halı hattı lansmanı; wellness DTC genişleme; Instagram; konfor odaklı"),

        ("Surya Rugs", "surya.com", "Global Tasarım Halı",
         "30,000+ ürün; makine ve el dokuması; modern ve geleneksel; uygun fiyat",
         "2018 sonrası DTC online pivot; Amazon ve Shopify; Instagram ev dekor; geniş ürün"),

        ("Momeni Rugs", "momeni.com", "Tasarımcı Koleksiyon Halı",
         "Novogratz işbirliği; el tuftingi; modern retro; cesur renkler; statement",
         "2019 sonrası DTC genişleme; influencer koleksiyonları; Instagram; tasarım odaklı"),

        ("Pappelina", "pappelina.com", "İsveç Plastik Halı",
         "Geri dönüştürülmüş plastik; iç ve dış mekan; İsveç tasarımı; dayanıklı; renkli",
         "2018 sonrası DTC global genişleme; İskandinav tasarım; Instagram; sürdürülebilir"),
    ],

    # ─────────────────────────────────────────────────────────────────
    # 6. DEKORATİF YASTIK & KIRLENT  (Throw Pillows / Cushion Covers)
    # ─────────────────────────────────────────────────────────────────
    "Dekoratif Yastık & Kırlent": [
        ("Calhoun & Co", "calhounandco.com", "Esprili Nakış Yastık",
         "El nakışı; esprili sloganlar; yün üzerine zincir dikiş; hediye odaklı",
         "2018'de kuruldu; Instagram viral; hediye segmenti DTC; Shopify; cesur marka dili"),

        ("Bolé Road Textiles", "boleroadtextiles.com", "Etiyopya El Dokuması Yastık",
         "Etiyopya zanaatkarları; el dokuması pamuk; geleneksel motifler; fair trade",
         "2018'de kuruldu; etik moda DTC; Instagram; hikaye odaklı pazarlama; Shopify"),

        ("Jungalow", "jungalow.com", "Boho Renkli Kırlent",
         "Justina Blakeney tasarımı; cesur renkler; boho desenler; saçaklı detay",
         "2018 sonrası DTC genişleme; Instagram boho ev dekor; influencer marka; Target işbirliği"),

        ("Society6 Pillows", "society6.com", "Sanatçı Tasarımı Yastık",
         "Bağımsız sanatçı desenleri; on-demand baskı; geniş seçenek; custom boyut",
         "2018 sonrası pillow hattı büyümesi; sanatçı topluluk DTC; Instagram"),

        ("Serena & Lily", "serenaandlily.com", "Kıyı Tarzı Lüks Kırlent",
         "Coastal chic tasarım; premium kumaşlar; mavi-beyaz paleti; el işçiliği",
         "2018 sonrası DTC genişleme; Instagram ev dekor; premium segment; California stili"),

        ("Parachute Pillows", "parachutehome.com", "Premium Dekoratif Yastık",
         "Keten ve pamuk; nötr tonlar; minimal tasarım; Oeko-Tex sertifikalı",
         "2019 sonrası yastık hattı genişlemesi; DTC ev tekstili; Instagram; LA kökenli"),

        ("The Citizenry Pillow", "the-citizenry.com", "El Yapımı Dünya Yastıkları",
         "Global zanaatkarlar; el dokuması; doğal boyama; fair trade sertifikalı",
         "2018 sonrası genişleme; etik DTC; Instagram; seyahat ilhamlı tasarım"),

        ("Studio Pillows", "studiopillows.com", "Tasarımcı Kumaş Kırlent",
         "Tasarımcı kumaşlardan üretim; özel kesim; premium dolgu; sipariş üzerine",
         "2019'da kuruldu; Etsy'den DTC Shopify'a geçiş; Instagram iç mimarlık"),

        ("MINNA", "mightyminna.com", "Etik El Dokuması Yastık",
         "Latin Amerika zanaatkarları; el dokuması; doğal boyama; modern geometrik",
         "2018'de kuruldu; B Corp DTC; Instagram; etik üretim hikayesi; Shopify"),

        ("CLO Studios", "clostudios.com", "3D Örgü Yastık",
         "Chunky örgü; el yapımı; merinos yünü; 3D dokulu; statement parça",
         "2020'de kuruldu; Instagram viral; el yapımı DTC; modern zanaat; Shopify"),

        ("Ferm Living Cushions", "fermliving.com", "Danimarka Tasarım Kırlent",
         "İskandinav minimalizm; organik pamuk; geometrik desenler; nötr tonlar",
         "2018 sonrası DTC online genişleme; Kopenhag tasarım; Instagram minimal ev"),

        ("Dusen Dusen", "dusendusen.com", "Pop Art Kırlent",
         "Cesur grafikler; canlı renkler; modern pop art; eğlenceli desenler",
         "2019'da ev hattı lansmanı; Brooklyn DTC; Instagram; renkli ev dekor trendi"),

        ("Eny", "enyliving.com", "Sürdürülebilir Pamuk Kırlent",
         "Organik pamuk; doğal boyama; minimalist; yeniden kullanılabilir ambalaj",
         "2021'de kuruldu; sürdürülebilir DTC; Instagram eco-living; Shopify"),

        ("Block Shop Textiles", "blockshoptextiles.com", "Baskı Tekniği Yastık",
         "Hindistan blok baskı; el yapımı; doğal boyama; sanatsal; fair trade",
         "2018'de kuruldu; artisan DTC; Instagram zanaat; etik üretim; Shopify"),

        ("Ara Collective", "aracollective.com", "Meksika El Dokuması Kırlent",
         "Oaxaca zanaatkarları; geleneksel dokuma; modern renkler; etik üretim",
         "2019'da kuruldu; Latin Amerika zanaat DTC; Instagram; fair trade odaklı"),

        ("Loomination", "loomination.com", "Makrome Dekoratif Yastık",
         "El yapımı makrome; boho tasarım; pamuk iplik; saçaklı detay; doğal",
         "2019'da kuruldu; boho DTC; Instagram ve Etsy; el yapımı zanaat"),

        ("Leah Singh", "leahsingh.com", "Punch Needle Yastık",
         "Punch needle tekniği; Hindistan üretimi; renkli geometrik; el yapımı",
         "2018'de kuruldu; artisan DTC; Instagram zanaat trendi; fair trade; Shopify"),

        ("Joanna Buchanan", "joannabuchanan.com", "Lüks Dekoratif Yastık",
         "İpek ve kadife; boncuk işleme; böcek ve hayvan motifleri; hediye odaklı",
         "2019'da kuruldu; lüks ev dekor DTC; Instagram; premium hediye segmenti"),

        ("Edie Parker Home", "edieparker.com", "Neon Akrilik Aksesuar Yastık",
         "Pop renk paleti; esprili tasarım; modern retro; statement dekor parçası",
         "2020'de ev hattı lansmanı; mevcut aksesuar markası genişleme; Instagram"),

        ("Justina Blakeney Pillow", "justinablakeney.com", "Boho Lüks Kırlent",
         "Tufted detay; çoklu doku; boho maximalist; doğal ve canlı renkler",
         "2019'da bağımsız koleksiyon lansmanı; Instagram ev dekor influencer; DTC"),

        ("Hawkins New York", "hawkinsnewyork.com", "Minimal Keten Yastık",
         "Yıkanmış keten; stonewashed; minimal renk paleti; Avrupa üretimi; premium",
         "2018'de kuruldu; minimal DTC; Instagram; New York tasarım; Shopify"),

        ("Kip & Co", "kipandco.com.au", "Renkli Avustralya Kırlent",
         "Cesur renkler; eğlenceli desenler; kadife ve pamuk; Avustralya tasarımı",
         "2018 sonrası DTC büyüme; Avustralya ev dekor; Instagram; neşeli estetik"),

        ("Goodee Pillows", "goodee.com", "Etik Seçilmiş Kırlent",
         "Küratörlü etik markalar; B Corp ürünler; sürdürülebilir; global zanaatkarlar",
         "2019'da kuruldu; etik marketplace DTC; Instagram; the Olsen twins desteği"),

        ("Casa Amarosa", "casaamarosa.com", "Hindistan El İşi Kırlent",
         "Block print; el nakışı; tye-dye; Hint zanaatkarları; boho-chic",
         "2020'de kuruldu; artisan DTC; Instagram boho ev; etik üretim; Shopify"),

        ("Lil Hummingbird", "lilhummingbird.com", "Organik Çocuk Kırlent",
         "GOTS organik; çocuk güvenli; hayvan figürleri; yumuşak dolgu; yıkanabilir",
         "2021'de kuruldu; çocuk odası DTC; Instagram anne influencer; Shopify"),

        ("Anchal Project", "anchalproject.org", "Sosyal Girişim Kırlent",
         "Eski sari kumaşından; kantha dikişi; sosyal fayda; Hindistan kadın kooperatifi",
         "2018'de kuruldu; sosyal girişim DTC; Instagram; etik moda ve ev dekor"),

        ("Morrow Soft Goods", "morrowsoftgoods.com", "Pamuk Yastık Koleksiyonu",
         "Stonewashed pamuk; rahat lüks; nötr tonlar; LA tarzı casual; premium",
         "2019'da kuruldu; LA lifestyle DTC; Instagram; hotel-chic ev dekor"),

        ("Celadon Home", "celadonhome.co", "Seramik İlhamlı Kırlent",
         "Seramik renk paleti; organik formlar; doğal kumaşlar; sanatsal",
         "2021'de kuruldu; sanat + ev dekor DTC; Instagram; niş tasarım"),
    ],

    # ─────────────────────────────────────────────────────────────────
    # 7. BEBEK UYKU & TEKSTİL  (Baby Bedding / Swaddles / Sleep Sacks)
    # ─────────────────────────────────────────────────────────────────
    "Bebek Uyku & Tekstil": [
        ("Kyte Baby", "kytebaby.com", "Bambu Uyku Tulumu",
         "Bambu rayon; TOG derecelendirmeli; nefes alabilir; termoregülasyon; yumuşak",
         "2018'de kuruldu; Instagram ve TikTok anne influencer viral; DTC Shopify; hızlı büyüme"),

        ("Solly Baby", "sollybaby.com", "Organik Bebek Kundağı",
         "GOTS organik pamuk; hafif wrap; nefes alabilir; çeşitli renk; bebek taşıma",
         "2018'de kuruldu; Instagram anne topluluğu; DTC Shopify; premium bebek markası"),

        ("Gunamuna", "gunamuna.com", "Fermuarlı Uyku Tulumu",
         "Alt değiştirme fermuarı; bambu kumaş; TOG seçenekleri; güvenli uyku",
         "2019'da kuruldu; Amazon ve DTC; anne influencer pazarlama; pratik tasarım"),

        ("Nested Bean", "nestedbean.com", "Ağırlıklı Uyku Tulumu",
         "Hafif göğüs ağırlığı; bebek güvenliği sertifikalı; rahatlatıcı etki; organik",
         "2018'de kuruldu; Shark Tank; DTC ve Amazon; uyku sağlığı odaklı"),

        ("Aden + Anais", "adenandanais.com", "Müslin Kundak Bezi",
         "Premium müslin; nefes alabilir; çok amaçlı; organik pamuk seçenekleri",
         "2018 sonrası DTC odaklı büyüme; Instagram anne; Amazon bestseller; klasik marka"),

        ("Dreamland Baby", "dreamlandbabyco.com", "Hafif Ağırlıklı Uyku Tulumu",
         "CalmWear teknolojisi; hafif boncuk ağırlığı; güvenli uyku; organik kumaş",
         "2019'da Shark Tank'ta lansmanlandı; TikTok anne viral; DTC; uyku çözümü"),

        ("Copper Pearl", "copperpearl.com", "Stretchy Kundak Bezi",
         "Ultra esnek jersey; çok amaçlı; emzirme örtüsü; araba koltuğu örtüsü",
         "2018'de kuruldu; Instagram ve TikTok; DTC; Amazon bestseller; uygun fiyat premium"),

        ("Mushie", "mushie.com", "İskandinav Bebek Tekstili",
         "Müslin kundak; silikon bebek ürünleri; minimal İskandinav tasarım; pastel",
         "2018'de kuruldu; Instagram anne estetik; DTC global; İskandinav tasarım trendi"),

        ("Halo Sleep", "halosleep.com", "SIDS Önleyici Uyku Tulumu",
         "Geri alınabilir kundak; SIDS güvenliği; hastane onaylı; pamuk ve polar",
         "2018 sonrası DTC genişleme; güvenlik odaklı pazarlama; Amazon ve Shopify; anne topluluğu"),

        ("Love to Dream", "lovetodream.com", "Kol Yukarı Kundak",
         "Patentli kol yukarı tasarım; doğal uyku pozisyonu; geçiş uyku tulumu",
         "2018 sonrası DTC global genişleme; Avustralya kökenli; Instagram; Amazon bestseller"),

        ("Woolino", "woolino.com", "Merinos Yünü Uyku Tulumu",
         "Merinos yünü; 4 mevsim; termoregülasyon; 2 ay - 2 yaş; doğal",
         "2018'de kuruldu; premium DTC; Amazon; doğal malzeme odaklı; anne blog yazarları"),

        ("Burt's Bees Baby", "burtsbeesbaby.com", "Organik Bebek Tekstili",
         "GOTS organik pamuk; onesie ve uyku tulumu; doğal boyama; çevre dostu",
         "2018 sonrası DTC genişleme; organik bebek segmenti; Instagram anne; Amazon"),

        ("Snoo by Happiest Baby", "happiestbaby.com", "Akıllı Bebek Beşiği Tekstili",
         "Snoo beşik kundağı; güvenlik klipsi; organik pamuk; otomatik sallanma sistemi ile uyumlu",
         "2018 sonrası DTC büyüme; Dr. Harvey Karp; TikTok viral; teknoloji + bebek"),

        ("Ely's & Co", "elysandco.com", "Bebek Çarşaf Seti",
         "Jersey pamuk; fitted çarşaf; su geçirmez pad; nefes alabilir; crib uyumlu",
         "2019'da kuruldu; Amazon DTC; anne topluluğu; uygun fiyat premium; Instagram"),

        ("Kantha Bae", "kanthabae.com", "Kantha Bebek Battaniyesi",
         "El dikişi kantha; vintage sari kumaş; her parça benzersiz; Hindistan üretimi",
         "2019'da kuruldu; artisan DTC; Instagram boho anne; etik üretim; Shopify"),

        ("Snuggle Me Organic", "snugglemeorganic.com", "Organik Bebek Yatağı",
         "Organik pamuk kapak; GOTS sertifikalı; sensörsüz rahatlatma; portatif",
         "2018 sonrası DTC büyüme; Instagram anne; Amazon bestseller; doğal malzeme"),

        ("Pehr", "pfrdesigns.com", "Organik Bebek Nevresim",
         "GOTS organik; el pompom detay; pastel renkler; İskandinav estetik",
         "2018 sonrası DTC büyüme; Instagram anne ve çocuk odası; Shopify; premium"),

        ("Little Unicorn", "littleunicorn.com", "Müslin Bebek Battaniye",
         "Pamuk müslin; cotton quilt; modern desenler; hediye seti; çok amaçlı",
         "2018 sonrası DTC genişleme; Instagram anne topluluğu; Amazon; hediye segmenti"),

        ("Malabar Baby", "malabarbaby.com", "Hindistan Baskı Bebek Tekstili",
         "Block print; el baskısı; organik pamuk; GOTS; geleneksel Hint desenleri",
         "2018'de kuruldu; artisan DTC; Instagram; kültürel zanaat + modern bebek"),

        ("Mebie Baby", "mebiebaby.com", "Muslin Kundak & Battaniye",
         "Bambu müslin; çiçek desenleri; vintage estetik; hediye kutusu; yumuşak",
         "2020'de kuruldu; Instagram anne influencer; DTC Shopify; boho bebek trendi"),

        ("Gathre", "gathre.com", "Deri Bebek Mat & Aksesuar",
         "Bonded deri; kolay temizleme; modern tasarım; oyun matı ve değiştirme pedi",
         "2018'de kuruldu; Instagram anne; DTC Shopify; pratik lüks bebek ürünleri"),

        ("Under the Nile", "underthenile.com", "Organik Mısır Pamuğu Bebek",
         "Mısır organik pamuk; fair trade; doğal boyama; hassas cilt; hipoalerjenik",
         "2018 sonrası DTC genişleme; organik bebek segmenti; Instagram; sürdürülebilir"),

        ("Toki Mats", "tokimats.com", "Bebek Oyun Matı",
         "Organik pamuk kaplama; köpük dolgu; modern desen; katlanabilir; kolay temizlik",
         "2019'da kuruldu; Instagram anne dekor; DTC Shopify; toksik olmayan malzeme"),

        ("Tiny Twinkle", "tinytwinkle.com", "Silikon Önlük & Kundak Seti",
         "Silikon önlük; müslin kundak; modern desen; pratik set; hediye odaklı",
         "2019'da kuruldu; Amazon ve DTC; Instagram anne; pratik bebek çözümleri"),

        ("Crane Baby", "cranebaby.com", "Kenten Bebek Nevresim",
         "Keten-pamuk karışım; nefes alabilir; nötr tonlar; modern minimal; premium",
         "2020'de kuruldu; premium bebek DTC; Instagram; modern çocuk odası estetiği"),

        ("Ollie World", "ollieworld.com", "Patentli Kundak Sistemi",
         "Moisture-wicking kumaş; özel velcro sistemi; kolay kullanım; güvenli",
         "2018'de kuruldu; DTC Shopify; anne influencer; Amazon; patent teknolojisi"),

        ("Saranoni", "saranoni.com", "Lüks Minky Bebek Battaniye",
         "Ultra yumuşak minky; saten kenar; lüks dokunuş; hediye odaklı; çeşitli boyut",
         "2018'de kuruldu; Instagram anne topluluğu; DTC; premium hediye segmenti"),

        ("Stokke Bedding", "stokke.com", "İskandinav Bebek Yatak Tekstili",
         "Stokke beşik uyumlu; organik pamuk; İskandinav tasarım; premium",
         "2019 sonrası DTC online genişleme; İskandinav bebek; Instagram; premium marka"),

        ("Natemia", "natemia.com", "Bambu Müslin Bebek",
         "Bambu müslin kundak; ultra yumuşak; nefes alabilir; pastel tonlar; 4 katman",
         "2019'da kuruldu; Amazon DTC; Instagram anne; doğal malzeme; uygun fiyat premium"),
    ],

    # ─────────────────────────────────────────────────────────────────
    # 8. EV GİYİM & LOUNGEWEAR  (Pajamas / Matching Sets / Homewear)
    # ─────────────────────────────────────────────────────────────────
    "Ev Giyim & Loungewear": [
        ("Lunya", "lunya.com", "Lüks Uyku Giyim",
         "Washable silk pijama; pima pamuk; restorative uyku teknolojisi; premium his",
         "2018 sonrası hızlı DTC büyüme; Instagram ve TikTok; kadın odaklı; LA kökenli"),

        ("Lake Pajamas", "lakepajamas.com", "Pima Pamuk Pijama",
         "Pima pamuk; klasik kesim; aile eşleştirme; monogram; pastel renkler",
         "2018'de kuruldu; Instagram viral; DTC Shopify; matching family trendi; hediye"),

        ("Leset", "lfreset.com", "Minimal Loungewear",
         "Jersey kumaş; minimal tasarım; mix-and-match; şık ve rahat; nötr tonlar",
         "2019'da kuruldu; Instagram minimal moda; DTC; NYC; work-from-home trendi"),

        ("Printfresh", "printfresh.com", "Renkli Baskılı Pijama",
         "Canlı desenler; organik pamuk; eğlenceli baskılar; hediye odaklı; çeşitli beden",
         "2019'da kuruldu; Instagram viral desenler; DTC Shopify; inklüzif beden aralığı"),

        ("Cozy Earth", "cozyearth.com", "Bambu Pijama",
         "Bambu viskoz; termoregülasyon; ultra yumuşak; Oprah's Favorite Things; premium",
         "2018 sonrası DTC büyüme; Oprah onayı ile viral; Instagram ve TikTok; Shopify"),

        ("Skims Loungewear", "skims.com", "Kim K Loungewear",
         "Fits Everybody koleksiyonu; inklüzif beden; nötr tonlar; body-hugging; modal",
         "2019'da kuruldu; Kim Kardashian; Instagram ve TikTok viral; mega DTC; global"),

        ("Eberjey", "eberjey.com", "Gisele Pijama",
         "Modal karışım; ikonik Gisele seti; yumuşak dokunuş; romantik tasarım",
         "2018 sonrası DTC pivot; önceden perakende ağırlıklı; Instagram; premium loungewear"),

        ("Petite Plume", "petiteplume.com", "Klasik Aile Pijama",
         "Pima pamuk; biyeli detay; klasik Amerikan tarzı; aile eşleştirme; monogram",
         "2018'de kuruldu; Instagram aile eşleştirme trendi; DTC; hediye ve tatil sezonu"),

        ("Feathers Sleepwear", "featherssleepwear.com", "Ultra Yumuşak Uyku Giyim",
         "Bambu rayon; termoregülasyon; minimalist kesim; nefes alabilir; unisex",
         "2020'de kuruldu; DTC Shopify; Instagram; konfor odaklı pazarlama; bambu trendi"),

        ("Stars Above", "starsabovesleep.com", "Modern Uyku Koleksiyonu",
         "Sürdürülebilir kumaşlar; modern kesimler; çeşitli beden; uygun fiyat premium",
         "2019'da kuruldu; Target DTC genişleme; Instagram; inklüzif marka"),

        ("Pour Les Femmes", "pourlesfemmes.com", "Etik Lüks Pijama",
         "Organik pamuk; Kongo Demokratik Cumhuriyeti kadın kooperatifi; sosyal fayda",
         "2018'de kuruldu; Robin Wright kurucusu; etik DTC; Instagram; lüks + sosyal misyon"),

        ("The Great Eros", "thegreateros.com", "İtalyan İpek Loungewear",
         "İtalyan ipek; lüks iç çamaşırı ve loungewear; minimal tasarım; NYC üretimi",
         "2018 sonrası DTC genişleme; Instagram lüks; NYC artisan; premium segment"),

        ("Bohemian Traders", "bohemiantraders.com", "Boho Ev Giyim",
         "Keten ve pamuk; boho kesimler; rahat ve şık; Avustralya tasarımı; tatil havası",
         "2018 sonrası DTC büyüme; Avustralya; Instagram lifestyle; boho moda"),

        ("Lahgo", "lahgo.com", "Restoratif Uyku Giyim",
         "Restorative fabric teknolojisi; uyku kalitesi artırıcı; modal karışım; nötr ton",
         "2019'da kuruldu; wellness DTC; uyku bilimi odaklı; Instagram; premium"),

        ("Cosabella", "cosabella.com", "İtalyan Loungewear",
         "İtalyan dantel; Curvy koleksiyonu; extended size; aile işletmesi; lüks",
         "2018 sonrası DTC online pivot; İtalyan kalite; Instagram; inklüzif moda"),

        ("Hanna Andersson", "hannaandersson.com", "Organik Pamuk Aile Pijama",
         "Organik pamuk; İsveç ilhamlı; aile eşleştirme; canlı çizgiler; dayanıklı",
         "2018 sonrası DTC odaklı büyüme; matching family trendi; Instagram; tatil sezonu"),

        ("Papinelle", "papinelle.com", "Avustralya Lüks Uyku Giyim",
         "Modal ve ipek; baskılı pijama; hediye kutusu; feminen tasarım; premium",
         "2018 sonrası DTC global genişleme; Avustralya kökenli; Instagram; hediye segmenti"),

        ("Roller Rabbit", "rollerrabbit.com", "Block Print Pijama",
         "Hindistan block print; canlı desenler; pamuk; aile eşleştirme; tatil koleksiyonu",
         "2018 sonrası DTC genişleme; Instagram; boho chic; tatil sezonu viral; hediye"),

        ("Sleepy Jones", "sleepyjones.com", "Modern Klasik Pijama",
         "Oxford shirting; menswear ilhamlı; unisex; premium pamuk; NYC tasarım",
         "2018 sonrası DTC büyüme; Instagram; modern erkek ve kadın pijama; Shopify"),

        ("Summersalt", "summersalt.com", "Sürdürülebilir Loungewear",
         "Geri dönüştürülmüş malzeme; on-demand üretim; modern kesim; inklüzif beden",
         "2018'de kuruldu; sürdürülebilir DTC; Instagram; body-positive pazarlama"),

        ("Hill House Home", "hillhousehome.com", "Nap Dress & Loungewear",
         "İkonik Nap Dress; ev ve dışarı giyilebilir; çiçek desenler; puf kol; viral",
         "2019'da Nap Dress lansmanı; TikTok ve Instagram mega viral; DTC; waitlist modeli"),

        ("Lunya Restore", "lunya.com", "Restore Uyku Kıyafeti",
         "Celliant fiber; kan dolaşımı destekleyen; uyku kalitesi artırıcı; bilimsel",
         "2020'de Restore hattı lansmanı; uyku teknolojisi DTC; Instagram wellness"),

        ("Naked Cashmere", "nakedcashmere.com", "Kaşmir Loungewear",
         "Saf kaşmir; günlük giyim; lüks his; uygun fiyat kaşmir; DTC direkt",
         "2019'da kuruldu; lüks DTC disruption; Instagram; fabrikadan direkt fiyat"),

        ("Skin", "skin.com", "Organik Pima Pijama",
         "Organik Pima pamuk; yumuşak jersey; ikinci cilt hissi; nötr tonlar; NYC",
         "2018 sonrası DTC genişleme; minimal estetik; Instagram; lüks iç çamaşırı geçmişi"),

        ("Londre Bodywear", "londrebodywear.com", "Sürdürülebilir Ev Giyim",
         "Geri dönüştürülmüş naylon; Kanada üretimi; modern kesim; çevre dostu",
         "2019'da kuruldu; sürdürülebilir DTC; Instagram; Kanada slow fashion"),

        ("Morgan Lane", "morgan-lane.com", "Ipek Lüks Pijama",
         "İpek charmeuse; dantel detay; feminen; lüks hediye; monogram opsiyonu",
         "2018 sonrası DTC genişleme; Instagram lüks yaşam; ünlü favorisi; Shopify"),

        ("Thakoon Loungewear", "thakoon.com", "Tasarımcı Ev Giyim",
         "Tasarımcı kesimler; ev ve dış giyim geçişli; premium kumaşlar; NYC moda",
         "2020'de DTC yeniden lansmanı; ünlü tasarımcı; Instagram; pandemic loungewear trendi"),

        ("The Nap Dress Shop", "napdress.shop", "Ev Elbisesi",
         "Rahat kesim ev elbisesi; çiçek baskı; pijama konforunda şık görünüm; pamuk",
         "2021'de kuruldu; nap dress trendi; DTC Shopify; Instagram; Hill House esinli"),
    ],
}
