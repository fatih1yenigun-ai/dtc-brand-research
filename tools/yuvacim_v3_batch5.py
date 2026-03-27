"""
Batch 5 — HOME / BEDDING / ORGANIZER niche
Post-2018, ecommerce-native DTC brands only.
No duplicates from batches 1-4.
"""

BATCH_DATA = {
    # =========================================================================
    # 1. Yatak Çarşafı & Nevresim (Innovative Sheet Brands)
    # =========================================================================
    "Yatak Çarşafı & Nevresim": [
        ("Purecare", "purecare.com", "Performans çarşaf", "OmniGuard kumaş teknolojisi ile antimikrobiyal ve nem yönetimi", "2019'da uyku hijyeni odaklı kuruldu, DTC ve Amazon'da hızla büyüdü"),
        ("Nollapelli", "nollapelli.com", "Bilim tabanlı çarşaf", "Cilt ve saç sağlığını koruyan patentli kumaş teknolojisi", "2019'da bir malzeme bilimci tarafından kuruldu, Shark Tank'ta yatırım aldı"),
        ("Flax Home", "flaxhome.com", "Keten çarşaf", "Avrupa keteni, önceden yıkanmış doğal dokular", "2020'de Kanada'da kuruldu, Instagram'da organik büyüme"),
        ("Crisp Sheets", "crispsheets.com", "Premium percale çarşaf", "Mısır pamuğu, percale dokuma, minimalist ambalaj", "2018'de Amsterdam'da kuruldu, Avrupa DTC pazarında öncü"),
        ("Tuft & Needle Sheets", "tuftandneedle.com", "T&N nevresim", "Nefes alan softweave kumaş, uygun fiyat odaklı", "2018'de yatak markasından çarşaf hattına genişledi"),
        ("Casaluna", "casaluna.com", "Organik çarşaf", "GOTS sertifikalı organik pamuk ve keten karışımları", "2020'de Target'ın DTC ilhamıyla piyasaya sürdüğü premium hat"),
        ("Aizome", "aizome.co", "Doğal boyalı çarşaf", "Japon indigo boyama tekniği, kimyasal içermeyen", "2019'da Japonya'dan ilham alınarak San Francisco'da kuruldu"),
        ("Rise & Fall", "riseandfallhome.com", "Lüks temel çarşaf", "İtalyan usta dokumacılardan organik pamuk", "2020'de Londra'da kuruldu, DTC lüks segmentine odaklanıyor"),
        ("Dip & Doze", "dipanddoze.com", "Keten nevresim", "Stonewashed Fransız keteni, soluk pastel tonlar", "2020'de İngiltere'de kuruldu, Instagram estetik odaklı"),
        ("Thuma Sheets", "thuma.co", "Organik yatak tekstili", "Percale ve sateen seçenekleri, platform yatak ekosistemi", "2018'de yatak çerçevesi markası olarak başladı, 2021'de çarşaf hattı ekledi"),
        ("Italic Sheets", "italic.com", "Fabrika direkt çarşaf", "Lüks markaların tedarikçilerinden markasız üretim", "2018'de kuruldu, üyelik bazlı DTC model"),
        ("Beddable", "beddable.com.au", "Bambu çarşaf", "Termoregüle eden bambu lyocell kumaş", "2021'de Avustralya'da kuruldu, TikTok viral büyüme"),
        ("Sunday Home", "sundayhome.co", "Percale çarşaf seti", "Tek parça fiyatlandırma, kolay set alışverişi", "2020'de kuruldu, Z kuşağı odaklı renk paleti"),
        ("Isle", "islecollection.com", "Lüks tropikal çarşaf", "Tencel ve organik pamuk karışımı, serinletici", "2021'de kuruldu, yaz iklimleri için optimize edilmiş"),
        ("Dearest", "dearestjuliet.com", "Vintage esintili nevresim", "Liberty baskı tarzı çiçek desenleri, pamuk percale", "2021'de kuruldu, cottagecore estetiğiyle Instagram'da popüler"),
        ("Pact Bedding", "wearpact.com", "Organik pamuk çarşaf", "Fair Trade sertifikalı, %100 organik pamuk", "2019'da giyim markası yatak tekstiline genişledi"),
        ("LANE by Sunday", "lanebysunday.com", "Serinletici çarşaf", "Faz değiştiren serinletme teknolojisi", "2022'de kuruldu, sıcak uyuyanlar için DTC çözüm"),
        ("Woolshire", "woolshire.co", "Merinos çarşaf", "Avustralya merinos yünü, tüm mevsim termoregülasyon", "2020'de kuruldu, doğal lif odaklı uyku çözümü"),
        ("Haven Mattress Sheets", "havenmattress.com", "Adaptif çarşaf", "Derin cep tasarımı, 360 derece elastik bant", "2019'da Kanada'da kuruldu, Amazon ve Shopify DTC"),
        ("Woven Nook", "wovennook.com", "Bohem nevresim", "El dokuması görünümlü baskılar, waffle dokuma", "2019'da kuruldu, Etsy'den Shopify'a geçiş hikayesi"),
        ("Mela Comfort", "melacomfort.co.uk", "Eucalyptus çarşaf", "Okaliptüs lifi, hipoalerjenik ve ipeksi dokunuş", "2020'de İngiltere'de kuruldu, sürdürülebilirlik odaklı DTC"),
        ("Olivia Rocco", "oliviarocco.com", "Bütçe dostu premium çarşaf", "TC 400 pamuk saten, online-first dağıtım", "2019'da İngiltere'de kuruldu, değer odaklı DTC yaklaşımı"),
        ("Shift Sleep", "shiftsleep.co", "Vardiyalı çalışan çarşafı", "Karartma ve serinletme özellikleri, gündüz uyku desteği", "2022'de kuruldu, niş hedef kitle stratejisi"),
        ("ZipSheet", "zipsheet.com", "Fermuarlı çarşaf", "Alt ve üst çarşaf bütünleşik, kolay yatak yapımı", "2020'de kuruldu, pratik çözüm odaklı DTC"),
        ("Loom & Last", "loomandlast.com", "İpek karışım çarşaf", "Mulberry ipek ve bambu karışımı, lüks dokunuş", "2021'de kuruldu, uyku kalitesi ve cilt bakımı birleşimi"),
        ("Sweave Bedding", "sweavebedding.com", "İki taraflı çarşaf", "Bir taraf saten diğer taraf percale, çift doku", "2019'da kuruldu, Instagram influencer pazarlama stratejisi"),
        ("Good Sheet", "goodsheet.co", "Minimalist çarşaf seti", "Oeko-Tex sertifikalı, nötr renk paleti", "2021'de kuruldu, DTC basitleştirilmiş alışveriş deneyimi"),
        ("Slumber Cloud Sheets", "slumbercloud.com", "NASA teknolojili çarşaf", "Outlast faz değiştirme teknolojisi ile sıcaklık düzenleme", "2018'de NASA lisanslı teknoloji ile kuruldu"),
        ("Primary Goods", "primarygoods.com", "Temel çarşaf seti", "Tek fiyat, tüm boyutlar aynı fiyat politikası", "2022'de kuruldu, şeffaf fiyatlandırma DTC modeli"),
    ],

    # =========================================================================
    # 2. Yastık & Uyku Aksesuarı (Ergonomic/Cooling/Custom Pillows)
    # =========================================================================
    "Yastık & Uyku Aksesuarı": [
        ("Panda Memory Foam", "pandastore.com", "Bambu memory foam yastık", "Bambu kaplama, jel infüzyonlu memory foam çekirdek", "2019'da kuruldu, İngiltere DTC pazarında hızlı büyüme"),
        ("Nuzzle Pillow", "nuzzlepillow.com", "Ayarlanabilir katmanlı yastık", "3 katmanlı yükseklik ayarı, TikTok viral ürün", "2021'de kuruldu, sosyal medya pazarlama odaklı"),
        ("Groove Pillow", "groovepillow.com", "Yan uyuyan yastığı", "Omuz oluğu tasarımı, ergonomik yan pozisyon desteği", "2020'de kuruldu, Kickstarter ile fonlama"),
        ("Earthfoam", "earthfoam.com", "Doğal lateks yastık", "%100 organik Dunlop lateks, sıfır atık üretim", "2021'de kuruldu, sürdürülebilirlik manifestosu ile viral"),
        ("Level Sleep Pillow", "levelsleep.com", "Bölgeli destek yastığı", "Restore teknolojisi, farklı bölgelerde farklı sertlik", "2018'de kuruldu, uyku bilimi araştırma temelli"),
        ("Napform", "napform.com", "Seyahat yastığı", "Katlanabilir ergonomik tasarım, bellek köpük", "2020'de kuruldu, dijital göçebe hedef kitlesi"),
        ("Sognare", "sognarepillow.com", "Lüks İtalyan yastık", "İtalyan el işçiliği, aşağı alternatif dolgulu", "2019'da kuruldu, lüks DTC segmentinde konumlandı"),
        ("Bear Pillow", "bearmattress.com", "Celliant yastık", "Celliant kaplama ile kan dolaşımı artırma teknolojisi", "2018'de yatak markasından yastık hattına genişledi"),
        ("Saybrook Pillow", "saybrookadjustablepillow.com", "Parçalı köpük yastık", "Ayarlanabilir parçalı CertiPUR-US memory foam", "2020'de kuruldu, Amazon ve DTC dual kanal"),
        ("Koala Pillow", "koala.com", "Gel-infused yastık", "Kutulu yastık konsepti, jel soğutma teknolojisi", "2019'da Avustralya yatak markasından yastık ürün hattı"),
        ("Sutera", "sutera.co", "Kelebek yastık", "Ergonomik kelebek şekli, servikal destek", "2020'de kuruldu, Facebook reklam ile DTC büyüme"),
        ("The Pillow Bar", "thepillowbar.com", "Kişiselleştirilmiş yastık", "Online quiz ile uyku pozisyonuna göre öneri", "2019'da kuruldu, kişiselleştirme trendi öncüsü"),
        ("Cloud Nine Pillow", "cloudninepillow.com", "Soğutma jelli yastık", "Dual taraflı: soğutma jel + memory foam", "2021'de kuruldu, yaz kampanyalarıyla Instagram büyüme"),
        ("Snorkel Pillow", "snorkelpillow.com", "Yüzüstü uyuyan yastığı", "Nefes alma kanalı, yüzüstü uyuyanlar için tasarım", "2020'de kuruldu, niş problem çözücü ürün"),
        ("Zamat", "zamat.com", "Ayarlanabilir servikal yastık", "Boyun eğrisi desteği, çıkarılabilir iç yastık", "2019'da kuruldu, Amazon bestseller DTC"),
        ("Cosy House Pillow", "cosyhousecollection.com", "Bambu yastık", "Bambu rayon kaplama, hipoalerjenik dolgu", "2018'de kuruldu, Amazon FBA başarı hikayesi"),
        ("Wamsutta Dream Zone", "wamsuttadreamzone.com", "Isı düzenleyici yastık", "37.5 teknolojisi ile aktif sıcaklık yönetimi", "2019'da DTC kanalına geçiş"),
        ("The Easy Breather", "nestbedding.com", "Parçalı lateks yastık", "Ayarlanabilir parçalı lateks, organik pamuk kılıf", "2018'de Nest Bedding'in DTC yastık alt markası"),
        ("Hullo Buckwheat", "hullopillow.com", "Karabuğday yastık", "Organik karabuğday kabuğu dolgu, sonsuz ayarlanabilirlik", "2018'de kuruldu, minimalist ürün felsefesi"),
        ("Casper Foam", "casper.com", "Üç katmanlı köpük yastık", "Üç farklı köpük katmanı, tüm uyku pozisyonları", "2019'da Casper yastık hattını genişletti"),
        ("Helix Adjustable Pillow", "helixsleep.com", "Ayarlanabilir yastık", "İç yastık çıkarılabilir, yükseklik özelleştirme", "2020'de yatak markasından yastık lansmanı"),
        ("Good Pillow", "goodpillow.co", "Geri dönüştürülmüş yastık", "Okyanus plastiğinden geri dönüştürülmüş dolgu", "2021'de kuruldu, çevresel misyon odaklı DTC"),
        ("Codi Pillow", "codipillow.com", "Hava katmanlı yastık", "Hava sirkülasyonu kanalları, nefes alan tasarım", "2022'de kuruldu, Indiegogo ile lansman"),
        ("RestWell Pillow", "restwellpillow.com", "Ortopedik yastık", "Doktor önerisi ile tasarlanmış servikal destek", "2020'de kuruldu, sağlık profesyoneli onaylı DTC"),
        ("Brentwood Pillow", "brentwoodhome.com", "Yoga yastığı & uyku", "Kapok ve lateks karışımı, çok amaçlı", "2019'da ev tekstili markasından yastık genişlemesi"),
        ("Night Pillow", "nightpillow.co", "İpek kaplı yastık", "Mulberry ipek kılıf, saç ve cilt koruma", "2021'de kuruldu, güzellik-uyku kesişim noktası"),
        ("SlumberEase", "slumberease.co", "Horlama önleyici yastık", "Eğimli tasarım, hava yolu açıklığı desteği", "2020'de kuruldu, spesifik sorun çözücü DTC"),
        ("Avocado Green Pillow", "avocadogreenmattress.com", "Organik lateks yastık", "GOLS sertifikalı organik lateks, GOTS pamuk", "2019'da sürdürülebilir yatak markasından yastık hattı"),
        ("Leesa Pillow", "leesa.com", "Hibrit yastık", "Avassa köpük çekirdek, soğutma jel kaplama", "2019'da DTC yatak markasından yastık genişlemesi"),
    ],

    # =========================================================================
    # 3. Ağırlıklı Battaniye (Weighted Blankets / Anxiety / Sleep)
    # =========================================================================
    "Ağırlıklı Battaniye": [
        ("Aricove", "aricove.com", "Ağırlıklı battaniye", "Cam boncuk dolgu, bambu kaplama, nefes alan", "2020'de kuruldu, Amazon ve Shopify DTC çift kanal"),
        ("Calming Comfort", "calmingcomfort.com", "Sharper Image ağırlıklı", "12 noktalı bağlama sistemi, eşit ağırlık dağılımı", "2019'da DTC lansmanı, TV reklamları ile büyüme"),
        ("Integrity Bedding", "integritybedding.com", "Premium ağırlıklı battaniye", "Amerikan üretimi, organik pamuk seçenekleri", "2018'de kuruldu, yerel üretim odaklı DTC"),
        ("Wemore", "wemore.com", "Çocuk ağırlıklı battaniyesi", "Çocuklara özel ağırlık oranları, eğlenceli desenler", "2021'de kuruldu, ebeveyn odaklı DTC marka"),
        ("Truhugs", "truhugs.com", "Giyilebilir ağırlıklı battaniye", "Giyilebilir tasarım, hareket özgürlüğü ile basınç terapisi", "2020'de Kickstarter ile lansmanı, yenilikçi form faktörü"),
        ("Napure", "napure.co", "Serinletici ağırlıklı", "Bambu tencel kaplama, cam boncuk, serinletici", "2021'de kuruldu, sıcak iklimlere özel DTC çözüm"),
        ("Rest Blanket", "restblanket.com", "Minimalist ağırlıklı", "Tek parça tasarım, iç-dış ayrılmaz yapı", "2020'de kuruldu, Scandinavian minimal estetik"),
        ("Sheltered Co", "shelteredco.com", "Sürdürülebilir ağırlıklı", "Geri dönüştürülmüş malzeme, karbon nötr üretim", "2021'de kuruldu, B Corp sertifikası hedefi"),
        ("Zzoma Weighted", "zzoma.com", "Pozisyonel ağırlıklı battaniye", "Yan uyuma teşviki, asimetrik ağırlık dağılımı", "2020'de kuruldu, uyku pozisyonu düzeltme odaklı"),
        ("Snuggly Weighted", "snugglyweighted.com", "Peluş ağırlıklı battaniye", "Sherpa dış kaplama, çift taraflı kullanım", "2021'de kuruldu, TikTok haul videoları ile büyüme"),
        ("Calmforter", "calmforter.com", "Premium ağırlıklı yorgan", "Yorgan ve ağırlıklı battaniye birleşimi, çift amaçlı", "2019'da kuruldu, ürün inovasyonu ile DTC farklılaşma"),
        ("DreamHug", "dreamhug.com", "Çift kişilik ağırlıklı", "İki farklı ağırlık bölgesi, çiftlere özel", "2022'de kuruldu, Indiegogo başarılı fonlama"),
        ("Slumbr Weighted", "slumbr.co", "Kişiselleştirme odaklı ağırlıklı", "Online quiz ile kilo ve tercihe göre ağırlık önerisi", "2020'de kuruldu, kişiselleştirme trendi DTC"),
        ("Weighted Comfort", "weightedcomfort.co", "Terapötik ağırlıklı battaniye", "Terapistlerle geliştirilmiş, klinik onaylı tasarım", "2019'da kuruldu, sağlık profesyoneli kanalı"),
        ("Huggaroo", "huggaroo.com", "Ağırlıklı omuz sargısı", "Boyun ve omuz bölgesine odaklı, lavanta dolgulu", "2019'da kuruldu, niş vücut bölgesi odaklı DTC"),
        ("Comma Weighted", "commaweighted.com", "Organik ağırlıklı battaniye", "%100 organik pamuk, çelik boncuk dolgu", "2020'de kuruldu, organik sertifikalı DTC marka"),
        ("BlanQuil", "blanquil.com", "Akıllı ağırlıklı battaniye", "Çıkarılabilir iç dolgu, yıkanabilir dış kılıf", "2019'da kuruldu, Facebook reklam ile hızlı DTC büyüme"),
        ("Cooshi", "cooshi.co", "Serinletici ağırlıklı battaniye", "Tencel lyocell kaplama, termoregülasyon", "2021'de kuruldu, sıcak uyuyanlar için niş DTC"),
        ("Gravis Weighted", "gravisweighted.com", "Ağır ağırlıklı battaniye", "25 lb'a kadar seçenekler, ekstra derin basınç", "2020'de kuruldu, yoğun basınç tercihi odaklı niş"),
        ("Nest Weighted", "nestweighted.com", "Çocuk odaklı ağırlıklı", "Pediatrist onaylı, yaşa göre ağırlık kılavuzu", "2021'de kuruldu, çocuk wellness pazarı DTC"),
        ("Solace Weighted", "solaceweighted.com", "Lüks ağırlıklı battaniye", "İpek-pamuk karışım kaplama, premium cam boncuk", "2020'de kuruldu, lüks DTC segment konumlandırma"),
        ("Calm Blanket Co", "calmblanketco.com", "Anksiyete odaklı ağırlıklı", "Anksiyete topluluğu ile birlikte geliştirilmiş", "2021'de kuruldu, mental sağlık misyonu DTC"),
        ("HushBaby Weighted", "hushbaby.co", "Bebek ağırlıklı uyku tulumu", "Hafif basınç, bebek güvenliği sertifikalı", "2022'de kuruldu, yeni ebeveyn niş pazarı"),
        ("Float Weighted", "floatweighted.com", "Su bazlı ağırlıklı battaniye", "Su dolgu sistemi, doğal ağırlık hissi", "2022'de Kickstarter ile lansmanı"),
        ("Serenity Weighted", "serenityweighted.co", "Meditasyon ağırlıklı yastık", "Meditasyon pratiği için tasarlanmış, kucak yastığı", "2021'de kuruldu, wellness kesişim noktası DTC"),
        ("Naptime Weighted", "naptimeweighted.com", "Ofis ağırlıklı battaniye", "Kompakt boyut, çanta ile taşınabilir, iş yeri kullanımı", "2022'de kuruldu, uzaktan çalışma trendi odaklı"),
        ("Cocooning", "cocooning.co", "Koza tipi ağırlıklı battaniye", "Vücudu saran koza tasarımı, 360 derece basınç", "2021'de kuruldu, benzersiz form faktörü ile viral"),
        ("Tranquility Weighted", "tranquilityweighted.com", "Soğutma jelli ağırlıklı", "Soğutma jel teknolojisi, cam boncuk ağırlık", "2019'da kuruldu, Target DTC ortaklığı ile büyüme"),
        ("SnugCozy Weighted", "snugcozy.com", "Hoodie battaniye ağırlıklı", "Giyilebilir hoodie battaniye + ağırlıklı basınç", "2022'de kuruldu, iki trend birleşimi konsept"),
    ],

    # =========================================================================
    # 4. Battaniye & Throw (Chunky Knit, Sherpa, Sustainable)
    # =========================================================================
    "Battaniye & Throw": [
        ("Ohhio", "ohhio.com", "Dev örgü battaniye", "Braid koleksiyonu, devasa tüp örgü tasarım", "2018'de Kickstarter ile kuruldu, chunky knit trendini başlattı"),
        ("Viso Project Home", "visoproject.com", "Sanat koleksiyonu throw", "Sanatçı işbirlikleri, sınırlı üretim desenler", "2019'da kuruldu, sanat-ev tekstili birleşimi"),
        ("Calhoun & Co", "calhounandco.com", "Jakar dokuma throw", "ABD yapımı jakar dokuma, pop kültür desenleri", "2019'da kuruldu, eğlenceli tasarım odaklı DTC"),
        ("Woolly Mammoth Merino", "woollymammothmerino.com", "Avustralya merinos throw", "Ekstra ince merinos yünü, ultrasoft dokunuş", "2019'da Avustralya'da kuruldu, DTC global genişleme"),
        ("Bohemian Mama", "bohemianmama.com", "Bohem throw battaniye", "El yapımı makrome ve dokuma, bohem estetik", "2020'de Etsy'den Shopify DTC'ye geçiş"),
        ("Well Being Blanket", "wellbeingblanket.com", "Wellness throw battaniye", "Ağırlıklı throw, meditasyon kullanımı", "2020'de kuruldu, wellness yaşam tarzı DTC"),
        ("Sackcloth Ashes", "sackclothandashes.com", "Sosyal misyon battaniye", "Her satışta bir battaniye bağışı", "2019'da kuruldu, buy-one-give-one DTC model"),
        ("Coyuchi Throws", "coyuchi.com", "Organik throw", "GOTS organik pamuk ve yün, doğal boyalar", "2019'da throw hattını DTC kanalında genişletti"),
        ("Thula Blanket", "thulablanket.com", "Güney Afrika throw", "Güney Afrika el dokuması, adil ticaret", "2019'da kuruldu, etik üretim hikayesi DTC"),
        ("Sobremesa Blanket", "sobremesablanket.com", "Meksika throw", "Geleneksel Meksika dokuma teknikleri, modern desenler", "2020'de kuruldu, kültürel miras odaklı DTC"),
        ("Atlas Blanket Co", "atlasblanketco.com", "Fas throw battaniye", "El dokuması Fas battaniyeleri, pom-pom detay", "2020'de kuruldu, Instagram influencer pazarlama"),
        ("Coterie Brooklyn", "coteriebrooklyn.com", "Alpaka throw", "Peru alpaka yünü, Brooklyn tasarımı", "2019'da kuruldu, lüks doğal lif DTC"),
        ("Wrap London", "wrapmewarm.com", "Giyilebilir throw", "Düğmeli ve kuşaklı tasarımlar, throw + giysi", "2021'de kuruldu, çok işlevli throw konsepti"),
        ("Kindfolk", "kindfolkyarn.com", "El örgüsü throw kiti", "DIY chunky knit kitleri, online workshop", "2020'de kuruldu, DIY trendi ve DTC birleşimi"),
        ("Tartan Blanket Co", "tartanblanketco.com", "İskoç tartan throw", "Geri dönüştürülmüş yün, İskoç tartan desenleri", "2019'da İskoçya'da kuruldu, Instagram estetik büyüme"),
        ("Forestry Wool", "forestrywool.com", "Orman temalı throw", "Yeni Zelanda yünü, doğa ilhamı tasarımlar", "2020'de Yeni Zelanda'da kuruldu, eko turizm DTC"),
        ("Baya", "baya.co", "Pamuklu ev battaniyesi", "Türk pamuğu, hafif dört mevsim kullanım", "2020'de kuruldu, Akdeniz yaşam tarzı DTC"),
        ("The Tartan Company", "thetartancompany.com", "Lüks yün throw", "El bitimi, İskoç heritage tasarım", "2019'da kuruldu, lüks hediye pazarı DTC"),
        ("Pure New Wool", "purenewwool.com", "Galler yünü throw", "Galler koyun yünü, geleneksel dokuma", "2020'de İngiltere'de kuruldu, yerel üretim DTC"),
        ("TerraKlay", "terraklay.com", "Sürdürülebilir pamuk throw", "Hindistan el dokuması, doğal boyalar", "2019'da kuruldu, etik moda-ev tekstili DTC"),
        ("Cabin Wool", "cabinwool.com", "Kabin tarzı throw", "Kalın yün, rustik estetik, kış konsepti", "2021'de kuruldu, kış mevsimi odaklı DTC"),
        ("Nuzzie Knit", "nuzzie.com", "Örgü ağırlıklı throw", "Chunky knit + ağırlıklı battaniye birleşimi", "2020'de kuruldu, iki trend birleşimi DTC hit"),
        ("Loomstead", "loomstead.co", "Atölye throw", "Küçük parti üretim, zanaat throw battaniye", "2022'de kuruldu, slow fashion DTC yaklaşımı"),
        ("Haven Blanket", "havenblanket.co", "Waffle throw", "Waffle dokuma, %100 pamuk, spa hissi", "2021'de kuruldu, otel deneyimi evde konsepti"),
        ("Heritage Wool Co", "heritagewoolco.com", "Miras yün throw", "100 yıllık dokuma tekniklerini modern DTC ile sunan", "2020'de kuruldu, geleneksel zanaat + e-ticaret"),
        ("Goodlinens Throw", "goodlinens.co", "Keten throw battaniye", "Yıkanmış keten, hafif ve nefes alan", "2021'de kuruldu, yaz throw pazarı DTC"),
        ("Aerende", "aerende.co.uk", "Sosyal girişim throw", "Dezavantajlı topluluklar tarafından üretim", "2019'da İngiltere'de kuruldu, sosyal etki DTC"),
        ("Homebird", "homebirdtextiles.com", "İrlanda yün throw", "İrlanda koyun yünü, geleneksel desenleri modern yorumlama", "2020'de kuruldu, İrlanda zanaatı DTC"),
        ("The Citizenry Throw", "the-citizenry.com", "Zanaat throw koleksiyonu", "Dünya çapında zanaatkarlarla üretim, adil ticaret", "2019'da throw koleksiyonu genişletildi, hikaye odaklı DTC"),
        ("Somewhere Throw", "somewheretoday.com", "Seyahat ilhamı throw", "Her throw bir destinasyon hikayesi, sınırlı seri", "2022'de kuruldu, seyahat nostalji DTC konsepti"),
    ],

    # =========================================================================
    # 5. Dolap & Giysi Organizeri (Closet Systems, Wardrobe Organizers)
    # =========================================================================
    "Dolap & Giysi Organizeri": [
        ("Neatfi", "neatfi.com", "Aydınlatmalı dolap sistemi", "LED entegre raf sistemleri, kurulum gerektirmez", "2019'da kuruldu, Amazon DTC aydınlatmalı organizasyon"),
        ("Closet Envy", "closetenvy.com", "Modüler dolap sistemi", "Tamamen özelleştirilebilir modüler paneller", "2020'de kuruldu, DTC modüler dolap çözümü"),
        ("Haven Closets", "havenclosets.co", "Lüks dolap düzenleme", "Kadife kaplama askılar ve kutular set halinde", "2021'de kuruldu, premium organizasyon DTC"),
        ("Neat Living", "neatliving.co", "Kapsül gardırop organizeri", "Kapsül gardırop planlama aracı + organizasyon ürünleri", "2020'de kuruldu, minimalizm trendi DTC"),
        ("Dotti", "dottistorage.com", "Renkli dolap kutuları", "Modüler renkli saklama kutuları, çocuk odası odaklı", "2021'de kuruldu, Instagram renk paleti estetiği"),
        ("Cloffice", "cloffice.co", "Dolap-ofis dönüşüm", "Dolap alanını home office'e dönüştüren sistem", "2021'de kuruldu, pandemi home office trendi DTC"),
        ("Tidy Home Happy", "tidyhomehappy.com", "Marie Kondo tarzı organizasyon", "Katlanabilir kutu ve bölme sistemi, bambu", "2019'da kuruldu, KonMari trendi ile büyüyen DTC"),
        ("The Wardrobe Edit", "thewardrobeedit.co", "Gardırop düzenleme kiti", "Profesyonel organizatör tasarımı, etiket sistemi dahil", "2022'de kuruldu, influencer organizatör işbirliği DTC"),
        ("Capsule Wardrobe Co", "capsulewardrobeco.com", "Kapsül gardırop askı sistemi", "Numaralı askılar, mevsimsel rotasyon sistemi", "2020'de kuruldu, kapsül gardırop akımı DTC"),
        ("Orderly", "orderly.co", "Akıllı dolap organizeri", "QR kodlu kutular, envanter takip uygulaması", "2022'de kuruldu, teknoloji-organizasyon birleşimi DTC"),
        ("The Holding Company", "theholdingcompany.co.uk", "Premium organizasyon kutuları", "Keten kaplama kutular, özel etiket sistemi", "2019'da İngiltere'de kuruldu, lüks organizasyon DTC"),
        ("Shelf Help", "shelfhelp.co", "Raf organizatörleri", "Genişletilebilir raf bölücüler, evrensel uyum", "2020'de kuruldu, pratik çözüm DTC markası"),
        ("Neatly Organized", "neatlyorganized.co", "Dolap başlangıç seti", "Herşey dahil starter kit, adım adım kılavuz", "2021'de kuruldu, yeni başlayanlar için DTC organizasyon"),
        ("Declutter Box", "declutterbox.com", "Aylık organizasyon kutusu", "Aylık abonelik kutusu, her ay farklı alan odaklı", "2021'de kuruldu, abonelik model DTC organizasyon"),
        ("Kolorit", "kolorit.co", "Renkli dolap aksesuarları", "Renk kodlu askı ve kutular, pastel tonlar", "2022'de kuruldu, Gen Z dolap estetiği DTC"),
        ("Smart Wardrobe", "smartwardrobe.co", "Dijital dolap yönetimi", "RFID etiketli askılar, giyim takip uygulaması", "2022'de kuruldu, akıllı ev entegrasyonu DTC"),
        ("The Laundry Society", "thelaundrysociety.com", "Çamaşır organizasyonu", "Premium çamaşır torbaları ve sepetleri, keten", "2019'da kuruldu, çamaşır deneyimini yükseltme DTC"),
        ("Klokke", "klokke.co", "İskandinav dolap sistemi", "Beyaz meşe ve beyaz metal, minimalist tasarım", "2020'de kuruldu, Scandi estetik DTC"),
        ("Halo Hangers", "halohangers.com", "Premium kadife askı", "Slim kadife askılar, alan tasarrufu, set halinde", "2019'da kuruldu, TikTok dolap dönüşüm trendi"),
        ("ClosetMate", "closetmate.co", "Yer tasarruflu askı sistemi", "Kademeli askı organizatörü, dikey alan kullanımı", "2020'de kuruldu, küçük alan çözümleri DTC"),
        ("Roomy", "roomy.co", "Mevsimsel saklama çözümü", "Vakumlu saklama + etiketleme sistemi", "2021'de kuruldu, mevsim geçişi organizasyon DTC"),
        ("Curate Home", "curatehome.co", "Açık gardırop sistemi", "Endüstriyel boru raf, görünür depolama estetiği", "2020'de kuruldu, açık dolap trendi DTC"),
        ("Clos-ette", "closette.com", "Lüks dolap tasarımı", "Özel tasarım lüks dolap iç düzeni, DTC genişleme", "2018'de DTC kanalını açtı, New York lüks segmenti"),
        ("The Tidy Files", "thetidyfiles.com", "Belge ve aksesuar organizeri", "Dolap içi belge cepleri ve aksesuar bölmeleri", "2021'de kuruldu, detay organizasyon DTC"),
        ("Hang It Up", "hangitup.co", "Kişiselleştirilmiş askı", "İsim baskılı ve renk seçimli özel askılar", "2022'de kuruldu, kişiselleştirme trendi DTC"),
        ("Wardrobist", "wardrobist.com", "Gardırop planlama servisi", "Online gardırop danışmanlığı + fiziksel ürün seti", "2021'de kuruldu, hizmet + ürün hibrit DTC model"),
        ("Tame", "tame.style", "Ayakkabı dolap organizeri", "Drop-front ayakkabı kutuları, şeffaf ve istiflenebilir", "2020'de kuruldu, sneaker kültürü DTC organizasyon"),
        ("Neat Soul", "neatsoul.com", "Bambu dolap organizeri", "Bambu bölücüler ve kutular, doğal estetik", "2021'de kuruldu, eko dostu organizasyon DTC"),
    ],

    # =========================================================================
    # 6. Mutfak Organizasyonu (Kitchen Storage, Pantry, Spice)
    # =========================================================================
    "Mutfak Organizasyonu": [
        ("Spice Labs", "spicelabs.co", "Akıllı baharat düzeni", "Manyetik baharat kavanozları, kapak etiketleme sistemi", "2020'de kuruldu, mutfak organizasyonu DTC"),
        ("Neat Stack", "neatstack.co", "İstiflenebilir konteyner", "Tamamen istiflenebilir modüler gıda kapları", "2021'de kuruldu, TikTok mutfak dönüşüm trendi"),
        ("Pantry Home", "pantryhome.co", "Kiler organizasyon seti", "Etiketli kavanoz seti, raf düzenleyici dahil", "2020'de kuruldu, kiler dönüşüm trendi DTC"),
        ("Savvy & Sorted", "savvyandsorted.com", "Premium kiler etiketleri", "Su geçirmez vinil etiketler, 300+ mutfak etiketi", "2019'da kuruldu, Etsy'den Shopify'a DTC büyüme"),
        ("The Pantry Co", "thepantryco.com.au", "Kiler kavanoz sistemi", "Uniform cam kavanozlar, bambu kapaklı", "2019'da Avustralya'da kuruldu, Instagram kiler estetiği"),
        ("Clever Storage", "cleverstorage.co", "Tezgah altı organizatör", "Yapışkanlı ve vidalı tezgah altı raf sistemleri", "2020'de kuruldu, küçük mutfak DTC çözümleri"),
        ("Stacked", "getstacked.co", "Buzdolabı organizatörü", "Şeffaf buzdolabı kutuları, sınıflandırma sistemi", "2021'de kuruldu, buzdolabı dönüşüm TikTok trendi"),
        ("Homeries", "homeries.com", "Çekmece ve dolap organizatörü", "Genişletilebilir çekmece bölücüler, mutfak özel", "2019'da kuruldu, Amazon ve DTC çift kanal"),
        ("ClearSpace", "clearspaceliving.com", "Şeffaf saklama kutusu", "Premium şeffaf plastik, istiflenebilir tasarım", "2020'de kuruldu, The Home Edit ilhamı DTC"),
        ("Pretzel Container", "pretzelcontainer.com", "Hava geçirmez konteyner", "Tek el ile açılır kapak, hava geçirmez seal", "2021'de kuruldu, pratik kullanım odaklı DTC"),
        ("Smeg Organizer", "smegorganizer.com", "Retro mutfak organizasyon", "Retro renklerde mutfak kutuları ve kavanozlar", "2022'de kuruldu, retro estetik mutfak DTC"),
        ("The Organised Pantry", "theorganisedpantry.com.au", "Kiler başlangıç kiti", "Her şey dahil kiler dönüşüm kiti, rehber dahil", "2020'de Avustralya'da kuruldu, komple çözüm DTC"),
        ("Lockabox", "lockabox.com", "Kilitli gıda saklama", "Kilitli mutfak kutusu, çocuk güvenliği ve diyet", "2019'da İngiltere'de kuruldu, niş güvenlik DTC"),
        ("Gneiss Spice", "gneisspice.com", "Manyetik baharat rafı", "Manyetik cam kavanozlar, buzdolabına yapışan", "2019'da kuruldu, alan tasarruflu DTC baharat çözümü"),
        ("Label Luxe", "labelluxe.co", "Lüks mutfak etiketleri", "Altın folyo etiketler, kişiselleştirilebilir", "2021'de kuruldu, lüks mutfak estetiği DTC"),
        ("Fridge Rescue", "fridgerescue.co", "Buzdolabı kurtarma seti", "Buzdolabı deodorant, organizatör, etiket combo", "2022'de kuruldu, buzdolabı bakım + organizasyon DTC"),
        ("Kilner Home DTC", "kilnerhome.com", "Cam saklama kavanozları", "Heritage cam kavanozlar, DTC direkt satış genişleme", "2019'da DTC kanalını açtı, mutfak organizasyon seti"),
        ("Sorted Food Org", "sortedfood.co", "Yemek hazırlık organizatörü", "Haftalık meal prep kutuları ve etiketleme sistemi", "2021'de kuruldu, meal prep topluluğu DTC"),
        ("Minimalist Pantry", "minimalistpantry.co", "Minimalist kiler çözümü", "Monokrom tasarım, sayılı temel parçalarla kiler", "2022'de kuruldu, minimalizm mutfak DTC"),
        ("Flip Stack", "flipstack.co", "Dikey mutfak rafı", "Dikey istiflenebilir raf, dar alanlara özel", "2021'de kuruldu, mikro mutfak DTC çözümü"),
        ("Tidy Jars", "tidyjars.com", "Bambu kapaklı kavanoz", "Borosilikat cam, bambu kapak, silikon conta", "2020'de kuruldu, sürdürülebilir mutfak DTC"),
        ("SpiceDeck", "spicedeck.com", "Çekmece içi baharat organizeri", "Eğimli baharat bölmeleri, çekmece özel tasarım", "2021'de kuruldu, çekmece baharat trendi DTC"),
        ("Kitchen Capsule", "kitchencapsule.co", "Kapsül mutfak seti", "30 parça temel mutfak organizasyon kiti", "2022'de kuruldu, kapsül konsept mutfak DTC"),
        ("The Container Lab", "thecontainerlab.com", "Laboratuvar tarzı saklama", "Bilim laboratuvarı esintili cam kaplar", "2021'de kuruldu, benzersiz estetik DTC"),
        ("Undercover Shelf", "undercovershelf.com", "Gizli raf sistemi", "Dolap kapağı arkası organizatör, görünmez depolama", "2020'de kuruldu, gizli depolama DTC çözümü"),
        ("Kinforest", "kinforest.co", "Ahşap mutfak organizatörü", "Akasya ahşap raf ve kutu, zanaat kalitesi", "2021'de kuruldu, doğal malzeme mutfak DTC"),
        ("Simply Stored", "simplystored.co", "Etiketli komple set", "Kavanoz + etiket + raf komple kiler seti", "2021'de kuruldu, tek seferde kiler dönüşüm DTC"),
        ("Pantri", "pantri.co", "Akıllı kiler yönetimi", "QR kodlu kavanozlar, son kullanma tarihi takip uygulaması", "2022'de kuruldu, teknoloji + organizasyon DTC"),
        ("Fresh & Tidy", "freshandtidy.co", "Sebze-meyve organizeri", "Havalandırmalı saklama kutuları, tazelik uzatma", "2021'de kuruldu, gıda israfı azaltma DTC"),
    ],

    # =========================================================================
    # 7. Banyo Organizasyonu (Bathroom Storage, Shower Caddies)
    # =========================================================================
    "Banyo Organizasyonu": [
        ("Shower Beer", "showerbeer.co", "Duş içecek tutucu", "Silikon yapışkanlı duş bardak tutucu", "2019'da kuruldu, viral TikTok banyo aksesuarı"),
        ("Glow Bathroom", "glowbathroom.co", "LED banyo organizeri", "LED aydınlatmalı ayna arkası dolap", "2021'de kuruldu, akıllı banyo DTC"),
        ("Shelfology", "shelfology.com", "Görünmez duvar rafı", "Gizli montaj sistemi, floating shelf banyo", "2019'da kuruldu, minimalist banyo DTC"),
        ("Tile Redi Niche", "tileredi.com", "Duş niş organizeri", "Fayans uyumlu gömme duş rafı, su geçirmez", "2019'da DTC kanalını açtı, duş renovasyon çözümü"),
        ("Mist Bath", "mistbath.co", "Bambu banyo organizatörü", "Su dayanıklı bambu raflar ve kutular", "2021'de kuruldu, doğal malzeme banyo DTC"),
        ("Soak Society", "soaksociety.co", "Banyo ritüeli organizeri", "Banyo tuzu ve yağ organizasyon kutusu", "2020'de kuruldu, self-care organizasyon DTC"),
        ("The Bath Box", "thebathbox.co", "Küvet tepsisi", "Bambu küvet tepsisi, tablet ve şarap tutucu", "2019'da kuruldu, Instagram banyo estetiği DTC"),
        ("Cloud Shower", "cloudshower.co", "Duş caddy sistemi", "Paslanmaz çelik, yapışkanlı montaj, modüler", "2020'de kuruldu, kiracı dostu banyo DTC"),
        ("Steam Shelf", "steamshelf.co", "Buhar dayanımlı raf", "Anti-buğu kaplama, paslanmaz çelik banyo rafı", "2021'de kuruldu, dayanıklılık odaklı DTC"),
        ("Bath Bliss DTC", "bathblissdtc.com", "Banyo başlangıç seti", "Komple banyo organizasyon kiti, 15 parça", "2020'de DTC kanalını açtı, set bazlı satış stratejisi"),
        ("Mirror Cabinet Co", "mirrorcabinetco.com", "Aynalı banyo dolabı", "Slim profil, LED aydınlatma, gömme montaj", "2021'de kuruldu, modern banyo DTC"),
        ("Vanity Vault", "vanityvault.co", "Tezgah altı makyaj organizeri", "Çekmeceli vanity organizatörü, akrilik", "2020'de kuruldu, banyo vanity DTC çözümü"),
        ("AquaShelf", "aquashelf.co", "Su geçirmez duş rafı", "Alüminyum alaşım, drenaj delikleri, yapışkanlı", "2021'de kuruldu, Amazon ve DTC çift kanal"),
        ("The Towel Bar", "thetowelbar.co", "Havlu organizasyon sistemi", "Katlı havlu rafı, ısıtıcılı seçenekler", "2020'de kuruldu, havlu depolama DTC"),
        ("Bask Bath", "baskbath.co", "Banyo köşe organizeri", "Köşe montaj, 3 katmanlı, paslanmaz", "2021'de kuruldu, köşe alan kullanımı DTC"),
        ("Pristine Bath", "pristinebath.co", "Beyaz banyo organizasyon", "Mat beyaz metal seri, tutarlı tasarım dili", "2022'de kuruldu, monokrom banyo estetiği DTC"),
        ("Nook Bathroom", "nookbathroom.co", "Niş banyo çözümleri", "Küçük banyo özel raf ve kanca sistemleri", "2021'de kuruldu, küçük banyo DTC"),
        ("Drip Dry", "dripdry.co", "Duş kurutma organizeri", "Duş içi kuru alan, kişisel bakım ürünleri rafı", "2020'de kuruldu, duş organizasyonu DTC"),
        ("Suds & Shelf", "sudsandshelf.co", "Sabunluk ve raf combo", "Entegre sabunluk + mini raf, tek parça tasarım", "2022'de kuruldu, kompakt çözüm DTC"),
        ("Wet Room Co", "wetroomco.com", "Islak zemin organizeri", "Drenajlı zemin organizatörü, duş köşesi", "2021'de kuruldu, ıslak alan uzmanı DTC"),
        ("Caddy King", "caddyking.co", "Premium duş caddy", "Teak ahşap duş caddy, lüks spa hissi", "2020'de kuruldu, premium duş aksesuarı DTC"),
        ("Tidy Tiles", "tidytiles.co", "Fayans uyumlu organizasyon", "Fayans rengine uyumlu yapışkanlı organizatörler", "2022'de kuruldu, estetik uyum DTC"),
        ("Splash Organizer", "splashorganizer.co", "Çocuk banyo organizeri", "Eğlenceli renkler ve şekiller, vantuzlu montaj", "2021'de kuruldu, çocuk banyosu DTC"),
        ("Stack Bath", "stackbath.co", "İstiflenebilir banyo kutusu", "Su geçirmez istiflenebilir kutular, banyo özel", "2020'de kuruldu, modüler banyo DTC"),
        ("Eco Bath Store", "ecobathstore.co", "Eko banyo organizasyon", "Geri dönüştürülmüş okyanus plastiği raflar", "2021'de kuruldu, sürdürülebilir banyo DTC"),
        ("Bathroom Butler DTC", "bathroombutlerdtc.co", "Havlu ısıtıcı organizeri", "Havlu ısıtıcı + organizatör kombo ünite", "2019'da DTC genişleme, ısıtıcı + depolama konsepti"),
        ("Rinse & Repeat", "rinseandrepeat.co", "Şampuan dispenser sistemi", "Duvara monteli dispenser, etiket sistemi dahil", "2020'de kuruldu, plastik azaltma + organizasyon DTC"),
        ("Fresh Bath Co", "freshbathco.com", "Banyo hava kalitesi", "Nem alıcı + organizatör ünite, küf önleyici", "2021'de kuruldu, banyo sağlığı DTC"),
        ("Lather Shelf", "lathershelf.co", "Tıraş organizeri", "Tıraş aksesuarları duvar organizatörü, manyetik", "2022'de kuruldu, erkek banyo organizasyon DTC"),
    ],

    # =========================================================================
    # 8. Çekmece Düzenleyici (Drawer Dividers, Compartments)
    # =========================================================================
    "Çekmece Düzenleyici": [
        ("Drawerganize", "drawerganize.com", "Ayarlanabilir çekmece bölücü", "Bahar mekanizmalı genişleyen bölücüler", "2020'de kuruldu, Amazon DTC çekmece çözümü"),
        ("Tidied Up", "tidiedup.co", "Bambu çekmece organizeri", "Bambu bölmeli insert, mutfak ve banyo uyumlu", "2021'de kuruldu, doğal malzeme DTC organizasyon"),
        ("The Drawer Edit", "thedraweredit.co", "Çekmece dönüşüm kiti", "Ölçüye göre kesilen bölücüler, online planlama aracı", "2022'de kuruldu, kişiselleştirme odaklı DTC"),
        ("Divvy", "divvy.co", "Modüler çekmece sistemi", "Lego mantığıyla birleşen modüler bölmeler", "2021'de kuruldu, modüler tasarım DTC"),
        ("Tidy Drawer", "tidydrawer.com", "İç çamaşırı çekmece organizeri", "Bal peteği bölme sistemi, katlanabilir", "2019'da kuruldu, TikTok çekmece dönüşüm viral"),
        ("Compartment Co", "compartmentco.com", "Özel ölçü çekmece bölmesi", "Online ölçü girişi ile özel üretim bölmeler", "2022'de kuruldu, made-to-measure DTC"),
        ("KonMari Drawer", "konmaridrawer.co", "KonMari çekmece kutusu", "Marie Kondo katlanma tekniğine uygun kutular", "2020'de kuruldu, KonMari lisanslı DTC ürün"),
        ("Junk Drawer Fix", "junkdrawerfix.com", "Karışık çekmece çözümü", "Evrensel junk drawer organizatörü, ayarlanabilir", "2021'de kuruldu, herkesin problemi DTC çözüm"),
        ("Insert Order", "insertorder.co", "Premium ahşap insert", "Ceviz ve meşe ahşap çekmece insertleri", "2020'de kuruldu, lüks mutfak DTC"),
        ("DrawerFit", "drawerfit.co", "Ölçüye göre bölücü", "3D tarama ile ölçü alma, özel üretim", "2022'de kuruldu, teknoloji destekli DTC"),
        ("Sorted Drawers", "sorteddrawers.co", "Renk kodlu çekmece sistemi", "Her renk farklı kategori, aile kullanımı", "2021'de kuruldu, aile odaklı DTC organizasyon"),
        ("Silicone Divider Co", "siliconedivider.com", "Silikon çekmece bölücü", "Esnek silikon, her çekmeceye uyum, yıkanabilir", "2020'de kuruldu, pratik materyal DTC"),
        ("The Neat Drawer", "theneatdrawer.co", "Kadife kaplı çekmece insertı", "Kadife kaplama, mücevher ve aksesuar odaklı", "2021'de kuruldu, lüks organizasyon DTC"),
        ("Stack & Slide", "stackandslide.co", "Kayan çekmece organizeri", "İç içe kayan bölmeler, derinlik kullanımı", "2022'de kuruldu, derin çekmece DTC çözümü"),
        ("Minimal Drawer", "minimaldrawer.co", "Minimalist çekmece seti", "Mat beyaz veya siyah, 5 parça temel set", "2021'de kuruldu, minimalist estetik DTC"),
        ("Flex Divide", "flexdivide.co", "Esnek bölücü sistem", "Mıknatıslı bağlantı, sonsuz konfigürasyon", "2022'de kuruldu, Kickstarter fonlanmış DTC"),
        ("Cutlery Nest", "cutlerynest.co", "Çatal bıçak çekmece organizeri", "Genişletilebilir çatal bıçak düzenleyici, bambu", "2020'de kuruldu, mutfak çekmece DTC"),
        ("Underwear Edit", "underwearedit.co", "İç çamaşırı organizeri", "Hücresel bölme sistemi, katlanma rehberi dahil", "2021'de kuruldu, niş çekmece DTC"),
        ("Little Drawer Co", "littledrawerco.com", "Çocuk çekmece organizeri", "Renkli hayvan figürlü bölmeler, çocuk dostu", "2022'de kuruldu, çocuk odası DTC"),
        ("Makeup Drawer Pro", "makeupdrawerpro.com", "Makyaj çekmece organizeri", "Akrilik bölmeler, ruj-fondöten-fırça ayrımı", "2020'de kuruldu, güzellik organizasyon DTC"),
        ("Tool Drawer Kit", "tooldrawerkit.com", "Alet çekmece organizeri", "Köpük kesimli özel alet yuvası, atölye odaklı", "2021'de kuruldu, DIY topluluğu DTC"),
        ("Office Drawer Set", "officedrawerset.co", "Ofis çekmece düzenleyici", "Kalem, zımba, post-it bölmeleri, masaüstü çekmece", "2020'de kuruldu, home office DTC organizasyon"),
        ("Sock Drawer Co", "sockdrawerco.com", "Çorap çekmece organizeri", "Altıgen bölmeler, çorap eşleştirme sistemi", "2021'de kuruldu, spesifik niş DTC çözüm"),
        ("Tidy Tech Drawer", "tidytechdrawer.co", "Teknoloji çekmece organizeri", "Kablo yönetimi, şarj cihazı bölmeleri", "2022'de kuruldu, teknoloji aksesuarı DTC"),
        ("Herb Drawer", "herbdrawer.co", "Baharat çekmece organizeri", "Eğimli kavanoz yuvası, çekmece içi baharat düzeni", "2021'de kuruldu, mutfak niş DTC"),
        ("Velvet Insert Co", "velvetinsertco.com", "Lüks kadife çekmece astarı", "Kadife rulo astar + bölücü kombo, kolay kesim", "2020'de kuruldu, DIY lüks DTC"),
        ("Honeycomb Drawer", "honeycombdrawer.co", "Bal peteği çekmece bölücü", "Altıgen modüler sistem, genişletilebilir", "2021'de kuruldu, benzersiz şekil DTC"),
        ("Smart Drawer", "smartdrawer.co", "QR kodlu çekmece sistemi", "QR etiketli bölmeler, içerik takip uygulaması", "2022'de kuruldu, akıllı organizasyon DTC"),
        ("The Divider Store", "thedividerstore.com", "Çekmece bölücü süpermarketi", "150+ farklı boyut ve malzeme seçeneği", "2020'de kuruldu, geniş seçim DTC"),
    ],
}


# ---------------------------------------------------------------------------
# Quick stats
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    total = 0
    for category, brands in BATCH_DATA.items():
        count = len(brands)
        total += count
        print(f"{category}: {count} marka")
    print(f"\nToplam: {total} marka")
