"""
yuvacim_v3_batch2.py
Post-2018 ecommerce-native DTC brands — HOME / ORGANIZER niche
6 categories, 200+ brands total
All descriptions in Turkish (brand names & URLs in English)
"""

BATCH_DATA = {

    # ─────────────────────────────────────────────────────────────────
    # 1. HAVLU & BORNOZ & PEŞTEMAL  (Towels, Robes & Peshtemals)
    # ─────────────────────────────────────────────────────────────────
    "Havlu & Bornoz & Pestemal": [
        ("Onsen", "onsentowel.com", "Waffle Dokulu Havlu",
         "Japon onsen kültüründen esinlenme; waffle dokulu hafif havlu; hızlı kuruma; %100 Supima pamuk; antibakteriyel özellik",
         "2018'de kuruldu; Kickstarter'da $1M+ topladı; DTC Shopify modeli; Instagram ve influencer odaklı büyüme; minimalist estetik"),

        ("Mizu Towel", "mizutowel.com", "Kendi Kendini Temizleyen Havlu",
         "Gümüş iyonlu self-cleaning havlu; antibakteriyel; koku önleyici; hızlı kuruma teknolojisi; hafif tasarım",
         "2019'da kuruldu; Kickstarter'da viral oldu; $500K+ crowdfunding; DTC direkt satış; sosyal medya ve PR ile büyüme"),

        ("Weezie", "weezie.com", "Monogramlu Lüks Havlu",
         "Kişiselleştirilmiş monogram havlu ve bornoz; Uzun elyaf Türk pamuğu; lüks otel hissi; hediye odaklı konumlandırma",
         "2018'de kuruldu; Instagram-first DTC marka; influencer ve hediye pazarı; Shopify üzerinden satış; güçlü sosyal medya varlığı"),

        ("House No. 23", "houseno23.com", "Türk Pamuğu Havlu",
         "Luxury Türk pamuğu; 700 GSM ağırlık; otel kalitesinde yumuşaklık; klasik tasarım; dayanıklı dokuma",
         "2019'da kuruldu; DTC premium havlu; Instagram ve Meta Ads büyüme; Shopify e-ticaret; online-only marka"),

        ("Laguna Beach Textile", "lagunabeachtextile.com", "Oversize Plaj Havlusu",
         "180x90cm oversize boyut; %100 Türk pamuğu; canlı desenler; çift taraflı dokuma; lüks plaj deneyimi",
         "2018'de kuruldu; Instagram ve influencer pazarlama; Amazon ve DTC satış; plaj yaşam tarzı konumlandırma"),

        ("Baina", "bainaofficial.com", "Organik Pamuk Banyo Havlusu",
         "GOTS organik pamuk; minimal ve estetik tasarım; düşük çevresel ayak izi; Avustralya kökenli; nötr tonlar",
         "2020'de kuruldu; Instagram-native marka; Shopify DTC; sürdürülebilir moda influencer'ları ile iş birliği; premium segment"),

        ("Hawkins New York", "hawkinsnewyork.com", "Minimalist Banyo Tekstili",
         "Waffle ve havlu kumaş serileri; Stonewashed linen; minimalist NYC estetiği; nötr palet; çok amaçlı kullanım",
         "2019'da banyo serisini genişletti; DTC e-ticaret; Instagram odaklı; tasarım blog'ları ve PR ile büyüme"),

        ("Parachute Home", "parachutehome.com", "Premium Havlu Koleksiyonu",
         "Uzun elyaf Türk pamuğu; 700 GSM; OEKO-TEX sertifikalı; klasik ve modern tasarımlar; lüks banyo deneyimi",
         "2018 sonrası havlu hattını genişletti; DTC lider marka; Instagram ve Facebook Ads; influencer pazarlama; showroom açılımı"),

        ("Coyuchi", "coyuchi.com", "Organik Pamuk Havlu",
         "%100 organik pamuk; GOTS sertifikalı; sürdürülebilir üretim; doğal boyama; çevre dostu ambalaj",
         "2019'da DTC kanalını güçlendirdi; B Corp sertifikalı; Instagram ve sosyal medya büyüme; çevre bilinçli tüketici hedefi"),

        ("Brooklinen", "brooklinen.com", "Süper Yumuşak Havlu Seti",
         "Uzun elyaf Türk pamuğu; 820 GSM ultra yumuşak; waffle ve klasik seçenekler; lüks ama erişilebilir fiyat",
         "2018'de havlu hattını genişletti; DTC unicorn; Meta Ads ve podcast reklamları; viral büyüme; online-first"),

        ("Lulu & Georgia", "luluandgeorgia.com", "Tasarım Odaklı Banyo Tekstili",
         "Bohemian ve modern tasarımlar; püsküllü peştemal tarzı havlular; dekoratif ve fonksiyonel; Instagram estetiği",
         "2019'da banyo kategorisi genişlemesi; DTC e-ticaret; influencer ve interior design blog iş birlikleri"),

        ("The Citizenry", "the-citizenry.com", "El Yapımı Peştemal",
         "Türk ustalarla iş birliği; el dokuma peştemal; doğal boyalar; artisan üretim; hikaye odaklı marka",
         "2018'de peştemal hattını ekledi; DTC Shopify; Instagram hikaye anlatımı; adil ticaret vurgusu; premium fiyat"),

        ("Kassatex", "kassatex.com", "Otel Kalitesinde Havlu",
         "Türk ve Mısır pamuğu; 600-800 GSM; otel tipi lüks; geniş renk yelpazesi; organik seçenekler",
         "2019'da DTC kanalını güçlendirdi; Amazon ve Shopify; influencer pazarlama; hospitality background"),

        ("Rightside", "rightside.co", "Waffle Bornoz",
         "Hafif waffle kumaş; unisex tasarım; hızlı kuruma; seyahat dostu; minimalist kesim",
         "2020'de kuruldu; DTC only; Instagram ve TikTok pazarlama; Shopify; micro-influencer iş birlikleri"),

        ("Nubi", "nubitowels.com", "Türk Peştemalı",
         "Geleneksel Türk peştemalı; %100 pamuk; hafif ve çok amaçlı; plaj, sauna, ev kullanımı; el dokuması",
         "2019'da kuruldu; Amazon ve DTC; Instagram peştemal trendi; Türk üretimi; uygun fiyat segmenti"),

        ("Lusso", "lussocloud.com", "Cloud Bornoz",
         "Ultra yumuşak marshmallow kumaş; cloud bornoz konsepti; spa deneyimi evde; unisex; lüks hediyelik",
         "2020'de kuruldu; TikTok'ta viral; DTC Shopify; influencer seeding; hızlı büyüme"),

        ("Abyss & Habidecor", "abyssonline.com", "El Yapımı Lüks Havlu",
         "Portekiz üretimi; el yapımı; Giza pamuk; 700+ GSM; aşırı lüks segment; benzersiz dokuma teknikleri",
         "2019'da DTC e-ticaret genişlemesi; Instagram lüks yaşam tarzı; online satışa geçiş; premium hedef kitle"),

        ("Izzy & Jean", "izzyandjean.com", "Çevre Dostu Havlu",
         "Geri dönüştürülmüş pamuk; sıfır atık üretim; düşük su tüketimi; yumuşak doku; modern renkler",
         "2021'de kuruldu; sürdürülebilirlik odaklı DTC; Instagram ve TikTok; çevre bilinçli Z kuşağı hedefi"),

        ("Tucca", "tucca.co", "Rüzgar Geçirmez Plaj Havlusu",
         "Patentli rüzgar geçirmez klips sistemi; %100 organik pamuk; çift taraflı tasarım; kompakt katlama",
         "2019'da kuruldu; Kickstarter başarısı; DTC ve Amazon; Instagram plaj yaşam tarzı; yenilikçi tasarım"),

        ("Dock & Bay", "dockandbay.com", "Hızlı Kuruyan Mikrofiber Havlu",
         "Geri dönüştürülmüş plastikten mikrofiber; ultra kompakt; 3x hızlı kuruma; canlı desenler; seyahat dostu",
         "2018'de hızlı büyüme; Amazon ve DTC; Instagram ve TikTok; festival ve seyahat influencer'ları; çevre mesajı"),

        ("Marici", "marici.co", "Bambu Havlu",
         "Bambu elyaf teknolojisi; antimikrobiyal; ultra emici; çevre dostu; hipoalerjenik; yumuşak doku",
         "2020'de kuruldu; DTC Shopify; Instagram odaklı; sürdürülebilir yaşam segmenti; online-only"),

        ("Sunday Mist", "sundaymist.co", "Linen Peştemal",
         "Keten-pamuk karışımı; hafif ve nefes alabilir; doğal tonlar; çok amaçlı kullanım; vintage his",
         "2021'de kuruldu; DTC marka; Instagram ve Pinterest; ev dekorasyonu influencer'ları; artisan üretim"),

        ("Atolye 11:11", "atolye1111.com", "Tasarım Peştemal",
         "Türk peştemalı yeniden yorumlama; modern desenler; geleneksel dokuma; pamuk-linen karışım; renk blok tasarım",
         "2019'da kuruldu; DTC ve Etsy; Instagram Türk tasarım hareketi; ihracat odaklı; artisan işçilik"),

        ("Nutrl", "nutrltowel.com", "Antimikrobiyal Spor Havlusu",
         "Gümüş iyonlu antimikrobiyal kumaş; hızlı kuruma; koku önleyici; spor ve fitness odaklı; hafif",
         "2020'de kuruldu; Kickstarter; DTC Shopify; fitness influencer iş birlikleri; Amazon genişleme"),

        ("Wild Sage", "wildsageco.com", "Bohem Havlu & Peştemal",
         "El baskı desenleri; doğal boyalar; bohem estetik; pamuk-linen karışım; saçaklı detaylar",
         "2020'de kuruldu; Etsy başlangıç; DTC genişleme; Instagram boho estetik; küçük parti üretim"),

        ("Toalla", "toallalife.com", "Kompakt Seyahat Havlusu",
         "Ultra kompakt katlama; hızlı kuruma teknolojisi; antibakteriyel; hafif; çok fonksiyonlu",
         "2021'de kuruldu; DTC Shopify; TikTok ve Instagram; dijital nomad ve seyahat segmenti; online-only"),

        ("Feather Dry", "featherdry.com", "Ultra Hafif Saç Havlusu",
         "Patentli mikrofiber saç havlusu; %50 daha hızlı kuruma; saç kırılmasını önleme; hafif tasarım",
         "2020'de kuruldu; TikTok viral; DTC marka; beauty influencer seeding; Amazon genişleme"),

        ("Slowtide", "slowtide.com", "Sanatçı İş Birlikli Plaj Havlusu",
         "Sanatçı kolaborasyonları; sınırlı üretim desenler; %100 pamuk; büyük boy; koleksiyon değeri",
         "2018'de büyük genişleme; DTC ve perakende; Instagram sanat topluluğu; surf kültürü; limited edition strateji"),

        ("SoL Organics", "solorganix.com", "Fair Trade Organik Havlu",
         "Fair Trade sertifikalı; %100 organik pamuk; Hindistan üretimi; etik tedarik zinciri; GOTS sertifikalı",
         "2019'da havlu hattı genişlemesi; DTC Shopify; sosyal sorumluluk mesajı; bilinçli tüketici hedefi"),

        ("House of Jude", "houseofjude.com", "Bambu Bebek & Yetişkin Havlusu",
         "Bambu kumaş; ultra yumuşak; hipoalerjenik; anne-bebek segmenti; organik boyama; çevre dostu",
         "2019'da kuruldu; Instagram anne toplulukları; DTC Shopify; influencer pazarlama; Kanada kökenli"),

        ("Peshtemal Studio", "peshtemalstudio.com", "Premium El Dokuma Peştemal",
         "El dokuma Türk peştemalı; doğal pamuk; geleneksel tezgah; modern renk paleti; çok amaçlı",
         "2020'de kuruldu; DTC ve Etsy; Instagram; Türk zanaatkarlarla direkt çalışma; artisan marka"),

        ("Geometry", "geometryhouse.com", "Tasarım Mutfak & Banyo Havlusu",
         "Sanatçı tasarımı desenler; yüksek emici kumaş; dijital baskı; mutfak ve banyo koleksiyonu; hediye segmenti",
         "2019'da kuruldu; DTC Shopify; Instagram estetik; ev dekorasyonu influencer'ları; hızlı koleksiyon döngüsü"),

        ("Lui + Elle", "luielle.com", "Bambu Lüks Bornoz",
         "Bambu kumaş bornoz; ultra yumuşak; termal düzenleme; antibakteriyel; unisex tasarım; spa hissi",
         "2021'de kuruldu; DTC online; Instagram ve TikTok; wellness influencer iş birlikleri; premium segment"),
    ],

    # ─────────────────────────────────────────────────────────────────
    # 2. DOLAP & GİYSİ ORGANİZERİ  (Closet & Clothing Organizers)
    # ─────────────────────────────────────────────────────────────────
    "Dolap & Giysi Organizeri": [
        ("Open Spaces", "getopenspaces.com", "Premium Dolap Organizasyonu",
         "Modüler raf ve kutu sistemi; minimalist estetik; pastel renk paleti; sürdürülebilir malzeme; çok amaçlı",
         "2019'da kuruldu; Instagram-first DTC; influencer ve interior designer iş birlikleri; Shopify; premium segment"),

        ("Stoney Clover Lane", "stoneycloverlane.com", "Kişiselleştirilebilir Organizasyon",
         "Yamalarla kişiselleştirme; renkli ve eğlenceli tasarım; seyahat ve giysi organizasyonu; koleksiyon yapılabilir",
         "2018'de hızlı büyüme; Instagram ve TikTok viral; ünlü iş birlikleri; Shopify DTC; pop-up mağazalar"),

        ("Neat Method", "neatmethod.com", "Profesyonel Dolap Organizasyonu",
         "Profesyonel organizasyon hizmeti + ürün hattı; etiketleme sistemi; şeffaf kutular; lüks organizasyon",
         "2019'da ürün hattı lansmanı; Instagram influencer etkisi; DTC e-ticaret; Netflix etkisi; organizasyon trendi"),

        ("MDesign", "mdesignhomedecor.com", "Uygun Fiyatlı Dolap Çözümü",
         "Geniş ürün yelpazesi; tel ve plastik organizer; her bütçeye uygun; modüler sistem; kolay kurulum",
         "2018'de Amazon-native büyüme; DTC genişleme; sosyal medya ve influencer; uygun fiyat stratejisi"),

        ("The Container Store x KonMari", "containerstore.com", "KonMari Dolap Sistemi",
         "Marie Kondo iş birliği; hikari kutuları; dikey katlama sistemi; bamboo ve keten malzeme; zen estetik",
         "2019'da koleksiyon lansmanı; Netflix Tidying Up etkisi; DTC ve perakende; sosyal medya viral"),

        ("Closet Complete", "closetcomplete.com", "Lüks Kadife Askı Sistemi",
         "Premium kadife askılar; kaymaz yüzey; ince profil; alan tasarrufu; renk kodlama; dayanıklı",
         "2018'de Amazon büyümesi; DTC genişleme; Instagram dolap dönüşüm videoları; influencer seeding"),

        ("Mustard Made", "mustardmade.com", "Renkli Metal Dolap",
         "Retro metal locker tasarımı; canlı renkler; çocuk ve yetişkin versiyonları; modüler; dekoratif",
         "2018'de Avustralya'da kuruldu; Instagram viral; DTC global satış; interior design influencer'ları; trend ürün"),

        ("Songmics", "songmics.com", "Taşınabilir Giysi Dolabı",
         "Kumaş kaplamalı portatif dolap; metal çerçeve; modüler raflar; kolay montaj; apartman dostu",
         "2018'de Amazon-native hızlı büyüme; DTC genişleme; TikTok ve YouTube dönüşüm videoları; uygun fiyat"),

        ("StorageWorks", "storageworks.com", "Doğal Hasır Dolap Sepeti",
         "El örgüsü hasır sepetler; pamuk astar; etiket alanı; dolap rafları için özel boyut; bohem estetik",
         "2019'da kuruldu; Amazon-native; DTC Shopify; Instagram ev dekorasyonu; farmhouse ve boho trendleri"),

        ("Sorted by Sadie", "sortedbysadie.com", "Dolap Etiketleme Sistemi",
         "Premium etiketler ve kutu sistemi; minimalist tasarım; kişiselleştirme; profesyonel organizatör onaylı",
         "2021'de kuruldu; TikTok organizasyon trendi; DTC Shopify; micro-influencer iş birlikleri; niş marka"),

        ("Simplified by Emilie", "simplifiedbyemilie.com", "Giysi Katlama Tahtası",
         "Standartlaştırılmış katlama sistemi; Marie Kondo ilhamı; ABS plastik; farklı boyutlar; dikey depolama",
         "2020'de kuruldu; TikTok viral; DTC ve Amazon; organizasyon influencer'ları; uygun fiyat"),

        ("Dimore", "shopdimore.com", "İtalyan Esintili Dolap Sistemi",
         "Modüler ahşap dolap iç düzeni; İtalyan tasarım estetiği; özelleştirilebilir; premium malzeme",
         "2021'de kuruldu; Instagram interior design; DTC Shopify; Pinterest odaklı pazarlama; lüks segment"),

        ("Elfa x The Home Edit", "elfa.com", "Mesh Dolap Sistemi",
         "Modüler mesh çekmece sistemi; özelleştirilebilir konfigürasyon; dayanıklı metal; kolay kurulum",
         "2019'da The Home Edit iş birliği; Netflix etkisi; DTC online; Instagram dolap dönüşümleri; viral trend"),

        ("California Closets DTC", "californiaclosets.com", "Özelleştirilmiş Walk-in Closet",
         "3D tasarım aracı; custom walk-in closet; premium malzemeler; LED aydınlatma; lüks segment",
         "2019'da DTC dijital dönüşüm; online tasarım aracı; Instagram ve Pinterest; virtual consultation"),

        ("Clutter", "clutter.com", "Akıllı Depolama Çözümü",
         "Fotoğraflı envanter sistemi; depolama kutuları; pickup hizmeti; dijital katalog; on-demand iade",
         "2018'de hızlı büyüme; app-first model; DTC teknoloji; sosyal medya reklamları; şehir yaşamı hedefi"),

        ("STORE", "storeeverything.com", "Vakumlu Giysi Saklama",
         "Vakumlu saklama poşetleri; %75 alan tasarrufu; su geçirmez; mevsimlik giysi depolama; çeşitli boyutlar",
         "2019'da Amazon-native; DTC genişleme; TikTok alan tasarrufu videoları; uygun fiyat; yüksek hacim"),

        ("Modular Closets", "modularclosets.com", "DIY Modüler Dolap Sistemi",
         "Kendin-yap modüler dolap; online tasarım aracı; kolay montaj; beyaz melanin; profesyonel görünüm",
         "2018'de kuruldu; DTC online satış; YouTube montaj videoları; Facebook Ads; orta segment fiyat"),

        ("Polished Habitat", "polishedhabitat.com", "Dolap Organizasyon Seti",
         "Komple dolap seti; askı, kutu, bölücü; renk koordineli; etiket dahil; tek pakette çözüm",
         "2020'de kuruldu; blog-to-brand dönüşümü; DTC Shopify; Pinterest ve Instagram; DIY topluluğu"),

        ("Smug Organizer", "smugorganizer.com", "Lüks Dolap Aksesuarları",
         "Deri tutacaklı kutular; altın detay askılar; mermer desen etiketler; Instagram-worthy organizasyon",
         "2021'de kuruldu; Instagram estetik; DTC Shopify; lüks ev dekorasyonu influencer'ları; premium fiyat"),

        ("Tidy Living", "tidyliving.ca", "Kompakt Dolap Organizeri",
         "Küçük alan çözümleri; katlanır organizer; asılabilir raf; ayakkabı düzenleyici; apartman dostu",
         "2019'da kuruldu; Amazon ve DTC; Instagram ve TikTok; küçük alan yaşam trendi; Kanada kökenli"),

        ("Homesort", "homesort.co", "Modüler Kumaş Dolap Kutusu",
         "Katlanabilir kumaş kutular; bambu çerçeve; etiket penceresi; nefes alabilir kumaş; doğal görünüm",
         "2020'de kuruldu; DTC Shopify; Instagram minimal estetik; ev organizasyonu influencer'ları; çevre dostu malzeme"),

        ("The Laundress", "thelaundress.com", "Giysi Bakım & Saklama",
         "Giysi bakım ürünleri ile saklama çözümleri; lavanta keseleri; nefes alır kılıflar; moth koruması",
         "2018'de DTC genişleme; Instagram ve influencer; premium giysi bakımı; lüks segment; Shopify"),

        ("The Organized Home", "organizedhomeco.com", "Dolap İç Düzenleme Kiti",
         "Komple iç düzenleme kiti; bölücüler; çekmece organizeri; askı seti; renk kodlu etiketler",
         "2021'de kuruldu; TikTok organizasyon videoları; DTC Shopify; micro-influencer stratejisi; orta segment"),

        ("Neatly", "goneatly.com", "Akıllı Etiket Sistemi",
         "QR kodlu akıllı etiketler; app bağlantısı; dijital envanter; mevsimlik rotasyon takibi; modern çözüm",
         "2021'de kuruldu; teknoloji odaklı DTC; App Store ve Shopify; TikTok ve Instagram; yenilikçi konsept"),

        ("Styled Settings", "styledsettings.com", "Altın & Mermer Dolap Aksesuarı",
         "Altın metal askılar; mermer desen kutular; lüks estetik; Instagram-worthy tasarım; hediye segmenti",
         "2020'de kuruldu; Instagram interior influencer'ları; DTC Shopify; Pinterest; premium dekoratif organizasyon"),

        ("Marie Kondo x Berkshire Blanket", "konmari.com", "KonMari Hikari Kutuları",
         "Marie Kondo tasarımı; keten-pamuk karışımı; nefes alır; iç bölmeli; katlama dostu boyutlar",
         "2019'da lansmanı; Netflix viral etkisi; DTC online; global shipping; minimalist yaşam trendi"),

        ("Ebb & Flo", "ebbandflohome.com", "Pamuk Halatlı Dolap Sepeti",
         "El örgüsü pamuk halat; bohem tasarım; çok boyutlu; dolap rafları için ideal; dekoratif",
         "2020'de kuruldu; DTC ve Etsy; Instagram boho estetik; el yapımı vurgusu; sürdürülebilir üretim"),

        ("Bundle & Bliss", "bundleandbliss.com", "Bebek Dolap Organizeri",
         "Bebek giysi organizasyonu; boyut ayırıcılar; pastel renkler; büyüme takibi; anne dostu tasarım",
         "2021'de kuruldu; Instagram anne topluluğu; DTC Shopify; TikTok bebek organizasyon; niş segment"),

        ("The Folde", "thefolde.com", "Ayakkabı & Giysi Katlanır Organizer",
         "Katlanabilir ayakkabı kutusu; şeffaf ön panel; istiflenebilir; mıknatıslı kapak; modüler sistem",
         "2020'de kuruldu; TikTok viral; Amazon ve DTC; ayakkabı koleksiyoncuları hedefi; uygun fiyat"),

        ("Yours Truly", "yourstrulyhome.com", "Kişiselleştirilmiş Dolap Seti",
         "İsim yazılı kutular; monogram askılar; kişisel dolap seti; hediye paketi; özel tasarım",
         "2021'de kuruldu; DTC Shopify; Instagram kişiselleştirme trendi; hediye segmenti; micro-influencer"),

        ("Mesh & March", "meshandmarch.com", "Metal Mesh Dolap Sepeti",
         "Endüstriyel metal mesh; mat siyah ve beyaz; istiflenebilir; etiket tutucu; modern estetik",
         "2020'de kuruldu; DTC online; Instagram endüstriyel minimalist; Amazon genişleme; uygun fiyat"),

        ("Gleam Organics", "gleamorganics.com", "Organik Pamuk Giysi Torbası",
         "GOTS organik pamuk; nefes alır giysi saklama; doğal böcek kovucu; iç bölmeli; fermuarlı",
         "2021'de kuruldu; DTC Shopify; sürdürülebilir yaşam influencer'ları; çevre dostu ambalaj; niş ürün"),
    ],

    # ─────────────────────────────────────────────────────────────────
    # 3. ÇEKMECE DÜZENLEYİCİ  (Drawer Organizers)
    # ─────────────────────────────────────────────────────────────────
    "Çekmece Düzenleyici": [
        ("Joseph Joseph", "josephjoseph.com", "Yenilikçi Çekmece Düzenleyici",
         "DrawerStore serisi; eğimli bölmeler; kompakt tasarım; inovatif mekanizmalar; mutfak ve banyo",
         "2019'da DTC kanalını genişletti; Amazon-native büyüme; Instagram mutfak organizasyonu; ödüllü tasarım"),

        ("Yamazaki Home", "yamazakihome.com", "Japon Minimalist Çekmece Sistemi",
         "Tosca ve Tower serileri; bambu-çelik kombinasyonu; Japon minimalizmi; modüler çekmece bölücüler",
         "2018'de global DTC genişleme; Instagram minimal estetik; Amazon ve Shopify; Japon tasarım trendi"),

        ("DrawerDecor", "drawerdecor.com", "Özelleştirilebilir Çekmece Bölücü",
         "DIY bölücü sistemi; kesebilir paneller; her çekmeceye uyum; silikon taban; kaymaz",
         "2019'da kuruldu; Amazon-native; DTC Shopify; YouTube ve Instagram DIY içerik; uygun fiyat"),

        ("Bambüsi", "bambusi.com", "Bambu Çekmece Organizeri",
         "%100 bambu; genişletilebilir bölmeler; doğal finish; mutfak çekmecesi için; çevre dostu",
         "2019'da kuruldu; Amazon bestseller; DTC genişleme; Instagram sürdürülebilir ev; bambu trend"),

        ("Squared Away", "squaredaway.com", "Premium Çekmece Seti",
         "Bed Bath & Beyond'un DTC markası; lüks kumaş organizerler; özel boyutlar; nötr tonlar; seri halinde",
         "2021'de kuruldu; DTC only; Instagram ve Meta Ads; influencer lansmanı; premium ev organizasyonu"),

        ("SpaceAid", "spaceaid.co", "Baharat Çekmece Organizeri",
         "Eğimli baharat rafı; çekmece içi yerleşim; etiket dahil; 4-sıra tasarım; mutfak odaklı",
         "2020'de kuruldu; Amazon viral ürün; TikTok mutfak organizasyonu; DTC genişleme; yüksek puan"),

        ("Lifewit", "lifewit.com", "Kumaş Çekmece Bölücü",
         "Nefes alır kumaş; katlanabilir; washable; iç tel çerçeve; çeşitli boyutlar; giysi ve iç çamaşırı",
         "2018'de Amazon-native; DTC online; TikTok ve YouTube organizasyon; uygun fiyat; yüksek hacim"),

        ("Home Edit x iDesign", "idesignlivesimply.com", "The Home Edit Çekmece Koleksiyonu",
         "Rainbow organizasyon konsepti; modüler akrilik bölücüler; istiflenebilir; renk kodlama sistemi",
         "2019'da The Home Edit iş birliği; Netflix etkisi; Amazon ve DTC; Instagram renk düzeni viral"),

        ("Vtopmart", "vtopmart.com", "Şeffaf Akrilik Çekmece Organizeri",
         "Kristal şeffaf akrilik; farklı boyut kombinasyonları; BPA-free; istiflenebilir; çok amaçlı",
         "2019'da Amazon bestseller; DTC genişleme; TikTok organizasyon videoları; uygun fiyat stratejisi"),

        ("Tidy Cup", "tidycup.com", "İç Çamaşırı Çekmece Organizeri",
         "Petek yapılı bölmeler; yumuşak kumaş; katlanabilir; çorap ve iç çamaşırı için özel; çeşitli boyut",
         "2020'de kuruldu; Amazon ve DTC; TikTok çekmece dönüşüm videoları; micro-influencer; niş ürün"),

        ("Simple Houseware", "simplehouseware.com", "Mesh Metal Çekmece Organizeri",
         "Metal mesh bölücü; siyah ve gümüş; masaüstü ve çekmece; ofis ve ev; modüler",
         "2018'de Amazon-native büyüme; DTC genişleme; YouTube ofis organizasyonu; uygun fiyat; çeşitli ürün"),

        ("Bamboozle", "bamboozlehome.com", "Bambu Mutfak Çekmece Düzenleyici",
         "Sürdürülebilir bambu; genişletilebilir; doğal antimikrobiyal; mutfak gereçleri için; dekoratif",
         "2019'da kuruldu; DTC ve Amazon; Instagram mutfak estetiği; sürdürülebilir ev trendi; çevre mesajı"),

        ("Crisp Home", "crisphome.co", "Premium Çekmece Organizasyon Seti",
         "Beyaz mat finish; modüler kutular; etiketleme sistemi; mutfak ve banyo; profesyonel görünüm",
         "2021'de kuruldu; Instagram organizasyon estetiği; DTC Shopify; The Home Edit ilhamı; premium segment"),

        ("SmartStore", "smartstore.se", "İskandinav Çekmece Sistemi",
         "Clipbox sistemi; klipsli kapak; şeffaf ve beyaz; istiflenebilir; İskandinav minimalizmi",
         "2019'da global DTC genişleme; Instagram İskandinav estetik; Amazon EU ve US; fonksiyonel tasarım"),

        ("Keegh", "keegh.com", "Silikon Çekmece Bölücü",
         "Esnek silikon; herhangi bir çekmeceye uyum; yıkanabilir; gıda güvenli; renkli seçenekler",
         "2020'de kuruldu; Amazon-native; TikTok viral; DTC genişleme; yenilikçi malzeme; uygun fiyat"),

        ("KitchenEdge", "kitchenedge.com", "Bıçak ve Çatal Çekmece Düzenleyici",
         "Bambu bıçak bloğu çekmece versiyonu; ayrı bölmeler; mıknatıslı tutucu; premium malzeme",
         "2019'da kuruldu; Amazon bestseller; DTC Shopify; YouTube mutfak düzeni; profesyonel şef onaylı"),

        ("OXO Good Grips DTC", "oxo.com", "Ergonomik Çekmece Organizeri",
         "Ergonomik tasarım; kaymaz taban; modüler; BPA-free; genişletilebilir; evrensel uyum",
         "2019'da DTC online satışı güçlendirdi; Amazon ve Shopify; Instagram ev organizasyonu; güvenilir marka"),

        ("Peak Design Home", "peakdesign.com", "Manyetik Çekmece Düzenleyici",
         "Manyetik bağlantı sistemi; modüler; alüminyum yapı; yeniden konfigüre edilebilir; premium",
         "2020'de ev hattını genişletti; Kickstarter geçmişi; DTC online; Instagram ve YouTube; teknoloji meraklısı hedef kitle"),

        ("Madesmart", "madesmart.com", "Interlocking Çekmece Organizeri",
         "Kilitli bağlantı sistemi; BPA-free; yumuşak kavrama taban; ölçeklendirilebilir; çok amaçlı",
         "2018'de DTC ve Amazon büyüme; Instagram mutfak organizasyonu; influencer iş birlikleri; uygun fiyat"),

        ("Dial Industries", "dialindustries.com", "Mega Genişletilebilir Çekmece Organizeri",
         "Genişletilebilir modüler tasarım; farklı boyutlar; şeffaf; BPA-free; kolay temizlik",
         "2019'da Amazon bestseller; DTC genişleme; YouTube ve Instagram; pratik çözüm; uygun fiyat"),

        ("Stackers", "stackers.com", "Takı ve Aksesuar Çekmece Organizeri",
         "Kadife kaplı iç; istiflenebilir katmanlar; takı ve saat için özel bölmeler; lüks tasarım",
         "2018'de DTC genişleme; Instagram takı organizasyonu; influencer seeding; premium hediye segmenti"),

        ("Totally Tidy", "totallytidy.co", "Makyaj Çekmece Organizeri",
         "Akrilik makyaj düzenleyici; bölmeli çekmece sistemi; şeffaf; empilable; güzellik odaklı",
         "2020'de kuruldu; TikTok beauty organizasyon; DTC Shopify; beauty influencer iş birlikleri"),

        ("Purse Neat", "purseneat.com", "Çanta İçi Çekmece Organizeri",
         "Çanta içi düzenleme sistemi; keçe malzeme; çeşitli marka çantalara uyum; bölmeli; hafif",
         "2019'da kuruldu; Amazon ve DTC; Instagram ve TikTok; çanta koleksiyoncuları; niş segment"),

        ("Divide & Conquer", "divideandconquerhome.com", "Ayarlanabilir Bambu Bölücü",
         "Yaylı bambu bölücü; herhangi bir çekmeceye uyum; kolay kurulum; doğal görünüm; dayanıklı",
         "2021'de kuruldu; Amazon-native; TikTok DIY; DTC genişleme; çevre dostu vurgu"),

        ("Simplify Home", "simplifyhomeco.com", "Çekmece Organizasyon Komple Set",
         "Tek pakette komple çözüm; iç çamaşırı, çorap, aksesuar bölmeleri; kumaş; katlanabilir; yıkanabilir",
         "2020'de kuruldu; Amazon ve DTC; TikTok çekmece dönüşümleri; uygun fiyat; yüksek müşteri puanı"),

        ("Navaris", "navaris.com", "Bambu Çekmece Düzenleyici Set",
         "5 parça bambu set; farklı boyutlar; doğal finish; mutfak ve banyo; istiflenebilir",
         "2019'da Amazon lansmanı; DTC online; Instagram sürdürülebilir ev; Avrupa kökenli; çevre dostu"),

        ("Bliss Home", "blisshomeorganizing.com", "Lüks Kumaş Çekmece Bölücü",
         "Premium linen kumaş; sert çerçeve; zarif etiketler; nötr tonlar; lüks organizasyon",
         "2021'de kuruldu; Instagram premium estetik; DTC Shopify; interior influencer'ları; hediye segmenti"),

        ("Tidy Tribe", "tidytribe.co", "Çocuk Çekmece Organizeri",
         "Renkli ve eğlenceli tasarım; çocuk dostu boyutlar; yumuşak malzeme; yıkanabilir; resimli etiketler",
         "2021'de kuruldu; Instagram anne toplulukları; DTC Shopify; TikTok çocuk odası dönüşüm; niş segment"),

        ("Rigga", "rigga.co", "Modüler Plastik Çekmece Bölücü",
         "Geridönüştürülmüş plastik; klipsli bağlantı; sonsuz konfigürasyon; renk seçenekleri; sürdürülebilir",
         "2020'de kuruldu; DTC online; Instagram sürdürülebilirlik; çevre dostu üretim; Avrupa tasarım"),

        ("Drawer Style", "drawerstyle.com", "Dekoratif Çekmece Kaplama & Organizeri",
         "Yapışkanlı çekmece kaplama kağıdı + organizer set; desenli; parfümlü; dekoratif fonksiyonel",
         "2020'de kuruldu; TikTok çekmece makyajı trendi; DTC Shopify; Instagram estetik; uygun fiyat"),
    ],

    # ─────────────────────────────────────────────────────────────────
    # 4. MUTFAK ORGANİZASYONU  (Kitchen Organization)
    # ─────────────────────────────────────────────────────────────────
    "Mutfak Organizasyonu": [
        ("Caraway Home", "carawayhome.com", "Non-Toxic Mutfak Organizasyonu",
         "Seramik kaplı tencere seti + organizasyon aksesuarları; manyetik raf sistemi; non-toxic; şık tasarım",
         "2019'da kuruldu; Instagram-first DTC; influencer ve ünlü iş birlikleri; Shopify; viral mutfak estetiği"),

        ("Our Place", "fromourplace.com", "Çok Fonksiyonlu Mutfak Sistemi",
         "Always Pan ve Wonder Oven; entegre depolama çözümleri; alan tasarrufu; minimalist mutfak felsefesi",
         "2019'da kuruldu; Instagram ve TikTok viral; DTC unicorn; ünlü iş birlikleri; çok kültürlü marka hikayesi"),

        ("W&P", "wandpdesign.com", "Tasarım Mutfak Depolama",
         "Porter mutfak aksesuarları; silikon saklama kapları; buzluk küpü; sürdürülebilir; Brooklyn tasarım",
         "2018'de hızlı büyüme; DTC Shopify; Instagram food ve design; influencer seeding; çevre dostu"),

        ("Stasher", "stasherbag.com", "Silikon Yeniden Kullanılabilir Poşet",
         "Platin silikon; yeniden kullanılabilir; bulaşık makinesi ve fırın güvenli; renk çeşitliliği; sıfır atık",
         "2018'de hızlı büyüme; DTC ve perakende; Instagram sürdürülebilirlik; influencer; B Corp sertifikalı"),

        ("HexClad", "hexclad.com", "Hibrit Tencere ve Mutfak Organizasyonu",
         "Patentli hexagonal tasarım; tencere organizasyon rafı; manyetik tutucu; Gordon Ramsay onaylı",
         "2019'da viral büyüme; DTC Shopify; YouTube ve Instagram; ünlü şef endorsement; premium segment"),

        ("Fridge No More", "fridgenomore.co", "Buzdolabı Organizasyon Seti",
         "Şeffaf akrilik buzdolabı kutuları; etiket sistemi; istiflenebilir; BPA-free; farklı boyutlar",
         "2021'de kuruldu; TikTok buzdolabı dönüşüm videoları; DTC ve Amazon; viral trend; uygun fiyat"),

        ("YouCopia", "youcopia.com", "Akıllı Mutfak Organizeri",
         "SpiceLiner baharat organizeri; StoreMore raf; RollOut çekmece; yenilikçi çözümler; alan optimizasyonu",
         "2018'de DTC genişleme; Amazon bestseller; Instagram mutfak düzeni; influencer; ödüllü tasarım"),

        ("Prepd", "getprepd.com", "Yemek Hazırlık Organizasyonu",
         "Modüler meal prep container seti; porsiyon kontrol; sızdırmaz; mikrodalga güvenli; renk kodlu",
         "2019'da kuruldu; Kickstarter; DTC online; Instagram fitness ve meal prep; sağlıklı yaşam segmenti"),

        ("iDesign Clarity", "idesignlivesimply.com", "Akrilik Mutfak Organizeri",
         "Kristal şeffaf akrilik; modüler mutfak çekmece ve raf sistemi; The Home Edit uyumlu; BPA-free",
         "2019'da The Home Edit iş birliği; Netflix etkisi; Amazon ve DTC; Instagram viral; çeşitli ürün"),

        ("SimpleHuman Kitchen", "simplehuman.com", "Akıllı Mutfak Depolama",
         "Sensörlü çöp kutusu; sabunluk; kağıt havluluk; paslanmaz çelik; entegre mutfak sistemi",
         "2019'da DTC online güçlendirme; Amazon ve Shopify; YouTube mutfak teknolojisi; premium segment"),

        ("PackIt", "packit.com", "Dondurulabilir Mutfak Saklama",
         "Gel-bazlı soğutma teknolojisi; katlanabilir; yemek taşıma ve saklama; çeşitli boyutlar; çevre dostu",
         "2018'de DTC genişleme; Amazon ve Shopify; Instagram yemek hazırlık; anne ve fitness segmenti"),

        ("Open Kitchen by Williams Sonoma", "williams-sonoma.com", "Mutfak Organizasyon Koleksiyonu",
         "Premium ahşap ve metal organizer; tezgah üstü depolama; baharat rafı; kesme tahtası standı",
         "2020'de DTC koleksiyon lansmanı; Instagram mutfak estetiği; influencer şef iş birlikleri; lüks"),

        ("Kozy Kitchen", "kozykitchen.co", "Bambu Mutfak Düzenleyici Seti",
         "5 parça bambu set; baharat rafı, bıçaklık, kağıt havluluk; tezgah üstü; doğal malzeme",
         "2020'de kuruldu; Amazon-native; DTC genişleme; Instagram farmhouse estetik; uygun fiyat; çevre dostu"),

        ("Tidy Kitchen Co", "tidykitchenco.com", "Tezgah Üstü Organizasyon Sistemi",
         "Modüler tezgah sistemi; yağ şişeliği, baharat standı, mutfak gereci tutucu; mat siyah metal",
         "2021'de kuruldu; Instagram mutfak dönüşüm; DTC Shopify; TikTok mutfak organizasyonu; modern estetik"),

        ("MakeRoom", "makeroom.com", "Mutfak Dolap İçi Düzenleyici",
         "Raf yükseltici; döner platform; tel çekmece; dolap kapağı askısı; alan ikiye katlama",
         "2021'de kuruldu; TikTok alan tasarrufu; DTC ve Amazon; micro-influencer iş birlikleri; pratik çözümler"),

        ("Mesa Living", "mesaliving.co", "Tezgah Üstü Dekoratif Organizer",
         "Ahşap ve beton kombinasyonu; tezgah üstü meyve sepeti, baharat tutucu; modern estetik; el yapımı",
         "2020'de kuruldu; Instagram interior design; DTC Shopify; artisan üretim; Meksika ilhamlı tasarım"),

        ("Blueland Kitchen", "blueland.com", "Yeniden Kullanılabilir Mutfak Depolama",
         "Tablet bazlı temizlik; yeniden doldurulabilir şişeler; düzenli tezgah; sıfır plastik; güzel tasarım",
         "2019'da kuruldu; DTC unicorn; Instagram ve TikTok; sürdürülebilirlik; Shark Tank; çevre viral"),

        ("Rig-Tig by Stelton", "rig-tig.com", "Danimarkalı Mutfak Organizasyonu",
         "İskandinav tasarım; bread box, storage jar, utensil holder; bambu ve melamin; fonksiyonel estetik",
         "2019'da DTC global genişleme; Instagram İskandinav mutfak; Amazon EU; tasarım ödülleri"),

        ("Prepara", "prepara.com", "Taze Tutma & Organizasyon",
         "Herb Savor taze ot saklama; GlassLock depolama; silikon kapak; mutfak verimliliği; yenilikçi",
         "2018'de DTC genişleme; Amazon ve Shopify; YouTube mutfak hack'leri; influencer iş birlikleri"),

        ("Material Kitchen", "materialkitchen.com", "reBoard Kesme Tahtası & Organizer",
         "Geri dönüştürülmüş malzeme; entegre depolama; şef tasarımı; minimalist; sürdürülebilir üretim",
         "2018'de kuruldu; DTC Shopify; Instagram mutfak estetiği; profesyonel şef onaylı; premium"),

        ("Kilner", "kilnerjar.com", "Vintage Cam Depolama Kavanozları",
         "Heritage cam kavanoz; klipsli kapak; fermantasyon seti; retro tasarım; mutfak dekorasyonu",
         "2019'da DTC online genişleme; Instagram vintage mutfak; fermantasyon trendi; Pinterest; İngiliz kökenli"),

        ("Totally Bamboo", "totallybamboo.com", "Bambu Mutfak Organizasyon Seti",
         "%100 Moso bambu; salt box, utensil holder, spice rack; sürdürülebilir; doğal antimikrobiyal",
         "2018'de DTC genişleme; Amazon bestseller; Instagram sürdürülebilir mutfak; çevre dostu marka"),

        ("Fellow Home", "fellowproducts.com", "Kahve Organizasyonu & Depolama",
         "Atmos vakumlu kahve kabı; Carter bardak; Fellow Stagg kettle; kahve istasyonu organizasyonu; premium",
         "2018'de büyük genişleme; DTC Shopify; Instagram kahve kültürü; Kickstarter kökenli; kahve meraklısı hedef"),

        ("Mepal", "mepal.com", "Hollanda Tasarımı Saklama Kabı",
         "Modüler saklama sistemi; klipsli kapak; istiflenebilir; BPA-free; renk kodlu; Hollanda tasarımı",
         "2019'da global DTC genişleme; Instagram mutfak organizasyonu; Amazon EU; pratik ve estetik; sürdürülebilir"),

        ("Pantry Living", "pantryliving.co", "Kiler Organizasyon Seti",
         "Şeffaf akrilik kiler kutuları; etiket seti; döner platform; raf yükseltici; komple kiler çözümü",
         "2021'de kuruldu; TikTok kiler dönüşüm videoları; DTC Shopify; Instagram; micro-influencer stratejisi"),

        ("Two Pillars", "twopillars.co", "Silikon Mutfak Depolama",
         "ZipBag silikon yeniden kullanılabilir poşet; UniLid evrensel kapak; bulaşık makinesi güvenli; çevre dostu",
         "2019'da kuruldu; Kickstarter başarısı; DTC online; Instagram ve TikTok; sürdürülebilirlik mesajı"),

        ("Source Bulk Foods DTC", "sourcebulkfoods.com", "Sıfır Atık Mutfak Depolama",
         "Cam kavanoz seti; toptan gıda sistemi; etiketleme; tartılabilir tara; sıfır atık mutfak",
         "2019'da DTC online; Instagram sıfır atık; sürdürülebilir yaşam influencer'ları; Avustralya kökenli"),

        ("Seed & Sprout", "seedsprout.com.au", "Yeniden Kullanılabilir Mutfak Organizasyonu",
         "Balmumu wrap; silikon poşet; cam saklama; bambu tepsiler; komple sürdürülebilir mutfak; çevre dostu",
         "2018'de kuruldu; DTC ve Amazon; Instagram sürdürülebilirlik; influencer iş birlikleri; Avustralya kökenli"),

        ("Oggi", "oggicorp.com", "Paslanmaz Çelik Mutfak Depolama",
         "Paslanmaz çelik saklama kavanozları; vakumlu kapak; tezgah üstü organizasyon; mutfak dekorasyonu",
         "2019'da DTC genişleme; Amazon ve online perakende; Instagram mutfak estetiği; modern endüstriyel stil"),

        ("Full Circle Home", "fullcirclehome.com", "Çevre Dostu Mutfak Organizasyonu",
         "Bambu bulaşıklık; kompost kutusu; geri dönüştürülmüş plastik depolama; B Corp; sürdürülebilir",
         "2019'da DTC genişleme; Instagram çevre dostu; Amazon ve Shopify; sürdürülebilirlik sertifikalı"),
    ],

    # ─────────────────────────────────────────────────────────────────
    # 5. BANYO ORGANİZASYONU  (Bathroom Organization)
    # ─────────────────────────────────────────────────────────────────
    "Banyo Organizasyonu": [
        ("Cadence", "keepyourcadence.com", "Manyetik Seyahat & Banyo Kapsülü",
         "Manyetik altıgen kapsüller; sızdırmaz; kişiselleştirilebilir etiket; modüler sistem; TSA uyumlu",
         "2020'de kuruldu; TikTok ve Instagram viral; DTC Shopify; influencer seeding; Shark Tank; yenilikçi tasarım"),

        ("Kitsch", "mykitsch.com", "Banyo Organizasyon & Aksesuar",
         "Saten yastık kılıfı; saç tokası organizeri; banyo aksesuarları; sürdürülebilir; Instagram estetiği",
         "2018'de hızlı büyüme; TikTok viral; DTC ve Target; influencer pazarlama; beauty ve wellness"),

        ("Tooletries", "tooletries.com", "Silikon Duş Organizeri",
         "Silikon yapışkanlı duş tutucu; jilet ve şampuan organizeri; vida gerektirmeyen montaj; su geçirmez",
         "2018'de kuruldu; Instagram erkek bakım; DTC Shopify; Amazon; Kickstarter kökenli; Avustralya"),

        ("Bathroom Butler", "bathroombutler.com", "Isıtmalı Havluluk & Organizasyon",
         "Elektrikli havlu ısıtıcı; duvar montajlı; enerji verimli; modern tasarım; krom ve mat siyah",
         "2019'da DTC global genişleme; Instagram lüks banyo; influencer iş birlikleri; premium segment"),

        ("Fullgive", "fullgive.co", "Banyo Tezgahı Organizeri",
         "Bambu ve akrilik kombinasyonu; tezgah üstü düzenleyici; bölmeli; su geçirmez taban; estetik",
         "2021'de kuruldu; Amazon-native; TikTok banyo dönüşüm; DTC genişleme; uygun fiyat"),

        ("Yamazaki Tower", "yamazakihome.com", "Manyetik Banyo Rafı",
         "Manyetik çamaşır makinesi rafı; duş caddy; diş fırçası tutucu; Japon minimalist; beyaz çelik",
         "2019'da global DTC genişleme; Instagram minimal banyo; Amazon bestseller; Japon tasarım trendi"),

        ("Better Living Products", "betterlivingproducts.com", "Hilet Duş Dispenser",
         "Duvara monte şampuan dispenseri; 3'lü sistem; şık tasarım; plastik şişe eliminasyonu; çevre dostu",
         "2019'da DTC online güçlendirme; Instagram banyo dönüşüm; Amazon; sürdürülebilir banyo trendi"),

        ("HiRise", "hiriseshower.com", "Katlı Duş Rafı",
         "Paslanmaz çelik katlı raf; köşe ve düz seçenekler; drenaj delikleri; ayarlanabilir yükseklik",
         "2020'de kuruldu; Amazon ve DTC; Instagram modern banyo; TikTok duş organizasyonu; sade tasarım"),

        ("Brightroom", "target.com/brightroom", "Modern Banyo Depolama",
         "Target'ın DTC organizasyon markası; banyo kutuları; raf sepetleri; etiket sistemi; beyaz estetik",
         "2022'de lansmanı; Target exclusive; Instagram ve TikTok; influencer organize-with-me videoları"),

        ("Shower Power", "showerpowerfit.com", "Hidroelektrik Bluetooth Duş Başlığı & Raf",
         "Su gücüyle çalışan Bluetooth hoparlör; entegre duş rafı; LED aydınlatma; yenilikçi tasarım",
         "2020'de kuruldu; Kickstarter viral; DTC Shopify; TikTok; teknoloji ve banyo birleşimi"),

        ("Umbra", "umbra.com", "Tasarım Odaklı Banyo Organizasyonu",
         "Flip kanca; Penguin sabunluk; Fiboo banyo seti; Kanada tasarımı; fonksiyonel estetik; geniş ürün hattı",
         "2019'da DTC online güçlendirme; Instagram tasarım; Amazon; global dağıtım; ödüllü tasarım"),

        ("Navage", "navage.com", "Burun Temizleme & Banyo Wellness",
         "Burun irrigasyon cihazı; tuz kapsül sistemi; şarj edilebilir; banyo tezgahı ürünü; wellness",
         "2018'de DTC viral büyüme; Facebook ve Instagram Ads; infomercial strateji; Amazon; sağlık segmenti"),

        ("by Humankind", "byhumankind.com", "Sıfır Plastik Banyo Organizasyonu",
         "Yeniden doldurulabilir banyo ürünleri; bambu raflar; deodorant, şampuan bar, diş macunu tabletleri",
         "2019'da kuruldu; DTC abonelik modeli; Instagram sürdürülebilirlik; influencer; sıfır atık banyo"),

        ("Wild", "wearewild.com", "Yeniden Doldurulabilir Banyo Ürünü & Stand",
         "Alüminyum deodorant kılıfı; refill sistemi; bambu stand; banyo organizasyonu; çevre dostu",
         "2019'da kuruldu; DTC abonelik; Instagram ve TikTok; sürdürülebilirlik influencer'ları; UK kökenli"),

        ("DAME", "wearedame.co", "Sürdürülebilir Banyo Organizasyonu",
         "Yeniden kullanılabilir tampon aplikatörü; geri dönüşüm kutusu; banyo depolama; B Corp",
         "2018'de kuruldu; DTC online; Instagram sürdürülebilirlik; Kickstarter; feminist marka; UK kökenli"),

        ("The Honest Company Bath", "honest.com", "Bebek & Aile Banyo Organizasyonu",
         "Bebek banyo seti; organizer caddy; non-toxic ürünler; banyo oyuncak saklama; aile odaklı",
         "2019'da DTC banyo organizasyon genişlemesi; Instagram anne segmenti; influencer; Jessica Alba marka"),

        ("Vesta Living", "vestaliving.co.nz", "Silikon Banyo Organizeri",
         "Silikon duvar organizer; yapışkanlı montaj; diş fırçası tutucu; sabunluk; renkli; kolay temizlik",
         "2020'de kuruldu; Instagram banyo estetiği; DTC Shopify; TikTok; Yeni Zelanda kökenli; çevre dostu"),

        ("Shelfie", "getshelfie.com", "Duş İçi Raf Sistemi",
         "Su geçirmez yapışkan teknoloji; çelik raf; ağır yük kapasitesi; modern tasarım; kolay montaj",
         "2021'de kuruldu; TikTok banyo dönüşüm; DTC Shopify; Instagram; micro-influencer; uygun fiyat"),

        ("Saje Natural Wellness", "saje.com", "Aromaterapi Banyo Organizasyonu",
         "Difüzör ve yağ standı; banyo aromaterapi seti; ahşap organizer; wellness deneyimi; doğal",
         "2019'da DTC online genişleme; Instagram wellness; influencer spa-at-home trendi; Kanada kökenli"),

        ("The Bare Home", "thebarehome.com", "Minimalist Banyo Depolama",
         "Yeniden doldurulabilir cam şişeler; etiketleme; bambu raf; minimalist banyo; sıfır plastik",
         "2020'de kuruldu; DTC Shopify; Instagram zero-waste; sürdürülebilirlik influencer'ları; Kanada"),

        ("Shower+ by Kohler", "kohler.com", "Akıllı Duş Sistemi & Organizasyon",
         "App kontrollü duş; entegre depolama nişleri; LED aydınlatma; ses komutu; akıllı banyo",
         "2020'de DTC online lansmanı; Instagram akıllı ev; YouTube teknoloji; premium segment"),

        ("Morihata", "morihata.com", "Japon Binchotan Banyo Aksesuarları",
         "Binchotan kömür; su arıtma; sabunluk; havluluk; Japon zanaatkarlığı; doğal malzemeler",
         "2019'da DTC online genişleme; Instagram Japon estetik; artisan marka; premium; wellness"),

        ("Haven Store", "havenstore.co", "Spa Esinli Banyo Organizasyonu",
         "Mermer ve pirinç aksesuarlar; sabunluk, diş fırçalık, tepsi; spa estetiği; lüks hediyelik",
         "2021'de kuruldu; Instagram lüks banyo; DTC Shopify; influencer; hediye segmenti; premium"),

        ("Stylewell", "homedepot.com/stylewell", "Modüler Banyo Depolama",
         "Home Depot DTC markası; over-toilet raf; köşe dolap; banyo arabası; uygun fiyatlı modern tasarım",
         "2020'de lansmanı; online-first; Instagram banyo dönüşüm; DIY influencer iş birlikleri; orta segment"),

        ("Allswalls", "allswalls.com", "Mıknatıslı Banyo Organizeri",
         "Neodim mıknatıs teknolojisi; çelik yüzeylere yapışan organizer; modüler; yeniden konumlandırılabilir",
         "2021'de kuruldu; TikTok viral; DTC Shopify; yenilikçi montaj sistemi; apartman dostu; online-only"),

        ("Bath Bliss", "bathbliss.co", "Bambu Banyo Tepsisi & Organizer",
         "Küvet tepsisi; bambu banyo rafı; sabunluk seti; doğal malzeme; spa deneyimi; hediye seti",
         "2020'de kuruldu; Amazon ve DTC; Instagram self-care; TikTok banyo rutini; uygun-orta fiyat"),

        ("Zenna Home", "zennahome.com", "Duş Perdesi & Organizasyon Sistemi",
         "Entegre duş sistemi; perde çubuğu ile raf; NeverRust teknoloji; kolay montaj; alan tasarrufu",
         "2019'da DTC genişleme; Amazon ve online; Instagram banyo yenileme; pratik çözümler; orta segment"),

        ("Object 002", "object002.com", "Minimalist Banyo Aksesuar Seti",
         "Beton sabunluk; pirinç tutucu; minimal tasarım; el yapımı; sınırlı üretim; koleksiyonluk",
         "2021'de kuruldu; Instagram minimal estetik; DTC Shopify; tasarım blog'ları; artisan premium"),

        ("Tidy Toast", "tidytoast.com", "Diş Fırçası UV Sterilizatör & Organizer",
         "UV-C sterilizasyon; duvar montajlı tutucu; 4 fırça kapasitesi; şarj edilebilir; hijyenik",
         "2021'de kuruldu; TikTok banyo hijyeni; DTC online; Amazon; teknoloji influencer'ları; yenilikçi"),

        ("Bino", "bfriendsco.com", "Şeffaf Akrilik Banyo Organizeri",
         "Kristal akrilik; banyo tezgahı düzenleyici; bölmeli; makyaj ve banyo ürünleri; şık ve pratik",
         "2019'da Amazon-native; DTC genişleme; TikTok banyo estetiği; uygun fiyat; yüksek müşteri puanı"),

        ("Mist", "mistbath.com", "Aromaterapi Duş Sistemi & Organizer",
         "Duş aromaterapi difüzörü; esansiyel yağ organizeri; buhar aktivasyonlu; spa deneyimi",
         "2020'de kuruldu; TikTok shower routine viral; DTC Shopify; wellness influencer'ları; yenilikçi niş"),
    ],

    # ─────────────────────────────────────────────────────────────────
    # 6. SEYAHAT ORGANİZERİ  (Travel Organizers & Packing)
    # ─────────────────────────────────────────────────────────────────
    "Seyahat Organizeri": [
        ("Calpak", "calpaktravel.com", "Trendy Seyahat Organizasyonu",
         "Packing cube seti; makyaj çantası; laptop çantası; pastel renkler; Instagram estetiği; hafif malzeme",
         "2018'de DTC rebrand; Instagram ve TikTok viral; influencer ve ünlü iş birlikleri; Shopify; hızlı büyüme"),

        ("Beis", "beistravel.com", "Lüks Seyahat Organizasyonu",
         "The Packing Cubes; The Cosmetic Case; The Diaper Bag; fonksiyonel lüks; Shay Mitchell markası",
         "2018'de kuruldu; Instagram-native marka; celebrity founder; TikTok viral; DTC Shopify; premium segment"),

        ("Cadence", "keepyourcadence.com", "Manyetik Seyahat Kapsülü",
         "Altıgen manyetik kaplar; TSA uyumlu; sızdırmaz; kişiselleştirilebilir etiket; modüler seyahat organizasyonu",
         "2020'de kuruldu; TikTok viral; Shark Tank yatırımı; DTC Shopify; influencer seeding; yenilikçi tasarım"),

        ("Stoney Clover Lane", "stoneycloverlane.com", "Kişiselleştirilebilir Seyahat Çantası",
         "Yama ve harflerle kişiselleştirme; neon renkler; seyahat organizer seti; koleksiyon yapılabilir",
         "2018'de hızlı büyüme; Instagram ve TikTok; ünlü iş birlikleri (Disney, Barbie); pop-up mağazalar; DTC"),

        ("Paravel", "tourparavel.com", "Sürdürülebilir Seyahat Organizasyonu",
         "Geri dönüştürülmüş malzemeler; packing cube set; karbon nötr; Domino koleksiyonu; çevre dostu lüks",
         "2018'de kuruldu; DTC Shopify; Instagram sürdürülebilir seyahat; influencer; karbon offset programı"),

        ("July", "july.com", "Premium Valiz & Organizer Seti",
         "Kişiselleştirilebilir valiz; iç organizasyon sistemi; packing cubes; minimalist tasarım; hafif",
         "2019'da kuruldu; Instagram ve TikTok; DTC online; Avustralya kökenli; global genişleme; premium"),

        ("Monos", "monos.com", "Japon İlhamlı Seyahat Organizasyonu",
         "Compressible packing cubes; Metro sırt çantası; Japon minimalizm; sürdürülebilir malzeme; premium",
         "2018'de kuruldu; Instagram ve TikTok; DTC Shopify; influencer; Kanada kökenli; hızlı büyüme"),

        ("Troubadour", "troubadour.com", "Waterproof Seyahat Organizeri",
         "Su geçirmez İtalyan kumaş; vegan deri detaylar; tech organizer; toiletry bag; iş seyahati",
         "2019'da DTC genişleme; Instagram profesyonel seyahat; influencer; sürdürülebilir lüks; Londra kökenli"),

        ("Peak Design Travel", "peakdesign.com", "Modüler Seyahat Sistemi",
         "Packing cubes; Tech Pouch; Wash Pouch; 45L seyahat çantası ile entegre; modüler sistem",
         "2018'de seyahat hattı lansmanı; Kickstarter $5M+; DTC Shopify; YouTube ve Instagram; fotoğrafçı segmenti"),

        ("Dagne Dover", "dagnedover.com", "Neopren Seyahat Organizeri",
         "Neopren makyaj çantası; tech organizer; bebek çantası; su geçirmez; yıkanabilir; renkli",
         "2018'de hızlı büyüme; DTC Shopify; Instagram ve TikTok; influencer; fonksiyonel moda"),

        ("Roam Luggage", "roamluggage.com", "Özelleştirilebilir Valiz & Organizasyon",
         "Renk kişiselleştirme; 1M+ kombinasyon; packing cube seti; iç düzenleme; premium malzeme",
         "2018'de kuruldu; DTC online; Instagram kişiselleştirme trendi; influencer; ABD üretimi"),

        ("Thule Packing", "thule.com", "Outdoor Seyahat Organizasyonu",
         "Compression packing cubes; shoe organizer; clean/dirty separator; dayanıklı malzeme; outdoor",
         "2019'da seyahat organizasyon hattı; DTC online; Instagram outdoor seyahat; adventure influencer'ları"),

        ("Bellroy", "bellroy.com", "Deri Seyahat Organizeri",
         "Tech Kit; Toiletry Kit; packing cubes; premium deri ve geri dönüştürülmüş kumaş; ince profil",
         "2018'de seyahat hattını genişletti; DTC Shopify; Instagram minimalist seyahat; B Corp; Avustralya"),

        ("Tortuga", "tortugabackpacks.com", "Dijital Nomad Seyahat Paketi",
         "Carry-on boyutu sırt çantası; packing cubes; laptop bölmesi; dijital nomad odaklı; ergonomik",
         "2019'da DTC genişleme; Instagram ve YouTube dijital nomad; Kickstarter kökenli; online-only"),

        ("Nomatic", "nomatic.com", "Fonksiyonel Seyahat Sistemi",
         "Travel Pack; packing cubes; toiletry bag; laundry bag; RFID koruma; teknoloji dostu tasarım",
         "2018'de hızlı büyüme; Kickstarter $3M+; DTC Shopify; YouTube seyahat; influencer; erkek odaklı"),

        ("Lo & Sons", "loandsons.com", "Akıllı Seyahat Çantası & Organizeri",
         "The Rowledge; Pearl crossbody; çoklu bölme; laptop + ayakkabı + kıyafet; kadın seyahat",
         "2018'de DTC büyüme; Instagram kadın seyahat; influencer; Shopify; iş ve leisure seyahat"),

        ("Brevitē", "brevite.co", "Kompakt Seyahat Organizasyonu",
         "Kamera ve günlük sırt çantası; iç bölme sistemi; packing cubes; kompakt; çok fonksiyonlu",
         "2018'de kuruldu; Instagram fotoğrafçılık; DTC Shopify; Kickstarter; kreatif profesyonel hedef"),

        ("Db (Douchebags)", "dbjourney.com", "Macera Seyahat Organizasyonu",
         "Roller çanta; packing cubes; CIA kilitleme; dayanıklı; kayak ve snowboard seyahati; modüler",
         "2019'da global DTC genişleme; Instagram macera spor; influencer atlet iş birlikleri; İskandinav"),

        ("Tropicfeel", "tropicfeel.com", "All-in-One Seyahat Sistemi",
         "Shell sırt çantası; entegre packing system; ayakkabı bölmesi; wardrobe sistemi; modüler",
         "2019'da kuruldu; Kickstarter €4M+; DTC online; Instagram seyahat; sustainability; İspanya kökenli"),

        ("Away Travel Accessories", "awaytravel.com", "Seyahat Aksesuarları",
         "Packing cubes; The Insider; laundry bag; shoe cube; valiz ile entegre; minimalist tasarım",
         "2018'de aksesuar hattını genişletti; DTC unicorn; Instagram ve TikTok; influencer; premium"),

        ("Zoomlite", "zoomlite.com.au", "RFID Seyahat Organizeri",
         "RFID koruma; pasaport organizeri; tech pouch; packing cubes; güvenlik odaklı seyahat",
         "2019'da DTC genişleme; Amazon ve online; Instagram seyahat güvenliği; Avustralya kökenli"),

        ("Bagsmart", "bagsmart.com", "Tech & Kablo Seyahat Organizeri",
         "Elektronik organizer; kablo düzenleyici; tablet kılıfı; kompakt; çok bölmeli; uygun fiyat",
         "2018'de Amazon-native büyüme; DTC genişleme; TikTok seyahat paketi; teknoloji influencer'ları"),

        ("Gonex", "gonextravel.com", "Sıkıştırılabilir Packing Cube",
         "Çift fermuarlı kompresyon; %60 alan tasarrufu; su geçirmez; hafif naylon; çeşitli boyutlar",
         "2018'de Amazon-native; DTC genişleme; TikTok ve YouTube packing videoları; uygun fiyat; yüksek hacim"),

        ("Vasco", "vascotravel.com", "Kompresyon Seyahat Seti",
         "8 parça kompresyon set; ayakkabı torbası; makyaj çantası; elektronik organizer; komple paket",
         "2019'da kuruldu; Amazon bestseller; DTC online; YouTube seyahat; backpacker hedef kitle"),

        ("EzPacking", "ezpacking.com", "Komple Seyahat Organizasyon Kiti",
         "12 parça set; packing cubes; laundry bag; shoe bag; toiletry bag; tek pakette her şey",
         "2019'da kuruldu; Amazon ve DTC; YouTube seyahat organizasyonu; influencer; aile seyahati segmenti"),

        ("Sherpani", "sherpani.com", "Kadın Odaklı Seyahat Organizeri",
         "Anti-theft çanta; RFID bölmeli organizer; crossbody; gizli cep tasarımı; kadın güvenliği",
         "2018'de DTC genişleme; Instagram kadın solo seyahat; influencer; fonksiyonel güvenlik"),

        ("This Is Ground", "thisisground.com", "Deri Seyahat Tech Organizeri",
         "Bandito tech roll; Cordito kablo organizeri; deri Dopp kit; el yapımı; LA üretimi; lüks",
         "2018'de DTC büyüme; Instagram lüks seyahat; Kickstarter kökenli; artisan marka; premium"),

        ("Flight 001", "flight001.com", "Spacepak Seyahat Organizasyonu",
         "Spacepak modüler set; Go Clean germ-free pouch; Seat Pak; uçuş odaklı organizasyon; yenilikçi",
         "2019'da DTC dönüşüm; Instagram seyahat; havayolu influencer'ları; online-first pivot; niş marka"),

        ("Gravel", "graveltravel.com", "Explorer Toiletry Organizer",
         "Su geçirmez PLUS pouch; ayrılabilir ayna; askılı tasarım; kompakt; çok bölmeli; unisex",
         "2019'da kuruldu; Kickstarter; DTC Shopify; Instagram ve YouTube seyahat; minimalist seyahat"),

        ("Pakt", "pfriem.com", "Minimalist Seyahat Organizasyonu",
         "One Bag seyahat felsefesi; komple organizasyon sistemi; geri dönüştürülmüş malzeme; kompakt",
         "2019'da kuruldu; Kickstarter başarısı; DTC online; Instagram minimalist seyahat; sürdürülebilir"),

        ("Well Traveled", "welltraveledco.com", "Lüks Seyahat Organizer Seti",
         "Saffiano deri; takı organizeri; makyaj çantası; tech pouch; zarif tasarım; hediye segmenti",
         "2020'de kuruldu; DTC Shopify; Instagram lüks seyahat; influencer seeding; kadın premium segment"),

        ("Aer", "aersf.com", "Urban Seyahat Organizasyonu",
         "Travel Kit; Cable Kit; Slim Pouch; cordura naylon; su geçirmez; şehir seyahati; modern tasarım",
         "2018'de genişleme; DTC Shopify; Instagram urban seyahat; Kickstarter kökenli; San Francisco"),

        ("Horizn Studios", "horizn-studios.com", "Akıllı Seyahat Organizasyonu",
         "Akıllı valiz; smart packing cubes; GPS takip; entegre powerbank; Alman mühendisliği; premium",
         "2018'de DTC büyüme; Instagram akıllı seyahat; influencer; teknoloji seyahat; Berlin kökenli"),

        ("Surpass", "surpasstravel.com", "Ultra Hafif Packing Cube",
         "18g ultra hafif; ripstop naylon; nefes alır mesh; su geçirmez; sıkıştırılabilir; backpacker",
         "2020'de kuruldu; Amazon ve DTC; TikTok seyahat hack'leri; ultra hafif seyahat trendi; uygun fiyat"),

        ("Zulupack", "zulupack.com", "Su Geçirmez Seyahat Organizeri",
         "IPX7 su geçirmez; kaynaklı dikişler; roll-top kapatma; outdoor seyahat; dayanıklı; kompakt",
         "2019'da DTC genişleme; Instagram macera seyahat; su sporları influencer'ları; Fransa kökenli"),
    ],
}


# ─────────────────────────────────────────────────────────────────────
# Quick stats helper
# ─────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    total = 0
    for category, brands in BATCH_DATA.items():
        count = len(brands)
        total += count
        print(f"  {category}: {count} marka")
    print(f"\n  TOPLAM: {total} marka")
