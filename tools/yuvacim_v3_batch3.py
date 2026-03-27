"""
yuvacim_v3_batch3.py
Post-2018 ecommerce-native DTC brands — ORGANIZER / HOME SCENT / RUG / TEXTILE niche
6 categories, 200+ brands total
All descriptions in Turkish (brand names & URLs in English)
"""

BATCH_DATA = {

    # ─────────────────────────────────────────────────────────────────
    # 1. MAKYAJ & KOZMETİK ORGANİZERİ  (Makeup & Vanity Organizers)
    # ─────────────────────────────────────────────────────────────────
    "Makyaj & Kozmetik Organizeri": [
        ("Luxe Makeup Organizer", "luxemakeuporganizer.com", "Akrilik Makyaj Organizeri",
         "Şeffaf akrilik tasarım; döner tabanlı modeller; modüler bölmeler; Instagram'da viral; çekmeceli ve açık rafli versiyonlar",
         "2019'da kuruldu; sosyal medya reklamlarıyla büyüdü; Shopify DTC; influencer unboxing videoları; Amazon'da da satış"),

        ("Sorbus", "sorbushome.com", "Çok Katmanlı Organizatör",
         "Bambu ve akrilik serileri; döner tezgah üstü organizer; makyaj çekmecesi sistemleri; uygun fiyatlı segment",
         "2018'de e-ticaret odaklı kuruldu; Amazon-native marka; TikTok organizasyon videoları; DTC genişleme; hızlı SKU çeşitlendirme"),

        ("Pocpik", "pocpik.com", "Taşınabilir Makyaj Organizeri",
         "LED aynalı seyahat makyaj çantası; su geçirmez bölmeler; USB şarjlı aydınlatma; kompakt tasarım",
         "2019'da kuruldu; Amazon ve Shopify DTC; TikTok reklamlarıyla viral; seyahat influencer'larıyla iş birliği; online-only"),

        ("Relavel", "relavel.com", "Seyahat Makyaj Çantası",
         "Profesyonel makyaj çantası; ayarlanabilir bölücüler; su geçirmez iç kaplama; omuz askılı ve el çantası versiyonları",
         "2018'de kuruldu; Amazon-first strateji; sosyal medya pazarlama; DTC Shopify mağazası; makyaj sanatçıları tarafından önerilen"),

        ("Glamlily", "glamlily.com", "Dekoratif Kozmetik Organizeri",
         "Mermer desenli akrilik organizerler; altın detaylı tasarım; Instagram estetik; modüler sistem",
         "2020'de kuruldu; Instagram ve Pinterest odaklı DTC; influencer pazarlama; Shopify native; estetik ambalaj"),

        ("Begin Magic", "beginmagic.com", "Döner Makyaj Organizeri",
         "360° döner taban; ayarlanabilir raflar; büyük kapasiteli tasarım; şeffaf akrilik",
         "2019'da kuruldu; Amazon-native DTC; TikTok organizasyon trendleriyle büyüdü; uygun fiyatlı lüks segment"),

        ("Elecdon", "elecdon.com", "LED Aynalı Makyaj Kutusu",
         "Dahili LED ayna; taşınabilir kozmetik saklama; şarj edilebilir; profesyonel kalite",
         "2020'de kuruldu; e-ticaret odaklı; TikTok viral ürün; Amazon ve DTC; teknoloji x güzellik kavramı"),

        ("STORi", "storibox.com", "Şeffaf Akrilik Organizer",
         "Stackable modüler kutular; çekmeceli sistemler; tezgah üstü düzenleyiciler; minimalist şeffaf tasarım",
         "2018'de kuruldu; Amazon ve DTC Shopify; organizasyon influencer'larıyla iş birliği; Marie Kondo trendi; online-only"),

        ("Vtopmart", "vtopmart.com", "Mutfak & Makyaj Organizeri",
         "Çok amaçlı şeffaf saklama kutuları; etiketli kavanoz setleri; banyo ve makyaj organizerleri; BPA-free",
         "2018'de kuruldu; Amazon-native; TikTok pantry/organization trendleriyle viral; DTC genişleme; hızlı büyüme"),

        ("Cosmocube", "cosmocube.co", "Lüks Akrilik Makyaj Dolabı",
         "Premium akrilik; kilit mekanizmalı; özel bölmeler; influencer tasarımlı koleksiyonlar",
         "2019'da kuruldu; Instagram influencer'larıyla iş birliği; DTC Shopify; lüks segment; özel sipariş seçenekleri"),

        ("Fabuday", "fabuday.com", "Kozmetik Tezgah Organizeri",
         "Altın metal çerçeve; cam raflı; modern minimalist tasarım; parfüm ve makyaj için",
         "2020'de kuruldu; Instagram estetik odaklı; DTC Shopify; Pinterest pazarlama; ev dekor x kozmetik kavramı"),

        ("Moosy Life", "moosylife.com", "Bambu Kozmetik Organizeri",
         "Doğal bambu malzeme; sürdürülebilir; çok bölmeli tasarım; banyo tezgahı için ideal",
         "2019'da kuruldu; çevre dostu mesajla DTC; Amazon ve Shopify; organik yaşam influencer'larıyla iş birliği"),

        ("Bino", "binoproducts.com", "Plastik Saklama Organizeri",
         "BPA-free plastik; stackable tasarım; banyo ve makyaj için; şeffaf ve renkli seçenekler",
         "2018'de kuruldu; Amazon-native marka; TikTok ve YouTube organizasyon videoları; uygun fiyatlı segment; DTC genişleme"),

        ("Acrylic Solutions", "acrylicsolutions.co", "Özel Tasarım Akrilik Organizer",
         "Kişiye özel boyut; modüler akrilik sistemler; profesyonel makyaj istasyonu; şeffaf premium kalite",
         "2019'da kuruldu; Instagram DTC; özel sipariş modeli; makyaj sanatçıları hedef kitle; Shopify native"),

        ("Makeup Miner", "makeupminer.com", "Mücevherli Kozmetik Kutusu",
         "Kristal detaylı organizerler; kadife iç kaplama; hediye paketi tasarım; vanity masası için",
         "2020'de kuruldu; TikTok viral ürün; DTC Shopify; Instagram influencer pazarlama; premium hediye segmenti"),

        ("Deco Haus", "decohaus.co", "Estetik Banyo Organizeri",
         "Terazzo ve mermer desenli; Instagram estetik; banyo tezgah düzenleyici; kozmetik ve cilt bakımı için",
         "2020'de kuruldu; Instagram-first DTC; Pinterest odaklı büyüme; Shopify native; ev dekor influencer'larıyla iş birliği"),

        ("GlamBox", "glamboxorganizer.com", "Döner Kozmetik Dolabı",
         "360° döner sistem; 8 kat kapasiteli; ayarlanabilir yükseklik; akrilik ve bambu versiyonları",
         "2019'da kuruldu; TikTok ve Instagram reklamlarıyla büyüdü; Amazon ve DTC; uygun fiyatlı lüks; online-only"),

        ("Miuopur", "miuopur.com", "Ruj & Fırça Standı",
         "Özel ruj bölmeleri; makyaj fırçası tutucusu; kompakt tezgah üstü tasarım; silikon ayak",
         "2019'da kuruldu; Amazon-native; sosyal medya reklamları; DTC genişleme; niş kozmetik saklama"),

        ("HBlife", "hblifetech.com", "Makyaj Fırçası Organizeri",
         "Şeffaf akrilik fırça tutucusu; inci ve boncuk dolgusu; dekoratif saklama; cam kapaklı modeller",
         "2018'de kuruldu; Amazon-native DTC; Pinterest organizasyon trendleri; güzellik blog'larıyla iş birliği; küresel e-ticaret"),

        ("Kingtop", "kingtoplife.com", "Masaüstü Kozmetik İstasyonu",
         "Çekmeceli masaüstü organizer; akrilik + bambu hibrit; USB şarj portlu modeller; modern tasarım",
         "2019'da kuruldu; Amazon ve Shopify DTC; teknoloji x organizasyon kavramı; YouTube review'larıyla büyüdü"),

        ("Love-KANKEI", "love-kankei.com", "Rustik Ahşap Organizer",
         "Doğal ahşap ve metal kombinasyonu; endüstriyel rustik stil; duvar montajlı; banyo ve makyaj için",
         "2018'de kuruldu; Amazon-native; Instagram rustik dekor trendleri; DTC Shopify; ev dekor crossover"),

        ("Cq Acrylic", "cqacrylic.com", "Büyük Kapasiteli Akrilik Organizer",
         "Profesyonel boyut; 6-9 çekmeceli; döner ve sabit modeller; şeffaf premium akrilik",
         "2018'de kuruldu; Amazon bestseller; YouTube makyaj koleksiyonu videoları; DTC genişleme; toptan da satış"),

        ("ILive Smartly", "ilivesmartly.com", "Akıllı Kozmetik Buzdolabı",
         "Mini buzdolabı kozmetik saklama; sıcaklık kontrollü; cilt bakımı ürünleri için ideal; sessiz çalışma",
         "2020'de kuruldu; TikTok skincare fridge trendleriyle viral; DTC Shopify; influencer pazarlama; niş segment"),

        ("Etoile Collective", "etoilecollective.com", "Lüks Makyaj Çantası & Organizer",
         "Vegan deri; altın fermuar detayları; modüler bölmeler; şık seyahat çantaları",
         "2018'de kuruldu; Instagram influencer DTC; Shopify native; premium segment; makyaj sanatçıları hedef kitle"),

        ("Pimoys", "pimoys.com", "Seyahat Kozmetik Organizeri",
         "Su geçirmez neopren; asılabilir tasarım; çoklu cepli; kompakt katlanır model",
         "2020'de kuruldu; Amazon-native; TikTok seyahat hack'leriyle viral; DTC; uygun fiyatlı segment"),

        ("Sunficon", "sunficon.com", "Silikon Makyaj Tutucusu",
         "Esnek silikon malzeme; kolay temizlenen; renkli tasarımlar; fırça ve ruj tutucusu",
         "2019'da kuruldu; Amazon DTC; TikTok güzellik trendleriyle büyüdü; uygun fiyatlı; online-only"),

        ("Chic Moda", "chicmoda.co", "Kadife Mücevher & Makyaj Kutusu",
         "Kadife kaplama; çok katmanlı; ayna dahil; hediye için ideal; zarif tasarım",
         "2020'de kuruldu; Instagram DTC; influencer unboxing; Shopify native; hediye segmenti"),

        ("Alima Beauty Tools", "alimabeautytools.com", "Manyetik Kapaklı Organizer",
         "Manyetik kapak sistemi; toz geçirmez; modüler iç bölmeler; premium akrilik",
         "2019'da kuruldu; Shopify DTC; Instagram güzellik influencer'larıyla iş birliği; Amazon genişleme; niş segment"),

        ("BoxLegend", "boxlegend.com", "Katlanır Saklama & Organizer",
         "Katlanabilir kumaş kutular; etiket pencereli; çekmece bölücüleri; çok amaçlı",
         "2018'de kuruldu; Amazon-native; TikTok organizasyon challenge'larıyla viral; DTC Shopify; hızlı SKU artışı"),

        ("Luvo Store", "luvostore.com", "Minimalist Kozmetik Rafı",
         "Metal tel raf tasarım; duvar montajlı ve tezgah üstü; Scandinavian estetik; mat siyah ve altın renkleri",
         "2020'de kuruldu; Instagram estetik DTC; Pinterest pazarlama; Shopify native; ev dekor odaklı"),

        ("Twillory Home", "twilloryhome.com", "Tezgah Üstü Cam Organizer",
         "Pirinç çerçeveli cam kutu; vintage estetik; parfüm ve kozmetik vitrini; Instagram'da trend",
         "2021'de kuruldu; Instagram-first DTC; Pinterest viral; Shopify native; dekoratif saklama niş'i"),

        ("Neatly Made", "neatlymade.co", "Modüler Çekmece Düzenleyici",
         "Özel ölçü çekmece bölücüleri; bambu ve akrilik; makyaj çekmecesi için ideal; DIY modüler sistem",
         "2020'de kuruldu; TikTok çekmece düzenleme videoları viral; DTC Shopify; Amazon genişleme; organizasyon niş'i"),

        ("Vanity Collections", "vanitycollections.com.au", "Lüks Vanity Organizeri",
         "Akrilik vanity ünitesi; entegre aydınlatma; Avustralya tasarımı; profesyonel makyaj istasyonu; modüler",
         "2019'da kuruldu; Instagram makyaj odası; DTC Shopify; Avustralya'dan küresel; influencer pazarlama; premium"),
    ],

    # ─────────────────────────────────────────────────────────────────
    # 2. EV KOKUSU & MUM  (Home Scent & Candles)
    # ─────────────────────────────────────────────────────────────────
    "Ev Kokusu & Mum": [
        ("Boy Smells", "boysmells.com", "Cinsiyet Nötr Mum",
         "Hindistan cevizi ve balmumu karışımı; cinsiyet normlarını yıkan kokular; ikonik pembe ambalaj; parfüm serisi de var",
         "2018'de Los Angeles'ta kuruldu; LGBTQ+ topluluk odaklı; Instagram ve influencer DTC; Shopify native; kült marka statüsü"),

        ("Otherland", "otherland.com", "Sanat Tasarımlı Lüks Mum",
         "Her mum bir sanat eseri; hindistan cevizi ve soya karışımı; tekrar doldurulabilir cam; mevsimsel koleksiyonlar",
         "2019'da New York'ta kuruldu; Instagram estetik DTC; sanatçı iş birlikleri; Shopify native; premium segment"),

        ("Homesick", "homesick.com", "Nostalji Temalı Mum",
         "Şehir ve ülke ilhamlı kokular; anı ve nostalji konsepti; hediye odaklı; kişiselleştirilebilir etiketler",
         "2018'de DTC olarak büyüdü; sosyal medya pazarlama; Shopify native; hediye ekonomisi; viral konsept"),

        ("Apotheke", "apothekeco.com", "Eczane İlhamlı Ev Kokusu",
         "Farmasi estetiği; soya bazlı; amber cam şişeler; difüzör ve mum serisi; minimalist ambalaj",
         "2018'de Brooklyn'de kuruldu; DTC Shopify; influencer ve otel iş birlikleri; Instagram estetik; premium segment"),

        ("Brooklyn Candle Studio", "brooklyncandlestudio.com", "Artisan Soya Mum",
         "%100 soya mumu; el yapımı; minimalist etiket tasarımı; sürdürülebilir ambalaj; mevsimsel kokular",
         "2018'de kuruldu; Etsy'den Shopify DTC'ye geçiş; Instagram ve Pinterest; çevre dostu mesaj; küçük seri üretim"),

        ("Keap Candles", "keapbk.com", "Sürdürülebilir Mum",
         "Hindistan cevizi mumu; geri dönüştürülebilir ambalaj; abonelik modeli; doğal fitil; temiz yanma",
         "2018'de kuruldu; B Corp sertifikalı; DTC abonelik; çevre dostu mesaj; Instagram ve influencer pazarlama"),

        ("Vitruvi", "vitruvi.com", "Esansiyel Yağ Difüzörü",
         "Seramik taş difüzörler; minimalist tasarım; doğal esansiyel yağlar; oda spreyleri; wellness odaklı",
         "2018'de DTC olarak genişledi; Instagram estetik; influencer pazarlama; Shopify native; premium wellness segment"),

        ("P.F. Candle Co.", "pfcandleco.com", "Amber Kavanoz Mum",
         "İkonik amber cam; soya bazlı; el yapımı; incense ve oda spreyleri; California ilhamlı kokular",
         "2018 sonrası DTC patlaması; Instagram kült marka; Shopify ve toptan; sürdürülebilir üretim; geniş perakende dağıtım"),

        ("Voluspa", "voluspa.com", "Dekoratif Teneke Mum",
         "Hindistan cevizi balmumu; dekoratif teneke ve cam kaplar; Japonica serisi; zengin koku paleti",
         "2018 sonrası DTC büyümesi; Instagram estetik; influencer pazarlama; Shopify ve Amazon; premium hediye segmenti"),

        ("Candle Collective", "candlecollective.co", "Abonelik Mum Kutusu",
         "Aylık mum aboneliği; indie markaları keşfet; küçük seri üretim; sürpriz koku konsepti",
         "2019'da kuruldu; DTC abonelik modeli; Instagram ve TikTok; Shopify native; indie mum topluluğu"),

        ("Mala the Brand", "malathebrand.com", "Vegan El Yapımı Mum",
         "Hindistan cevizi soya karışımı; %100 vegan; pamuk fitil; minimalist tasarım; Kanada kökenli",
         "2018'de kuruldu; Instagram DTC; çevre dostu mesaj; Shopify native; influencer pazarlama; küçük seri"),

        ("Wax & Wane", "waxandwane.co", "Kristal İçerikli Mum",
         "Doğal kristaller içeren mumlar; soya bazlı; niyet belirleme ritüeli; wellness konsepti",
         "2019'da kuruldu; TikTok wellness trendleriyle viral; DTC Shopify; Instagram spirituel topluluk; niş segment"),

        ("Yield Design Co.", "yielddesign.co", "Minimalist Cam Mum",
         "Borosilikat cam kaplar; tekrar doldurulabilir; double-wall tasarım; minimalist etiket",
         "2018'de kuruldu; DTC Shopify; Instagram minimalist estetik; sürdürülebilir tasarım; premium segment"),

        ("Anecdote Candles", "anecdotecandles.com", "Espri Etiketli Mum",
         "Komik ve ilişkilendirilebilir etiketler; soya bazlı; hediye odaklı; pop kültür referansları",
         "2019'da kuruldu; Instagram ve TikTok viral; DTC Shopify; hediye ekonomisi; Gen Z hedef kitle"),

        ("Nette", "nette.nyc", "Lüks Sanat Mum",
         "Balmumu ve soya karışımı; sanatçı iş birlikleri; parfüm kalitesinde kokular; lüks ambalaj",
         "2019'da New York'ta kuruldu; Instagram estetik DTC; sanat dünyası crossover; Shopify native; ultra-premium segment"),

        ("Overose", "overose.com", "Pembe Balmumu Mum",
         "İmza pembe balmumu; güçlü koku yayılımı; ikonik pembe cam; parfüm ilhamlı; Parisian estetik",
         "2018'de kuruldu; Instagram estetik; influencer DTC; küresel gönderim; premium niş segment"),

        ("Maison Louis Marie", "maisonlouismarie.com", "Botanik Parfüm & Mum",
         "Botanik bahçe ilhamlı kokular; No.04 en çok satan; doğal bileşenler; difüzör ve parfüm serisi",
         "2018 sonrası DTC büyümesi; TikTok No.04 viral; Instagram DTC; Shopify native; kült ürün statüsü"),

        ("Harlem Candle Co.", "harlemcandlecompany.com", "Lüks Harlem İlhamlı Mum",
         "Harlem Rönesansı ilhamlı; lüks cam kaplar; zengin koku profilleri; kültürel hikaye anlatımı",
         "2018'de DTC genişleme; Instagram ve influencer pazarlama; Shopify native; Siyah sahipli marka; premium segment"),

        ("Flamingo Estate", "flamingoestate.com", "Çiftlik İlhamlı Ev Kokusu",
         "Los Angeles çiftlik ilhamlı; doğal bileşenler; mum, sabun, yağ; organik ve sürdürülebilir",
         "2020'de kuruldu; Instagram estetik; celebrity influencer'lar; DTC Shopify; farm-to-home konsepti"),

        ("Loewe Perfumes Home", "loewe.com", "Botanik Seramik Mum",
         "Seramik sebze ve meyve şeklinde mumlar; sanat eseri ambalaj; botanik kokular; koleksiyon parçası",
         "2020'de ev kokusu DTC lansmanı; Instagram viral; TikTok estetik; lüks segment; koleksiyoner hedef kitle"),

        ("Earl of East", "earlofeast.com", "Londra İlhamlı Soya Mum",
         "%100 soya; pamuk fitil; stoneware kaplar; Londra'dan ilham alan kokular; geri dönüştürülebilir ambalaj",
         "2018'de DTC genişleme; Instagram estetik; Shopify native; sürdürülebilir mesaj; Birleşik Krallık ve küresel"),

        ("Lulu Candles", "lulucandles.com", "Kişiselleştirilmiş Etiketli Mum",
         "Kişiselleştirilmiş etiketler; soya bazlı; düğün ve etkinlik hediyeleri; çoklu koku seçenekleri",
         "2019'da kuruldu; TikTok kişiselleştirme trendleriyle büyüdü; Etsy ve DTC; düğün sektörü; online-only"),

        ("Cancelled Plans", "cancelledplans.com", "İçe Dönük Yaşam Mumu",
         "Ev kalma temalı komik etiketler; soya-hindistan cevizi karışımı; rahat akşam konsepti; hediye odaklı",
         "2019'da kuruldu; Instagram ve TikTok viral; DTC Shopify; pandemi döneminde patlama; Gen Z/Millennial hedef"),

        ("Backdrop", "backdropcandles.com", "Film & Dizi İlhamlı Mum",
         "Popüler film ve dizi kokularını yakalayan mumlar; nostalji konsepti; soya bazlı; koleksiyon serisi",
         "2020'de kuruldu; TikTok popüler kültür; Instagram DTC; Shopify native; niş hediye segmenti"),

        ("Forvr Mood", "forvrmood.com", "Jackie Aina Mum Markası",
         "Beauty influencer Jackie Aina'nın markası; zengin kokular; lüks cam kaplar; çeşitlilik odaklı",
         "2020'de kuruldu; YouTube/Instagram influencer DTC; Shopify native; influencer-founded brand; hızlı büyüme"),

        ("Bijou Candles", "bijoucandles.co", "Mücevher Gizli Mum",
         "Her mumda gizli mücevher; sürpriz yüzük veya kolye; soya bazlı; hediye konsepti; heyecan faktörü",
         "2019'da kuruldu; TikTok unboxing videoları viral; DTC Shopify; hediye ekonomisi; eğlenceli konsept"),

        ("Spoken Flames", "spokenflames.com", "Motivasyon Sözlü Mum",
         "Pozitif afirmasyonlu etiketler; soya-hindistan cevizi karışımı; wellness x koku konsepti; Instagram estetik",
         "2019'da kuruldu; Instagram ve TikTok; Siyah sahipli marka; DTC Shopify; wellness trend; hediye segmenti"),

        ("Skylar Body", "skylar.com", "Temiz Parfüm & Ev Kokusu",
         "Hipoalerjenik; vegan; temiz formül; ev difüzörleri ve mumları; hassas cilt için güvenli",
         "2018'de kuruldu; clean beauty trendi; DTC Shopify; Instagram ve influencer; hassas cilt topluluğu"),

        ("Replica by Maison Margiela Home", "maisonmargiela.com", "Anı Bazlı Ev Kokusu",
         "Gerçek yaşam anlarını koku olarak yakalayan mumlar; By the Fireplace, Beach Walk gibi ikonik kokular",
         "2019'da ev kokusu DTC genişleme; TikTok viral; Instagram estetik; parfüm crossover; premium segment"),

        ("Cirillo Home", "cirillohome.com", "İtalyan İlhamlı Soya Mum",
         "İtalyan yaşam tarzı kokular; limon bahçesi, espresso gibi; soya bazlı; el yapımı; seramik kaplar",
         "2020'de kuruldu; Instagram İtalyan estetik; DTC Shopify; influencer pazarlama; hediye segmenti"),

        ("WXYCANDLES", "wxy.com", "Minimalist Tasarım Mumu",
         "Dikdörtgen beton kaplar; uzun yanan formül; doğal soya; İngiliz tasarım; urban kokular",
         "2018'de DTC genişleme; Instagram minimalist estetik; Shopify native; tasarım odaklı; premium segment"),

        ("Lite + Cycle", "liteandcycle.com", "Beeswax Doğal Mum",
         "%100 balmumu; pamuk fitil; toksik olmayan; doğal koku; temiz yanma; çevre dostu ambalaj",
         "2019'da kuruldu; clean living trendi; DTC Shopify; Instagram ve Pinterest; sürdürülebilir mesaj"),

        ("Siblings", "siblings.co", "Tekrar Doldurulabilir Mum",
         "Seramik kap + yedek mum sistemi; sürdürülebilir; minimalist Scandinavian tasarım; mevsimsel kokular",
         "2020'de kuruldu; çevre dostu DTC; Instagram estetik; refill modeli; Shopify native; premium segment"),

        ("Aerangis", "aerangis.com", "Tropikal Botanik Mum",
         "Egzotik çiçek kokular; el dökümü; doğal balmumu; tropikal estetik; lüks ambalaj",
         "2019'da kuruldu; Instagram botanik estetik; DTC Shopify; niş koku topluluğu; premium segment"),

        ("Moodcast Fragrance", "moodcastfragrance.com", "Ruh Hali Bazlı Mum",
         "Koku x ruh hali eşleştirmesi; soya karışımı; renk kodlu sistem; wellness konsepti; difüzör serileri",
         "2019'da kuruldu; Instagram wellness DTC; Shopify native; hedef kitle: ruh hali bilinçli tüketici; niş segment"),
    ],

    # ─────────────────────────────────────────────────────────────────
    # 3. HALI & KİLİM  (Rugs & Carpets)
    # ─────────────────────────────────────────────────────────────────
    "Halı & Kilim": [
        ("Ruggable", "ruggable.com", "Yıkanabilir Halı",
         "Makinede yıkanabilir iki parçalı sistem; ped + kapak; su ve leke geçirmez; çoklu desen seçenekleri",
         "2018 sonrası DTC patlaması; Instagram ve Facebook Ads; Shopify native; yıkanabilir halı kategorisini yarattı; $100M+ gelir"),

        ("Tumble", "tumblelivingco.com", "Yıkanabilir Çocuk Dostu Halı",
         "Makinede yıkanabilir; çift yüzlü tasarım; çocuk ve evcil hayvan dostu; leke tutmaz",
         "2020'de kuruldu; Shark Tank yatırımı; Instagram DTC; aile odaklı mesaj; Shopify native; hızlı büyüme"),

        ("Rugs USA", "rugsusa.com", "Online Halı Pazaryeri",
         "Binlerce desen; uygun fiyatlı; trendlere hızlı adaptasyon; her boyut ve stil; hızlı kargo",
         "2018 sonrası DTC dönüşümü; Instagram ve TikTok; influencer ev dekor; sosyal medya pazarlama; e-ticaret odaklı"),

        ("Boutique Rugs", "boutiquerugs.com", "Butik Tasarım Halı",
         "El yapımı görünümlü; vintage ve modern tasarımlar; uygun fiyatlı lüks; geniş renk paleti",
         "2018'de DTC lansmanı; Instagram influencer ev dekor; Shopify native; blog ve Pinterest pazarlama; hızlı SKU artışı"),

        ("Loloi Rugs", "loloirugs.com", "Tasarımcı İş Birlikli Halı",
         "Chris Loves Julia, Amber Lewis koleksiyonları; el yapımı ve makine üretimi; geniş ürün yelpazesi",
         "2018 sonrası DTC ve influencer patlaması; Instagram influencer koleksiyonları; Shopify ve Amazon; premium segment"),

        ("Miss Amara", "missamara.com", "Küratörlü Online Halı",
         "AI destekli halı eşleştirme; AR oda önizleme; küratörlü koleksiyon; Avustralya kökenli",
         "2018'de DTC genişleme; teknoloji x ev dekor; Instagram ve Facebook; AR deneyimi; online-only"),

        ("Revival Rugs", "revivalrugs.com", "Vintage Türk Halısı",
         "Otantik vintage Türk halıları; her biri tek; el yapımı; doğal boyalar; antik halı restorasyonu",
         "2019'da kuruldu; Instagram vintage estetik; DTC Shopify; Türk halısı kültürü; Austin, Texas merkezli"),

        ("Loomah", "loomah.com", "Lüks El Yapımı Halı",
         "Premium el dokuma; özel sipariş; doğal lifler; minimalist tasarımlar; tasarımcı iş birlikleri",
         "2019'da kuruldu; Instagram iç mimari; DTC Shopify; lüks segment; tasarım stüdyoları hedef kitle"),

        ("Jaipur Living", "jaipurliving.com", "Sürdürülebilir El Dokuma Halı",
         "Adil ticaret; geri dönüştürülmüş malzeme; el dokuma; Hindistan zanaatkarları; GoodWeave sertifikalı",
         "2018 sonrası DTC büyümesi; sürdürülebilir mesaj; Instagram ve Pinterest; Shopify ve toptan; B Corp hedefli"),

        ("Parachute Rugs", "parachutehome.com", "Premium Ev Tekstili Halı",
         "El tutufted; doğal yün; minimalist tasarım; nötr tonlar; yatak odası ve oturma odası boyutları",
         "2019'da halı kategorisine giriş; DTC Shopify; Instagram estetik; premium segment; ev tekstili crossover"),

        ("Rug Culture", "rugculture.com.au", "Avustralya DTC Halı",
         "Modern ve vintage tasarımlar; geniş boyut seçenekleri; Avustralya'dan küresel gönderim; uygun fiyat",
         "2018'de DTC genişleme; Instagram ve Facebook Ads; Shopify native; influencer ev dekor; online-only"),

        ("Cali Fabrics Rugs", "califabricsrugs.com", "Kaliforniya Yaşam Tarzı Halı",
         "Boho ve coastal tasarımlar; doğal jüt ve pamuk; makine yıkanabilir serileri; California ilhamlı",
         "2019'da kuruldu; Instagram California estetik; DTC Shopify; influencer pazarlama; rahat yaşam konsepti"),

        ("House of Rugs", "houseofrugs.co", "Modern Geometrik Halı",
         "Geometrik ve soyut desenler; yüksek trafik alanları için dayanıklı; modern renk paletleri",
         "2020'de kuruldu; Instagram modern iç mekan; DTC Shopify; Pinterest pazarlama; online-only marka"),

        ("Burrow Rugs", "burrow.com", "Modüler Ev Mobilya & Halı",
         "Modüler mobilya markasından halı serisi; leke tutmaz; kolay temizlenen; modern tasarım",
         "2020'de halı kategorisine giriş; DTC Shopify; Instagram ve Facebook; modüler ev kavramı; VC destekli"),

        ("Bien Rugs", "bienrugs.com", "Meksika El Yapımı Halı",
         "Meksika zanaatkarları; doğal lifler; geleneksel teknikler; modern yorumlar; adil ticaret",
         "2019'da kuruldu; Instagram el yapımı zanaat; DTC Shopify; kültürel hikaye; sürdürülebilir üretim"),

        ("The Citizenry Rugs", "the-citizenry.com", "Zanaatkar İş Birlikli Halı",
         "Dünya genelinde zanaatkar iş birlikleri; el yapımı; doğal malzeme; adil ücret; hikaye anlatımı",
         "2018 sonrası DTC büyümesi; Instagram zanaatkar hikaye; Shopify native; sosyal etki mesajı; premium segment"),

        ("Armadillo Rugs", "armadillo-co.com", "Sürdürülebilir El Yapımı Halı",
         "Adil ticaret; organik pamuk ve yün; el yapımı; doğal boyalar; çevre dostu ambalaj; Avustralya kökenli",
         "2018 sonrası DTC genişleme; Instagram sürdürülebilir dekor; Shopify native; B Corp sertifikalı; premium"),

        ("Aelfie", "aelfie.com", "Renkli Eklektik Halı",
         "Canlı renkler ve desenler; eklektik tasarım; Hindistan'da el yapımı; oyunbaz estetik",
         "2018'de DTC genişleme; Instagram renkli iç mekan; Shopify native; Brooklyn tasarım; niş segment"),

        ("Nani Marquina", "nanimarquina.com", "İspanyol Tasarım Halı",
         "İspanyol tasarım; el yapımı; sanatçı iş birlikleri; sürdürülebilir üretim; modern sanat estetiği",
         "2018 sonrası DTC genişleme; Instagram tasarım dünyası; Shopify; iç mimari odaklı; premium segment"),

        ("Angela Adams", "angelaadams.com", "Maine İlhamlı Tasarım Halı",
         "Maine doğasından ilham; el tuftlanmış yün; cesur geometrik desenler; sanatsal tasarım",
         "2018 sonrası DTC genişleme; Instagram sanatsal dekor; Shopify native; el yapımı ABD üretim; premium segment"),

        ("Dash & Albert", "dashandalbert.com", "Indoor-Outdoor Halı",
         "İç ve dış mekan halıları; yıkanabilir; dayanıklı; pamuk ve polipropilen; çeşitli desenler",
         "2018 sonrası DTC büyümesi; Instagram ev ve bahçe; Shopify native; Annie Selke markası; geniş dağıtım"),

        ("Lulu and Georgia", "luluandgeorgia.com", "Küratörlü Tasarım Halı",
         "Influencer koleksiyonları; vintage ve modern mix; premium tasarım; geniş boyut seçenekleri",
         "2018 sonrası DTC patlaması; Instagram influencer ev dekor; Shopify native; küratörlü koleksiyon; hızlı büyüme"),

        ("Southwestern Rugs Depot", "southwesternrugsdepot.com", "Güneybatı Motifli Halı",
         "Navajo ve güneybatı desenleri; el dokuma; doğal yün; geleneksel motifler; Amerikan kökenli",
         "2019'da kuruldu; Instagram western dekor trendleri; DTC Shopify; niş segment; boho-western crossover"),

        ("Kalim", "kalim.co", "Modern Kilim",
         "Geleneksel kilim tekniği; modern renk paletleri; düz dokuma; hafif ve esnek; çift taraflı kullanım",
         "2020'de kuruldu; Instagram modern kilim trendleri; DTC Shopify; Türk zanaatkarları; online-only"),

        ("RugPad USA", "rugpadusa.com", "Halı Altı Kaydırmaz Ped",
         "Kaydırmaz ped; doğal kauçuk; keçe dolgulu; özel boyut kesim; halı ömrünü uzatır",
         "2018'de DTC genişleme; Amazon ve Shopify; halı aksesuarı niş'i; online-only; pratik ürün"),

        ("Coral & Hive", "coralandhive.com", "Güney Afrika El Yapımı Halı",
         "Mohair ve yün; Güney Afrika zanaatkarları; doğal boyalar; modern tasarım; sürdürülebilir",
         "2018'de DTC genişleme; Instagram zanaatkar hikaye; Shopify native; adil ticaret; premium niş segment"),

        ("Lila Valadan", "lilavaladan.com", "Minimalist Yün Halı",
         "El dokuma yün; minimalist tek renk tasarımlar; Yeni Zelanda yünü; yumuşak doku; modern estetik",
         "2019'da kuruldu; Instagram minimalist dekor; DTC Shopify; tasarımcı odaklı; lüks segment"),

        ("Hook & Loom", "hookandloom.com", "Eco-Cotton Halı",
         "Geri dönüştürülmüş pamuk; sıfır atık üretim; çift taraflı; el yapımı; ABD üretim",
         "2018'de DTC genişleme; sürdürülebilir mesaj; Shopify native; çevre dostu tüketici hedef; uygun fiyatlı"),

        ("Knotty Rugs", "knottyrugs.com", "El Düğümlü Vintage Halı",
         "El düğümlü; vintage İran ve Türk halıları; restorasyon; doğal boyalar; koleksiyon parçaları",
         "2019'da kuruldu; Instagram vintage halı; DTC Shopify; tek ve benzersiz parçalar; koleksiyoner hedef kitle"),

        ("Kush Handmade Rugs", "kushrugs.com", "Fas El Yapımı Halı",
         "Beni Ourain ve Fas halıları; doğal yün; el yapımı; berberi motifler; boho estetik",
         "2019'da kuruldu; Instagram boho dekor; DTC Shopify; Fas zanaatkarları; influencer ev dekor"),

        ("Companion Rugs", "companionrugs.com", "Sanatçı Tasarımlı Modern Halı",
         "Sanatçı iş birlikleri; cesur desenler; Yeni Zelanda yünü; el tuftlanmış; sınırlı üretim",
         "2020'de kuruldu; Instagram sanat x dekor; DTC Shopify; niş sanat topluluğu; premium segment"),

        ("Beni Rugs", "bfrugdecor.com", "Otantik Fas Berberi Halı",
         "Otantik Beni Ourain; Atlas Dağları zanaatkarları; doğal yün; her biri tek; geleneksel dokuma",
         "2019'da kuruldu; Instagram boho estetik; DTC Shopify; kültürel hikaye anlatımı; online-only"),

        ("Flor", "flor.com", "Modüler Karo Halı",
         "Birleştirilebilir halı karoları; kendi tasarımını yap; kolay değiştirilebilir; geri dönüştürülebilir",
         "2018 sonrası DTC genişleme; Instagram DIY dekor; Shopify native; modüler konsept; çevre dostu mesaj"),
    ],

    # ─────────────────────────────────────────────────────────────────
    # 4. PERDE & EV TEKSTİLİ  (Curtains & Home Textiles)
    # ─────────────────────────────────────────────────────────────────
    "Perde & Ev Tekstili": [
        ("Wovn Home", "wovnhome.com", "Keten Perde",
         "Doğal keten; ışık filtreleyen; yumuşak dökümlü; nötr tonlar; özel boyut seçenekleri",
         "2019'da kuruldu; Instagram ev dekor; DTC Shopify; organik keten vurgusu; minimalist estetik"),

        ("Solino Home", "solinohome.com", "Keten Ev Tekstili",
         "100% Avrupa keteni; masa örtüsü, peçete, perde; OEKO-TEX sertifikalı; doğal renkler",
         "2018'de kuruldu; Amazon-native DTC; Instagram organik yaşam; Shopify genişleme; premium keten segment"),

        ("Half Price Drapes", "halfpricedrapes.com", "Uygun Fiyatlı Lüks Perde",
         "İpek, kadife, keten görünümlü perdeler; geniş renk seçenekleri; ücretsiz kumaş örneği; özel boyut",
         "2018 sonrası DTC büyümesi; Instagram ev dekor; Shopify native; Facebook Ads; uygun fiyatlı lüks segment"),

        ("Loom Decor", "loomdecor.com", "Kişiselleştirilmiş Perde & Tekstil",
         "Online tasarım aracı; kumaş seçimi; özel ölçü perde; yastık ve masa örtüsü; profesyonel sonuç",
         "2018'de DTC lansmanı; Instagram iç tasarım; Shopify native; özel tasarım aracı; premium segment"),

        ("Barn & Willow", "barnandwillow.com", "Online Özel Perde",
         "Lüks kumaşlar; blackout ve sheer seçenekleri; ücretsiz kumaş swatch; kolay ölçüm rehberi",
         "2018'de DTC genişleme; Instagram ev dekor influencer'ları; Shopify native; online perde devrimi"),

        ("Quince Linens", "onequince.com", "Uygun Fiyatlı Lüks Keten",
         "Fabrikadan doğrudan; Avrupa keteni; çarşaf, havlu, perde; lüks kalite uygun fiyat; şeffaf fiyatlandırma",
         "2018'de kuruldu; DTC Shopify; Instagram ve Facebook; şeffaf tedarik zinciri; hızlı büyüme"),

        ("The Shade Store", "theshadestore.com", "Premium Özel Perde & Stor",
         "El yapımı ABD üretim; 1000+ kumaş seçeneği; ücretsiz tasarım danışmanlığı; özel ölçü",
         "2018 sonrası DTC dijital dönüşüm; Instagram iç tasarım; Shopify; showroom + online; premium segment"),

        ("Pure Linen", "purelinen.co", "Organik Keten Perde",
         "GOTS organik sertifikalı; Litvanya keteni; doğal tonlar; el yıkamalı; sürdürülebilir ambalaj",
         "2019'da kuruldu; Instagram organik yaşam; DTC Shopify; Avrupa üretim; çevre dostu mesaj"),

        ("Inside Weather", "insideweather.com", "Sürdürülebilir Ev Mobilyası & Tekstil",
         "Özel sipariş; sürdürülebilir malzeme; ABD üretim; perde ve yastık; modern tasarım",
         "2019'da kuruldu; DTC Shopify; Instagram modern ev; sürdürülebilir üretim mesajı; premium segment"),

        ("Silk & Willow", "silkandwillow.com", "Doğal Boyalı İpek Tekstil",
         "Bitkisel boyalar; organik ipek ve pamuk; masa örtüsü, peçete, runner; düğün ve ev dekoru",
         "2018'de kuruldu; Instagram düğün dekor; DTC Shopify; doğal boya vurgusu; boho estetik; niş segment"),

        ("Piglet in Bed", "pigletinbed.com", "Keten Yatak & Ev Tekstili",
         "100% doğal keten; çarşaf, perde, bornoz; stonewashed yumuşaklık; geniş renk paleti",
         "2018'de kuruldu; Instagram estetik; DTC Shopify; İngiltere kökenli küresel; influencer pazarlama; premium keten"),

        ("Coyuchi", "coyuchi.com", "Organik Pamuk Ev Tekstili",
         "GOTS organik; havlu, çarşaf, perde; sürdürülebilir; geri dönüşüm programı; doğal renkler",
         "2018 sonrası DTC dönüşümü; Instagram sürdürülebilir yaşam; Shopify native; B Corp sertifikalı; premium segment"),

        ("Cultiver", "cultiver.com.au", "Avustralya Keten Ev Tekstili",
         "Premium Avrupa keteni; çarşaf, yastık kılıfı, perde; doğal ve nötr tonlar; zamansız tasarım",
         "2018'de DTC genişleme; Instagram minimalist estetik; Shopify native; Avustralya'dan küresel; premium segment"),

        ("Pepper Home", "pepperhome.co", "Modern Perde & Döşemelik",
         "Instagram estetik perdeler; kolay montaj; modern desen ve renkler; uygun fiyatlı; hızlı kargo",
         "2020'de kuruldu; Instagram ev dekor; DTC Shopify; Facebook ve TikTok Ads; online-only marka"),

        ("Linen Tales", "linentales.com", "Litvanya Keten Perde",
         "El yapımı Litvanya keteni; stonewashed; perde, masa örtüsü, mutfak tekstili; OEKO-TEX",
         "2018'de kuruldu; Instagram organik yaşam; DTC Shopify; Etsy crossover; Avrupa üretim; çevre dostu"),

        ("Rough Linen", "roughlinen.com", "El Yapımı Amerikan Keten",
         "ABD'de el dikimi; premium Avrupa keteni; çarşaf, masa örtüsü, perde; artisan yaklaşım",
         "2018'de DTC genişleme; Instagram artisan estetik; Shopify native; premium segment; slow fashion"),

        ("Haven Earth", "havenearth.com", "Organik Bambu Perde",
         "Bambu lifi perde; doğal ışık filtresi; termal yalıtım; çevre dostu; hypoallerjenik",
         "2020'de kuruldu; Instagram çevre dostu ev; DTC Shopify; sürdürülebilir mesaj; online-only"),

        ("Deiji Studios", "deijistudios.com", "Keten Yaşam Tarzı Tekstil",
         "Keten pijama, perde, ev tekstili; pastel tonlar; minimalist Avustralya tasarımı; unisex",
         "2019'da kuruldu; Instagram estetik; DTC Shopify; slow fashion; influencer pazarlama; premium segment"),

        ("Wilet", "wilet.co", "Kırışmaz Keten Çarşaf & Perde",
         "Keten-Tencel karışımı; kırışma dirençli; yumuşak doku; modern renkler; kolay bakım",
         "2020'de kuruldu; DTC Shopify; Instagram modern ev; keten x teknoloji kavramı; Shark Tank görünümü"),

        ("Ella Moss Home", "ellamosshome.com", "Boho Ev Tekstili",
         "Boho chic perdeler; pom-pom detaylar; doğal dokular; makrome ve püsküllü; çeşitli desenler",
         "2019'da ev tekstili lansmanı; Instagram boho dekor; DTC Shopify; moda markası crossover; influencer pazarlama"),

        ("MINNA", "mightyminna.com", "Etik El Dokuma Tekstil",
         "Latin Amerika zanaatkarları; el dokuma battaniye, yastık, perde; adil ticaret; sürdürülebilir",
         "2018'de DTC genişleme; Instagram etik moda; Shopify native; sosyal etki hikayesi; B Corp sertifikalı"),

        ("Fig Linens", "figlinens.com", "Lüks İtalyan Keten",
         "İtalyan üretim; lüks çarşaf, havlu, masa örtüsü; monogram hizmeti; hediye ambalaj",
         "2018 sonrası DTC genişleme; Instagram lüks yaşam; Shopify native; kişiselleştirme; ultra-premium segment"),

        ("Citizenry Textiles", "the-citizenry.com", "Zanaatkar El Dokuma Perde",
         "Küresel zanaatkar iş birlikleri; el dokuma; doğal lifler; kültürel desenler; adil ticaret",
         "2018 sonrası perde genişlemesi; Instagram zanaatkar hikaye; DTC Shopify; sosyal etki; premium segment"),

        ("Matteo Los Angeles", "matteola.com", "Vintage Yıkamalı Keten Perde",
         "Vintage wash keten; California estetik; yumuşak dokulu perdeler; nötr ve toprak tonları",
         "2019'da DTC genişleme; Instagram California yaşam tarzı; Shopify native; premium segment; iç mimarlar hedef"),

        ("Hawkins New York", "hawkinsnewyork.com", "Basit Tasarım Ev Tekstili",
         "Basit ve işlevsel; stonewashed keten ve pamuk; mutfak, banyo, yatak tekstili; nötr renkler",
         "2018'de DTC genişleme; Instagram minimalist estetik; Shopify native; tasarım odaklı; premium segment"),

        ("Libeco", "libecohomestores.com", "Belçika Premium Keten",
         "1858'den beri keten üretimi; DTC kanalı yeni; premium çarşaf, perde, masa örtüsü; sürdürülebilir",
         "2019'da DTC online mağaza; Instagram premium keten; Shopify native; Belçika üretim; heritage x modern"),

        ("Casa Amarosa", "casaamarosa.com", "El Baskılı Hint Tekstili",
         "El baskısı blok print; pamuk perde ve masa örtüleri; canlı renkler; Hint zanaatkarları; boho estetik",
         "2019'da kuruldu; Instagram boho dekor; DTC Shopify; el yapımı vurgusu; kültürel hikaye; niş segment"),

        ("Mae Engelgeer", "mae-engelgeer.nl", "Hollanda Tasarım Tekstili",
         "Cesur geometrik desenler; el dokuma; tasarımcı serisi; perde ve battaniye; modern sanat estetiği",
         "2018 sonrası DTC genişleme; Instagram tasarım dünyası; Shopify native; premium niş; iç mimarlar hedef"),

        ("Tandem Home", "tandemhome.co", "Minimalist Keten Perde Takımı",
         "Paket halinde perde çözümleri; ölçüye göre kesim; doğal keten; kolay montaj; uygun fiyat",
         "2021'de kuruldu; Instagram ev dekor; DTC Shopify; online perde kolaylığı; startup; hızlı iterasyon"),

        ("Marialma", "marialma.com", "Teknolojik Ev Tekstili",
         "Bakır ve gümüş infüzyonlu kumaşlar; antimikrobiyal; termal düzenleme; çarşaf ve perde; Portekiz üretim",
         "2019'da kuruldu; teknoloji x tekstil; DTC Shopify; Instagram sağlık bilinçli; Avrupa üretim; premium"),
    ],

    # ─────────────────────────────────────────────────────────────────
    # 5. DEKORATİF YASTIK & KIRLENT  (Decorative Pillows & Cushions)
    # ─────────────────────────────────────────────────────────────────
    "Dekoratif Yastık & Kırlent": [
        ("Studio Pillows", "studiopillows.com", "Tasarımcı Dekoratif Kırlent",
         "İç mimarlar için özel kumaşlar; kesme kadife, keten, ipek; geniş desen kütüphanesi; özel boyut",
         "2019'da kuruldu; Instagram iç tasarım; DTC Shopify; tasarımcı hedef kitle; premium segment"),

        ("The Citizenry Pillows", "the-citizenry.com", "Zanaatkar El Yapımı Kırlent",
         "Dünya genelinde zanaatkar iş birlikleri; el dokuma; doğal lifler; kültürel desenler; adil ticaret",
         "2018 sonrası DTC genişleme; Instagram zanaatkar hikaye; Shopify native; sosyal etki; premium segment"),

        ("Bolé Road Textiles", "boleroadtextiles.com", "Etiyopya El Dokuma Kırlent",
         "Etiyopya zanaatkarları; el dokuma pamuk; canlı renkler; geometrik desenler; adil ticaret",
         "2018'de DTC genişleme; Instagram Afrika tasarım; Shopify native; kültürel hikaye; sosyal etki mesajı"),

        ("Jungalow", "jungalow.com", "Boho Tropikal Kırlent",
         "Justina Blakeney markası; tropik desenler; canlı renkler; pom-pom ve püskül detaylar; makrome",
         "2018 sonrası DTC genişleme; Instagram boho dekor; Shopify native; influencer-founded; Target iş birliği"),

        ("Sunday Citizen Pillows", "sundaycitizen.co", "Snug Bambu Yastık",
         "Bambu kumaş; ultra yumuşak; kristal infüzyon serileri; dekoratif ve fonksiyonel; hypoallerjenik",
         "2019'da yastık kategorisine giriş; DTC Shopify; Instagram wellness; sürdürülebilir mesaj; premium segment"),

        ("Bole Road Home", "boleroadtextiles.com", "Afrika İlhamlı Dekoratif Yastık",
         "Etiyopya el dokuma; organik pamuk; cesur geometrik desenler; dünya renkleri; zanaatkar hikayesi",
         "2018'de DTC genişleme; Instagram kültürel dekor; Shopify native; adil ticaret; sosyal etki vurgusu"),

        ("Block Shop Textiles", "blockshoptextiles.com", "El Baskılı Hint Kırlent",
         "El baskısı blok print; Rajasthan zanaatkarları; organik mürekkep; kadın zanaatkar güçlendirme",
         "2018 sonrası DTC büyümesi; Instagram el yapımı estetik; Shopify native; sosyal etki; premium segment"),

        ("Casa Amarosa Pillows", "casaamarosa.com", "El Baskılı Dekoratif Kırlent",
         "Hint blok print; pamuk ve keten; canlı renkler; boho ve coastal tasarımlar; el yapımı",
         "2019'da kuruldu; Instagram boho dekor; DTC Shopify; kültürel hikaye; uygun fiyatlı premium"),

        ("Quiet Town", "quiet-town.com", "Modern Duş Perdesi & Yastık",
         "Su geçirmez kumaş duş perdesi markasından kırlent serisi; grafik desenler; modern renkler",
         "2018'de kuruldu; Instagram modern banyo/ev; DTC Shopify; Brooklyn tasarım; çevre dostu malzeme"),

        ("Morrow Soft Goods", "morrowsoftgoods.com", "Organik Pamuk Kırlent",
         "GOTS organik pamuk; stonewashed yumuşaklık; nötr tonlar; minimalist tasarım; kaliforniya estetik",
         "2019'da kuruldu; Instagram minimalist ev; DTC Shopify; organik sertifika; premium segment"),

        ("Ara Collective", "aracollective.com", "Latin Amerika El Yapımı Kırlent",
         "Meksika zanaatkarları; el dokuma; doğal boyalar; geleneksel motifler; sürdürülebilir",
         "2019'da kuruldu; Instagram etik dekor; DTC Shopify; zanaatkar hikayesi; sosyal etki; premium niş"),

        ("Leah Singh", "leahsingh.com", "El İşlemeli Lüks Kırlent",
         "El işlemesi; Kashmir ve Hindistan zanaatkarları; lüks ipek ve yün; sanatsal desenler; sınırlı üretim",
         "2018'de DTC genişleme; Instagram lüks ev dekor; Shopify native; el yapımı vurgusu; ultra-premium"),

        ("Society6 Pillows", "society6.com", "Sanatçı Tasarımlı Kırlent",
         "Bağımsız sanatçı tasarımları; print-on-demand; sınırsız desen; uygun fiyatlı sanat dekor",
         "2018 sonrası DTC büyümesi; Instagram sanat topluluğu; print-on-demand modeli; küresel sanatçı ağı"),

        ("Loloi Pillows", "loloirugs.com", "Tasarımcı Koleksiyon Kırlent",
         "Amber Lewis, Chris Loves Julia koleksiyonları; dokumalı ve baskılı; halı-kırlent koordinasyonu",
         "2019'da kırlent kategorisine giriş; Instagram influencer dekor; DTC ve toptan; premium segment"),

        ("Heather Taylor Home", "heathertaylorhome.com", "El Boyalı Çizgili Kırlent",
         "El boyalı kumaşlar; canlı çizgili desenler; LA üretim; organik pamuk; her biri benzersiz",
         "2018'de DTC genişleme; Instagram LA estetik; Shopify native; el yapımı vurgusu; premium segment"),

        ("Dusen Dusen", "dufrfrufsen.com", "Grafik Desenli Modern Kırlent",
         "Cesur grafik desenler; pop art estetiği; parlak renkler; ev tekstili x sanat; yatak ve kırlent",
         "2018'de DTC genişleme; Instagram grafik tasarım; Shopify native; Brooklyn tasarım; Gen Z hedef kitle"),

        ("St. Frank", "stfrankcloth.com", "Afrika İlhamlı Vintage Kırlent",
         "Afrika'dan vintage kumaşlar; el dikimi; her parça tek; kültürel hikaye anlatımı; karışık medya",
         "2018'de DTC genişleme; Instagram kültürel dekor; Shopify native; sosyal etki; premium segment"),

        ("Bfgf", "bfrfrfgf.com", "El İşlemeli Yüz Kırlent",
         "Punch needle tekniği; portre ve yüz desenleri; el yapımı; organik pamuk; sanatsal ve eğlenceli",
         "2019'da kuruldu; Instagram el yapımı sanat; DTC Shopify; TikTok viral; niş sanat kırlent; online-only"),

        ("Studio Arethusa", "studioarethusa.com", "Botanik Desenli Keten Kırlent",
         "Botanik illüstrasyonlar; doğal keten; dijital baskı; İngiltere tasarım; doğa ilhamlı",
         "2019'da kuruldu; Instagram botanik estetik; DTC Shopify; niş tasarım; sanat x ev dekor"),

        ("Minna Goods", "mightyminna.com", "Etik El Dokuma Kırlent",
         "Latin Amerika zanaatkarları; el dokuma pamuk; geometrik desenler; adil ticaret; B Corp",
         "2018'de DTC genişleme; Instagram etik dekor; Shopify native; sosyal etki; sürdürülebilir mesaj"),

        ("Revival Home", "revivalhomegoods.com", "Vintage Kumaş Kırlent",
         "Geri dönüştürülmüş vintage kumaşlar; her biri tek; sürdürülebilir; eklektik tasarım; el yapımı",
         "2019'da kuruldu; Instagram vintage estetik; DTC Shopify; upcycle konsepti; çevre dostu; niş segment"),

        ("Ochre Handmade", "ochrehome.com", "El Yapımı Kadife Kırlent",
         "Crushed kadife; el dikimi; zengin renkler; lüks doku; İngiltere atölyesi; dekoratif püsküller",
         "2018 sonrası DTC genişleme; Instagram lüks ev dekor; Shopify native; el yapımı İngiliz kalitesi"),

        ("Anchal Project", "anchalproject.org", "Sosyal Etki Kırlent",
         "Hindistan kadın zanaatkarları; el dikişi; geri dönüştürülmüş malzeme; kantha tekniği; sosyal misyon",
         "2018'de DTC genişleme; Instagram sosyal etki; Shopify native; kar amacı gütmeyen; etik tüketim"),

        ("Joanna Buchanan", "joannabuchanan.com", "Lüks Dekoratif Yastık",
         "Mücevherli ve boncuklu detaylar; lüks kumaşlar; koleksiyon parçaları; hediye segmenti; el işçiliği",
         "2018'de DTC genişleme; Instagram lüks dekor; Shopify native; premium hediye; influencer pazarlama"),

        ("Apiece Apart Home", "apieceapart.com", "Doğal Dokulu Kırlent",
         "Organik pamuk ve yün; doğal dokular; nötr tonlar; Meksika ve Peru üretim; sürdürülebilir",
         "2019'da ev kategorisine giriş; Instagram slow fashion; DTC Shopify; moda crossover; premium segment"),

        ("Goodee Pillows", "goodeeworld.com", "Küratörlü Etik Kırlent",
         "Etik ve sürdürülebilir marka seçkisi; B Corp küratörlüğü; dünya genelinde zanaatkarlar; premium kalite",
         "2019'da kuruldu; Instagram etik yaşam; DTC Shopify; küratörlü pazar yeri; sosyal etki mesajı"),

        ("Justina Blakeney Home", "jungalow.com", "Boho Maksimalist Kırlent",
         "Cesur desenler; tropikal ve boho; farbala ve püskül; canlı renkler; katmanlı dekor konsepti",
         "2018 sonrası DTC büyümesi; Instagram boho queen; Shopify native; influencer-founded; kitap ve marka"),

        ("Meso Goods", "mesogoods.com", "Guatemala El Dokuma Kırlent",
         "Maya zanaatkarları; geleneksel tezgah dokuma; doğal renkler; kültürel miras; adil ticaret",
         "2019'da kuruldu; Instagram kültürel dekor; DTC Shopify; zanaatkar hikaye; sosyal etki; niş segment"),

        ("Eny Lee Parker Home", "enyleeparker.com", "Seramik & Tekstil Sanat Kırlent",
         "Sanat eseri kırlentler; seramik sanatçısı bakış açısı; sınırlı üretim; cesur formlar ve renkler",
         "2020'de ev tekstili lansmanı; Instagram sanat dünyası; DTC Shopify; niş sanat segment; premium"),

        ("Pom Pom at Home", "pompomsathome.com", "Pom-Pom Detaylı Keten Kırlent",
         "Stonewashed keten; pom-pom ve püskül detaylar; doğal tonlar; rahat boho lüks; el yapımı detaylar",
         "2018 sonrası DTC genişleme; Instagram boho lüks; Shopify native; influencer ev dekor; premium segment"),
    ],

    # ─────────────────────────────────────────────────────────────────
    # 6. BEBEK UYKU & TEKSTİL  (Baby Sleep & Textiles)
    # ─────────────────────────────────────────────────────────────────
    "Bebek Uyku & Tekstil": [
        ("Kyte Baby", "kytebaby.com", "Bambu Bebek Tulumu",
         "Bambu rayon kumaş; ultra yumuşak; nefes alan; termal düzenleyici; TOG derecelendirmeli uyku tulumu",
         "2018'de kuruldu; Instagram anne influencer'lar; DTC Shopify; bambu bebek giyimi kategorisini büyüttü; $100M+ gelir"),

        ("Dreamland Baby", "dreamlandbabyco.com", "Ağırlıklı Bebek Uyku Tulumu",
         "Hafif ağırlıklı uyku tulumu; derin basınç stimülasyonu; daha uzun uyku; pediatrist onaylı",
         "2019'da kuruldu; Shark Tank yatırımı; anne-kuruculu marka; DTC Shopify; Instagram ve TikTok; hızlı büyüme"),

        ("Copper Pearl", "copperpearl.com", "Bambu Kundak & Önlük",
         "Bambu viskoz; kundak battaniye; bandana önlükler; bebek şapkaları; yumuşak ve nefes alan",
         "2018'de DTC büyümesi; Instagram anne; Amazon ve Shopify; influencer pazarlama; uygun fiyatlı premium"),

        ("Solly Baby", "sollybaby.com", "Yumuşak Bebek Taşıyıcı Wrap",
         "TENCEL modal kumaş; ultra hafif; nefes alan; ergonomik tasarım; güzel renk seçenekleri",
         "2018'de DTC genişleme; Instagram anne influencer; Shopify native; babywearing topluluğu; premium segment"),

        ("Gunamuna", "gunamuna.com", "Fermuarlı Uyku Tulumu",
         "Alt değiştirme fermuarı; bambu viskoz; TOG derecelendirmeli; ayak açılabilir tasarım",
         "2018'de kuruldu; DTC Shopify; Instagram anne; Amazon genişleme; pediatrist onaylı; online-only"),

        ("Little Sleepies", "littlesleepies.com", "Bambu Bebek & Çocuk Pijama",
         "Bambu viskoz; yumuşak ve esnek; ebeveyn-çocuk eşleşen setler; canlı desenler; geniş beden aralığı",
         "2019'da kuruldu; TikTok ve Instagram viral; DTC Shopify; anne topluluğu; hızlı büyüme; $50M+ gelir"),

        ("Woolino", "woolino.com", "Merinos Yünü Uyku Tulumu",
         "4 mevsim merinos yünü; termal düzenleme; 2 ay-2 yaş tek beden; OEKO-TEX sertifikalı",
         "2018'de DTC genişleme; Amazon ve Shopify; doğal kumaş vurgusu; Instagram anne; premium segment"),

        ("Snuggle Me Organic", "snugglemeorganic.com", "Organik Bebek Yatağı",
         "GOTS organik pamuk; ergonomik bebek yatağı; anne karnı simülasyonu; hypoallerjenik; ABD üretim",
         "2018'de DTC patlaması; Instagram anne influencer; Shopify native; organik vurgusu; viral ürün"),

        ("Nested Bean", "nestedbean.com", "Ağırlıklı Bebek Giyimi",
         "Zen uyku tulumu ve Zen bodysuit; hafif ağırlıklı ön panel; anne dokunuşu simülasyonu; patentli tasarım",
         "2018'de DTC genişleme; Amazon ve Shopify; anne topluluğu; Instagram; bilimsel araştırma destekli"),

        ("Pehr", "pfrfrfrhr.com", "El Boyalı Bebek Tekstili",
         "El boyalı desenler; organik pamuk; bebek çarşafı, battaniye, saklama sepetleri; pastel tonlar",
         "2018 sonrası DTC büyümesi; Instagram bebek odası dekor; Shopify native; Kanada kökenli; premium segment"),

        ("Under the Nile", "underthenile.com", "Mısır Organik Pamuk Bebek",
         "%100 Mısır organik pamuğu; GOTS sertifikalı; adil ticaret; bebek giyim ve tekstil; doğal boyalar",
         "2018'de DTC genişleme; çevre dostu anne topluluğu; Amazon ve Shopify; sürdürülebilir üretim"),

        ("Solly Baby Wrap", "sollybaby.com", "Modal Bebek Taşıma Kumaşı",
         "TENCEL Modal; ultra yumuşak ve hafif; 4.5m uzunluk; nefes alan; çeşitli renkler; kolay sarım",
         "2018 sonrası DTC genişleme; Instagram babywearing; Shopify native; anne influencer pazarlama; premium"),

        ("Burt's Bees Baby", "burtsbeesbaby.com", "Organik Pamuk Bebek Tekstili",
         "GOTS organik pamuk; bebek çarşafı, pijama, battaniye; doğal boya; uygun fiyatlı organik",
         "2018 sonrası DTC dönüşümü; Instagram organik bebek; Shopify ve Amazon; anne influencer; geniş dağıtım"),

        ("SwaddleDesigns", "swaddledesigns.com", "Hemşire Tasarımlı Kundak",
         "Hemşire kuruculu; muslin ve flannel kundak; uyku tulumu; güvenli uyku odaklı; çeşitli boyutlar",
         "2018'de DTC genişleme; Amazon ve Shopify; anne topluluğu; güvenli uyku mesajı; online-only"),

        ("Loulou Lollipop", "louloulollipop.com", "Muslin Battaniye & Bebek Aksesuarı",
         "Bambu muslin battaniye; silikon diş kaşıyıcı; modern desenler; yumuşak pastel renkler",
         "2018'de kuruldu; Instagram anne estetik; DTC Shopify; Kanada kökenli; influencer pazarlama; hızlı büyüme"),

        ("Mebie Baby", "mebiebaby.com", "Muslin Kundak & Şapka Seti",
         "Organik muslin; modern desenler; yenidoğan setleri; hastane çıkış kıyafetleri; hediye kutuları",
         "2019'da kuruldu; Instagram yenidoğan; DTC Shopify; anne influencer; hediye segmenti; online-only"),

        ("Saranoni", "saranoni.com", "Lüks Bebek Battaniyesi",
         "Ultra yumuşak minky kumaş; çift katlı; saten kenar; lüks dokunuş; yetişkin boyutları da mevcut",
         "2018'de DTC genişleme; Instagram lüks bebek; Shopify native; anne influencer; hediye premium segment"),

        ("Lou Lou & Company", "loulouandcompany.com", "Düğümlü Bebek Tulumu",
         "Bambu ve pamuk karışımı; düğümlü tasarım; yumuşak ve esnek; yenidoğan favori; modern desenler",
         "2018'de kuruldu; Instagram yenidoğan; DTC Shopify; anne influencer; online-only; hızlı büyüme"),

        ("Halo Sleep", "halosleep.com", "Giyilebilir Uyku Battaniyesi",
         "Ters fermuar güvenlik tasarımı; TOG seçenekleri; muslin ve pamuk; AAP önerisi; güvenli uyku",
         "2018 sonrası DTC dönüşümü; Instagram güvenli uyku; Amazon ve Shopify; pediatrist önerili; geniş dağıtım"),

        ("Aden + Anais", "adenandanais.com", "Muslin Kundak Battaniye",
         "Muslin kundak öncüsü; bambu ve pamuk; çeşitli desen ve boyut; uyku tulumu; bebek havlusu",
         "2018 sonrası DTC genişleme; Instagram anne; Shopify native; Disney iş birlikleri; küresel marka"),

        ("Tubby Todd", "tubbytodd.com", "Doğal Bebek Cilt Bakımı & Tekstil",
         "All-Over Ointment viral ürün; doğal bileşenler; bebek banyosu; hassas cilt; dermatolojist test",
         "2018'de DTC patlaması; Instagram anne; TikTok viral; Shopify native; cilt bakımı x bebek tekstil"),

        ("Goumi Kids", "goumikids.com", "Organik Bebek Eldiveni & Ayakkabı",
         "Organik pamuk + bambu; yenidoğan eldiveni; yumuşak patik; jammerz pantolon; kolay giydirme",
         "2018'de DTC genişleme; Instagram yenidoğan; Shopify native; organik vurgusu; uygun fiyatlı premium"),

        ("Gathre", "gathre.com", "Deri Oyun Matı & Bebek Tekstili",
         "Mikrofiber deri; kolay temizlenen; modern tasarım; oyun matı, önlük, çanta; çeşitli renkler",
         "2018'de DTC genişleme; Instagram modern anne; Shopify native; çok amaçlı ürünler; premium segment"),

        ("Caden Lane", "cadenlane.com", "Bebek Çarşafı & Kundak Seti",
         "Modern desenler; pamuk ve bambu; bebek yatağı çarşafı; kundak; isim işlemeli battaniye",
         "2018'de DTC genişleme; Instagram bebek odası dekor; Shopify native; kişiselleştirme; hediye segmenti"),

        ("Baby Brezza Sleep", "babybrezza.com", "Akıllı Bebek Uyku Sistemi",
         "Otomatik bebek salıncağı; beyaz gürültü; hareket sensörü; app kontrollü; uyku takibi",
         "2018 sonrası DTC genişleme; teknoloji x bebek; Instagram ve YouTube; Amazon ve Shopify; yenilikçi"),

        ("Hatch Baby", "hfrfrfratch.co", "Akıllı Bebek Gece Lambası",
         "Ses makinesi + gece lambası; app kontrollü; uyku eğitimi; renk değiştiren; Hatch Rest bestseller",
         "2018 sonrası DTC patlaması; TikTok ve Instagram viral; Shopify native; anne topluluğu; teknoloji segment"),

        ("Burp Cloth Co.", "burpclothco.com", "Premium Burp Bezi & Bebek Tekstili",
         "Organik muslin; büyük boyut; emici; modern desenler; hediye seti; yenidoğan temel ihtiyaç",
         "2019'da kuruldu; Instagram yenidoğan; DTC Shopify; anne influencer; hediye segmenti; online-only"),

        ("Crane Baby", "cranebaby.com", "Bebek Odası Dekor & Tekstil",
         "Organik pamuk çarşaf; nemlendirici ve hava temizleyici; bebek odası dekor setleri; modern tasarım",
         "2019'da bebek tekstil lansmanı; Instagram bebek odası; DTC Shopify; wellness x bebek; premium segment"),

        ("Owlet x Dreamland", "dreamlandbabyco.com", "Akıllı Uyku İzleme Tulumu",
         "Uyku izleme sensörlü; ağırlıklı tulum; SpO2 ve nabız takibi; app bağlantılı; güvenli uyku",
         "2020'de iş birliği; teknoloji x uyku; DTC Shopify; anne teknoloji topluluğu; premium güvenlik segment"),

        ("Tealbee", "tealbee.com", "Toddler Uyku Tulumu",
         "2-5 yaş için uyku tulumu; ayaklı tasarım; TOG derecelendirmeli; yürüyen çocuk güvenliği; kolay fermuarı",
         "2019'da kuruldu; Amazon-native DTC; Instagram anne; toddler uyku niş'i; Shopify genişleme; online-only"),

        ("Rookie Humans", "rookiehumans.com", "Sanatçı Tasarımlı Bebek Çarşafı",
         "Bağımsız sanatçı tasarımları; organik pamuk; canlı illüstrasyonlar; çarşaf ve battaniye; hediye",
         "2018'de kuruldu; Instagram bebek odası sanat; DTC Shopify; sanatçı iş birlikleri; premium niş segment"),

        ("Milk Snob", "milksnob.com", "Çok Amaçlı Bebek Örtüsü",
         "Emzirme örtüsü, araba koltuğu kapağı, alışveriş arabası örtüsü; tek ürün çoklu kullanım; modern desenler",
         "2018'de Shark Tank sonrası DTC büyümesi; Instagram anne; Shopify native; çok amaçlı konsept; viral ürün"),

        ("Poppy + Sage", "poppyandsage.com", "El Dikişli Muslin Battaniye",
         "El dikişli; organik muslin; modern tropik desenler; yenidoğan hediye; hafif ve nefes alan",
         "2019'da kuruldu; Instagram boho anne; DTC Shopify; Bali üretim; el yapımı vurgusu; hediye segmenti"),

        ("Finn + Emma", "finnandemma.com", "Organik & Adil Ticaret Bebek",
         "%100 organik pamuk; doğal ahşap oyuncaklar; adil ticaret; GOTS sertifikalı; modern tasarım",
         "2018'de DTC genişleme; Instagram çevre dostu anne; Shopify native; sosyal etki; premium organik"),

        ("Natemia", "natemia.com", "Bambu Muslin & Bebek Havlusu",
         "Bambu muslin battaniye; kapüşonlu bebek havlusu; ultra yumuşak; doğal ve hypoallerjenik; pastel tonlar",
         "2019'da kuruldu; Amazon-native DTC; Instagram anne; bambu bebek tekstili niş'i; uygun fiyatlı premium"),
    ],
}
