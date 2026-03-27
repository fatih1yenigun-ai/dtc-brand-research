"""
yuvacim_v3_batch4.py
Post-2018 ecommerce-native DTC brands — MOM / LOUNGEWEAR / SLIPPERS / SLEEP TECH / SUSTAINABLE HOME / PET / OUTDOOR
7 categories, 200+ brands total
All descriptions in Turkish (brand names & URLs in English)
"""

BATCH_DATA = {

    # ─────────────────────────────────────────────────────────────────
    # 1. ANNE & HAMİLE ÜRÜNLERİ  (Mom & Maternity Products)
    # ─────────────────────────────────────────────────────────────────
    "Anne & Hamile Ürünleri": [
        ("Momcozy", "momcozy.com", "Giyilebilir Göğüs Pompası",
         "Hands-free giyilebilir süt pompası; sessiz motor; app bağlantılı modeller; portatif tasarım",
         "2018'de kuruldu; Amazon ve TikTok üzerinden viral büyüme; influencer mama kampanyaları; global DTC; yıllık $100M+ gelir"),

        ("Frida Mom", "frida.com", "Doğum Sonrası Bakım",
         "Postpartum recovery kit; perineal buz pedi; göğüs bakım ürünleri; tabusuz pazarlama",
         "2019'da Frida Baby'nin uzantısı olarak lansmanlandı; sosyal medyada cesur reklam kampanyaları; Amazon bestseller; DTC Shopify"),

        ("Willow", "onewillow.com", "Akıllı Süt Pompası",
         "Tamamen kablosuz pompa; süt torbasız versiyon; uygulama ile süt takibi; giyilebilir teknoloji",
         "2018'de piyasaya çıktı; CES ödülü kazandı; DTC ve perakende genişleme; teknoloji odaklı anne markası"),

        ("Elvie", "elvie.com", "Sessiz Giyilebilir Pompa",
         "Ultra sessiz motor; akıllı sensörler; app bağlantılı; sütyenin içine sığan tasarım",
         "2019'da ABD pazarına girdi; Londra kökenli; DTC ve Target genişleme; kadın sağlığı teknolojisi öncüsü"),

        ("Bodily", "itsbodily.com", "Doğum Sonrası Giyim",
         "Emzirme sütyeni; postpartum iç çamaşırı; doğum elbisesi; fonksiyonel tasarım",
         "2019'da kuruldu; doğum ve sonrasına özel giyim DTC; Instagram ve influencer odaklı; tabusuz marka dili"),

        ("Kindred Bravely", "kindredbravely.com", "Hamile & Emzirme Giyim",
         "Emzirme pijaması; hamile taytı; postpartum destek kıyafetleri; yumuşak kumaşlar",
         "2018 sonrası hızlı DTC büyüme; Amazon bestseller; Facebook Ads ve anne toplulukları; sadık müşteri kitlesi"),

        ("Hatch", "hatchcollection.com", "Lüks Hamile Giyim",
         "Premium hamile elbiseleri; doğum sonrası da giyilebilir tasarım; şık ve fonksiyonel",
         "2019 sonrası DTC büyümesi; Instagram influencer kampanyaları; ünlü anneler ile işbirliği; premium segment"),

        ("KeaBabies", "keababies.com", "Bebek Taşıma & Emzirme",
         "Bebek wrap taşıyıcı; emzirme örtüsü; muslin battaniye; organik malzemeler",
         "2018'de kuruldu; Amazon-first strateji; 50.000+ yorum; DTC genişleme; uygun fiyatlı kalite"),

        ("Natalist", "natalist.com", "Doğurganlık & Gebelik Testi",
         "Ovülasyon testi; gebelik testi; prenatal vitamin; doğurganlık desteği",
         "2019'da kuruldu; modern doğurganlık markası; DTC Shopify; Instagram ve TikTok; temiz tasarım"),

        ("Needed", "thisisneeded.com", "Prenatal Vitamin & Takviye",
         "Klinik formüllü prenatal; postnatal vitamin; doğurganlık desteği; temiz içerik",
         "2019'da kuruldu; fonksiyonel tıp destekli; DTC abonelik modeli; Instagram ve podcast pazarlama"),

        ("Perelel", "perelelhealth.com", "Dönemsel Prenatal Vitamin",
         "Trimestere özel vitamin paketleri; OB-GYN geliştirmeli; temiz içerik formülü",
         "2020'de kuruldu; Tia Mowry yatırımcı; DTC abonelik; Instagram ve influencer; hekim onaylı"),

        ("Mommy's Bliss", "mommysbliss.com", "Bebek Sindirim & Rahatlatıcı",
         "Gripe water; probiyotik damla; organik bebek bakım; anne takviye ürünleri",
         "2018 sonrası DTC genişleme; Amazon bestseller; organik sertifikalı; anne topluluk pazarlaması"),

        ("Babymoov", "babymoov.com", "Akıllı Bebek Ürünleri",
         "Cosypad bebek yatağı; Dream Belt hamile yastığı; UV koruma çadırı; ergonomik tasarımlar",
         "2019 sonrası DTC e-ticaret büyümesi; Fransız kökenli; Amazon ve Shopify; inovatif bebek ürünleri"),

        ("TushBaby", "tushbaby.com", "Ergonomik Kalça Taşıyıcı",
         "Kalça oturak taşıyıcı; omuz ağrısını önler; cepli tasarım; 0-36 ay kullanım",
         "2019'da kuruldu; Shark Tank'ta yatırım aldı; TikTok viral; DTC ve Amazon; anne influencer kampanyaları"),

        ("Ryan & Rose", "ryanandrose.com", "Emzik & Diş Kaşıyıcı",
         "Cutie PAT emzik+diş kaşıyıcı kombo; medikal silikon; minimalist tasarım; pastel renkler",
         "2018'de kuruldu; Instagram viral; anne DTC markası; Shopify; organik büyüme ve influencer"),

        ("Lalo", "meetlalo.com", "Modern Mama Çocuk Ürünleri",
         "The Chair mama sandalyesi; The Daily stroller; oyun parkı; şık tasarım",
         "2019'da kuruldu; DTC modern ebeveyn markası; Instagram odaklı; premium segment; modüler ürünler"),

        ("Tubby Todd", "tubbytodd.com", "Doğal Bebek Cilt Bakım",
         "All Over Ointment; botanik içerikli; bebek egzama çözümü; temiz formül",
         "2018 sonrası viral büyüme; anne toplulukları; TikTok ve Instagram; DTC Shopify; kült ürün statüsü"),

        ("Baby Brezza", "babybrezza.com", "Otomatik Mama Hazırlayıcı",
         "Formula Pro; otomatik sterilizatör; akıllı biberon ısıtıcı; tek tuşla mama",
         "2018 sonrası DTC genişleme; Amazon bestseller; TikTok anne içerikleri; teknoloji odaklı bebek markası"),

        ("Copper Pearl", "copperpearl.com", "Bebek Bandana & Muslin",
         "Bebek bandana önlüğü; muslin örtü; knit battaniye; modern desen tasarımları",
         "2018'de kuruldu; Shopify DTC; Instagram mama influencer; Amazon bestseller; uygun fiyatlı şık bebek aksesuarı"),

        ("Posh Peanut", "poshpeanut.com", "Lüks Bebek & Anne Giyim",
         "Bambu kumaş tulumlar; eşleşen aile pijamaları; çiçekli desenler; ultra yumuşak",
         "2018'de kuruldu; Instagram-first büyüme; anne-bebek eşleşen setler trendi; DTC; influencer kampanyaları"),

        ("Kyte Baby", "kytebaby.com", "Bambu Bebek Giyim",
         "Bambu rayon kumaş; sleep bag; toddler giyim; organik yumuşaklık",
         "2018'de kuruldu; TikTok viral; DTC Shopify; anne topluluğu; bambu bebek giyim kategorisinin lideri"),

        ("Gunamuna", "gunamuna.com", "Bebek Uyku Tulumu",
         "Fermuarlı uyku tulumu; WONDERZiP teknoloji; alt değiştirme kolaylığı; TOG seçenekleri",
         "2019'da kuruldu; DTC Shopify; Amazon genişleme; Instagram anne influencer; inovatif uyku çözümü"),

        ("Little Sleepies", "littlesleepies.com", "Bambu Bebek Pijama",
         "Bambu viskoz pijama; eşleşen aile seti; canlı desenler; ultra esnek kumaş",
         "2019'da kuruldu; TikTok ve Instagram viral; drop model satış; DTC; sadık anne topluluğu"),

        ("Solly Baby", "sollybaby.com", "Premium Bebek Wrap",
         "TENCEL modal wrap; hafif ve nefes alır; yenidoğan için ideal; şık renkler",
         "2018 sonrası DTC büyüme; Instagram anne influencer; premium wrap segmenti; Shopify; ünlü tercihi"),

        ("Gathre", "gathre.com", "Deri Oyun Matı & Önlük",
         "Suni deri mat; kolay temizlenir; su geçirmez önlük; şık ev estetiği",
         "2018'de kuruldu; Instagram-first; anne yaşam tarzı markası; DTC Shopify; fonksiyonel+estetik"),

        ("Caden Lane", "cadenlane.com", "Hastane & Yenidoğan Seti",
         "Hastane çıkış kıyafeti; swaddle seti; kişiselleştirilebilir ürünler; modern desenler",
         "2018 sonrası DTC genişleme; Instagram ve TikTok; bebek hediye pazarı; Shopify; anne influencer"),

        ("Bonsie", "bonsiewear.com", "Cilt Teması Tulumu",
         "Skin-to-skin bodysuit; NICU ilhamlı; ön açılır tasarım; organik pamuk",
         "2019'da kuruldu; anne-bebek cilt teması odaklı; DTC niş marka; Instagram; hastane ve ev kullanımı"),

        ("Nesting Days", "nestingdays.com", "Cilt Teması Giyim",
         "Kanguru bakım gömleği; baba ve anne versiyonları; preemie uyumlu; tıbbi onaylı",
         "2019'da kuruldu; NICU ailelerine yönelik; DTC Shopify; sosyal medya ve hastane partnerlikleri"),

        ("WeeSprout", "weesprout.com", "Silikon Bebek Beslenme",
         "Silikon tabak; emme tabanlı kase; bebek kaşığı seti; BPA-free",
         "2018'de kuruldu; Amazon-first DTC; anne topluluğu pazarlama; uygun fiyatlı; 10.000+ Amazon yorum"),

        ("Loveevery", "lovevery.com", "Gelişim Odaklı Oyuncak",
         "Yaş bazlı oyun kiti; Montessori ilhamlı; nörobilim destekli; ahşap oyuncaklar",
         "2018'de kuruldu; DTC abonelik kutusu; Instagram ve podcast; $100M+ değerleme; ebeveyn eğitimi içeriği"),

    ],

    # ─────────────────────────────────────────────────────────────────
    # 2. EV GİYİM & LOUNGEWEAR  (Homewear & Loungewear)
    # ─────────────────────────────────────────────────────────────────
    "Ev Giyim & Loungewear": [
        ("Skims", "skims.com", "Shapewear & Loungewear",
         "Cozy koleksiyon; yumuşak modal kumaş; geniş beden aralığı; nötr tonlar",
         "2019'da Kim Kardashian tarafından kuruldu; sosyal medya ve ünlü etkisi; DTC; $4B+ değerleme; global viral"),

        ("Lunya", "lunya.com", "Lüks Uyku Giyim",
         "Washable silk pijama; Restore pijama seti; anti-mikrobiyel kumaş; şık uyku giyim",
         "2018 sonrası hızlı DTC büyüme; Instagram ve podcast reklamları; kadın odaklı; premium segment"),

        ("Eberjey", "eberjey.com", "Premium Pijama & Loungewear",
         "Gisele pijama seti; modal kumaş; romantik tasarımlar; lüks ev giyim",
         "2018 sonrası DTC dönüşümü; Instagram influencer; resort ve loungewear; Nordstrom genişleme"),

        ("Lake Pajamas", "lakepajamas.com", "Pima Pamuk Pijama",
         "Pima pamuk; eşleşen aile pijamaları; klasik tasarım; pastel renkler",
         "2018'de kuruldu; Instagram-first viral büyüme; DTC Shopify; anne influencer; aile pijama trendi öncüsü"),

        ("Printfresh", "printfresh.com", "Desenli Pijama & Ev Giyim",
         "Cesur baskılar; eşleşen setler; bambu ve organik pamuk; sanatsal tasarım",
         "2019'da kuruldu; Philadelphia kökenli; Instagram ve TikTok; DTC; benzersiz desen tasarımları"),

        ("Barefoot Dreams", "barefootdreams.com", "CozyChic Loungewear",
         "CozyChic kumaş; battaniye hırka; ultra yumuşak; ev keyfi odaklı",
         "2018 sonrası viral DTC büyüme; Oprah's Favorite Things; TikTok ve Instagram; Nordstrom bestseller"),

        ("Cozy Earth", "cozyearth.com", "Bambu Loungewear & Pijama",
         "Bambu viskoz; termal düzenleme; Oprah onaylı; premium ev giyim",
         "2019'da viral oldu; Oprah's Favorite Things 4 yıl üst üste; DTC; Shark Tank; bambu lüks segment"),

        ("Lahgo", "lahgo.com", "Restoratif Uyku Giyim",
         "Supportive uyku sutyeni; restoratif pijama; organik Pima; wellness odaklı",
         "2019'da kuruldu; uyku bilimi ile tasarım; DTC Shopify; Instagram; kadın wellness markası"),

        ("The Great Eros", "thegreateros.com", "İpek Loungewear",
         "Saf ipek iç çamaşırı ve loungewear; minimal tasarım; lüks ev giyim; zamansız estetik",
         "2018 sonrası DTC büyüme; Brooklyn kökenli; Instagram ve editorial; niş lüks segment"),

        ("Smash + Tess", "smashtess.com", "Romper Tulum Ev Giyim",
         "Yetişkin romper; yumuşak modal; cepli tulum; rahat ev giyim",
         "2018'de kuruldu; Kanada kökenli; Instagram viral; anne loungewear; DTC Shopify; eşleşen aile seti"),

        ("Softies", "softiesonline.com", "Marshmallow Ev Giyim",
         "Ultra yumuşak polar kumaş; sabahlık; ev pantolonu; kışlık loungewear",
         "2018 sonrası DTC genişleme; QVC ve online; yumuşaklık odaklı marka; uygun fiyatlı konfor"),

        ("Stars Above", "target.com/brand/stars-above", "Target Uyku Giyim",
         "Uygun fiyatlı pijama; termal seçenekler; modern tasarım; geniş beden aralığı",
         "2019'da Target'ın DTC markası olarak lansmanlandı; sosyal medya pazarlama; e-ticaret odaklı"),

        ("Cou Cou Intimates", "coucouintimates.com", "Organik Pamuk İç Giyim",
         "%100 organik pamuk; minimal tasarım; şeffaf üretim; sürdürülebilir iç giyim",
         "2019'da kuruldu; Instagram ve TikTok viral; DTC; Gen-Z hedef kitle; sürdürülebilirlik mesajı"),

        ("À La Bonne Franquette", "bonnefranquette.co", "Fransız Tarzı Loungewear",
         "Fransız esintili ev giyim; çizgili pijama; rahat şıklık; doğal kumaşlar",
         "2020'de kuruldu; Instagram estetik; DTC niş marka; Fransız yaşam tarzı konsepti"),

        ("Petite Plume", "petiteplume.com", "Klasik Pijama Seti",
         "Pima pamuk; düğmeli klasik pijama; monogram seçeneği; aile eşleşme",
         "2018'de kuruldu; preppy pijama trendi; Instagram ve influencer; DTC Shopify; hediye pazarı"),

        ("Sant and Abel", "santandabel.com", "Erkek & Kadın Pijama",
         "Liberty kumaş; ekose ve çiçekli desenler; Avustralya tarzı; unisex seçenekler",
         "2018 sonrası DTC genişleme; Avustralya kökenli; Instagram ve Facebook; premium pijama markası"),

        ("Lusomé", "lusomedesigns.com", "Gece Terlemesi Pijama",
         "Nem yönetimi teknolojisi; menopoz dostu; Xirotex kumaş; fonksiyonel uyku giyim",
         "2019'da DTC genişleme; Kanada kökenli; Amazon ve Shopify; kadın sağlığı odaklı niş"),

        ("Dagsmejan", "dagsmejan.com", "İsviçre Uyku Teknoloji Giyim",
         "Nattwell kumaş; termal düzenleme; nem yönetimi; bilimsel olarak test edilmiş",
         "2018'de kuruldu; İsviçre kökenli; Kickstarter başarısı; DTC; uyku bilimi odaklı pijama"),

        ("Jambys", "jambys.com", "Dışarı Çıkmaz Şort",
         "Ev şortu; cepli; iç çamaşırsız giyilebilir; modal kumaş; unisex",
         "2019'da kuruldu; viral TikTok; DTC Shopify; pandemi döneminde patlama; ev giyim niş"),

        ("The Cloud Set", "thecloudset.com", "Bulut Yumuşaklık Loungewear",
         "Ultra yumuşak polar set; crop hoodie ve jogger; pastel renkler; TikTok estetik",
         "2020'de kuruldu; TikTok viral; Gen-Z DTC; Shopify; influencer kampanyaları"),

        ("Little Bipsy", "littlebipsy.com", "Modern Bebek & Aile Loungewear",
         "Eşleşen aile seti; modern renkler; yumuşak kumaş; minimalist tasarım",
         "2019'da kuruldu; Instagram mama influencer; DTC Shopify; aile eşleşme trendi; hızlı büyüme"),

        ("Saltair", "saltair.com", "Vücut Bakım & Loungewear Yaşam",
         "Vücut yıkama; serum losyon; aromaterapi; loungewear yaşam tarzı",
         "2021'de kuruldu; Iskra Lawrence destekli; DTC ve Target; TikTok viral; kişisel bakım + ev hissi"),

        ("Roller Rabbit", "rollerrabbit.com", "Block Print Pijama",
         "El baskısı Hint kumaş; pijama seti; kaftan; egzotik loungewear",
         "2018 sonrası DTC genişleme; Instagram estetik; resort wear x loungewear; Shopify; premium"),

        ("Only Hearts", "onlyhearts.com", "İç Giyim Loungewear",
         "Dantel iç giyim; organik pamuk; rahat bralette; loungewear geçişi",
         "2018 sonrası DTC dijital dönüşüm; NYC kökenli; Instagram; vintage estetik; online-first yeniden yapılanma"),

        ("Sleeper", "the-sleeper.com", "Parti Pijama & Feather",
         "Tüylü pijama; linen loungewear; Instagram-worthy tasarım; Ukrayna yapımı",
         "2019'da viral oldu; Ukrayna kökenli; Instagram fenomen; DTC; lüks pijama-as-fashion konsepti"),

        ("Entireworld", "theentireworld.com", "Renkli Sweatsuit",
         "Tip A sweatpant; organik pamuk; gökkuşağı renk seçenekleri; unisex",
         "2018'de kuruldu; Scott Sternberg (Band of Outsiders); DTC; pandemi viral; renkli loungewear öncüsü"),

        ("Pangaia", "thepangaia.com", "Bilim Bazlı Loungewear",
         "Biyo-boyalı hoodie; geri dönüştürülmüş pamuk; FLWRDWN teknoloji; sürdürülebilir",
         "2019'da kuruldu; bilim x moda; DTC; Instagram viral; $70 hoodie ikonu; sürdürülebilirlik öncüsü"),

        ("Summersalt", "summersalt.com", "Seyahat & Loungewear",
         "Cloud 9 pijama; seyahat dostu kumaş; kompakt katlama; çok amaçlı tasarım",
         "2018 sonrası loungewear genişlemesi; DTC Shopify; Instagram; seyahat+ev giyim kavramı"),

        ("Cozywear Co", "cozywearco.com", "Sherpa Loungewear Set",
         "Sherpa hoodie set; kışlık ev giyim; oversize fit; TikTok trendi",
         "2020'de kuruldu; TikTok ve Instagram; DTC Shopify; Gen-Z hedef kitle; viral cozy estetik"),

        ("Girlfriend Collective", "girlfriend.com", "Sürdürülebilir Loungewear",
         "Geri dönüştürülmüş PET kumaş; LITE koleksiyon; şeffaf üretim; inclusive boyutlandırma",
         "2018 sonrası loungewear genişleme; DTC; Instagram viral; sürdürülebilir activewear-to-loungewear"),

    ],

    # ─────────────────────────────────────────────────────────────────
    # 3. TERLİK & EV AYAKKABISI  (Slippers & House Shoes)
    # ─────────────────────────────────────────────────────────────────
    "Terlik & Ev Ayakkabısı": [
        ("Comfy", "comfy.com", "Oversize Hayvan Terlik",
         "Devasa hayvan figürlü terlik; memory foam; eğlenceli tasarım; viral TikTok ürünü",
         "2019'da kuruldu; TikTok ve Instagram viral; DTC Shopify; lisanslı tasarımlar; hediye pazarı"),

        ("Bombas Slippers", "bombas.com", "Gripper Terlik",
         "Merino yünü; kaymaz taban; ev çorabı-terlik arası; sıcak tutma teknolojisi",
         "2019'da terlik lansmanı; DTC; bir al bir bağışla modeli; Shark Tank markası; sosyal medya"),

        ("Mahabis", "mahabis.com", "Modüler Tasarım Terlik",
         "Değiştirilebilir taban; iç mekan+dış mekan; yün üst; Skandinav tasarım",
         "2018 sonrası DTC genişleme; Avrupa viral; Instagram; modüler terlik konsepti öncüsü"),

        ("Naadam Slippers", "naadam.co", "Kaşmir Terlik",
         "Moğol kaşmir; lüks ev terliği; sürdürülebilir kaynak; el yapımı kalite",
         "2019'da terlik lansmanı; DTC; Instagram influencer; lüks segment; sürdürülebilir kaşmir"),

        ("Dearfoams", "dearfoams.com", "Köpük Teknoloji Terlik",
         "DF Adapt memory foam; Fresh Feel anti-mikrobiyel; makinede yıkanabilir",
         "2018 sonrası DTC dijital dönüşüm; Amazon bestseller; TikTok viral; uygun fiyatlı konfor"),

        ("Olukai Slippers", "olukai.com", "Hawaii Tarzı Ev Terliği",
         "Premium deri; anatomik ayak yatağı; Hawaii ilhamlı tasarım; iç-dış kullanım",
         "2018 sonrası DTC genişleme; Hawaii kültürü; Instagram; premium terlik segmenti"),

        ("Cobian", "cobian.com", "Surf & Ev Terliği",
         "Draino terlik; su geçirmez; hafif EVA; plaj-ev geçişli tasarım",
         "2018 sonrası DTC online büyüme; Kaliforniya surf kültürü; Amazon ve Shopify; outdoor-indoor terlik"),

        ("Toesox", "toesox.com", "Parmaklı Ev Terliği",
         "Beş parmaklı terlik; kaymaz taban; yoga ve Pilates uyumlu; organik pamuk",
         "2018 sonrası DTC genişleme; yoga topluluğu; Amazon; niş wellness terlik; Instagram fitness"),

        ("Cushionaire", "cushionaire.com", "Mantar Tabanlı Terlik",
         "Cork footbed; ayarlanabilir kayış; hafif; Birkenstock alternatifi",
         "2019'da kuruldu; Amazon-native marka; TikTok viral; uygun fiyatlı mantar terlik; $30M+ Amazon satışı"),

        ("Halflinger", "haflinger.com", "Yün Keçe Terlik",
         "Doğal yün keçe; el yapımı; Alman mühendisliği; anatomik taban",
         "2018 sonrası DTC online büyüme; Almanya kökenli; Amazon genişleme; premium yün terlik"),

        ("Kyrgies", "kyrgies.com", "El Yapımı Keçe Terlik",
         "Kırgız el yapımı; doğal yün keçe; deri taban; adil ticaret",
         "2019'da kuruldu; sosyal girişim; DTC Shopify; Instagram; fair trade terlik markası"),

        ("Baabuk", "baabuk.com", "İsviçre Yün Terlik",
         "Merino yünü; dış mekan tabanlı; GUS ev terliği; sürdürülebilir üretim",
         "2018'de kuruldu; İsviçre tasarımı; Kickstarter başarısı; DTC; doğal yün terlik"),

        ("Nootkas", "nootkas.com", "Merino Yün Terlik",
         "Felted merino; muleton astar; deri taban; Pacific Northwest ilhamı",
         "2018'de kuruldu; el yapımı ABD üretimi; DTC Shopify; Etsy genişleme; yerel üretim"),

        ("Subu", "subujapan.com", "Japon Kış Terliği",
         "SUBU sandal; -4°C'ye dayanıklı; Japon minimalizm; iç-dış kullanım",
         "2019'da global lansman; Japon kökenli; Instagram viral; DTC; kış sandalet konsepti"),

        ("Oomshoes", "oomshoes.com", "Ortopedik Ev Terliği",
         "OOfoam teknoloji; plantar fasiit desteği; ortopedik terlik; ultra hafif",
         "2019'da ev terliği lansmanı; Amazon ve DTC; podiatrist önerili; recovery terlik"),

        ("PR Soles", "prsoles.com", "Recovery Terlik",
         "ACUPOINT masaj noktaları; sporcu recovery; hafif; NCAA lisanslı",
         "2019'da kuruldu; sporcu topluluğu; Amazon ve DTC; recovery wellness; Instagram fitness"),

        ("Grateful Dead x Chinatown Market", "chinatownmarket.com", "Koleksiyon Terlik",
         "Grateful Dead ayı terlik; fuzzy kürk; koleksiyon parçası; streetwear x ev giyim",
         "2019'da viral lansman; Nike SB ilhamı; DTC drop model; Instagram ve StockX; streetwear terlik"),

        ("Arcopedico", "arcopedicousa.com", "Portekiz Ortopedik Terlik",
         "Light Line teknoloji; esnek taban; anatomik destek; vegan seçenekler",
         "2018 sonrası DTC ABD genişleme; Portekiz üretimi; Amazon; ortopedik niş; online büyüme"),

        ("Freedom Moses", "freedommoses.com", "Tasarım Slayt Terlik",
         "PVC-free; vegan; metalik ve desen seçenekleri; İsrail tasarımı; hafif",
         "2019'da viral oldu; Instagram ve TikTok; DTC global; eşleşen aile seti; yaz terlik trendi"),

        ("Cougar Pillow", "cougarshoes.com", "Yastık Tabanlı Terlik",
         "Pillow slide teknoloji; EVA köpük; su geçirmez; iç-dış mekan",
         "2020'de lansmanlandı; TikTok pillow slide trendi; DTC ve perakende; Kanada kökenli"),

        ("Roam", "roamfootwear.com", "Özelleştirilebilir Slayt Terlik",
         "Custom renk seçimi; sürdürülebilir malzeme; unisex; köpük taban",
         "2020'de kuruldu; TikTok ve Instagram; DTC Shopify; kişiselleştirme trendi; Gen-Z hedef kitle"),

        ("Lusso Cloud", "lussocloud.com", "Lüks Bulut Terlik",
         "Puff slide; ultra hafif köpük; şeker renkleri; oversize taban",
         "2020'de kuruldu; TikTok viral; DTC; pillow slide trendi; premium versiyonu"),

        ("Cloud Slides by Pillow", "pillowslides.com", "Yastık Slayt Terlik",
         "EVA kalın taban; masaj dokusu; 4.5cm platform; su geçirmez",
         "2020'de kuruldu; TikTok mega viral; Amazon bestseller; DTC; milyonlarca satış; uygun fiyat"),

        ("Hoka Ora Recovery", "hoka.com", "Spor Recovery Terlik",
         "Oversize köpük; meta-rocker taban; sporcu recovery; maksimalist konfor",
         "2019'da recovery slide lansmanı; DTC ve perakende; TikTok ve Instagram fitness; koşucu topluluğu"),

        ("Kane Footwear", "kanefootwear.com", "Sürdürülebilir Recovery Terlik",
         "Şeker kamışı bazlı EVA; karbon-negatif; recovery slide; hafif tasarım",
         "2021'de kuruldu; sürdürülebilir sporcu terliği; DTC; Instagram fitness; çevre dostu recovery"),

        ("Nuubu Slides", "nuubu.com", "Refleksoloji Terlik",
         "Akupressür noktaları; masaj tabanlı; sağlık odaklı terlik; detox iddiası",
         "2020'de kuruldu; Facebook ve Instagram Ads viral; DTC; sağlık wellness niş; global e-ticaret"),

        ("Archies Footwear", "archiesfootwear.com.au", "Arch Support Flip Flop",
         "Podiatrist tasarımı; 2.2cm kemer desteği; ultra hafif; tıbbi sınıf",
         "2018'de kuruldu; Avustralya kökenli; TikTok viral; Amazon bestseller; ortopedik parmak arası"),

        ("OOFOS", "oofos.com", "OOfoam Recovery Terlik",
         "OOfoam %37 darbe emilimi; biomechanical olarak test edilmiş; spor sonrası recovery",
         "2018 sonrası viral DTC büyüme; podiatrist önerili; TikTok ve Instagram; spor recovery terlik"),

        ("Meetься Slides", "meethese.com", "Minimalist Ev Slayt",
         "Linen üst yüzey; bambu taban; doğal malzeme; Japon minimalizmi",
         "2020'de kuruldu; DTC Amazon; doğal malzeme terlik; minimalist tasarım; online-only"),

        ("Quince Slippers", "onequince.com", "Kaşmir Ev Terliği",
         "Moğol kaşmir; fabrikadan direkt fiyat; lüks kalite; uygun fiyat",
         "2019'da kuruldu; DTC lüks demokratikleştirme; Instagram; Everlane alternatifi; terlik koleksiyonu"),

    ],

    # ─────────────────────────────────────────────────────────────────
    # 4. UYKU TEKNOLOJİSİ  (Sleep Technology)
    # ─────────────────────────────────────────────────────────────────
    "Uyku Teknolojisi": [
        ("Eight Sleep", "eightsleep.com", "Akıllı Yatak Isıtma/Soğutma",
         "Pod yatak kapağı; çift taraflı ısı kontrolü; uyku takibi; AI iklim ayarı; app bağlantılı",
         "2018 sonrası hızlı DTC büyüme; teknoloji VC yatırımı; podcast reklamları; $500M+ değerleme; uyku optimizasyonu"),

        ("Hatch", "hatchsleep.com", "Akıllı Uyku Işığı & Ses Makinesi",
         "Hatch Restore; gün doğumu alarmı; meditasyon sesleri; uyku rutini oluşturma; app kontrolü",
         "2019'da Restore lansmanı; DTC; TikTok viral; uyku wellness trendi; $200M+ gelir"),

        ("Loftie", "byloftie.com", "Akıllı Alarm & Beyaz Gürültü",
         "Telefonsuz uyku; beyaz gürültü; uyku hikayeleri; nefes egzersizi; çift alarm",
         "2020'de kuruldu; Kickstarter başarısı; DTC; uyku hijyeni hareketi; Instagram ve podcast"),

        ("Oura Ring", "ouraring.com", "Uyku Takip Yüzüğü",
         "Parmak yüzüğü formunda uyku takibi; REM analizi; kalp hızı; vücut sıcaklığı; SpO2",
         "2018 sonrası viral büyüme; Gen 3 lansmanı; DTC; ünlü ve sporcu kullanımı; $2.5B+ değerleme"),

        ("Whoop", "whoop.com", "Uyku & Recovery Takip Bandı",
         "Bilek bandı; uyku koçu; strain skoru; HRV takibi; abonelik modeli",
         "2018 sonrası tüketici genişleme; DTC abonelik; sporcu ve CEO topluluğu; podcast viral"),

        ("Muse S", "choosemuse.com", "Meditasyon & Uyku EEG Bandı",
         "Beyin dalgası ölçümü; uyku meditasyonu; biofeedback; dijital uyku hapı",
         "2019'da Muse S lansmanı; Kanada kökenli; DTC; uyku teknolojisi niş; meditasyon topluluğu"),

        ("Dreem", "dreem.com", "EEG Uyku Kafa Bandı",
         "Klinik sınıf EEG; uyku evresi analizi; ses stimülasyonu; derin uyku artırma",
         "2019'da tüketici versiyonu; Fransa kökenli; DTC; bilimsel araştırma destekli; uyku kliniği"),

        ("Sleepme", "sleep.me", "Yatak Soğutma/Isıtma Sistemi",
         "Dock Pro; su bazlı ısı kontrolü; 55-115°F aralığı; app kontrolü; çift bölgeli",
         "2018'de ChiliSleep olarak başladı; 2022 rebrand; DTC; podcast reklamları; uyku biyohacking"),

        ("BedJet", "bedjet.com", "Klima Sistemi Yatak İçin",
         "Hava bazlı soğutma/ısıtma; çift bölge; biorhythm uyku programı; app kontrolü",
         "2018 sonrası DTC büyüme; Shark Tank; Amazon ve Shopify; uyku iklim kontrolü; niş teknoloji"),

        ("Casper Glow Light", "casper.com", "Uyku Işığı",
         "Kademeli kararan ışık; hareket sensörü; portatif; sıcak amber ton",
         "2019'da lansmanlandı; Casper'ın teknoloji genişlemesi; DTC; Instagram; uyku ritueli ürünü"),

        ("Dodow", "mydodow.com", "Uyku Metronom Işığı",
         "Nefes senkronize ışık; melatonin tetikleyici; ilaçsız uyku yardımı; kompakt",
         "2018 sonrası viral DTC büyüme; Fransa kökenli; Amazon bestseller; 1M+ kullanıcı"),

        ("Morphée", "morphee.co", "Ekransız Meditasyon Cihazı",
         "210 meditasyon; ekransız; doğa sesleri; çocuk versiyonu; portatif",
         "2019'da kuruldu; Fransa kökenli; Kickstarter rekor; DTC; telefonsuz uyku çözümü"),

        ("Withings Sleep", "withings.com", "Yatak Altı Uyku Sensörü",
         "Yatak altı mat; uyku analizi; horlama algılama; kalp hızı; akıllı ev entegrasyonu",
         "2019'da yeni versiyon; Fransa kökenli; DTC ve Amazon; sağlık ekosistemi; tıbbi sınıf veri"),

        ("Sleepace", "sleepace.com", "Uyku İzleme Yastık Sensörü",
         "RestOn uyku şeridi; gerçek zamanlı izleme; uyku raporu; akıllı alarm",
         "2018 sonrası DTC genişleme; Çin kökenli; Amazon global; uygun fiyatlı uyku teknolojisi"),

        ("Somnox", "meetsomnox.com", "Robot Uyku Yastığı",
         "Nefes simülasyonu; CO2 sensörü; bilişsel davranışçı terapi; yumuşak robot",
         "2019'da lansmanlandı; Hollanda kökenli; Kickstarter; DTC; uyku robotiği niş"),

        ("SleepScore Labs", "sleepscore.com", "Non-Contact Uyku Takibi",
         "SleepScore Max; sonar teknoloji; temassız; uyku değerlendirmesi; kişisel öneriler",
         "2018 sonrası tüketici DTC lansmanı; ResMed spin-off; bilimsel altyapı; app bazlı"),

        ("Philips SmartSleep", "philips.com", "Derin Uyku Kafa Bandı",
         "SleepMapper sensörleri; ses stimülasyonu; derin uyku artırma; klinik kanıtlı",
         "2018'de lansmanlandı; DTC ve perakende; CES ödülü; klinik uyku teknolojisi"),

        ("Hapbee", "hapbee.com", "Ultra Düşük Frekanslı Uyku Bandı",
         "ULF manyetik sinyal; melatonin simülasyonu; kafein efekti; uyku modu; app kontrolü",
         "2020'de kuruldu; crowdfunding başarısı; DTC; biyohacking topluluğu; deneysel uyku teknolojisi"),

        ("Cove", "feelcove.com", "Nöro-Uyku Bandı",
         "Mastoid kemik stimülasyonu; stres azaltma; uyku kalitesi artırma; FDA kayıtlı",
         "2019'da kuruldu; nörobilim destekli; DTC; wellness topluluğu; wearable uyku teknolojisi"),

        ("Sleep Number 360", "sleepnumber.com", "Akıllı Yatak",
         "SleepIQ teknoloji; otomatik sertlik ayarı; anti-horlama; uyku takibi",
         "2018'de 360 serisi; DTC genişleme; NFL ortaklığı; uyku kişiselleştirme; IoT yatak"),

        ("Manta Sleep", "mantasleep.com", "Modüler Uyku Maskesi",
         "100% karanlık; ayarlanabilir göz kupası; soğutma versiyonu; yan yatış uyumlu",
         "2018'de kuruldu; Kickstarter viral; DTC; Amazon bestseller; biohacker topluluğu; uyku maskesi yeniden tanımlama"),

        ("Ostrichpillow", "ostrichpillow.com", "İmmersif Uyku Yastığı",
         "Original kafa yastığı; Loop göz yastığı; Heatbag; seyahat uykusu; tasarım ödüllü",
         "2018 sonrası genişleme; İspanya kökenli; Kickstarter viral; DTC; şekerleme kültürü; ofis uykusu"),

        ("Derila", "derila.com", "Ergonomik Hafızalı Köpük Yastık",
         "Servikal destek; butterfly tasarım; memory foam; boyun ağrısı çözümü",
         "2020'de kuruldu; Facebook Ads viral; DTC global; uygun fiyatlı ergonomik yastık; online-only"),

        ("Purple Harmony", "purple.com", "GelFlex Grid Yastık",
         "Hyper-elastic polymer; hava kanalları; boyun desteği; sıcaklık nötr",
         "2019'da Harmony lansmanı; DTC; YouTube viral; yastık teknolojisi; premium segment"),

        ("Tempur-Pedic Breeze", "tempurpedic.com", "Soğutmalı Hafızalı Köpük",
         "PureCool+ teknoloji; SmartClimate örtü; 3 derece daha serin; premium segment",
         "2019'da LUXEbreeze lansmanı; DTC genişleme; uyku soğutma teknolojisi; bilimsel AR-GE"),

        ("Helix Midnight Luxe", "helixsleep.com", "Kişiselleştirilmiş Yatak",
         "Uyku testi ile eşleşme; zoned destek; GlacioTex soğutma; hibrit yatak",
         "2018 sonrası genişleme; DTC; uyku anketi bazlı kişiselleştirme; podcast reklamları"),

        ("Pluto Pillow", "plutopillow.com", "Kişiselleştirilmiş Yastık",
         "35+ kombinasyon; uyku pozisyonu analizi; kişiye özel dolgu; 3D mesh",
         "2019'da kuruldu; DTC Shopify; kişiselleştirilmiş yastık konsepti; Instagram; niş uyku teknolojisi"),

        ("Panda London", "pandalondon.com", "Bambu Hafızalı Köpük Yastık",
         "Bambu örtü; Hydro Foam; hipoalerjenik; gel-infüzyon; İngiltere tasarımı",
         "2018'de kuruldu; İngiltere kökenli; Amazon ve DTC; bambu yastık; sürdürülebilir uyku"),

        ("Coop Sleep Goods", "coopsleepgoods.com", "Ayarlanabilir Dolgu Yastık",
         "Crosscut memory foam; dolgu ekle/çıkar; yıkanabilir; GREENGUARD sertifikalı",
         "2018 sonrası DTC büyüme; Amazon bestseller #1 yastık; uyku topluluğu; kişiselleştirilebilir"),

        ("Lagoon Pillow", "lagoonpillow.com", "Uyku Quiz Bazlı Yastık",
         "13 farklı model; quiz ile eşleşme; uyku pozisyonu + vücut tipi analizi",
         "2021'de kuruldu; DTC Shopify; uyku kişiselleştirme trendi; podcast reklamları; niş"),

    ],

    # ─────────────────────────────────────────────────────────────────
    # 5. SÜRDÜRÜLEBİLİR EV ÜRÜNLERİ  (Sustainable Home Products)
    # ─────────────────────────────────────────────────────────────────
    "Sürdürülebilir Ev Ürünleri": [
        ("Blueland", "blueland.com", "Yeniden Doldurulabilir Temizlik",
         "Tablet temizlik ürünü; yeniden doldurulabilir şişe; plastik-free; çevre dostu formül",
         "2019'da kuruldu; Shark Tank yatırımı; DTC; TikTok viral; $200M+ satış; sürdürülebilir temizlik öncüsü"),

        ("Earth Breeze", "earthbreeze.com", "Çamaşır Deterjan Yaprağı",
         "Kağıt yaprak deterjan; plastik-free ambalaj; biyolojik bozunur; ultra kompakt",
         "2019'da kuruldu; DTC abonelik; TikTok ve Facebook viral; her satışta bağış; 1M+ abone"),

        ("Tru Earth", "tru.earth", "Eco-Strip Deterjan Şeridi",
         "Deterjan şeridi; sıfır plastik; HE uyumlu; vegan; kompakt ambalaj",
         "2019'da kuruldu; Kanada kökenli; DTC abonelik; sosyal medya viral; sürdürülebilir çamaşır"),

        ("Dropps", "dropps.com", "Pod Deterjan & Temizlik",
         "Bitkisel pod deterjan; kompostlanabilir ambalaj; B Corp; şeffaf içerik",
         "2018 sonrası DTC rebrand; abonelik modeli; Instagram ve Facebook; çevre dostu ev bakım"),

        ("Grove Collaborative", "grove.co", "Doğal Ev Bakım Platformu",
         "Doğal temizlik ürünleri; kişisel bakım; plastik-nötr; aylık kutu",
         "2018 sonrası hızlı DTC büyüme; $1B+ değerleme; abonelik; sürdürülebilir ev bakım marketplace"),

        ("Ethique", "ethique.com", "Katı Bar Ev & Kişisel Bakım",
         "Katı şampuan bar; katı bulaşık deterjanı; plastik-free; kompostlanabilir",
         "2018 sonrası global DTC genişleme; Yeni Zelanda kökenli; B Corp; 10M+ plastik şişe tasarrufu"),

        ("Package Free Shop", "packagefreeshop.com", "Sıfır Atık Ev Ürünleri",
         "Plastik-free her şey; bambu diş fırçası; balmumu wrap; sıfır atık başlangıç kiti",
         "2018 sonrası DTC genişleme; Lauren Singer markası; Instagram; sıfır atık yaşam topluluğu"),

        ("Marley's Monsters", "marleysmonsters.com", "Yeniden Kullanılabilir Ev Tekstil",
         "UNpaper havlu; yıkanabilir makyaj pedi; bez peçete; organik pamuk",
         "2018 sonrası DTC büyüme; Etsy'den Shopify'a geçiş; sıfır atık ev; Instagram; el yapımı ABD üretimi"),

        ("If You Care", "ifyoucare.com", "Doğal Mutfak Sarf Malzemesi",
         "Geri dönüştürülmüş alüminyum folyo; kompostlanabilir pişirme kağıdı; FSC sertifikalı",
         "2018 sonrası DTC online genişleme; sürdürülebilir mutfak; Amazon ve doğal marketler; Avrupa kökenli"),

        ("Stasher", "stasherbag.com", "Silikon Gıda Saklama Poşeti",
         "Platin silikon; yeniden kullanılabilir; bulaşık makinesi uyumlu; BPA-free",
         "2018 sonrası viral DTC büyüme; B Corp; TikTok ve Instagram; plastik poşet alternatifi; $30M+ gelir"),

        ("Bee's Wrap", "beeswrap.com", "Balmumu Gıda Ambalajı",
         "Organik pamuk+balmumu; plastik streç alternatifi; kompostlanabilir; Vermont yapımı",
         "2018 sonrası DTC genişleme; B Corp; sürdürülebilir mutfak; Amazon ve Shopify; organik büyüme"),

        ("Pela", "pelacase.com", "Kompostlanabilir Telefon Kılıfı & Ev",
         "Flax Straw malzeme; kompostlanabilir; Lomi komposter; sürdürülebilir teknoloji aksesuarı",
         "2018'de kuruldu; Kanada kökenli; DTC; Instagram viral; 1M+ kılıf satışı; Lomi komposter genişleme"),

        ("Lomi", "lfromi.com", "Ev Tipi Komposter",
         "Tek tuşla kompostlama; 24 saatte gübre; koku filtresi; mutfak tezgahı boyutu",
         "2021'de lansmanlandı; Pela markasından; Indiegogo $9M+ topladı; DTC; TikTok viral; yeşil yaşam"),

        ("FoodCycler", "foodcycler.com", "Gıda Atık Geri Dönüştürücü",
         "Gıda atığını toprağa dönüştürme; koku-free; kompakt; 3 saatte işlem",
         "2019'da tüketici lansmanı; Kanada kökenli; DTC ve Amazon; sürdürülebilir mutfak teknolojisi"),

        ("Papaya Reusables", "papayareusables.com", "Yeniden Kullanılabilir Kağıt Havlu",
         "Bambu havlu; yıkanabilir; 2000 kullanım; kompostlanabilir; rulo tasarım",
         "2019'da kuruldu; DTC Shopify; Amazon genişleme; plastik-free mutfak; uygun fiyatlı sürdürülebilirlik"),

        ("Public Goods", "publicgoods.com", "Minimalist Sürdürülebilir Ev",
         "Üyelik bazlı; doğal temizlik; bambu tuvalet kağıdı; şeffaf fiyatlandırma",
         "2019'da lansmanlandı; DTC abonelik; minimalist ambalaj; sürdürülebilir ev ürünleri marketplace"),

        ("Who Gives A Crap", "whogivesacrap.org", "Bambu Tuvalet Kağıdı",
         "Bambu ve geri dönüştürülmüş; plastik-free ambalaj; karların %50'si bağış; renkli tasarım",
         "2018 sonrası viral DTC büyüme; Avustralya kökenli; sosyal medya viral; B Corp; komik marka sesi"),

        ("Cleancult", "cleancult.com", "Karton Ambalaj Temizlik",
         "Karton milk-carton ambalaj; doğa dostu içerik; yeniden doldurulabilir; hindistancevizi bazlı",
         "2018'de kuruldu; DTC abonelik; Instagram ve Facebook; plastik-free temizlik; modern tasarım"),

        ("Branch Basics", "branchbasics.com", "Konsantre Temizlik Kiti",
         "Tek konsantre; tüm ev temizliği; toksik-free; cam şişe sistemi",
         "2019'da viral DTC büyüme; Instagram ve TikTok; temiz yaşam topluluğu; anne influencer; kült marka"),

        ("Common Good", "commongoodandco.com", "İstasyonlu Temizlik Sistemi",
         "Refill istasyonu; bitkisel içerik; cam pompa şişe; toplu yeniden dolum",
         "2018 sonrası DTC genişleme; NYC kökenli; sürdürülebilir temizlik; Instagram; yerel market refill"),

        ("Supernatural", "supernatural.com", "Cam Şişe Temizlik Seti",
         "Essential oil bazlı; cam şişe+tablet; uçucu yağ kokuları; Instagram estetik",
         "2019'da kuruldu; DTC abonelik; Instagram-first; güzel tasarımlı temizlik; wellness x temizlik"),

        ("Outlines", "liveoutlines.com", "Tablet Temizlik Sistemi",
         "Çözünür tablet; alüminyum şişe; 4 farklı ürün; su ekle ve kullan",
         "2020'de kuruldu; DTC; Blueland alternatifi; sürdürülebilir temizlik; Instagram ve TikTok"),

        ("ZeroWasteStore", "zerowastestore.com", "Sıfır Atık Ev Mağazası",
         "Bambu mutfak seti; doğal sünger; balmumu wrap; sıfır atık başlangıç paketi",
         "2019'da kuruldu; DTC e-ticaret; sıfır atık topluluk; Instagram ve Facebook; niş marketplace"),

        ("Bite", "bitetoothpastebits.com", "Tablet Diş Macunu & Ev Bakım",
         "Diş macunu tableti; plastik-free; tüp atığı azaltma; çevre dostu banyo",
         "2018'de kuruldu; DTC abonelik; TikTok viral; plastik-free kişisel bakım; ev banyosu dönüşümü"),

        ("Net Zero Company", "netzerocompany.com", "Karbon-Nötr Ev Ürünleri",
         "Plastik-free bulaşık deterjanı; kompostlanabilir sünger; karbon ofset; doğal içerik",
         "2020'de kuruldu; DTC Shopify; sürdürülebilir ev bakım; Instagram; çevre bilinçli tüketici"),

        ("No Tox Life", "notoxlife.com", "Toksik-Free Ev Temizlik",
         "Vegan bulaşık bloğu; bitkisel temizlik; zero waste; el yapımı",
         "2018'de kuruldu; LA kökenli; DTC ve Etsy; sıfır atık topluluk; Instagram; doğal yaşam"),

        ("Seventh Generation", "seventhgeneration.com", "Bitkisel Ev Temizlik",
         "Bitkisel deterjan; EPA Safer Choice; USDA biobased; geri dönüştürülmüş ambalaj",
         "2018 sonrası DTC e-ticaret genişleme; B Corp; Amazon ve DTC; sürdürülebilir ev bakım lideri"),

        ("Meliora", "meliorameansbetter.com", "Sıfır Atık Deterjan",
         "Toz deterjan; sıfır atık ambalaj; 3 malzeme; B Corp; ABD yapımı",
         "2018 sonrası DTC büyüme; Chicago kökenli; sıfır atık niş; Shopify; sürdürülebilir çamaşır"),

        ("Supernatural Clean", "supernaturalclean.com", "VR Fitness & Temiz Ev",
         "Essential oil temizlik; kristal taş difüzör; ev wellness; doğal koku",
         "2020'de kuruldu; DTC; wellness x temizlik konsepti; Instagram estetik; niş sürdürülebilir"),

        ("Dirty Labs", "dirtylabs.com", "Biyo-Enzim Deterjan",
         "Biyo-enzim teknoloji; soğuk su yıkama; %80 daha az su; sürdürülebilir bilim",
         "2020'de kuruldu; DTC; TikTok viral; bilim bazlı deterjan; sürdürülebilir çamaşır yeniliği"),

    ],

    # ─────────────────────────────────────────────────────────────────
    # 6. EVCİL HAYVAN YATAK & TEKSTİL  (Pet Bed & Textile)
    # ─────────────────────────────────────────────────────────────────
    "Evcil Hayvan Yatak & Tekstil": [
        ("FunnyFuzzy", "funnyfuzzy.com", "Tasarım Köpek Yatağı",
         "Donut yatak; kalın peluş; anti-anksiyete tasarım; makinede yıkanabilir; şık renkler",
         "2020'de kuruldu; TikTok ve Instagram viral; DTC Shopify; Çin üretimi; uygun fiyatlı lüks pet"),

        ("Wild One", "wildone.com", "Modern Evcil Hayvan Aksesuar",
         "Minimalist tasarım; köpek yatağı; tasma ve oyuncak; mat renkler; DTC premium",
         "2018'de kuruldu; NYC kökenli; Instagram-first; DTC; Target genişleme; modern pet yaşam tarzı"),

        ("Fable Pets", "fablepets.com", "Mimar Tasarımı Pet Ürün",
         "The Game oyuncak; The Crate köpek kafesi; fonksiyonel estetik; premium malzeme",
         "2020'de kuruldu; tasarım odaklı; DTC Shopify; Instagram; modern ev estetiğine uyumlu pet ürün"),

        ("Paw.com", "paw.com", "Ortopedik Köpek Yatağı",
         "PupRug faux kürk; memory foam; mobilya koruyucu; ortopedik destek",
         "2018'de kuruldu; DTC; TikTok ve Instagram; PupRug viral ürün; Amazon genişleme; $50M+ gelir"),

        ("Casper Dog Mattress", "casper.com", "Köpek Yatağı",
         "Basınç dağıtıcı foam; yıkanabilir kılıf; dayanıklı; Casper teknolojisi pet versiyonu",
         "2018'de lansmanlandı; Casper markası genişlemesi; DTC; premium köpek yatağı; Instagram"),

        ("Big Barker", "bigbarker.com", "Büyük Irk Ortopedik Yatak",
         "10 yıl garanti; ortopedik foam; büyük köpekler için; klinik test edilmiş",
         "2018 sonrası DTC büyüme; Amazon ve Shopify; büyük ırk özelleşmesi; köpek sağlığı odaklı"),

        ("Snoozer Pet Products", "snoozerpetproducts.com", "Lüks Pet Yatak",
         "Cozy Cave mağara yatak; Overstuffed lüks yatak; ABD yapımı; premium kumaşlar",
         "2018 sonrası DTC genişleme; ABD üretimi; Amazon bestseller; lüks köpek yatağı segmenti"),

        ("Molly Mutt", "mollymutt.com", "Sürdürülebilir Köpek Yatağı Kılıfı",
         "Dolgu-kendin-yap kılıf; geri dönüştürülmüş malzeme; makinede yıkanabilir; şık desenler",
         "2018 sonrası DTC büyüme; sürdürülebilir pet; Amazon ve Shopify; çevre dostu köpek yatağı"),

        ("Bark Box Bed", "barkshop.com", "Dayanıklı Köpek Yatağı",
         "Yıkılmaz tasarım; çiğneme dayanıklı; eğlenceli şekiller; BarkBox ekosistemi",
         "2019'da yatak lansmanı; DTC abonelik markası; TikTok ve Instagram; oyuncak+yatak kombo"),

        ("Maxbone", "maxbone.com", "Premium Modern Pet Yatak",
         "Ortopedik köpük; makinede yıkanabilir; minimalist tasarım; premium kumaş",
         "2018'de kuruldu; LA kökenli; DTC; Instagram influencer; modern pet yaşam tarzı markası"),

        ("P.L.A.Y.", "petplay.com", "Eco-Friendly Pet Yatak",
         "Geri dönüştürülmüş PET dolgu; makinede yıkanabilir; dünya desenleri; B Corp",
         "2018 sonrası DTC genişleme; sürdürülebilir pet; Instagram; B Corp sertifikalı; eco pet yatak"),

        ("Fi Collar", "tryfi.com", "Akıllı Köpek Tasması & Yatak Takibi",
         "GPS takip; aktivite monitörü; uyku takibi; kayıp köpek bulma; LTE bağlantı",
         "2019'da lansmanlandı; DTC abonelik; köpek wellness teknolojisi; Instagram; pet-tech"),

        ("Diggs", "diggs.pet", "Modern Köpek Kafesi & Yatak",
         "Revol katlanır kafes; Snooz yatak pedi; güvenlik odaklı; modern tasarım",
         "2018'de kuruldu; DTC; bebek güvenliği standartlarında pet ürün; Instagram; premium segment"),

        ("Pet Fusion", "petfusion.com", "Memory Foam Pet Yatak",
         "CertiPUR-US foam; su geçirmez astar; kaymaz taban; yükseltilmiş yatak",
         "2018 sonrası DTC büyüme; Amazon bestseller; uygun fiyatlı kalite; köpek ve kedi; 50K+ yorum"),

        ("Zee.Dog", "zee.dog", "Brezilya Tasarım Pet Aksesuar",
         "Neopren tasma; desenli yatak; köpek bandana; renkli koleksiyonlar",
         "2018 sonrası global DTC genişleme; Brezilya kökenli; Instagram; streetwear pet markası"),

        ("The Foggy Dog", "thefoggydog.com", "El Yapımı Pet Yatak & Bandana",
         "El yapımı ABD üretimi; botanik desenler; lüks köpek yatağı; seyahat pedi",
         "2019'da kuruldu; Etsy'den DTC'ye; Instagram; el yapımı pet tekstil; niş lüks"),

        ("Laylo Pets", "laylopets.com", "Tasarım Köpek Yatağı",
         "Modüler yatak; değiştirilebilir kılıf; su geçirmez liner; modern desenler",
         "2020'de kuruldu; DTC Shopify; Instagram; ev dekoruyla uyumlu pet yatak; tasarım odaklı"),

        ("Harry Barker", "harrybarker.com", "Eco-Chic Pet Yatak",
         "Geri dönüştürülmüş pamuk; ABD yapımı; preppy tasarım; monogram seçeneği",
         "2018 sonrası DTC genişleme; sürdürülebilir pet; Shopify ve Amazon; klasik Amerikan pet markası"),

        ("K9 Ballistics", "k9ballistics.com", "Çiğneme Dayanıklı Yatak",
         "Chew proof kumaş; TUFF yatak; ortopedik; garanti programı; büyük ırk odaklı",
         "2018 sonrası DTC büyüme; dayanıklılık niş; Amazon ve Shopify; köpek yatağı garantisi"),

        ("Waggo", "waggo.com", "Seramik & Tasarım Pet Ürün",
         "Seramik mama kabı; tasarım yatak; renkli koleksiyon; hediye seti",
         "2019'da kuruldu; tasarım odaklı pet markası; DTC Shopify; Instagram; modern pet aksesuar"),

        ("Canada Pooch", "canadapooch.com", "Köpek Giyim & Yatak",
         "Kış montu; yağmurluk; köpek yatağı; Kanada soğuğuna dayanıklı; şık tasarım",
         "2018 sonrası DTC genişleme; Kanada kökenli; Instagram; fonksiyonel köpek giyim; kışlık pet tekstil"),

        ("MiaCara", "miacara.com", "Lüks Alman Pet Mobilya",
         "Tasarım köpek yatağı; kedi kulesi; premium deri; Alman mühendisliği",
         "2018 sonrası DTC global genişleme; Almanya kökenli; Instagram; ultra premium pet mobilya"),

        ("Animals Matter", "animalsmatter.com", "Lüks Pet Yatak & Battaniye",
         "Faux kürk yatak; lüks battaniye; ortopedik; ABD el yapımı; kişiselleştirme",
         "2018 sonrası DTC büyüme; lüks pet segment; Instagram; ünlü evcil hayvan sahipleri; premium niş"),

        ("Armarkat", "armarkat.com", "Kedi Ağacı & Pet Yatak",
         "Faux kürk kedi yatağı; kedi ağacı; makinede yıkanabilir; uygun fiyatlı",
         "2018 sonrası Amazon DTC büyüme; uygun fiyatlı pet mobilya; bestseller; online-first marka"),

        ("Tall Tails", "talltailsdog.com", "Köpek Battaniye & Yatak",
         "Sherpa battaniye; su geçirmez yatak; seyahat pedi; dayanıklı kumaş",
         "2019'da genişleme; DTC ve perakende; köpek tekstil uzmanı; Instagram; uygun fiyatlı kalite"),

        ("Teddy Maximus", "teddymaximus.com", "Lüks Köpek Tasma & Yatak",
         "El yapımı deri tasma; premium yatak; İngiliz tasarımı; pastel renkler",
         "2019'da kuruldu; İngiltere kökenli; Instagram; lüks pet aksesuar; DTC Shopify; niş premium"),

        ("BuddyRest", "buddyrest.com", "Ortopedik Pet Yatak",
         "TruCool memory foam; antimikrobiyel; ABD yapımı; büyük ırk desteği",
         "2018'de kuruldu; DTC; Amazon genişleme; bilimsel destekli pet yatak; köpek sağlığı odaklı"),

        ("Pendleton Pet", "pendleton-usa.com", "Miras Desenlı Pet Yatak",
         "Pendleton yün desenleri; klasik Amerikan miras; premium pet yatak; hediye",
         "2019'da pet koleksiyon lansmanı; DTC genişleme; Instagram; miras marka x pet tekstil"),

        ("Bowl & Bone Republic", "bowlandbone.com", "Polonya Tasarım Pet Yatak",
         "Modern oval yatak; İskandinav tasarım; yıkanabilir kılıf; renkli koleksiyon",
         "2018 sonrası DTC genişleme; Polonya kökenli; Instagram; Avrupa tasarım pet mobilya"),

        ("Barkbox Super Chewer Bed", "barkbox.com", "Ultra Dayanıklı Yatak",
         "Ballistic nylon; kevlar dikişler; çiğneme garantisi; eğlenceli tasarım",
         "2020'de lansmanlandı; BarkBox ekosistemi; DTC abonelik; çiğneme dayanıklı yatak; sosyal medya"),

    ],

    # ─────────────────────────────────────────────────────────────────
    # 7. OUTDOOR & PİKNİK TEKSTİL  (Outdoor & Picnic Textile)
    # ─────────────────────────────────────────────────────────────────
    "Outdoor & Piknik Tekstil": [
        ("Rumpl", "rumpl.com", "Teknik Outdoor Battaniye",
         "Geri dönüştürülmüş sentetik izolasyon; su geçirmez; paketlenebilir; NanoLoft dolgu",
         "2018 sonrası hızlı DTC büyüme; Portland kökenli; Instagram ve TikTok; kamp battaniyesi kategorisi yaratıcı"),

        ("Nomadix", "nomadix.co", "Geri Dönüştürülmüş Outdoor Havlu",
         "%100 geri dönüştürülmüş; çok amaçlı havlu/battaniye; kum tutmaz; hızlı kuruyan",
         "2018'de kuruldu; B Corp; DTC Shopify; sürdürülebilir outdoor; plaj ve kamp; Instagram"),

        ("Coalatree", "coalatree.com", "Kompakt Kamp Battaniyesi",
         "Packable battaniye; hamak; poncho; geri dönüştürülmüş malzeme; çok fonksiyonlu",
         "2018 sonrası DTC büyüme; Kickstarter başarısı; outdoor macera markası; Instagram"),

        ("Kelty", "kelty.com", "Kamp Örtü & Piknik Battaniye",
         "Lowdown piknik örtüsü; su geçirmez taban; katlanabilir; uygun fiyat",
         "2019 sonrası DTC e-ticaret genişleme; kamp ekipmanı markası; Amazon ve Shopify; outdoor yaşam"),

        ("Voited", "vofrted.com", "Giyilebilir Kamp Battaniyesi",
         "CloudTouch battaniye; poncho dönüşebilir; REPREVE geri dönüştürülmüş; çok fonksiyonlu",
         "2019'da kuruldu; Hollanda kökenli; Kickstarter; DTC; outdoor x sürdürülebilirlik; Instagram"),

        ("Kachula by Coalatree", "coalatree.com", "Adventure Battaniye",
         "Snap-together birleşebilir; su dayanıklı; sleeping bag dönüşüm; poncho modu",
         "2019'da viral Kickstarter; DTC; çok fonksiyonlu outdoor battaniye; kamp topluluğu"),

        ("Tesalate", "tesalate.com", "Kum Tutmaz Plaj Havlusu",
         "AbsorbLite teknoloji; kum yapışmaz; hızlı kuruyan; çift kişilik seçenek",
         "2018'de kuruldu; Avustralya kökenli; DTC global; Instagram viral; plaj havlusu yeniliği"),

        ("Sand Cloud", "sandcloud.com", "Türk Plaj Havlusu",
         "Türk pamuğu; çok amaçlı; her satışta deniz koruma bağışı; bohem tasarım",
         "2018 sonrası viral DTC büyüme; Shark Tank; Instagram; sürdürülebilir plaj markası; okyanus koruma"),

        ("Dock & Bay", "dockandbay.com", "Mikrofiber Hızlı Kuruyan Havlu",
         "Süet mikrofiber; kompakt paketleme; canlı renkler; plaj ve spor; seyahat dostu",
         "2018'de global DTC genişleme; İngiltere kökenli; Amazon bestseller; Instagram; plaj havlusu trendi"),

        ("Slowtide", "slowtide.com", "Sanatçı İşbirliği Plaj Havlusu",
         "Sanatçı koleksiyonları; premium pamuk; büyük boy; ev ve plaj; sürdürülebilir",
         "2018 sonrası DTC büyüme; Kaliforniya kökenli; Instagram; surf kültürü; sanat x tekstil"),

        ("Yeti Lowlands", "yeti.com", "Premium Piknik Battaniyesi",
         "Su geçirmez taban; makinede yıkanabilir; dayanıklı; kamp ve piknik; premium",
         "2019'da lansmanlandı; Yeti outdoor ekosistemi; DTC; premium piknik segment; Instagram"),

        ("Matador", "matadorup.com", "Ultralight Piknik Örtü",
         "Pocket blanket; 10x10 feet; su geçirmez; 3.5oz ultra hafif; paketlenebilir",
         "2018 sonrası DTC genişleme; ultralight outdoor markası; Kickstarter; Amazon ve Shopify"),

        ("BEARZ Outdoor", "bearzoutdoor.com", "Kompakt Piknik Battaniyesi",
         "Su geçirmez; cep boyutuna katlanır; kum çivili; plaj ve park; uygun fiyat",
         "2019'da kuruldu; Amazon-native; DTC genişleme; outdoor piknik niş; uygun fiyatlı kalite"),

        ("Wise Owl Outfitters", "wiseowloutfitters.com", "Hamak & Kamp Battaniye",
         "Hafif hamak; kamp battaniyesi; paketlenebilir; bütçe dostu outdoor",
         "2018'de kuruldu; Amazon-first strateji; kamp topluluğu; DTC; uygun fiyatlı outdoor tekstil"),

        ("Nemo Puffin", "nemoequipment.com", "İzolasyonlu Kamp Battaniye",
         "Primaloft izolasyon; su dayanıklı; 2-in-1 battaniye/uyku tulumu; premium kamp",
         "2019'da lansmanlandı; DTC outdoor; premium kamp ekipmanı; Instagram; dört mevsim kullanım"),

        ("ENO DoubleNest", "eaglesnestoutfitters.com", "Hamak & Outdoor Tekstil",
         "Paraşüt naylon hamak; hafif; kompakt; renkli kombinasyonlar; outdoor yaşam",
         "2018 sonrası DTC e-ticaret büyüme; hamak kültürü; Instagram; üniversite ve kamp; festival"),

        ("Therm-a-Rest Honcho Poncho", "thermarest.com", "Giyilebilir Kamp Battaniye",
         "Sentetik izolasyon; poncho+battaniye; kapüşonlu; kamp ve festival; çok amaçlı",
         "2019'da lansmanlandı; DTC genişleme; kamp teknolojisi; outdoor topluluğu; çok fonksiyonlu"),

        ("Pendleton Outdoor", "pendleton-usa.com", "Miras Outdoor Battaniye",
         "National Park koleksiyon; yün outdoor battaniye; piknik örtüsü; ikonik desenler",
         "2018 sonrası DTC online genişleme; Amerikan mirası; Instagram; piknik ve kamp; hediye pazarı"),

        ("CGear Sand-Free", "cgear.com", "Kum Geçirmez Piknik Örtü",
         "Patentli mesh teknoloji; kum aşağı geçer; su geçirmez; askeri teknoloji adaptasyonu",
         "2018 sonrası tüketici DTC lansmanı; Avustralya askeri teknoloji; plaj örtüsü; Amazon"),

        ("Original Puffy Blanket by Rumpl", "rumpl.com", "İkonik Puffy Battaniye",
         "Uyku tulumu teknolojisi battaniyede; DWR kaplama; 20D ripstop; makinede yıkanabilir",
         "2018 sonrası kült ürün; Portland kökenli; DTC; Instagram; outdoor battaniye kategorisi lideri"),

        ("PackTowl", "packtowl.com", "Ultralight Seyahat Havlusu",
         "Süper emici; antimikrobiyel; kompakt; 4 farklı seri; outdoor ve seyahat",
         "2018 sonrası DTC genişleme; MSR/Cascade Designs; Amazon bestseller; ultralight topluluğu"),

        ("Horizon Hound", "horizonhound.com", "Dayanıklı Piknik Örtüsü",
         "600D Oxford taban; su geçirmez; taşıma kayışlı; aile boyutu; dayanıklı",
         "2019'da kuruldu; Amazon-native; DTC; piknik ve park; uygun fiyatlı; outdoor aile ürünü"),

        ("ECCOSOPHY", "eccosophy.com", "Mikrofiber Plaj Battaniyesi",
         "Kum tutmaz; hızlı kuruyan; geniş boy; renkli desenler; seyahat çantası dahil",
         "2019'da kuruldu; Amazon-first DTC; plaj ve outdoor; sürdürülebilir malzeme; online-only"),

        ("CGEAR Multimats", "cgear.com", "Çok Amaçlı Outdoor Mat",
         "Kamp matı; çadır önü; su geçiren mesh; kolay temizleme; katlanabilir",
         "2019'da tüketici DTC genişleme; Avustralya kökenli; kamp ve karavan; Amazon; outdoor tekstil"),

        ("Veer Basecamp", "goveer.com", "Cruiser Wagon & Outdoor Örtü",
         "Kamp vagonu örtüsü; su geçirmez; UV koruma; outdoor aile; all-terrain",
         "2019'da kuruldu; DTC; premium outdoor aile markası; Instagram; modern aile outdoor yaşamı"),

        ("Meadow Mat", "meadowmat.com", "Ekstra Büyük Piknik Örtüsü",
         "10x10 feet; su geçirmez taban; makinede yıkanabilir; aile boyutu; kompakt",
         "2020'de kuruldu; DTC Shopify; piknik niş; Instagram; aile outdoor ürünü"),

        ("Sackcloth & Ashes", "sackclothandashes.com", "Sosyal Sorumluluk Battaniye",
         "Polar battaniye; bir al bir bağışla; evsiz barınakları desteği; ABD üretimi",
         "2018 sonrası DTC büyüme; sosyal girişim; Instagram; outdoor ve ev battaniye; B Corp"),

        ("Blanket Tek", "blankettek.com", "Stadyum & Outdoor Battaniye",
         "Stadyum battaniyesi; su geçirmez taban; polar üst; rüzgar geçirmez; portatif",
         "2019'da kuruldu; Amazon DTC; outdoor etkinlik; spor ve konser; fonksiyonel battaniye"),

        ("Lagu", "lfragu.com", "Kum Tutmaz Plaj Örtüsü",
         "Özel doku teknoloji; kum yapışmaz; antibakteriyel; hafif; renkli desenler",
         "2019'da kuruldu; Avustralya kökenli; DTC global; plaj tekstil yeniliği; Instagram"),

        ("Wise Owl Hammock", "wiseowloutfitters.com", "Bütçe Dostu Hamak",
         "Paket naylon; çift kişilik; ağaç kayışlı; ultra hafif; 30 renk seçeneği",
         "2018'de kuruldu; Amazon bestseller; DTC; kamp ve backyard; uygun fiyatlı outdoor tekstil"),

    ],

}
