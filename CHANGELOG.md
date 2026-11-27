# Changelog

## 2026-11-27 – napi ellenőrzés: vagyonosodási törvény kihirdetése, Hankó-ügy

Státuszváltozás nincs.

- **9. Vagyonnyilatkozatok** – a november 24-én újra megszavazott vagyonosodási törvényt 2026. évi LXXIV. törvényként november 25-én kihirdették (Magyar Közlöny 180. szám); november 26-án hatályba lépett, a nem nyilvános vagyonnyilatkozatok megismeréséről szóló 4. § december 1-jén. Új forrás: NJT és Magyar Közlöny. ↗️ marad: a NAV-vizsgálatok még nem indultak el, eredmény nincs.
- **11. Elszámoltatás** – a Kecskeméti Törvényszék november 11-én másodfokon elutasította Hankó Balázs és Varga-Bajusz Veronika fellebbezését, letartóztatásuk februárig tart; vádemelés nincs. ⚠️ marad.

## 2026-11-24 – CONTRIBUTING.md frissítése

Az „Adatmódosítás pull requestben” lépései igazodnak a jelenlegi szabályokhoz és ellenőrzésekhez: a darabszámokat a README státusztáblájában is frissíteni kell; a `Vállalt határidő` kötelező és csak programbeli határidő lehet; minden módosításhoz CHANGELOG-bejegyzés kell; a színezést és a rögzítést a `sync_xlsx.ps1` intézi; a validátor által ellenőrzött elemek felsorolva.

## 2026-11-24 – Excel: színezés, rögzített oszlopok, Forrásjegyzék és Projekt lap

Tartalmi (státusz-) változás nincs.

- **Színezés státusz szerint:** a *TISZA mérleg* soraiban a 39–56. tételek színezetlenek voltak, a 24. és 25. tétel pedig még a régi ❌-es pirosat viselte. Mostantól minden sor a státusza színét kapja az *Összefoglaló* darabszám-táblázata alapján; a jelmagyarázat elcsúszott színei (🟦 sárga, ⏳ piros, ⚠️ és ❌ színtelen) is javultak. A színezést a `scripts/sync_xlsx.ps1` minden futáskor beállítja, így státuszváltáskor automatikusan követi.
- **Rögzítés:** a *TISZA mérleg* lapon a fejléc mellett az A–B oszlop (ID, Terület) is rögzített, jobbra görgetéskor látható marad.
- **Forrásjegyzék:** a nyolc alapforrás (program PDF, kormany.hu, NJT / Magyar Közlöny / parlament.hu, NNGYK, KSH, 24.hu, Telex, HVG / 444 / Portfolio) linket kapott; eddig üres volt.
- **Projekt lap:** új sorok – *Tételek száma* (a szkript frissíti), *Határidők* (a szabály rövid leírása, link a README szakaszára), *Adatok* (CSV → Excel).
- **Összefoglaló:** a „2026.10.19-i frissítés…” elavult megjegyzés helyett hivatkozás a CHANGELOG-ra és a dátumoszlopokra.
- `scripts/validate.ps1`: ellenőrzi a sorok színét, a rögzítést és a *Projekt* lap tételszámát.

## 2026-11-24 – határidők minden tételnél, életciklus-szabály

Státuszváltozás nincs.

- **`Vállalt határidő` mind az 56 tételnél**, kizárólag a program PDF-jének szövege alapján: `első intézkedések` (20 tétel: 1, 3, 4, 5, 8, 11, 17, 19, 21, 24, 25, 47, 49–56), `2030-ra` (16, 18), `azonnal` (38), `2030-ig, évente` (39), a többi 32 tételnél `nincs megadva` – a program szakpolitikai fejezetei ezekhez nem adnak határidőt.
- **18. Költségvetés állapota** – a programidézet a pontos programmondatra cserélődött („2030-ra teljesítjük a maastricht-i kritériumokat: 3% alatti költségvetési hiány…”), határideje `2030-ra`.
- **README, módszertan** – új „Határidők és a tételek életciklusa” szakasz: ❌ csak a határidő után jár (`első intézkedések` és `azonnal`: a kormányzati év vége, 2027. május 12.; dátumos vállalás: a dátum után; `nincs megadva`: a ciklus vége, 2030); a ✅ tételek a listában maradnak; ciklus végi zárás 2030-ban. Ez felváltja az „első intézkedések” korábbi szabályát.
- **24. Nemzeti megbékélés** – a december 31-i felülvizsgálat megmarad, ❌ azonban az új szabály szerint legkorábban 2027. május 12. után adható.
- `scripts/validate.ps1`: a `Vállalt határidő` mező nem lehet üres.

## 2026-11-24 – 24. tétel: ❌ → ⏳, felülvizsgálat év végén

- **24. Nemzeti megbékélés** – ❌ Nem teljesült → **⏳ Nyitott, 2026. december 31-i felülvizsgálattal**. A szöveg kiegészült a gesztusokkal (a miniszterelnök május 9-i bocsánatkérése és felhívása, a parlamenti bizottsági tisztségekről kötött megállapodás, az ÁBTL bocsánatkérése Lezsák Sándortól). A program „a nemzeti megbékélés programját” vállalta; ilyen program vagy intézmény még nem jött létre, a gesztusok miatt azonban a tétel az év végéig nyitott. Ha addig sem indul el érdemi program, ❌ lesz.
- README, módszertan: az „első intézkedések” szabálya kiegészült – csak gesztusok vagy előkészítés esetén ⏳ jár, a tételben megadott felülvizsgálati dátumig.

Státuszok: ✅ 8, ↗️ 15, 🔄 8, 🟦 16, ⏳ 5 (+1), ⚠️ 4, ❌ 0 (−1).

## 2026-11-24 – 25. tétel: ❌ → ↗️

- **25. Orosz befolyás feltárása** – ❌ Nem teljesült → **↗️ Előrelépés**. Konkrét kormányzati lépések: 2026 májusában egy, szeptember 8-án tíz orosz diplomata kiutasítása; a Szijjártó Péter orosz kapcsolatairól indított vizsgálat (július); a Külügyminisztérium november 18-i belső vizsgálata a választás előtt Moszkvából állítólag hazahozatott iratok ügyében, amely Panyi Szabolcs könyvének állításai nyomán indult. A programban vállalt átfogó, nyilvános jelentés még nem készült el, ezért nem ✅. Panyi könyve újságírói munka: a tétel háttérként hivatkozik rá, a státuszt a kormányzati lépések adják.

Státuszok: ✅ 8, ↗️ 15 (+1), 🔄 8, 🟦 16, ⏳ 4, ⚠️ 4, ❌ 1 (−1).

## 2026-11-24 – 55. tétel: ❌ → ↗️

- **55. Önkormányzati hatáskörök és források** – ❌ Nem teljesült → **↗️ Előrelépés**. A források terén konkrét lépések történtek: a kormány júniusban átütemezte a főváros szolidaritási hozzájárulásának befizetését, októberben mintegy 120 milliárd Ft támogatást adott Budapestnek. A 2027-es költségvetési javaslat 62 milliárd Ft-tal (~15%) csökkentené a hozzájárulást, és megszüntetné az iparűzési adó növekményének elvonását; ezek az elfogadással lépnek hatályba. A hatáskörök visszaadása még nem kezdődött el, ezért nem ✅ és nem 🔄.

Státuszok: ✅ 8, ↗️ 14 (+1), 🔄 8, 🟦 16, ⏳ 4, ⚠️ 4, ❌ 2 (−1).

## 2026-11-24 – a 2026-10-19 óta nem ellenőrzött tételek

28 tétel újraellenőrizve (`Utolsó ellenőrzés`: 2026-11-24), státuszváltozás nélkül. A fejlemények a szöveg végére, a források a programlink elé kerültek.

- **2. KEKVA-k** – a kekvák állami vagyonát augusztusban visszavették; az állami vagyonról szóló törvényt az államfő az Alkotmánybírósághoz küldte (döntés december 9-ig);
- **3, 4, 5, 6** – minisztériumi átvilágítások; a Balásy-szerződések felmondása (kb. 150 Mrd Ft); 41 Mrd Ft-os megtakarítás az állami járműbeszerzésen; újabb 16 beruházás vesztette el a kiemelt státuszát (1144/2026. Korm. határozat);
- **8. Integritás Hatóság** – 2026. évi XVIII. törvény, bővített jogkörök, két új elnökhelyettes;
- **10. Magántőkealapok** – 2026. évi LXVIII. törvény, kihirdetve november 13-án;
- **11. Elszámoltatás** – Hankó Balázs volt miniszter letartóztatása (NKA-ügy); jogerős ítélet még nincs, ezért ⚠️ marad;
- **12. Kórházi fertőzések** – az adatbázis szeptember 4-én indult, egyelőre szűkített adatkörrel;
- **16. Egészségügyi reform** – várólista-program kormányhatározattal; az október 15-i indulás csúszott;
- **17. EU-politika** – a kormány 2026-ban háromszor blokkolta Ukrajna csatlakozási tárgyalásainak előrehaladását;
- **18, 29. Költségvetés** – a 2026-os módosított költségvetés hiánycélja a GDP 7,5%-a; a tízhavi hiány 5462,5 Mrd Ft (ebből 2186 Mrd Ft RRF-előfinanszírozás);
- **20, 24, 25** – nincs lezáró vizsgálat, megbékélési program, illetve nyilvános jelentés; a 25.-nél a Szijjártó-vizsgálat megerősítése;
- **27. Iskolakezdési támogatás** – a második részlet kézbesítése november 10–30.; 🔄 marad, amíg le nem zárul;
- **43. NOKS** – a kormányfő által október közepére ígért első béremelési lépcső elmaradt;
- **15, 22, 23, 33, 34, 37, 38, 39, 44, 45** – a 2027-es költségvetést még nem fogadták el (zárószavazás decemberben).

Nem ellenőrizve (jogszabállyal teljesült ✅ tételek, `Utolsó ellenőrzés` 2026-10-19 maradt): 7, 13, 14, 26, 28.

## 2026-11-24 – README és eszközök

Tartalmi (státusz-) változás nincs.

- README: a státusztábla aktuális darabszámokat kapott; a CSV az elsődleges, szerkesztett adat, az Excel ebből készül (a korábbi szöveg fordítva állította); frissítési lépések; a program „első intézkedései” és a jogi lépésre adott ✅ értelmezése a módszertanban; a programidézetek elsődleges forrása a program PDF-je;
- **A munkafüzet új neve `data/TISZA_kormany_merleg.xlsx`** (korábban `TISZA_kormany_merleg_2026-10-19.xlsx`). Egyszeri átnevezés: a dátum nélküli név frissítéskor sem változik, így a rá mutató linkek ezentúl stabilak. A régi névre mutató közvetlen linkek a `main` ágon nem működnek tovább, a régi commitokban a fájl a régi néven elérhető marad;
- `data/README.md`: a `Vállalt határidő` oszlop `első intézkedések` értéke; az új fájlnév;
- `scripts/sync_xlsx.ps1`: az *Összefoglaló* címébe és a *Projekt* munkalap „Utolsó frissítés” sorába a legkésőbbi „Utolsó ellenőrzés” dátuma kerül (eddig 2026.10.19. maradt);
- `scripts/validate.ps1`: ellenőrzi a README státusztáblájának darabszámait, az „Utolsó frissítés” dátumát és a munkafüzet frissítési dátumait.

## 2026-11-24 – a program „első intézkedések” listájának tételei

48 → **56 tétel**. A program (*A működő és emberséges Magyarország alapjai*, 10–11. o.) „Az első intézkedések azonnal érezhető változást hoznak” listájának mérlegbe még fel nem vett pontjai. A programidézetek szó szerintiek, a forrás a program PDF-je.

Új tételek:
- **49. Európai Ügyészség** (✅) – a Bizottság 2026. július 10-én jóváhagyta a csatlakozást; a hazai hivatal 2027 elején állhat fel;
- **50. Vagyonvisszaszerzési Hivatal** (↗️) – 2026. évi XXXIV. törvény, hatályos július 29-től, elnöke Unger Anna; visszaszerzett vagyonról még nincs adat;
- **51. Kormányülések és titkos döntések nyilvánossága** (🔄) – 1143/2026. (V. 14.) Korm. határozat; a nem nyilvános határozatok listája szeptember 29-én megjelent, a kormányülési összefoglalók még nem;
- **52. Gyermekvédelmi vizsgálóbizottság** (↗️) – 19/2026. (V. 27.) OGY határozat; jelentés még nincs;
- **53. MNB-alapítványok iratai** (⏳) – az iratok a vizsgálóbizottsághoz és az ügyészséghez kerültek, a nyilvánossághoz nem;
- **54. Állami vagyonkataszter** (⏳) – október 19-i bejelentés egységes nyilvános adatbázisról, elkészült kataszter nincs;
- **55. Önkormányzati hatáskörök és források** (❌) – a szolidaritási hozzájárulás marad, visszaadó törvény nincs. A 24. és 25. tétellel egységesen az „első intézkedések” közé sorolt, érdemben el nem indult vállalás ❌;
- **56. Miniszteri vétójog** (✅) – 90/2026. (V. 13.) Korm. rendelet; a programban megnevezett négy tárca kapta meg.

Bővített tételek (státuszváltozás nélkül):
- **19. Közmédia** – a programidézet kiegészült a hírsugárzás felfüggesztésével; a hírszolgáltatás 2026. július 7-én leállt, azóta csak hírsáv működik;
- **21. Oktatás** – a Klebelsberg Központ átalakítása és az igazgatói munkáltatói jogok; a programidézet kiegészült.

Státuszok: ✅ 8 (+2), ↗️ 13 (+2), 🔄 8 (+1), 🟦 16, ⏳ 4 (+2), ⚠️ 4, ❌ 3 (+1).

## 2026-11-24 – elfogadott adó- és nyugdíjcsomag, új tétel: sport-összeférhetetlenség

47 → **48 tétel**. Az Országgyűlés 2026. november 24-én elfogadta az adócsomagot és a nyugdíjtörvényt (T/932.), de egyiket sem hirdették ki, és 2027 előtt nem hatályosak, ezért a 🟦 tételek 🟦-k maradnak. Kivétel a 42. tétel, mert a módszertan szerint a programtól kedvezőtlenül eltérő intézkedés elfogadáskor ⚠️ lesz.

- **42. Anyák SZJA-mentessége** – 🟦 → ⚠️: elfogadva (135–47–7); a kétgyermekes anyák mentessége a 40. életév évéig jár, havi 1 millió Ft-os korláttal, és elmarad a 2025-ben törvénybe írt kiterjesztés;
- **30, 31** – családi pótlék és adójóváírás elfogadva (135–47–7); egygyermekes családi pótlék 2027-től 24 400 Ft/hó; a jóváírás kifutási határát a források eltérően közlik (680 vagy 694 ezer Ft);
- **32, 35, 36, 46** – nyugdíjtörvény elfogadva (189–1); a folyósított kisnyugdíjak felzárkóztatásáról szóló kormányrendelet még nem jelent meg;
- **40. Vagyonadó** (134–49–7), **41. KATA** (135–48–6) – elfogadva;
- **1. EU-pénzek** – az Európai Bizottság 2026. november 20-án előzetesen, részben jóváhagyta a ~10 Mrd eurós RRF-kérelmet (126/127 mérföldkő); 🔄 marad;
- **9. Vagyonnyilatkozatok** – a vagyonosodási törvényt a köztársasági elnök visszaküldte, a parlament november 24-én változatlanul újra elfogadta (139–49); ↗️ marad;
- **19. Közmédia** – a Független Közmédia Testület tagjait november 24-én nem sikerült megválasztani; 🔄 marad;
- **47. Ügynökakták** – a nyitottaktak.hu november 4-én elindult; a B-dossziék 2027. február 25-ig követik, ezért ↗️ marad;
- **új: 48. Sport és politika összeférhetetlensége** (↗️) – elfogadva (140–49), még nincs kihirdetve. A konkrét szabály kormányzati bejelentésből ered, nem szó szerinti programpont;
- Forrásjegyzék: +3 forrás (kormany.hu, Infostart, Euronews).

Státuszok: ✅ 6, ↗️ 11 (+1), 🔄 7, 🟦 16 (−1), ⏳ 2, ⚠️ 4 (+1), ❌ 2.

## 2026-11-02 – benyújtott 2027-es törvényjavaslatok, új tétel: ügynökakták

46 → **47 tétel**. A meglévő tételek státusza nem változott: a törvényjavaslatokat benyújtották, de még nem fogadták el, ezért a 🟦 tételek 🟦-k maradnak.

- 2027-es költségvetés: T/950., benyújtva 2026. október 30-án; az adócsomag (T/927–T/931) és a nyugdíjtörvény-javaslat (T/932) 2026. október 27-én;
- **30. Családi pótlék** – T/929: az emelt összegű pótlék is duplázódik (pl. 23 300 → 46 600 Ft), már a januárban folyósított decemberi pótléktól;
- **31. SZJA / alacsony keresetűek** – a 9% új, 6%-os bérjóváírással jönne ki, az általános kulcs 15% marad;
- **32. Minimálnyugdíj** – T/932: a 120 000 Ft-os minimum csak a 2027-től megállapított nyugdíjakra; a folyósított nyugdíjak emeléséhez kormányrendelet kell, még nem jelent meg;
- **35, 36, 46** – benyújtva a T/932.-ben;
- **39. Egészségügyi többletforrás** – 500 → **510 milliárd Ft** a T/950.-ben;
- **40. Vagyonadó** – T/930: menetrend, 200 milliárd Ft várt bevétel, amelyet a Költségvetési Tanács bizonytalannak tart;
- **41. KATA** – T/928: nem főfoglalkozásban 50 ezer Ft/hó; a 75 ezer Ft a régi katában maradóké (2027), a régi rendszer 2029 elején megszűnik;
- **42. Anyák SZJA-mentessége** – T/929: a kétgyermekes anyák mentessége a 40. életév évéig jár (a már jogosultak egy része is elveszíti), a három- és többgyermekeseké élethosszig, havonta legfeljebb 1 millió Ft munkajövedelemig; továbbra is 🟦, elfogadáskor ⚠️;
- **új: 47. Ügynökakták** (↗️) – programvállalás („Megnyitjuk az ügynökaktákat.”), 2026. évi LI. törvény, a 2026. évi LX. törvény szerint november 4-én indul a nyilvánosságra hozatal (nyitottaktak.hu), a B-dossziék 2027. február 25-ig. A tábla csak az intézkedést követi, az iratokban szereplő személyeket nem.

## 2026-10-19 – elsődleges jogi források, új oszlopok, CSV → Excel szinkron

Elsődleges források (Nemzeti Jogszabálytár / Magyar Közlöny) a jogszabállyal teljesült ✅ tételekhez, státuszváltozás nélkül:
- **7. Szuverenitásvédelmi Hivatal** – 2026. évi XXIV. törvény (kihirdetve 2026. június 30.);
- **13. Kórházi klímák** – 1191/2026. (VI. 11.) Korm. határozat (3,67 Mrd Ft), plusz HVG és Economx; a befejezett helyszínek száma miniszteri közlésen alapul;
- **14. Vényköteles gyógyszerek** – 2026. évi XL. törvény (hatályos 2026. szeptember 1.), NAV-tájékoztató, BAMA; a 0% a humán magisztrális készítményekre is vonatkozik;
- **26. Tűzifa áfája** – 2026. évi XLVIII. törvény (elfogadva augusztus 28., hatályos szeptember 15.), Portfolio, Világgazdaság; a pelletre és brikettre is kiterjed;
- **28. Politikusi javadalmazások** – 2026. évi XXV. törvény (hatályos 2026. július 1.); a vármegyei közgyűlési elnökökre is vonatkozik.

Új oszlopok a `TISZA_merleg.csv`-ben és a munkafüzetben:
- `Vállalt határidő` – csak ahol a programidézet határidőt nevez meg (16: 2030-ra, 38: azonnal, 39: 2030-ig, évente);
- `Utolsó ellenőrzés` – minden tételnél 2026-10-19 (a snapshot dátuma).

Eszközök:
- `scripts/sync_xlsx.ps1`: a CSV-kből frissíti a munkafüzetet, így azt nem kell Excelben szerkeszteni;
- a validátor ellenőrzi az `Utolsó ellenőrzés` dátumformátumát, és elutasítja az érzékenységi címkével (MSIP) ellátott vagy titkosított munkafüzetet;
- `.gitignore`: az Excel zárolófájljai (`~$*`).

## 2026-10-19 – Excel-javítás

- Az *Összefoglaló* munkalapon az `A1:B1` egyesítés átfedte az `A1:F1` címsor-egyesítést (már az első feltöltött munkafüzetben is), ezért az Excel megnyitáskor „javította” a fájlt. Az átfedő egyesítés kikerült, a validátor mostantól ellenőrzi az átfedéseket.

## 2026-10-19 – tételösszevonások és státuszpontosítások

49 → **46 tétel**. A sorszámok átszámozódtak (régi → új): 1–2 → 1–2, 4–24 → 3–23, 27–49 → 24–46.

Összevonások (egy esemény = egy tétel):
- régi 2 + 3 → **2. KEKVA-k és kiszervezett állami vagyon** (↗️) – mindkettő ugyanarra a KEKVA-törvényre épült;
- régi 12 + 26 → **11. Elszámoltatás, korrupciós büntetőügyek** (⚠️) – azonos programidézet és forrás, eddig ⚠️ és ❌ párhuzamosan;
- régi 21 + 25 → **20. Sajtószabadság, KNBSZ / Szuvák-ügy** (⚠️) – a sajtószabadság-tétel hiányosságát eddig is csak ez az ügy indokolta.

Pontosítások:
- **15. Egészségügyi bérek** (régi 16): az +500 milliárd Ft kikerült, azt a 39. tétel követi (eddig kétszer szerepelt);
- **22. Pedagógusbérek** (régi 23): 🔄 → ⏳ – a 2027-es költségvetés a meglévő 80%-os garanciát tartja fenn, új pedagógus-béremelést nem jelentett be (Eduline, Portfolio);
- **42. Anyák SZJA-mentessége** (régi 45): ⚠️ → 🟦 – a bejelentés még nem elfogadott szabály; elfogadáskor ⚠️ lesz. A README-ben rögzítve: a bejelentés ≠ teljesítés szabály a kedvezőtlen eltérésre is vonatkozik.

Státuszok: ✅ 6, ↗️ 9 (−1), 🔄 7 (−1), 🟦 17 (+1), ⏳ 2 (+1), ⚠️ 3 (−2), ❌ 2 (−1).

## 2026-10-19 – technikai javítások

Tartalmi (státusz-) változás nincs, a 49 tétel és a darabszámok változatlanok.

- `TISZA_merleg.csv`: az első oszlop neve `#` helyett `ID` (a PowerShell `Import-Csv` a `#`-cal kezdődő fejlécet megjegyzésként eldobta);
- a 40. tétel első forrása `Program-összefoglaló:` címkét kapott;
- `Osszefoglalo.csv`: valódi táblázat lett (`Státusz`, `Jelentés`, `Tételek száma`); a módszertani megjegyzések a README-ben vannak;
- `Forrasjegyzek.csv`: új `Link` oszlop, a linkek a `Megjegyzés` mezőből ide kerültek;
- Excel: a fenti változások a munkafüzetben is; az *Összefoglaló* `COUNTIF` képlete a teljes *Státusz* oszlopra számol (eddig az 50. sornál megállt); semleges cím a „személyes mérleg (videó alapján)” helyett;
- `scripts/validate.ps1` és GitHub Actions: CSV-szerkezet, státuszok, darabszámok, README-tételszám és CSV–xlsx egyezés ellenőrzése;
- `LICENSE`: a CC BY 4.0 teljes jogi szövege; a README-ben javasolt forrásmegjelölés;
- `.gitattributes`: LF sorvég minden platformon, az xlsx bináris.

## 2026-10-19

Első nyilvánosra előkészített verzió.

Főbb frissítések:
- 2027-es költségvetési bejelentések beemelése;
- családi pótlék;
- minimálnyugdíj és kisnyugdíjak;
- nyugdíjas SZÉP-kártya;
- otthonápolási díj;
- egészségügyi +500 milliárd Ft-os vállalás;
- vagyonadó;
- KATA-tervek;
- szociális, NOKS-, kulturális és állami egyetemi bérek;
- anyák SZJA-mentességének tervezett korlátozása;
- időskorúak járadéka;
- public GitHub repo és CC BY 4.0 licenc hozzáadása;
- az Excel-fájl új **Projekt** munkalapot kapott a repo-, licenc- és módszertani információkkal.

A „2027-re bejelentve” tételek külön státuszt kaptak, hogy ne keveredjenek a már megvalósult intézkedésekkel.
