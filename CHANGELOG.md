# Changelog

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
