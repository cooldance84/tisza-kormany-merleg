# Changelog

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
