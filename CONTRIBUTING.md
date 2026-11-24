# Közreműködés

Köszönjük a javításokat és forrásokat.

## Jó hibajegy / Issue

Kérjük, add meg:

1. **Téma / sor neve**
2. **Jelenlegi állítás**
3. **Javasolt javítás**
4. **Forrás URL**
5. **Forrás dátuma**
6. Röviden: miért támasztja alá a forrás a módosítást?

## Forrás-prioritás

Előnyben részesítjük:
1. jogszabály / hivatalos dokumentum;
2. hivatalos statisztika;
3. hatósági vagy bírósági közlemény;
4. több, egymástól független sajtóforrás.

Pártkommunikáció és véleménycikk használható annak dokumentálására, hogy **ki mit állított**, de önmagában nem elegendő vitatott tény bizonyítására.

## Adatmódosítás pull requestben

1. A CSV-t szerkeszd, majd futtasd: `powershell -File scripts\sync_xlsx.ps1` – ez a CSV-ből frissíti az Excel-munkafüzetet (cellák, új vagy törölt sorok, darabszámok, frissítési dátum, tételszám, a sorok státusz szerinti színezése, a rögzített fejléc és A–B oszlop). A munkafüzetet ne Excelben szerkeszd, a színezést se állítsd kézzel.
2. Új tétel a `TISZA_merleg.csv` végére kerül, a következő `ID`-val. A státusz csak az [`Osszefoglalo.csv`](data/Osszefoglalo.csv) listájából választható.
3. **Státusz változásakor vagy új tételnél** frissítsd a darabszámokat az [`Osszefoglalo.csv`](data/Osszefoglalo.csv)-ben **és** a [README](README.md) „Státuszok” táblázatában; új tételnél a README-kben szereplő tételszámot is („… tételes mérleg”).
4. A `Vállalt határidő` mező kötelező, és csak a program szövegében szereplő határidő kerülhet bele (`első intézkedések`, `azonnal`, konkrét év, pl. `2030-ra`), különben `nincs megadva`. ❌ státusz csak a határidő után adható – lásd a README „Határidők és a tételek életciklusa” szakaszát.
5. A forrás cellában soronként egy link szerepeljen, opcionális címkével (`Program: https://…`, `Jogszabály – 2026. évi … törvény: https://…`).
6. Ha egy tételt a forrásokkal újra összevetettél, frissítsd az `Utolsó ellenőrzés` dátumát – akkor is, ha a státusz nem változott. A README „Utolsó frissítés” sorának a legkésőbbi ilyen dátumot kell mutatnia.
7. Írj bejegyzést a [`CHANGELOG.md`](CHANGELOG.md) elejére: mi változott, melyik tételnél, és ha a státusz is, akkor miért (a módszertan melyik szabálya alapján).
8. Az Excel-munkafüzetre ne tegyél érzékenységi címkét (pl. céges gépen az Office kérésére); a validátor elutasítja.
9. Futtasd: `powershell -File scripts\validate.ps1` (Windows) vagy `pwsh -File scripts/validate.ps1`. A GitHub Actions ugyanezt futtatja minden PR-en. A validátor jelzi, ha a darabszámok, a dátumok, a határidő-mező, a színezés vagy a CSV–Excel egyezés nincs rendben.
