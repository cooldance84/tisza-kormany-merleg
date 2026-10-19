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

1. Módosítsd a CSV-t **és** az Excel-munkafüzetet is – a kettőnek egyeznie kell.
2. Új tétel a `TISZA_merleg.csv` végére kerül, a következő `ID`-val; a státusz csak az [`Osszefoglalo.csv`](data/Osszefoglalo.csv) listájából választható, és ott a darabszámot is frissíteni kell.
3. A forrás cellában soronként egy link szerepeljen, opcionális címkével (`Program: https://…`).
4. Futtasd: `powershell -File scripts\validate.ps1` (Windows) vagy `pwsh -File scripts/validate.ps1`. A GitHub Actions ugyanezt futtatja minden PR-en.
