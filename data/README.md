# Adatfájlok

A könyvtár a TISZA-kormány mérleg aktuális, géppel is könnyen feldolgozható CSV-változatait tartalmazza.

- **TISZA_merleg.csv** – a teljes mérleg, 49 tétellel
- **Osszefoglalo.csv** – státuszok, jelentésük és tételszámuk
- **Forrasjegyzek.csv** – kiemelt források
- **TISZA_kormany_merleg_2026-10-19.xlsx** – ugyanez Excel-munkafüzetként (plusz *Projekt* munkalap)

**Snapshot dátuma:** 2026-10-19

A sorokban szereplő státuszok az adott snapshot időpontjának állapotát mutatják. Egy későbbi bejelentés, jogszabály vagy végrehajtási adat a státuszt megváltoztathatja.

A CSV-formátum szándékos: GitHubon jól diffelhető, kereshető és géppel feldolgozható. Kódolás: UTF-8, sorvég: LF, elválasztó: vessző.

## TISZA_merleg.csv

| Oszlop | Tartalom |
|---|---|
| `ID` | Sorszám, 1-től hézag nélkül |
| `Terület` | Téma rövid neve |
| `Probléma / vállalás` | A kiinduló probléma vagy a vállalás |
| `Mi történt a kormányváltás óta?` | A snapshot-dátumig ellenőrizhető fejlemények |
| `Státusz` | Az `Osszefoglalo.csv` egyik státusza, emojival együtt |
| `Kapcsolódó TISZA-programvállalás` | Idézet vagy összefoglaló a programból |
| `Forrás / ellenőrzés` | Soronként egy link, opcionális címkével: `Program: https://…`, `Program-összefoglaló: https://…` |

Egy cellában több forrás is lehet, ilyenkor soronként egy szerepel (idézőjeles, többsoros CSV-mező).

## Osszefoglalo.csv

`Státusz`, `Jelentés`, `Tételek száma` – a státuszok listája egyben az engedélyezett értékkészlet a `TISZA_merleg.csv` *Státusz* oszlopához. A darabszámokat a [`scripts/validate.ps1`](../scripts/validate.ps1) összeveti a mérleggel. A módszertani megjegyzések a fő [README](../README.md)-ben vannak.

## Forrasjegyzek.csv

`Forrás`, `Típus`, `Mire használjuk?`, `Megjegyzés`, `Link` – a `Megjegyzés` szöveges, a linkek a `Link` oszlopban vannak (több link esetén soronként egy).

## Excel-munkafüzet

A *TISZA mérleg* és a *Forrásjegyzék* munkalapnak cellára egyeznie kell a megfelelő CSV-vel; az *Összefoglaló* munkalap a darabszámokat `COUNTIF` képlettel, a teljes *Státusz* oszlopra számolja. Ha a munkafüzetet módosítod, a CSV-ket is frissíteni kell, és fordítva – az eltérést a validátor jelzi.
