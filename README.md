# TISZA-kormány mérleg – 2026

Nyilvános, forrásolt és folyamatosan frissíthető mérleg a 2026-ban hivatalba lépett TISZA-kormány programvállalásairól, intézkedéseiről, bejelentéseiről és hiányosságairól.

**Utolsó frissítés:** 2026. október 19.

> Ez a projekt nem pártkiadvány és nem kampányanyag.  
> Célja, hogy ugyanazzal a mércével különválassza a programígéretet, a bejelentést, a jogszabályi döntést és a ténylegesen megvalósult eredményt.

## Aktuális adatok

A repository diffelhető, nyílt formátumban is tartalmazza az aktuális adatokat:

- [`data/TISZA_merleg.csv`](data/TISZA_merleg.csv) – a teljes, **49 tételes mérleg**
- [`data/Osszefoglalo.csv`](data/Osszefoglalo.csv) – státuszok és összefoglaló
- [`data/Forrasjegyzek.csv`](data/Forrasjegyzek.csv) – kiemelt forrásjegyzék

A CSV-k a 2026.10.19-én frissített Excel-munkafüzet tartalmából készültek. A GitHubon a CSV a praktikusabb alapformátum, mert a változások soronként visszakövethetők és összehasonlíthatók.

## Státuszok

| Jelölés | Jelentés |
|---|---|
| ✅ **Eredmény** | Kézzelfoghatóan megvalósult intézkedés vagy eredmény |
| ↗️ **Előrelépés** | Konkrét lépés történt |
| 🔄 **Folyamatban** | Elindult, de még nem lezárt |
| 🟦 **2027-re bejelentve** | Konkrét intézkedést / összeget bejelentettek, de még nem hatályos |
| ⏳ **Nyitott** | Van programvállalás, de még nincs elég konkrét végrehajtási információ |
| ⚠️ **Hiányosság** | Problémás, részleges, késedelmes vagy az eredeti vállalástól eltérő végrehajtás |
| ❌ **Nem teljesült** | A vállalás érdemben még nem valósult meg |

## Módszertan

A tételeket lehetőség szerint ebben a sorrendben ellenőrizzük:

1. **Eredeti program / vállalás** – pontosan mit és mikorra ígértek?
2. **Hivatalos döntés** – született-e törvény, rendelet, költségvetési döntés vagy más jogi aktus?
3. **Végrehajtás** – ténylegesen hatályba lépett-e, történt-e kifizetés vagy gyakorlati változás?
4. **Eltérés** – változott-e az összeg, a határidő vagy a célcsoport az eredeti vállaláshoz képest?
5. **Forrás** – elsődlegesen hivatalos dokumentum, jogszabály vagy statisztika; ezek mellett több, egymástól független sajtóforrás.

### Fontos értelmezési szabály

A **bejelentés nem azonos a teljesítéssel**.  
Ezért például egy 2027-re bejelentett intézkedés addig nem kap ✅ státuszt, amíg ténylegesen nem lép hatályba / nem valósul meg.

Ugyanígy egy folyamatban lévő nyomozás, vizsgálat vagy jogalkotási folyamat sem tekinthető lezárt eredménynek.

## Forráskezelés

A mérleg soronként tartalmazza az ellenőrzéshez használt linkeket. Elsődleges forrásként előnyben részesítjük:

- Magyar Közlöny / Nemzeti Jogszabálytár
- KSH
- kormany.hu
- Országgyűlés
- ügyészségi és bírósági közlemények
- EU-intézmények hivatalos dokumentumai

Sajtóforrásoknál törekszünk több, egymástól független szerkesztőség összevetésére. Véleménycikket nem kezelünk tényforrásként.

## Mit nem állít ez a projekt?

- Nem ad általános „jó / rossz kormány” minősítést.
- Nem tekint minden bejelentést kész eredménynek.
- Nem tekint minden vádat bizonyított ténynek.
- Nem állítja, hogy a lista teljes; új tételek folyamatosan felvehetők.
- A státuszok mindig az adott frissítési dátum állapotát mutatják.

## Frissítés és közreműködés

Ha hibás adatot, elavult státuszt vagy jobb elsődleges forrást találsz, nyiss **Issue-t**, és add meg:

- melyik sor / téma érintett;
- mi a javítandó állítás;
- a forrás linkjét;
- a forrás dátumát.

A cél nem az, hogy valamelyik politikai oldal „nyerjen”, hanem hogy **egy állítás mellé ellenőrizhető forrás kerüljön**.

Részletesebb közreműködési szabályok: [`CONTRIBUTING.md`](CONTRIBUTING.md)

## Verziózás

A fontosabb módosításokat a [`CHANGELOG.md`](CHANGELOG.md) tartalmazza.

## Licenc

A dokumentáció és az adatállományok **Creative Commons Attribution 4.0 International (CC BY 4.0)** licenc alatt használhatók fel. Részletek: [LICENSE](LICENSE).
