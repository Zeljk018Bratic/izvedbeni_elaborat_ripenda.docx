# Izvedbeni Elaborat — Boutique Apartman Ripenda

> **Adaptacija autohtone istarske kamene kuće u Ripendi (Labin)**  
> Investitorica: **Sandra** · Izvođač radova: **Dragoslav**

Ovaj repozitorij sadrži Node.js build skriptu i svu prateću dokumentaciju za automatsko generiranje
profesionalnog Word dokumenta (`izvedbeni_elaborat_ripenda.docx`) — izvedbenog elaborata za projekt
adaptacije kamene kuće u boutique apartman za kratkoročni turistički najam.

---

## 🗂️  Arhitektura repozitorija

```
izvedbeni_elaborat_ripenda.docx/
├── build.js          ← Node.js build skripta (generira .docx)
├── package.json      ← Konfiguracija projekta i ovisnosti
├── README.md         ← Ovaj dokument — projektni dashboard
└── .gitignore        ← Isključuje node_modules i generirani .docx
```

---

## 🚀  Brzi početak

```bash
# 1. Instalirajte ovisnosti
npm install

# 2. Generirajte Word dokument
npm run build
# ili direktno:
node build.js

# Rezultat: izvedbeni_elaborat_ripenda.docx (u korijenu projekta)
```

**Zahtjevi:** Node.js ≥ 18

---

## 📐  Parametri prostora

| Parametar | Vrijednost |
|-----------|------------|
| Unutarnje dimenzije | **7,00 m × 4,36 m** (dužina × širina) |
| Tlocrtna površina | ≈ 30,52 m² |
| Slobodna visina stropa | ≥ 2,55 m (ovisno o zoni krova) |
| Ulaz | Sjeverni zid |
| Vidljivi rožnjaci | Raspon po kraćoj dimenziji (4,36 m) |

---

## 🗺️  Tlocrtne zone (bez hodnika)

Prostor je projektiran **bez hodnika** — direktan, otvoren protok između zona maksimizira
korisnu površinu i postiže boutique ambijent:

| Zona | Namjena | Površina | Položaj |
|------|---------|----------|---------|
| **Z-1** | Dnevni boravak / lounge | ≈ 13,0 m² | Središnji, orijentacija jug |
| **Z-2** | Kuhinja s blagovaonicom | ≈ 7,5 m² | Sjever, **lijeva strana** |
| **Z-3** | Spavaća soba | ≈ 7,0 m² | Sjeverozapad/sjever, **desna strana** |
| **Z-4** | Kupaonica | ≈ 3,0 m² | **Desni zid** — vodovodni razvod |

> **Orijentacija:** Gledano od ulaznih vrata — lijevo je zapad (kuhinja), desno je istok
> (kupaonica + vodovod). Sva dokumentacija u elaboratu konzistentno koristi ovaj koordinatni
> sustav. **Nema zrcalnih inačica (Spiegelverkehrt) tlocrta.**

---

## 🏗️  Tehničke specifikacije (pregled)

### Suha gradnja — Knauf CW75 + kamena vuna 50 mm
- Profil: **Knauf CW 75** (os 600 mm) + UW 75 vodilice
- Izolacija: **Rockwool Airrock LD 50 mm** (klasa A1)
- Obloge: Knauf GKB 12,5 mm (unutra) + GKFI 12,5 mm (prema vanjskom zidu)
- Ukupna debljina zida: ≈ **100 mm**
- Razina obrade: **Q3** (glatko, bez vidljivih spojeva)

### Otvoreni krovni rožnjaci (kraći raspon 4,36 m)
- Rožnjaci prate **kraću dimenziju** (4,36 m) — sljeme paralelno s dužom osi (7,00 m)
- Presjek: **14 × 8 cm**, hrast ili četinjača C24, blanjana vidljiva površina
- Os: ≤ 90 cm · Nagib: 27°–32° (vernakularni istarski standard)
- Pokrov: **Istarska kupa kanalica** — restauracija
- Izolacija između rožnjaka: **Rockwool Multirock Roll 15 cm**

### Vodovod — razvod uz desni (istočni) zid
- Hladna i topla voda: **PP-R PN 20, ∅ 20 mm** (izolacija toplih voda 13 mm Armaflex)
- Odvodnja: **PVC ∅ 110 mm** (WC), **∅ 50 mm** (tuš/sudoper)
- Tlačna proba: 10 bar / 30 min; protokol priložiti elaboratu
- Bojler: električki 100 L ili protočni grijač — smješten u Z-4 (desni zid)

### Elektroinstalacije — mikrolokacije
Kompletna tablica s visinama ugradnje od gotovog poda (GP):

| Element | Visina ugradnje |
|---------|-----------------|
| Utičnice — pod (dnevni, spavaća) | **30 cm** od GP |
| Utičnice — radna ploča kuhinje | **110 cm** od GP |
| Utičnice — kupaonica (IP44) | **110 cm** od GP |
| Prekidači | **105 cm** od GP |
| Rasvjeta (LED spot GU10, 2700 K) | Na rožnjaku (−5 cm od donje ivice) |
| Noćna orijentacijska rasvjeta | **10 cm** od GP |
| Termostat klime | **150 cm** od GP |
| Razvodni ormar (RO) | Poklopac na **150 cm** od GP |

---

## 🏡  Boutique Apartman — stilske smjernice

Koncept: **Rustic Luxe / Mediterranean Boutique**  
Kombinacija autohtonog istarskog nasljeđa (kamen, drvo, glinena opeka) s premium materijalima i suvremenim sadržajima.

| Kategorija | Rješenje |
|------------|----------|
| Pod | Tradicionalna istarska kamena ploča **40×40 cm** |
| Zidovi | Vidljivi originalni kameni zid + vapnena/kreč boja |
| Stolarija | Puna drvena vrata (hrast/bor) · Drveni prozori s kapcima **RAL 6009** |
| Kuhinja | Po mjeri, frontovi masivno drvo · Radna ploča bijeli kvarc |
| Kupaonica | Walk-in tuš · Crna armatura · Mozaik kamen (pod) · Zelena pločica (zid) |
| Rasvjeta | Vintage/industrijski lusteri · LED 2200–2700 K |
| Klimatizacija | Split klima 2,5 kW (A+++) · Tiha noćna funkcija |
| Smart amenities | Wi-Fi AX · Bluetooth zvučnik · Sef · Lokalni welcome basket |

**Cjenovna pozicija:** 90–140 EUR/noć (visoka sezona) · Minimalni boravak: 3 noći

---

## 📄  Sadržaj generiranog elaborata

Dokument `izvedbeni_elaborat_ripenda.docx` sadrži sljedeće sekcije:

1. Opći podaci o projektu (tablica)
2. Orijentacija prostora i tlocrtne zone
3. Suha gradnja — Knauf CW75 + kamena vuna (tablice + upute)
4. Otvoreni krovni rožnjaci — kraći raspon 4,36 m (tablice + upute)
5. Vodovodne instalacije — razvod uz desni zid (tablice + upute)
6. Elektroinstalacije — tablica mikrolokacija i visina ugradnje
7. Boutique stilski elementi i opremanje (tablice + smjernice)
8. Opće napomene i obveze izvođača

---

## 🎨  Dizajn dokumenta

- **Paleta:** Navy `#1F3864` + Gold `#C9A84C` — korporativni premium izgled
- **Font:** Calibri (naslov 26 pt, tijelo 10 pt)
- **Zaglavlje:** Naziv projekta + lokacija
- **Podnožje:** Stranica X od Y

---

## 📦  Ovisnosti

| Paket | Verzija | Svrha |
|-------|---------|-------|
| [`docx`](https://www.npmjs.com/package/docx) | `^9.7.1` | Generiranje Word (.docx) dokumenata |

---

*Repozitorij: `Zeljk018Bratic/izvedbeni_elaborat_ripenda.docx`*
