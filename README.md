# Boutique Apartman Ripenda — Izvedbeni Elaborat

Profesionalni repozitorij za izradu službenog izvedbenog elaborata i tehničke dokumentacije projekta adaptacije autohtone istarske kamene kuće u Ripendi (Labin).

## Sadržaj repozitorija

- `build.js` — Node.js skripta koja pomoću biblioteke `docx` generira visoko formatirani Word dokument:
  - izlazna datoteka: `izvedbeni_elaborat_ripenda.docx`
- `package.json` — konfiguracija projekta i popis ovisnosti
- `README.md` — glavni dashboard projekta

## Arhitektura dokumenta

Generirani dokument je strukturiran kao izvedbeni elaborat za realnu izvedbu na terenu:

1. **Osnovni podaci i orijentacija**
   - čiste unutarnje dimenzije: **7.00 m × 4.36 m**
   - jedinstvena, nezrcaljena orijentacija (bez Spiegelverkehrt konflikata)
2. **Layout zone bez hodnika**
   - ulaz/garderoba
   - dnevni boravak
   - kuhinja + blagovanje
   - spavaća niša
   - kupaonica
3. **Suha gradnja**
   - sustav: **Knauf CW75**
   - ispuna: **kamena vuna 50 mm**
4. **Krovna logika**
   - otvoreno krovište
   - rešetkasti nosači na **kraćem rasponu (4.36 m)**
5. **Instalacije vode**
   - vodne i odvodne linije grupirane uz **desni zid**
6. **Elektro mikro-lokacije**
   - tablica visina elemenata u cm od gotovog poda

## Dizajnerski standard (“Boutique Apartment”)

Dokument je vizualno i sadržajno prilagođen boutique standardu:

- korporativna paleta: **Navy + Gold**
- tipografija: **Calibri**
- slojevita rasvjeta i premium atmosfera
- maksimalno funkcionalan plan bez hodnika

## Upute za korištenje

### 1) Instalacija

```bash
npm install
```

### 2) Generiranje elaborata

```bash
npm run build
```

### 3) Rezultat

Nakon pokretanja build skripte, u root direktoriju nastaje:

- `izvedbeni_elaborat_ripenda.docx`

## Napomena za izvođača Dragoslava

Svi tehnički opisi, zone i instalacijske mikro-lokacije postavljeni su konzistentno u jednoj stvarnoj orijentaciji prostora. Repozitorij i generirani dokument **ne sadrže konfliktne mirror-image (Spiegelverkehrt) podatke**.

HTML pregled i ispis grafičkih priloga
Repo sada sadrži mini dashboard i odvojene ulazne točke za sheetove:

index.html — glavni dashboard
A-101.html — wrapper za A-101
A-102.html — wrapper za A-102
A-103.html — wrapper za A-103
A-301.html — wrapper za A-301
Prilozi_Instalacije_i_Fasade.html — glavni sadržaj priloga
Kako otvoriti projekt (HTML)
Opcija A — lokalno (preporučeno)
U root folderu projekta pokreni:

python -m http.server 8080
Zatim otvori:

http://localhost:8080/index.html
Opcija B — direktno iz repoa
Otvori index.html na GitHubu i koristi Raw prikaz (ili GitHub Pages ako je uključen).

Kako isprintati u PDF
Otvori željeni sheet (A-101, A-102, A-103, A-301) preko index.html
Pritisni Ctrl + P
Postavke:
Destination: Save as PDF
Layout: Landscape
Scale: 100%
Margins: None (ili Minimum)
Background graphics: ON
Preporučeni nazivi izlaza:

A-101_Tlocrt.pdf
A-102_Instalacije_Voda_Odvod.pdf
A-103_Instalacije_Elektro.pdf
A-301_Fasade.pdf
Napomena
A-101/A-102/A-103/A-301 su wrapper stranice koje preusmjeravaju na sekcije unutar Prilozi_Instalacije_i_Fasade.html putem anchor linkova (#A-101, #A-102, itd.).
