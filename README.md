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
