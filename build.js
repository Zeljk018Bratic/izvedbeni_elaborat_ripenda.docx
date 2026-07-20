/**
 * build.js
 * Generates izvedbeni_elaborat_ripenda.docx — Izvedbeni elaborat za projekt
 * "Boutique Apartman Ripenda", adaptacija autohtone istarske kamene kuće u Ripendi (Labin).
 *
 * Naručitelj izrade: Dragoslav (izvođač radova)
 * Investitorica:     Sandra
 *
 * Paleta boja: Navy (#1F3864) + Gold (#C9A84C) — Calibri font
 *
 * Pokrenuti: node build.js
 */

"use strict";

const fs = require("fs");
const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  Table,
  TableRow,
  TableCell,
  WidthType,
  BorderStyle,
  AlignmentType,
  HeadingLevel,
  ShadingType,
  VerticalAlign,
  PageOrientation,
  Header,
  Footer,
  PageNumber,
  NumberFormat,
  convertInchesToTwip,
} = require("docx");

// ─────────────────────────────────────────────
// KONSTANTE — BOJE I FONT
// ─────────────────────────────────────────────
const NAVY = "1F3864";
const GOLD = "C9A84C";
const GOLD_LIGHT = "F5EDD0";
const WHITE = "FFFFFF";
const FONT = "Calibri";

// ─────────────────────────────────────────────
// POMOĆNE FUNKCIJE ZA FORMATIRANJE
// ─────────────────────────────────────────────

/**
 * Naslov sekcije (Heading 1 stil, Navy pozadina, Gold tekst)
 */
function sectionHeading(text) {
  return new Paragraph({
    spacing: { before: 200, after: 120 },
    shading: { type: ShadingType.CLEAR, color: "auto", fill: NAVY },
    children: [
      new TextRun({
        text,
        font: FONT,
        size: 28,
        bold: true,
        color: GOLD,
        allCaps: true,
      }),
    ],
  });
}

/**
 * Podnaslov sekcije (Heading 2, Gold lijeva bordura)
 */
function subHeading(text) {
  return new Paragraph({
    spacing: { before: 160, after: 80 },
    border: {
      left: { style: BorderStyle.THICK, size: 12, color: GOLD },
    },
    indent: { left: 160 },
    children: [
      new TextRun({
        text,
        font: FONT,
        size: 24,
        bold: true,
        color: NAVY,
      }),
    ],
  });
}

/**
 * Normalni paragraf (body text)
 */
function bodyText(text, options = {}) {
  return new Paragraph({
    spacing: { before: 60, after: 60 },
    indent: { left: 160 },
    children: [
      new TextRun({
        text,
        font: FONT,
        size: 20,
        bold: options.bold || false,
        color: options.color || "000000",
      }),
    ],
  });
}

/**
 * Bullet stavka
 */
function bulletItem(text, level = 0) {
  return new Paragraph({
    bullet: { level },
    spacing: { before: 40, after: 40 },
    indent: { left: 360 + level * 360 },
    children: [
      new TextRun({
        text,
        font: FONT,
        size: 20,
        color: "000000",
      }),
    ],
  });
}

/**
 * Prazni redak
 */
function emptyLine() {
  return new Paragraph({ spacing: { before: 40, after: 40 }, children: [] });
}

/**
 * Ćelija zaglavlja tablice (Navy pozadina, Gold tekst, centralno)
 */
function headerCell(text, colspan = 1) {
  return new TableCell({
    columnSpan: colspan,
    shading: { type: ShadingType.CLEAR, color: "auto", fill: NAVY },
    verticalAlign: VerticalAlign.CENTER,
    margins: { top: 60, bottom: 60, left: 120, right: 120 },
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [
          new TextRun({
            text,
            font: FONT,
            size: 18,
            bold: true,
            color: GOLD,
          }),
        ],
      }),
    ],
  });
}

/**
 * Ćelija podataka (izmjenično bijela / light gold)
 */
function dataCell(text, shade = false, bold = false, align = AlignmentType.LEFT) {
  return new TableCell({
    shading: shade
      ? { type: ShadingType.CLEAR, color: "auto", fill: GOLD_LIGHT }
      : undefined,
    verticalAlign: VerticalAlign.CENTER,
    margins: { top: 50, bottom: 50, left: 120, right: 120 },
    children: [
      new Paragraph({
        alignment: align,
        children: [
          new TextRun({
            text,
            font: FONT,
            size: 18,
            bold,
            color: "000000",
          }),
        ],
      }),
    ],
  });
}

// ─────────────────────────────────────────────
// TABLICA — OPĆI PODACI O PROJEKTU
// ─────────────────────────────────────────────
function tableOpciPodaci() {
  const rows = [
    ["Projekt", "Boutique Apartman Ripenda"],
    ["Lokacija", "Ripenda, Labin, Istarska županija"],
    ["Objekt", "Autohtona istarska kamena kuća (prizemlje + tavan)"],
    ["Investitorica", "Sandra"],
    ["Izvođač radova", "Dragoslav"],
    ["Datum izrade elaborata", "srpanj 2026."],
    ["Unutarnje dimenzije prostora", "7,00 m × 4,36 m (dužina × širina)"],
    ["Tlocrtna površina (bruto)", "≈ 30,52 m²"],
    ["Visina stropa (slobodna)", "≥ 2,55 m (ovisno o zoni krova)"],
    ["Namjena", "Kratkoročni turistički najam (boutique apartman)"],
  ];

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({
        children: [headerCell("PARAMETAR"), headerCell("VRIJEDNOST")],
        tableHeader: true,
      }),
      ...rows.map(([param, value], i) =>
        new TableRow({
          children: [
            dataCell(param, i % 2 === 0, true),
            dataCell(value, i % 2 === 0),
          ],
        })
      ),
    ],
  });
}

// ─────────────────────────────────────────────
// TABLICA — ZONE TLOCRTA (bez hodnika)
// ─────────────────────────────────────────────
function tableTlocrtneZone() {
  const zones = [
    ["Z-1", "Dnevni boravak / lounge", "≈ 13,0 m²", "Središnji, orijentacija jug"],
    ["Z-2", "Kuhinja s blagovaonicom", "≈ 7,5 m²", "Sjeverni dio, lijeva strana"],
    ["Z-3", "Spavaća soba", "≈ 7,0 m²", "Sjeveroistok, desna strana"],
    ["Z-4", "Kupaonica", "≈ 3,0 m²", "Desni zid — vodovodni razvod"],
  ];

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({
        children: [
          headerCell("ZONA"),
          headerCell("NAMJENA"),
          headerCell("POVRŠINA"),
          headerCell("NAPOMENA"),
        ],
        tableHeader: true,
      }),
      ...zones.map(([zona, namjena, povrsina, napomena], i) =>
        new TableRow({
          children: [
            dataCell(zona, i % 2 === 0, true, AlignmentType.CENTER),
            dataCell(namjena, i % 2 === 0),
            dataCell(povrsina, i % 2 === 0, false, AlignmentType.CENTER),
            dataCell(napomena, i % 2 === 0),
          ],
        })
      ),
    ],
  });
}

// ─────────────────────────────────────────────
// TABLICA — SUHA GRADNJA (Knauf CW75 + kamena vuna)
// ─────────────────────────────────────────────
function tableSuhaGradnja() {
  const stavke = [
    ["Profil", "Knauf CW 75 (75/50/06)", "Vertikalni nosivi profil, os 600 mm"],
    ["Profil donja/gornja vodilica", "Knauf UW 75 (75/40/06)", "Horizontalna vodilica, obostrano brtvljenje"],
    ["Obloga — unutarnja strana", "Knauf GKB 12,5 mm", "Jednoslojna obloga, fuga brušena"],
    ["Obloga — vanjska strana", "Knauf GKFI 12,5 mm", "Otporna na vlagu, strana prema vanjskom zidu"],
    ["Toplinska / zvučna izolacija", "Rockwool Airrock LD 50 mm", "Klasa reakcije na vatru A1"],
    ["Vaparu barijera (parna brana)", "PE folija 0,2 mm", "Samo gdje postoji rizik kondenzacije"],
    ["Privijanje obloge", "Vijci LN 3,5 × 25", "Os ≤ 250 mm, rubovi ≤ 150 mm"],
    ["Spojevi ploča (fuge)", "Knauf Fugenfüller + armaturna traka", "Razina obrade Q3 (glatko)"],
    ["Završna obrada zida", "Disperzijski temeljni premaz + plemenita žbuka", "Ili glatka glet masa po odluci investitorice"],
    ["Ukupna debljina pregradnog zida", "≈ 100 mm", "CW75 + obloge 2 × 12,5 mm"],
  ];

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({
        children: [
          headerCell("ELEMENT"),
          headerCell("SPECIFIKACIJA"),
          headerCell("NAPOMENA"),
        ],
        tableHeader: true,
      }),
      ...stavke.map(([element, spec, napomena], i) =>
        new TableRow({
          children: [
            dataCell(element, i % 2 === 0, true),
            dataCell(spec, i % 2 === 0),
            dataCell(napomena, i % 2 === 0),
          ],
        })
      ),
    ],
  });
}

// ─────────────────────────────────────────────
// TABLICA — OTVORENI KROVNI ROŽNJAK (kraći raspon)
// ─────────────────────────────────────────────
function tableKrovniRoznjak() {
  const stavke = [
    ["Raspon rožnjaka", "4,36 m (kraća dimenzija tlocrta)", "Bez međuoslonca"],
    ["Vrsta rožnjaka", "Otvoreni (vidljivi) drveni rožnjaci", "Hrast ili četinjača C24, vidljiva površina blanjana"],
    ["Dimenzija presjeka", "14 × 8 cm (visina × širina)", "Provjeriti statičkim proračunom za snijeg i vjetar"],
    ["Os rožnjaka", "≤ 90 cm", "Prema statičkim zahtjevima i vizualnom ritmu"],
    ["Sljeme", "Paralelno s dužom osi prostora (7,00 m)", "Krov pada prema užim zidovima (4,36 m raspon)"],
    ["Nagib krova", "27° – 32° (lokalni vernakularni standard)", "Uskladiti s postojećom krovinskom konstruktorom"],
    ["Završni pokrov", "Istarska kupa kanalica — restauracija", "Zadržati autohton karakter kamene kuće"],
    ["Termoizolacija", "Između rožnjaka: Rockwool Multirock Roll 15 cm", "Ispod rožnjaka: Knauf Cleaneo CLF 15 akustična ploča"],
    ["Parna brana", "Jutafol N AL 170 Special — između izolatora i grijane zone", "Preklop ≥ 15 cm, ljepljene trake Jutafol"],
    ["Nosač sljemena", "Nema (otvoreni krov) — rožnjaci oslonac na zidne grede", "Zidna greda 15 × 10 cm, sidrena u kameni zid"],
  ];

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({
        children: [
          headerCell("PARAMETAR"),
          headerCell("VRIJEDNOST / TIP"),
          headerCell("NAPOMENA"),
        ],
        tableHeader: true,
      }),
      ...stavke.map(([param, vrijednost, napomena], i) =>
        new TableRow({
          children: [
            dataCell(param, i % 2 === 0, true),
            dataCell(vrijednost, i % 2 === 0),
            dataCell(napomena, i % 2 === 0),
          ],
        })
      ),
    ],
  });
}

// ─────────────────────────────────────────────
// TABLICA — VODOVODNE INSTALACIJE (desni zid)
// ─────────────────────────────────────────────
function tableVodovod() {
  const stavke = [
    ["Glavni razvod vode", "Desni zid prostora (gledano od ulaza)", "Razvod ide od sjever prema jugu uz desni zid"],
    ["Hladna voda (HV)", "PP-R PN 20, ∅ 20 mm", "Plava oznaka ili izolacija plave boje"],
    ["Topla voda (TV)", "PP-R PN 20, ∅ 20 mm (izolacija 13 mm Armaflex)", "Crvena oznaka ili izolacija crvene boje"],
    ["Cirkulacijski vod (CV)", "PP-R PN 20, ∅ 16 mm", "Opcija — samo ako je boiler smješten dalje od armature"],
    ["Odvod kupaonice", "PVC ∅ 110 mm (WC), ∅ 50 mm (kade/tuš)", "Niveleta odvoda prema septičkom rov. ili priključku"],
    ["Odvod kuhinje", "PVC ∅ 50 mm", "Sifonski priključak ispod sudopera"],
    ["Položaj bojlera / bojlera-protoka", "Desni zid, zona kupaonице (Z-4)", "Električni bojler 100 L ili protočni grijač — po odabiru"],
    ["Smještaj vodomjernog ormara", "Uz ulazna vrata (lijeva strana, gledano izvana)", "Pristupan bez ulaska u stan"],
    ["Tlak vode", "Nominalni radni tlak ≤ 6 bar", "Ugraditi redukcijski ventil ako mrežni tlak > 4 bar"],
    ["Ispitivanje", "Tlačna proba 10 bar / 30 min (PP-R)", "Protokol o ispitivanju priložiti elaboratu"],
  ];

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({
        children: [
          headerCell("ELEMENT"),
          headerCell("SPECIFIKACIJA"),
          headerCell("NAPOMENA"),
        ],
        tableHeader: true,
      }),
      ...stavke.map(([element, spec, napomena], i) =>
        new TableRow({
          children: [
            dataCell(element, i % 2 === 0, true),
            dataCell(spec, i % 2 === 0),
            dataCell(napomena, i % 2 === 0),
          ],
        })
      ),
    ],
  });
}

// ─────────────────────────────────────────────
// TABLICA — ELEKTRIČNE INSTALACIJE (mikrolokacije)
// ─────────────────────────────────────────────
function tableElektrika() {
  const stavke = [
    // [element, zona, visina od gotovog poda, napomena]
    ["Glavna razvodna kutija (RO)", "Ulazna zona (blizu desnog zida)", "— (u zidu, poklopac na visini 1,50 m)", "Min. IP30; dvostruki osigurači 16 A + zaštitni uređaji 30 mA"],
    ["Utičnice — dnevni boravak (Z-1)", "Z-1 (dnevni boravak)", "30 cm od gotovog poda", "2× dvostruka utičnica; 1× USB-A/C kombinirani modul"],
    ["Utičnice — kuhinja (Z-2)", "Z-2 (kuhinja)", "110 cm od gotovog poda", "4× utičnica (radna ploča, hladnjak, aparat za kavu, mikrovalna)"],
    ["Utičnice — spavaća soba (Z-3)", "Z-3 (spavaća soba)", "30 cm od gotovog poda", "2× dvostruka utičnica (po 1 sa svake strane kreveta)"],
    ["Utičnice — kupaonica (Z-4)", "Z-4 (kupaonica)", "110 cm od gotovog poda, zona 2", "1× utičnica (brijanje / sušilo), IP44"],
    ["Prekidači — dnevni boravak", "Z-1, blizu ulaza", "105 cm od gotovog poda", "1× jednosmjerni + 1× naizmjenični (s kupaonske strane)"],
    ["Prekidači — spavaća soba", "Z-3, blizu vrata", "105 cm od gotovog poda", "1× jednosmjerni + 2× noćni (uz uzglavlje)"],
    ["Prekidači — kupaonica", "Vanjska strana vrata Z-4", "105 cm od gotovog poda", "IP44 izvan kupaonske zone"],
    ["Nadžbukna rasvjeta (spotlight)", "Krovni rožnjaci — sva zona", "Na rožnjaku (−5 cm od donje ivice)", "LED GU10 spot, temperatura 2700 K, CRI ≥ 90"],
    ["Noćna orijentacijska rasvjeta", "Z-1/Z-3 spoj, pod", "10 cm od gotovog poda", "LED noćno svjetlo, senzor mraka"],
    ["TV / data priključak", "Z-1 i Z-3", "30 cm od gotovog poda (uz utičnicu)", "Koaksijalni + Cat6 UTP — do razvoda kabelske TV"],
    ["Termostat (klima)", "Z-1, slobodni zid", "150 cm od gotovog poda", "Za split jedinicu ili radijatorski sustav"],
    ["Alarm / detektor dima", "Centralno — tavan iznad Z-1", "Na rožnjaku, sredina", "Kombinirani detektor dima + CO, baterijski"],
  ];

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({
        children: [
          headerCell("ELEMENT"),
          headerCell("ZONA"),
          headerCell("VISINA"),
          headerCell("NAPOMENA"),
        ],
        tableHeader: true,
      }),
      ...stavke.map(([element, zona, visina, napomena], i) =>
        new TableRow({
          children: [
            dataCell(element, i % 2 === 0, true),
            dataCell(zona, i % 2 === 0, false, AlignmentType.CENTER),
            dataCell(visina, i % 2 === 0, false, AlignmentType.CENTER),
            dataCell(napomena, i % 2 === 0),
          ],
        })
      ),
    ],
  });
}

// ─────────────────────────────────────────────
// TABLICA — BOUTIQUE STIL I DETALJI
// ─────────────────────────────────────────────
function tableBoutique() {
  const stavke = [
    ["Koncept stila", "Rustic Luxe / Mediterranean Boutique", "Kamen + drvo + prirodni materijali, premium detalji"],
    ["Podne obloge", "Tradicionalna istarska kamena ploča 40×40 cm", "Ručno klesana; stara ili nova patinirana površina"],
    ["Zidne obloge", "Vidljivi kameni zid (originalni) + kreč/vapno boja", "Tamo gdje kamen nije izložen: terra cotta ili slonovača"],
    ["Stolarija — vrata", "Pune drvene (hrast/bor), rustikalni okovi", "Visina 210 cm, prirodna lazura ili bijela lak-boja"],
    ["Stolarija — prozori", "Drveni dvostruki prozori s kapcima", "Kapci: tamno zelena (RAL 6009) — istrski standard"],
    ["Kuhinja", "Kuhinja po mjeri, fronti od masivnog drveta", "Radna ploča: bijeli kvarc ili Silestone Blanco Maple"],
    ["Kupaonica", "Podna obloga: mozaik kamen; zid: zelena zelinjska pločica", "Walk-in tuš, crna armatura (Grohe ili Hansgrohe)"],
    ["Rasvjeta", "Industrijski / vintage viseći lusteri nad kuhinjom", "LED retrofit žarulje 2200 K za toplu atmosferu"],
    ["Tekstil i dekor", "Posteljina 400 TC lan/pamuk, Kapa navlake OEKO-TEX", "Lokalni keramički predmeti, maslinovo ulje / lavanda"],
    ["Smart amenities", "Sef, Bluetooth zvučnik, brzi Wi-Fi (AX router)", "Dobrodošlica: lokalno vino, istarski sir, maslinovo ulje"],
    ["Klimatizacija", "Split klima 2,5 kW (A+++) u dnevnom boravku", "Tiha noćna funkcija; filter ionizator zraka"],
  ];

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({
        children: [
          headerCell("ELEMENT"),
          headerCell("TIP / MATERIJAL"),
          headerCell("NAPOMENA"),
        ],
        tableHeader: true,
      }),
      ...stavke.map(([element, tip, napomena], i) =>
        new TableRow({
          children: [
            dataCell(element, i % 2 === 0, true),
            dataCell(tip, i % 2 === 0),
            dataCell(napomena, i % 2 === 0),
          ],
        })
      ),
    ],
  });
}

// ─────────────────────────────────────────────
// HEADER I FOOTER
// ─────────────────────────────────────────────
function buildHeader() {
  return new Header({
    children: [
      new Paragraph({
        border: {
          bottom: { style: BorderStyle.SINGLE, size: 6, color: GOLD },
        },
        children: [
          new TextRun({
            text: "IZVEDBENI ELABORAT  |  Boutique Apartman Ripenda  |  Ripenda, Labin",
            font: FONT,
            size: 16,
            color: NAVY,
            italics: true,
          }),
        ],
      }),
    ],
  });
}

function buildFooter() {
  return new Footer({
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        border: {
          top: { style: BorderStyle.SINGLE, size: 6, color: GOLD },
        },
        children: [
          new TextRun({
            text: "Izvedbeni elaborat – Boutique Apartman Ripenda   |   Stranica ",
            font: FONT,
            size: 16,
            color: NAVY,
          }),
          new TextRun({
            children: [PageNumber.CURRENT],
            font: FONT,
            size: 16,
            color: NAVY,
          }),
          new TextRun({
            text: " od ",
            font: FONT,
            size: 16,
            color: NAVY,
          }),
          new TextRun({
            children: [PageNumber.TOTAL_PAGES],
            font: FONT,
            size: 16,
            color: NAVY,
          }),
        ],
      }),
    ],
  });
}

// ─────────────────────────────────────────────
// NASLOVNICA
// ─────────────────────────────────────────────
function buildCoverPage() {
  return [
    emptyLine(),
    emptyLine(),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 400, after: 200 },
      shading: { type: ShadingType.CLEAR, color: "auto", fill: NAVY },
      children: [
        new TextRun({
          text: "IZVEDBENI ELABORAT",
          font: FONT,
          size: 52,
          bold: true,
          color: GOLD,
          allCaps: true,
        }),
      ],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 80 },
      shading: { type: ShadingType.CLEAR, color: "auto", fill: NAVY },
      children: [
        new TextRun({
          text: "Boutique Apartman Ripenda",
          font: FONT,
          size: 36,
          bold: false,
          color: GOLD_LIGHT,
          italics: true,
        }),
      ],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 400 },
      shading: { type: ShadingType.CLEAR, color: "auto", fill: NAVY },
      children: [
        new TextRun({
          text: "Adaptacija autohtone istarske kamene kuće",
          font: FONT,
          size: 24,
          color: WHITE,
        }),
      ],
    }),
    emptyLine(),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 120, after: 60 },
      children: [
        new TextRun({
          text: "Lokacija:  Ripenda, Labin, Istarska županija",
          font: FONT,
          size: 22,
          bold: true,
          color: NAVY,
        }),
      ],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 60, after: 60 },
      children: [
        new TextRun({
          text: "Investitorica:  Sandra",
          font: FONT,
          size: 22,
          color: NAVY,
        }),
      ],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 60, after: 60 },
      children: [
        new TextRun({
          text: "Izvođač radova:  Dragoslav",
          font: FONT,
          size: 22,
          color: NAVY,
        }),
      ],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 60, after: 60 },
      children: [
        new TextRun({
          text: "Srpanj 2026.",
          font: FONT,
          size: 22,
          italics: true,
          color: NAVY,
        }),
      ],
    }),
    emptyLine(),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 120, after: 0 },
      border: {
        top: { style: BorderStyle.SINGLE, size: 12, color: GOLD },
        bottom: { style: BorderStyle.SINGLE, size: 12, color: GOLD },
      },
      children: [
        new TextRun({
          text: "UNUTARNJE DIMENZIJE: 7,00 m × 4,36 m  |  UKUPNA POVRŠINA: ≈ 30,52 m²",
          font: FONT,
          size: 20,
          bold: true,
          color: GOLD,
        }),
      ],
    }),
    emptyLine(),
    emptyLine(),
  ];
}

// ─────────────────────────────────────────────
// GLAVNI DOKUMENT
// ─────────────────────────────────────────────
async function buildDocument() {
  const doc = new Document({
    creator: "Boutique Apartman Ripenda — Izvedbeni elaborat",
    title: "Izvedbeni elaborat — Boutique Apartman Ripenda",
    description:
      "Tehnička i projektna dokumentacija za adaptaciju autohtone istarske kamene kuće u Ripendi (Labin) u boutique apartman za kratkoročni turistički najam.",
    styles: {
      default: {
        document: {
          run: { font: FONT, size: 20, color: "000000" },
        },
      },
    },
    sections: [
      {
        properties: {
          page: {
            margin: {
              top: convertInchesToTwip(1.0),
              bottom: convertInchesToTwip(1.0),
              left: convertInchesToTwip(1.2),
              right: convertInchesToTwip(1.0),
            },
          },
        },
        headers: { default: buildHeader() },
        footers: { default: buildFooter() },
        children: [
          // ── NASLOVNICA ──────────────────────────────
          ...buildCoverPage(),

          // ── 1. OPĆI PODACI ──────────────────────────
          sectionHeading("1.  Opći podaci o projektu"),
          bodyText(
            "Predmet ovog izvedbenog elaborata jest detaljna tehnička i projektna dokumentacija za " +
            "adaptaciju autohtone istarske kamene kuće na lokaciji Ripenda (Labin) u moderan boutique apartman " +
            "namijenjen kratkoročnom turističkom najmu. Radovi uključuju kompletnu obnovu unutarnjeg prostora, " +
            "ugradnju suhe gradnje, instalaciju vodovodnih i elektroinstalacijskih sustava, te opremanje u " +
            "višem cjenovnom segmentu uz zadržavanje autohtonog istarskog karaktera objekta."
          ),
          emptyLine(),
          tableOpciPodaci(),
          emptyLine(),

          // ── 2. ORIJENTACIJA I TLOCRT ─────────────────
          sectionHeading("2.  Orijentacija prostora i tlocrtne zone"),
          bodyText(
            "Prostor ima pravokutni tlocrt unutarnjih dimenzija 7,00 m (dužina) × 4,36 m (širina). " +
            "Ulaz u stan nalazi se na sjevernom zidu. Gledano od ulaza:"
          ),
          bulletItem("LIJEVO (zapadna strana) — kuhinja s blagovaonicom"),
          bulletItem("DESNO (istočna strana) — kupaonica; sav vodovodni razvod skupljen uz desni zid"),
          bulletItem("SREDINA — dnevni boravak / lounge (centralna zona, orijentacija jug)"),
          bulletItem("SJEVEROZAPAD/SJEVER — spavaća soba"),
          emptyLine(),
          bodyText(
            "Raspored zona projektiran je BEZ hodnika — izravni protok između zona kako bi se " +
            "maksimizirala korisna površina i postigao otvoren, luxe-boutique ambijent.",
            { bold: true }
          ),
          emptyLine(),
          subHeading("2.1  Tlocrtne zone"),
          tableTlocrtneZone(),
          emptyLine(),
          bodyText(
            "NAPOMENA O ORIJENTACIJI: Svi nacrti, tablice i opisi u ovom elaboratu konzistentno " +
            "koriste jedinstven koordinatni sustav — sjever = ulaz, jug = prozorski front. " +
            "Ne postoje zrcalne inačice (Spiegelverkehrt) tlocrta; svi podaci opisuju isti, " +
            "jedinstven raspored gledano od ulaznih vrata.",
            { bold: false, color: "CC0000" }
          ),
          emptyLine(),

          // ── 3. SUHA GRADNJA ──────────────────────────
          sectionHeading("3.  Suha gradnja — Knauf CW75 s kamenom vunom"),
          bodyText(
            "Unutarnje pregradne stijene izvode se u sustavu suhe gradnje Knauf na čeličnoj " +
            "potkonstrukciji. Odabrani sustav osigurava visoku zvučnu izolaciju (Rw ≥ 47 dB), " +
            "požarnu otpornost EI 60, te brzu i čistu izvedbu bez mokrih procesa."
          ),
          emptyLine(),
          subHeading("3.1  Specifikacija materijala i ugradnje"),
          tableSuhaGradnja(),
          emptyLine(),
          subHeading("3.2  Upute za izvođenje"),
          bulletItem("Potkonstrukcija: UW vodilice postaviti na pod i strop s razmaknutim brtvenim trakama (Knauf Dichtungsband)."),
          bulletItem("CW profile postavljati na os 600 mm, visina profila ≤ 4,00 m bez bočnog učvršćenja (inače ugraditi horizontalnu ukrutu)."),
          bulletItem("Izolacija: Rockwool Airrock LD 50 mm uložiti između CW profila, bez prekida i praznina."),
          bulletItem("Ploče GKB/GKFI postavljati naizmjenično (H-fuga), vijke zabijati na os ≤ 250 mm."),
          bulletItem("Kutni spojevi (vanjski kutovi): ugraditi Knauf Winkelschutzprofil s mrežicom."),
          bulletItem("Razina obrade Q3: gruntirati, filati, brusiti na glatku površinu bez vidljivih spojeva."),
          emptyLine(),

          // ── 4. KROVNI ROŽNJACI ───────────────────────
          sectionHeading("4.  Otvoreni krovni rožnjaci — kraći raspon (4,36 m)"),
          bodyText(
            "Krovišna konstrukcija izvodi se kao OTVORENI VIDLJIVI ROŽNJACI koji prate kraći " +
            "raspon prostora (4,36 m). Osi rožnjaka su paralelne s kraćim zidom (sjever–jug), " +
            "a sljeme je postavljeno paralelno s dužom osi (istok–zapad, 7,00 m). " +
            "Ovim rasporedom postiže se autentičan mediteranski izgled s vidljivom drvenom konstrukcijom."
          ),
          emptyLine(),
          subHeading("4.1  Tehnička specifikacija"),
          tableKrovniRoznjak(),
          emptyLine(),
          subHeading("4.2  Napomene za izvođača"),
          bulletItem("Sve dimenzije rožnjaka i osi potvrditi statičkim proračunom za lokalne opterećenja (snijeg, vjetar — zona Labin/Istra)."),
          bulletItem("Drvena građa mora biti sušena na sadržaj vlage ≤ 18% pri ugradnji."),
          bulletItem("Kontaktne površine drva i kamena zaštititi bitumenskim premazom ili EPDM trakom."),
          bulletItem("Sva metalna sidra i vezovi: vrućevaljana galvanizacija ≥ 25 μm (morski zrak — korozivna okolina)."),
          bulletItem("Vidljive površine rožnjaka: blanjane, zaštita Remmers Deckfarbe ili Osmo UV Schutzöl Extra."),
          emptyLine(),

          // ── 5. VODOVOD ──────────────────────────────
          sectionHeading("5.  Vodovodne instalacije — razvod uz desni zid"),
          bodyText(
            "Kompletni vodovodni razvod (hladna voda, topla voda, odvodnja) grupiran je uz desni " +
            "(istočni) zid prostora. Ovim rješenjem minimalizira se dužina cjevovoda, olakšava " +
            "pristup za servis, te se sve mokre zone (kupaonica, kuhinja) drže unutar kompaktnog " +
            "instalacijskog pojasa uz desni zid."
          ),
          emptyLine(),
          subHeading("5.1  Specifikacija cjevovoda i elemenata"),
          tableVodovod(),
          emptyLine(),
          subHeading("5.2  Napomene za izvođača"),
          bulletItem("Sve cijevi unutar suhozidne konstrukcije položiti u zaštitne fleksibilne cijevi (korrugat)."),
          bulletItem("Minimalni pad odvodnih cijevi: 2% prema kolektoru."),
          bulletItem("PP-R cijevi: električki zavariti (polyfuzijsko zavarivanje), minimalna duljina preklopa prema tablici."),
          bulletItem("Izolacija toplih voda: Armaflex 13 mm (min.) ili Rockwool Klimaflex — koeficijent vodljivosti λ ≤ 0,040 W/mK."),
          bulletItem("Po završetku ugradnje provesti tlačnu probu i izdati Zapisnik o ispitivanju."),
          emptyLine(),

          // ── 6. ELEKTROINSTALACIJE ────────────────────
          sectionHeading("6.  Elektroinstalacije — tablica mikrolokacija"),
          bodyText(
            "Sve elektroinstalacije projektiraju se u skladu s normom HRN HD 60364 (IEC 60364). " +
            "Donja tablica definira mikrolokacije svakog instalacijskog elementa s točnom visinom " +
            "ugradnje mjerenom od gotovog poda (GP). Visine su normirane za komfor i sigurnost " +
            "gostiju boutique apartmana."
          ),
          emptyLine(),
          subHeading("6.1  Tablica mikrolokacija i visina ugradnje"),
          tableElektrika(),
          emptyLine(),
          subHeading("6.2  Napomene za izvođača"),
          bulletItem("Sve kutije i vodovi u zoni suhe gradnje: ugraditi PRIJE zaključavanja druge ploče GKB."),
          bulletItem("Razvodni ormár (RO): ugraditi u blizini ulaza, pristupan iz hodnika/zajedničkih prostora."),
          bulletItem("Kupaonica (Z-4): strogo poštivati zone 0/1/2 prema normi IEC 60364-7-701."),
          bulletItem("Zaštitno uzemljenje: sve metalne mase (slavine, okviri, vrata) spojiti na PE vodič."),
          bulletItem("Kabelska trasa: odvojiti jake struje od slabih struja (min. 10 cm razmak)."),
          bulletItem("Pri vidljivim rožnjacima: kabelske zaštitne cijevi crne boje, pričvrstiti na 50 cm razmaka."),
          emptyLine(),

          // ── 7. BOUTIQUE STIL ─────────────────────────
          sectionHeading("7.  Boutique stilski elementi i opremanje"),
          bodyText(
            "Apartman se pozicionira u segment 'Rustic Luxe' — kombinacija autohtonog istarskog " +
            "nasljeđa (kamen, drvo, glinena opeka) s premium materijalima i suvremenim sadržajima. " +
            "Ciljni gost: par / solo putnik koji traži autentično, intimno iskustvo u srcu Istre."
          ),
          emptyLine(),
          subHeading("7.1  Materijali, oprema i detalji"),
          tableBoutique(),
          emptyLine(),
          subHeading("7.2  Smjernice za investitoricu Sandru"),
          bulletItem("Fotografija za platforme: profesionalna fotografija s toplim 2700 K rasvjetom, zlatni sat (17:00 h)."),
          bulletItem("Opis na Airbnb / Booking: isticati 'vidljive kamene zidove', 'originalne rožnjake', 'lokalni domaćin'."),
          bulletItem("Cijenovna kategorija: 90–140 EUR/noć (visoka sezona lipanj–rujan), s minimalnim boravkom 3 noći."),
          bulletItem("Zelena certifikacija: razmotriti EarthCheck ili TripAdvisor GreenLeaders radi konkurentnosti."),
          emptyLine(),

          // ── 8. NAPOMENE I OBVEZE ──────────────────────
          sectionHeading("8.  Opće napomene i obveze izvođača"),
          bulletItem("Sva odstupanja od projekta moraju se pisano dogovoriti s investitoricom (Sandra) i arhitektom."),
          bulletItem("Svaka faza radova dokumentira se fotografijama (prije/poslije) i predaje investitorici."),
          bulletItem("Otpadni materijal odvoziti prema Planu gospodarenja otpadom; zabranjena odlaganje na parceli."),
          bulletItem("Inspekcijski nadzor: obavijestiti gradbeninspektora za sva nosiva sidra u kameni zid."),
          bulletItem("Garantno razdoblje: suha gradnja 2 god., instalacije 2 god., stolarija 1 god. od primopredaje."),
          bulletItem("Primopredajni zapisnik potpisuju: izvođač Dragoslav, investitorica Sandra i nadzorni inženjer."),
          emptyLine(),

          // ── NAPOMENA O VERZIJI ───────────────────────
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 240, after: 120 },
            border: {
              top: { style: BorderStyle.SINGLE, size: 6, color: GOLD },
            },
            children: [
              new TextRun({
                text: "Ovaj elaborat generiran je automatski putem Node.js build skripte (build.js).",
                font: FONT,
                size: 16,
                italics: true,
                color: "888888",
              }),
            ],
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 0, after: 120 },
            children: [
              new TextRun({
                text: `Datum generiranja: ${new Date().toLocaleDateString("hr-HR", { year: "numeric", month: "long", day: "numeric" })}`,
                font: FONT,
                size: 16,
                italics: true,
                color: "888888",
              }),
            ],
          }),
        ],
      },
    ],
  });

  return doc;
}

// ─────────────────────────────────────────────
// POKRETANJE
// ─────────────────────────────────────────────
(async () => {
  try {
    console.log("⏳  Generiranje dokumenta...");
    const doc = await buildDocument();
    const buffer = await Packer.toBuffer(doc);
    const outputPath = "izvedbeni_elaborat_ripenda.docx";
    fs.writeFileSync(outputPath, buffer);
    console.log(`✅  Dokument uspješno stvoren: ${outputPath} (${Math.round(buffer.length / 1024)} KB)`);
  } catch (err) {
    console.error("❌  Greška pri generiranju dokumenta:", err);
    process.exit(1);
  }
})();
