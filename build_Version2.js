const fs = require("fs");
const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  HeadingLevel,
  Table,
  TableRow,
  TableCell,
  WidthType,
  ShadingType,
  AlignmentType,
  BorderStyle,
  PageBreak,
  VerticalAlign,
} = require("docx");

const FONT = "Calibri";
const NAVY = "2C3E50";
const GOLD = "8A6D3A";
const LIGHT = "F4F1EA";

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 400, after: 200 },
    border: { bottom: { color: NAVY, space: 4, style: BorderStyle.SINGLE, size: 12 } },
    children: [new TextRun({ text, bold: true, color: NAVY, size: 30, font: FONT })],
  });
}

function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 260, after: 140 },
    children: [new TextRun({ text, bold: true, color: GOLD, size: 24, font: FONT })],
  });
}

function p(text, opts = {}) {
  return new Paragraph({
    spacing: { after: 120 },
    children: [new TextRun({ text, size: 21, font: FONT, ...opts })],
  });
}

function bullet(text) {
  return new Paragraph({
    spacing: { after: 60 },
    bullet: { level: 0 },
    children: [new TextRun({ text, size: 21, font: FONT })],
  });
}

function cell(text, opts = {}) {
  const { width = 2000, header = false, shade = null, bold = false, align = AlignmentType.LEFT } = opts;
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    shading: shade ? { type: ShadingType.CLEAR, fill: shade } : undefined,
    verticalAlign: VerticalAlign.CENTER,
    margins: { top: 80, bottom: 80, left: 100, right: 100 },
    children: [
      new Paragraph({
        alignment: align,
        children: [
          new TextRun({
            text: String(text),
            bold: header || bold,
            size: 20,
            font: FONT,
            color: header ? "FFFFFF" : "000000",
          }),
        ],
      }),
    ],
  });
}

function table(headers, rows, widths) {
  const total = widths.reduce((a, b) => a + b, 0);
  return new Table({
    width: { size: total, type: WidthType.DXA },
    columnWidths: widths,
    rows: [
      new TableRow({
        tableHeader: true,
        children: headers.map((hh, i) => cell(hh, { width: widths[i], header: true, shade: NAVY })),
      }),
      ...rows.map((r, ri) =>
        new TableRow({
          children: r.map((c, i) => cell(c, { width: widths[i], shade: ri % 2 ? LIGHT : "FFFFFF" })),
        })
      ),
    ],
  });
}

const sections = [];

// NASLOVNICA
sections.push({
  properties: {
    page: {
      size: { width: 11906, height: 16838 },
      margin: { top: 1000, bottom: 1000, left: 1100, right: 1100 },
    },
  },
  children: [
    new Paragraph({
      spacing: { before: 1200 },
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: "PROJEKT ADAPTACIJE", bold: true, size: 34, color: NAVY, font: FONT })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 100 },
      children: [new TextRun({ text: "AUTOHTONE ISTARSKE KAMENE KUĆE", bold: true, size: 34, color: NAVY, font: FONT })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 600 },
      children: [new TextRun({ text: "BOUTIQUE APARTMAN – RIPENDA", bold: true, size: 44, color: GOLD, font: FONT })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 800 },
      children: [new TextRun({ text: "IZVEDBENI ELABORAT", size: 26, color: "666666", font: FONT, italics: true })],
    }),
    new Table({
      alignment: AlignmentType.CENTER,
      width: { size: 6500, type: WidthType.DXA },
      columnWidths: [2500, 4000],
      rows: [
        new TableRow({ children: [cell("Investitor:", { width: 2500, bold: true, shade: LIGHT }), cell("Sandra Antunović", { width: 4000 })] }),
        new TableRow({ children: [cell("Izvođač radova:", { width: 2500, bold: true, shade: LIGHT }), cell("Dragoslav Lazić", { width: 4000 })] }),
        new TableRow({ children: [cell("Lokacija:", { width: 2500, bold: true, shade: LIGHT }), cell("Ripenda, Labin, Istarska županija", { width: 4000 })] }),
        new TableRow({ children: [cell("Tip zahvata:", { width: 2500, bold: true, shade: LIGHT }), cell("Adaptacija/rekonstrukcija kamene kuće", { width: 4000 })] }),
        new TableRow({ children: [cell("Neto površina:", { width: 2500, bold: true, shade: LIGHT }), cell("≈ 30,50 m²", { width: 4000 })] }),
        new TableRow({ children: [cell("Status dokumenta:", { width: 2500, bold: true, shade: LIGHT }), cell("FINALNI – spreman za izvođenje", { width: 4000 })] }),
        new TableRow({ children: [cell("Verzija:", { width: 2500, bold: true, shade: LIGHT }), cell("V4 – All-In sintetizirani elaborat", { width: 4000 })] }),
        new TableRow({ children: [cell("Datum izrade:", { width: 2500, bold: true, shade: LIGHT }), cell("20. srpnja 2026.", { width: 4000 })] }),
      ],
    }),
    new Paragraph({
      spacing: { before: 1200 },
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: "Ovaj dokument sintetizira sve dosadašnje idejne skice, mjere s terena i tehničke upute u jedan jedinstveni izvedbeni projekt.", size: 19, italics: true, color: "777777", font: FONT })],
    }),
    new Paragraph({
      spacing: { before: 200 },
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: "Sve mjere obavezno provjeriti na licu mjesta prije početka radova.", size: 19, italics: true, color: "A33333", font: FONT })],
    }),
    new Paragraph({ children: [new PageBreak()] }),
  ],
});

// TEHNIČKI SADRŽAJ
sections.push({
  properties: {
    page: {
      size: { width: 11906, height: 16838 },
      margin: { top: 1200, bottom: 1200, left: 1200, right: 1200 },
    },
  },
  children: [
    h1("1. LEGENDA PROSTORIJA I NETO KVADRATURA"),
    p("Unutarnji neto prostor je 7,00 x 4,36 m. Raspored je bez klasičnih hodnika radi maksimalne iskoristivosti, s jasnom i nezrcaljenom orijentacijom."),
    table(
      ["Naziv prostorije", "Dimenzije (m)", "Površina (m²)", "Tip poda i obloga"],
      [
        ["Kitchen & Dining", "3,06 × 4,36", "13,34", "Mikrocement / pločice 60x120"],
        ["Master Bedroom", "3,94 × 2,94", "11,58", "Pod u ravnini, francuski prozor"],
        ["Bathroom / WC", "2,40 × 1,42", "3,40", "Hidroizolacija, walk-in, pad 2%"],
        ["Anteroom / Passage", "1,54 × 1,42", "2,18", "Otvoreni komunikacijski prolaz"],
      ],
      [3200, 2000, 2000, 3800]
    ),
    p("Ukupna neto površina: ≈ 30,50 m².", { bold: true, italics: true }),

    h1("2. SUHA GRADNJA I SANACIJA"),
    bullet("Pregradni sustav: Knauf CW75 + UW profili."),
    bullet("Ispuna: kamena vuna 50 mm za akustiku i toplinsku stabilnost."),
    bullet("Obostrano vlagootporne ploče u zoni kupaonice."),
    bullet("Pod ravnan u nulu, bez pragova i prijelaza."),

    h1("3. KROV I GREDE"),
    bullet("Otvoreno krovište u rustikalnom stilu."),
    bullet("Smjer greda na kraćem rasponu 4,36 m."),
    bullet("Izolacija iznad rogova; bez vidljive PUR pjene."),

    h1("4. VODOVOD I ODVODNJA"),
    bullet("Sve vodne i odvodne linije grupirane uz desni zid."),
    bullet("Minimalni pad odvodnje 2%."),
    bullet("Bez prebacivanja vodne jezgre na lijevu stranu (bez Spiegelverkehrt konflikata)."),

    h1("5. ELEKTRO MIKRO-LOKACIJE"),
    table(
      ["Uređaj / Priključak", "Lokacija", "Visina od poda (cm)", "Napomena"],
      [
        ["Klima (unutarnja)", "Desni kuhinjski zid", "250", "Napajanje + kondenzat"],
        ["Smart TV + internet", "Pregradni zid sobe", "120", "Skrivena instalacija"],
        ["LED kuhinjska rasvjeta", "Ispod hrastovih polica", "140", "3000K topla"],
        ["Kupaonske utičnice", "Desni zid kupaonice", "110", "IP44"],
        ["Utičnice radne ploče", "Desni kuhinjski zid", "110", "Trostruki blok"],
        ["Utičnice uz krevet", "Oba boka kreveta", "40", "Izmjenični prekidači"],
      ],
      [3000, 2500, 2000, 3500]
    ),

    h1("6. BOUTIQUE INTERIJER"),
    h2("6.1. Rasvjeta i pod"),
    bullet("Topla rasvjeta 3000K za kamen i drvo."),
    bullet("Jedinstveni pod 60x120 ili mikrocement kroz cijeli prostor."),
    h2("6.2. Namještaj"),
    bullet("Kuhinja bez gornjih elemenata, dvije hrastove police s LED trakom."),
    bullet("Ugradbeni ormar do stropa (cca 270 cm), push-open fronte."),
    bullet("Walk-in tuš 90x120 s kanalicom u ravnini poda."),
  ],
});

const doc = new Document({ sections });

Packer.toBuffer(doc)
  .then((buffer) => {
    fs.writeFileSync("izvedbeni_elaborat_ripenda.docx", buffer);
    console.log("✅ Dokument uspješno generiran: izvedbeni_elaborat_ripenda.docx");
  })
  .catch((err) => {
    console.error("Greška pri generiranju dokumenta:", err);
  });