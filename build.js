const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell,
  WidthType, ShadingType, AlignmentType, BorderStyle, PageBreak, VerticalAlign
} = require('docx');

const FONT = "Calibri";
const NAVY = "2C3E50";
const GOLD = "8A6D3A";
const LIGHT = "F4F1EA";

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 400, after: 200 },
    border: { bottom: { color: NAVY, space: 4, style: BorderStyle.SINGLE, size: 12 } },
    children: [ new TextRun({ text, bold: true, color: NAVY, size: 30, font: FONT }) ]
  });
}

function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 260, after: 140 },
    children: [ new TextRun({ text, bold: true, color: GOLD, size: 24, font: FONT }) ]
  });
}

function p(text, opts={}) {
  return new Paragraph({
    spacing: { after: 120 },
    children: [ new TextRun({ text, size: 21, font: FONT, ...opts }) ]
  });
}

function bullet(text) {
  return new Paragraph({
    spacing: { after: 60 },
    bullet: { level: 0 },
    children: [ new TextRun({ text, size: 21, font: FONT }) ]
  });
}

function cell(text, opts={}) {
  const { width=2000, header=false, shade=null, bold=false, align=AlignmentType.LEFT } = opts;
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    shading: shade ? { type: ShadingType.CLEAR, fill: shade } : undefined,
    verticalAlign: VerticalAlign.CENTER,
    margins: { top: 80, bottom: 80, left: 100, right: 100 },
    children: [ new Paragraph({
      alignment: align,
      children: [ new TextRun({ text: String(text), bold: header || bold, size: 20, font: FONT, color: header ? "FFFFFF" : "000000" }) ]
    }) ]
  });
}

function table(headers, rows, widths) {
  const total = widths.reduce((a,b)=>a+b,0);
  return new Table({
    width: { size: total, type: WidthType.DXA },
    columnWidths: widths,
    rows: [
      new TableRow({
        tableHeader: true,
        children: headers.map((hh,i)=>cell(hh, {width: widths[i], header:true, shade: NAVY}))
      }),
      ...rows.map((r,ri)=> new TableRow({
        children: r.map((c,i)=>cell(c, {width: widths[i], shade: ri%2? LIGHT : "FFFFFF"}))
      }))
    ]
  });
}

const sections = [];

// ===================== STRANICA 1: NASLOVNICA =====================
sections.push({
  properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1000, bottom: 1000, left: 1100, right: 1100 } } },
  children: [
    new Paragraph({ spacing:{before:1200}, alignment: AlignmentType.CENTER, children:[
      new TextRun({ text:"PROJEKT ADAPTACIJE", bold:true, size:34, color:NAVY, font:FONT })
    ]}),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing:{after:100}, children:[
      new TextRun({ text:"AUTOHTONE ISTARSKE KAMENE KUĆE", bold:true, size:34, color:NAVY, font:FONT })
    ]}),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing:{after:600}, children:[
      new TextRun({ text:"BOUTIQUE APARTMAN \u2013 RIPENDA", bold:true, size:44, color:GOLD, font:FONT })
    ]}),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing:{after:800}, children:[
      new TextRun({ text:"IZVEDBENI ELABORAT", size:26, color:"666666", font:FONT, italics:true })
    ]}),
    new Table({
      alignment: AlignmentType.CENTER,
      width: { size: 6500, type: WidthType.DXA },
      columnWidths:,
      rows: [
        new TableRow({ children:[cell("Investitor:", {width:2500, bold:true, shade:LIGHT}), cell("Sandra Antunović", {width:4000})]}),
        new TableRow({ children:[cell("Izvođač radova:", {width:2500, bold:true, shade:LIGHT}), cell("Dragoslav Lazić", {width:4000})]}),
        new TableRow({ children:[cell("Lokacija:", {width:2500, bold:true, shade:LIGHT}), cell("Ripenda, Labin, Istarska županija", {width:4000})]}),
        new TableRow({ children:[cell("Tip zahvata:", {width:2500, bold:true, shade:LIGHT}), cell("Adaptacija/rekonstrukcija kamene kuće", {width:4000})]}),
        new TableRow({ children:[cell("Neto površina:", {width:2500, bold:true, shade:LIGHT}), cell("\u2248 30,5 m\u00b2", {width:4000})]}),
        new TableRow({ children:[cell("Status dokumenta:", {width:2500, bold:true, shade:LIGHT}), cell("FINALNI \u2013 spreman za izvođenje", {width:4000})]}),
        new TableRow({ children:[cell("Verzija:", {width:2500, bold:true, shade:LIGHT}), cell("V4 \u2013 All-In sintetizirani elaborat", {width:4000})]}),
        new TableRow({ children:[cell("Datum izrade:", {width:2500, bold:true, shade:LIGHT}), cell("20. srpnja 2026.", {width:4000})]}),
      ]
    }),
    new Paragraph({ spacing:{before:1200}, alignment: AlignmentType.CENTER, children:[
      new TextRun({ text:"Ovaj dokument sintetizira sve dosadašnje idejne skice, mjere s terena i tehničke upute u jedan jedinstveni izvedbeni projekt.", size:19, italics:true, color:"777777", font:FONT })
    ]}),
    new Paragraph({ spacing:{before:200}, alignment: AlignmentType.CENTER, children:[
      new TextRun({ text:"Sve mjere obavezno provjeriti na licu mjesta prije početka radova.", size:19, italics:true, color:"a33333", font:FONT })
    ]}),
    new Paragraph({ children:[ new PageBreak() ] })
  ]
});

// ===================== STRANICA 2 I NASTAVAK: TEHNIČKI SADRŽAJ =====================
const mainSectionChildren = [
  h1("1. LEGENDA PROSTORIJA I NETO KVADRATURA"),
  p("U skladu s izmjerama na terenu (7,00 x 4,36 m unutrašnjeg neto prostora), kuća je organizirana u 'L' obliku bez klasičnih hodnika radi maksimalne iskoristivosti kvadrata. S ceste se ulazi direktno u kuhinjski dio."),
  table(
    ["Naziv prostorije (Boutique Code)", "Dimenzije (m)", "Površina (m²)", "Tip poda i obloga"],
    [
      ["Kitchen & Dining (Prednji desni dio)", "3,06 × 4,36", "13,34 m²", "Mikrocement / Velike pločice 60x120 cm"],
      ["Master Bedroom (Prednji lijevi dio)", "3,94 × 2,94", "11,58 m²", "Pod u ravnini, veliki francuski prozor"],
      ["Bathroom / WC (Stražnji desni dio)", "2,40 × 1,42", "3,40 m²", "Hidroizolacija, walk-in tuš kanalica, pad 2%"],
      ["Anteroom / Passage (Stražnji produžetak)", "1,54 × 1,42", "2,18 m²", "Otvoreni komunikacijski prolaz iz kuhinje"]
    ],
    [3200, 2000, 2000, 3800]
  ),
  p("Ukupna neto unutrašnja stambena površina iznosi ≈ 30,50 m².", { bold: true, italics: true }),

  h1("2. GRAĐEVINSKI I SANACIJSKI RADOVI NA TERENU"),
  p("Svi radovi moraju očuvati vanjsku kamenu strukturu (50-60 cm debljine). Bočni nosivi zidovi su slijepi i ne smiju se štemati."),
  h2("2.1. Sanacija poda i ravnanje u 'nulu'"),
  bullet("Uklanjanje postojeće stepenice od 17 cm. Cijeli pod se ravna u idealnu nulu bez prijelaza i pragova."),
  bullet("Postavljanje hidroizolacijske trake V4 s podizanjem uz kamene zidove minimalno 15 cm."),
  bullet("Ugradnja toplinske izolacije od tvrdog Styrodura (XPS) debljine 5 cm i izlijevanje cementne glazure s vlaknima."),
  h2("2.2. Rušenje kamina i pregradni zidovi"),
  bullet("Potpuno uklanjanje starog otvorenog kamina u desnom kutu radi oslobađanja prostora za blagovaonski stol."),
  bullet("Zidanje novih pregradnih zidova kupaonice sustavom Knauf CW75 debljine 10 cm, s obostrano dvostrukim vlagootpornim zelenim pločama i 50 mm kamenom vunom unutra zbog maksimalne akustične izolacije prema spavaćoj sobi."),

  h1("3. TEHNIČKA SPECIFIKACIJA KROVIŠTA I GREDA"),
  p("Konstrukcija krova ostaje potpuno otvorena do prve grede u rustikalnom 'konoba' stilu."),
  bullet("Smjer greda: Glavne masivne drvene grede ugrađuju se na kraći raspon od 4,36 m kako bi statičko opterećenje bilo minimalno, a prostor vizualno širi."),
  bullet("Čiste visine stropa: Visina do greda u kuhinji iznosi 2,74 m, dok u spavaćoj sobi iznosi 2,91 m."),
  bullet("Toplinska izolacija: Kompletna izolacija (kamena vuna ili PIR) postavlja se isključivo IZNAD rogova. Između greda idu ugradbene tradicijske ton-ploče (tavelle) ili brodski pod. Vidljiva pur pjena je strogo zabranjena."),

  h1("4. VODOVODNE I KANALIZACIJSKE INSTALACIJE"),
  bullet("Sve vodovodne i odvodne linije (WC Ø110, tuš i umivaonik Ø50) izvode se duž desnog zida s minimalnim padom od 2%."),
  bullet("Kupaonica se naslanja direktno desno na postojeći WC od majke, što osigurava najkraći i najjeftiniji put spajanja na kanalizacijsku vertikalu."),
  bullet("Topla voda se rješava preko plosnatog električnog bojlera (npr. Ariston Velis 80L) smještenog u kupaonici iznad perilice rublja."),

  h1("5. ELEKTROINSTALACIJE I TABLICA MIKRO-LOKACIJA"),
  table(
    ["Uređaj / Priključak", "Lokacija na nacrtu", "Visina od poda (cm)", "Tehnička napomena"],
    [
      ["Klimatizacija (Unutarnja jedinica)", "Desni kuhinjski zid", "250 cm", "Izvod za napajanje i odvod kondenzata direktno van"],
      ["Televizor i Internet (Smart TV)", "Pregradni Knauf zid u sobi", "120 cm", "Kablovi i utičnice se skrivaju unutar konstrukcije zida"],
      ["LED rasvjeta (Kuhinjski niz)", "Ispod hrastovih polica", "140 cm", "Topla ugradbena LED traka (3000 K) duž cijele radne ploče"],
      ["Kupaonske utičnice i prekidači", "Desni zid kupaonice", "110 cm", "Obavezna ugradnja utičnica s IP44 zaštitom od vlage"],
      ["Utičnice iznad radne ploče", "Desni kuhinjski zid", "110 cm", "Trostruki ugradbeni blok za male kućanske aparate"],
      ["Utičnice uz bračni krevet", "Lijevi i desni bok kreveta", "40 cm", "Izmjenični prekidači za gašenje glavnog svjetla iz kreveta"]
    ],
    [3000, 2500, 2000, 3500]
  ),

  h1("6. RASVJETA, PODNE OBLOGE I DIZAJN INTERIJERA"),
  h2("6.1. Ambiantna rasvjeta i pod"),
  bullet("Obavezno korištenje tople rasvjete od 3000 K za istarski kamen i grede. Ugradbeni spotovi idu u prolaz kupaonice, a visilica iznad stola."),
  bullet("Jedinstveni pod bez pragova i prijelaza kroz cijelu kuću – velike pločice 60x120 cm ili mikrocement."),
  h2("6.2. Specifikacija 'Boutique' namještaja"),
