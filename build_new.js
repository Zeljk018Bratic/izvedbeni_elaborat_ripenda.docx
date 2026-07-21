const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell,
  WidthType, ShadingType, AlignmentType, BorderStyle, PageBreak, PageOrientation,
  Header, Footer, PageNumber, VerticalAlign, convertInchesToTwip
} = require('docx');

const FONT = "Calibri";
const NAVY = "2C3E50";
const GOLD = "8A6D3A";
const LIGHT = "F4F1EA";
const GREEN = "1E6B4F";

function h1(text, opts={}) {
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
      children: [ new TextRun({ text: String(text), bold: header || bold, size: header ? 20 : 20, font: FONT, color: header ? "FFFFFF" : "000000" }) ]
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

// ===================== COVER PAGE =====================
sections.push({
  properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1000, bottom: 1000, left: 1100, right: 1100 } } },
  children: [
    new Paragraph({ spacing:{before:1200}, alignment: AlignmentType.CENTER, children:[
      new TextRun({ text:"ELABORAT A", bold:true, size:26, color:"a33333", font:FONT })
    ]}),
    new Paragraph({ spacing:{before:200}, alignment: AlignmentType.CENTER, children:[
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
      columnWidths: [2500, 4000],
      rows: [
        new TableRow({ children:[cell("Investitor:", {width:2500, bold:true, shade:LIGHT}), cell("Sandra Antunović", {width:4000})]}),
        new TableRow({ children:[cell("Izvođač radova:", {width:2500, bold:true, shade:LIGHT}), cell("Dragoslav Lazić", {width:4000})]}),
        new TableRow({ children:[cell("Lokacija:", {width:2500, bold:true, shade:LIGHT}), cell("Ripenda, Labin, Istarska županija", {width:4000})]}),
        new TableRow({ children:[cell("Tip zahvata:", {width:2500, bold:true, shade:LIGHT}), cell("Adaptacija/rekonstrukcija postojeće kamene kuće", {width:4000})]}),
        new TableRow({ children:[cell("Neto tlocrtna površina:", {width:2500, bold:true, shade:LIGHT}), cell("\u2248 30,5 m\u00b2", {width:4000})]}),
        new TableRow({ children:[cell("Status dokumenta:", {width:2500, bold:true, shade:LIGHT}), cell("FINALNI \u2013 spreman za izvođenje", {width:4000})]}),
        new TableRow({ children:[cell("Verzija:", {width:2500, bold:true, shade:LIGHT}), cell("R3 \u2013 finalni jedinstveni izvedbeni projekt (Opcija A)", {width:4000})]}),
        new TableRow({ children:[cell("Datum izrade:", {width:2500, bold:true, shade:LIGHT}), cell("21. srpnja 2026.", {width:4000})]}),
      ]
    }),
    new Paragraph({ spacing:{before:1200}, alignment: AlignmentType.CENTER, children:[
      new TextRun({ text:"Ovaj dokument sintetizira sve dosadašnje idejne skice, mjere sa terena, tehničke upute i prijedloge (uklj. HomeByMe model \"Ripenda\") u jedan jedinstveni izvedbeni elaborat.", size:19, italics:true, color:"777777", font:FONT })
    ]}),
    new Paragraph({ spacing:{before:200}, alignment: AlignmentType.CENTER, children:[
      new TextRun({ text:"Sve mjere obavezno provjeriti na licu mjesta prije početka radova.", size:19, italics:true, color:"a33333", font:FONT })
    ]}),
    new Paragraph({ spacing:{before:120}, alignment: AlignmentType.CENTER, children:[
      new TextRun({ text:"Ovo je jedini važeći izvedbeni raspored prostora (\u201eOpcija A\u201c: ulaz u kuhinju, spavaća soba lijevo naprijed, kupaonica iza sobe, portal i terasa na ulaznoj strani). Ranije razmatrana \u201eOpcija B\u201c (soba straga) je odbačena i NIJE dio ovog dokumenta.", size:18, italics:true, color:"a33333", font:FONT })
    ]}),
    new Paragraph({ children:[ new PageBreak() ] })
  ]
});

// ===================== UVOD / SAŽETAK =====================
sections.push({
  properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1000, bottom: 1000, left: 1100, right: 1100 } } },
  headers: { default: new Header({ children: [ new Paragraph({ alignment: AlignmentType.RIGHT, children:[ new TextRun({text:"Boutique Apartman Ripenda \u2013 Izvedbeni elaborat", size:16, color:"999999", font:FONT}) ]}) ] }) },
  footers: { default: new Footer({ children: [ new Paragraph({ alignment: AlignmentType.CENTER, children:[
    new TextRun({text:"Stranica ", size:16, color:"999999", font:FONT}),
    new TextRun({children:[PageNumber.CURRENT], size:16, color:"999999", font:FONT})
  ]}) ] }) },
  children: [
    h1("UVOD"),
    p("Ovaj elaborat predstavlja završni izvedbeni paket za adaptaciju autohtone istarske kamene kuće u boutique apartman s naglaskom na kvalitetu detalja, funkcionalnost i arhitektonsku čitljivost."),
    p("Dokument je osmišljen kao službena podloga za investitora, izvođača, stolariju, instalatere i sve subizvođače koji sudjeluju u realizaciji zahvata."),

    h2("Sažetak projekta"),
    table(
      ["Element","Opis","Napomena"],
      [
        ["Koncept","Luksuzni boutique studio apartman uz očuvanje karaktera istarske kamene kuće","Cilj je očuvati autentičnost i dodati suvremenu funkcionalnost"],
        ["Glavna ideja","Otvoreni prostor bez klasičnih hodnika, jasna komunikacija kuhinja \u2013 spavaća soba \u2013 kupaonica","Prostor organiziran radi maksimalne fluidnosti"],
        ["Kvalitativni standard","Visokokvalitetni materijali, precizna izvedba i kontrola na licu mjesta","Materijali i detalji moraju biti usklađeni prije narudžbe"],
      ],
      [2200, 4200, 3100]
    ),

    h2("Napomene za izvođača"),
    bullet("Radove započeti tek po provjeri postojećih kotnih odnosa, visina i stanja konstrukcije na licu mjesta."),
    bullet("Sve izvedbene detalje, spojeve, završne obrade i priključke potrebno je uskladiti s arhitektonskom podlogom prije finalne narudžbe materijala."),
    bullet("U slučaju odstupanja od nacrta, obvezno zatražiti pisanu potvrdu investitora i projektanta prije nastavka radova."),
    bullet("Prije ugradnje stolarije i sanitarija provjeriti sve dimenzije, nagibe poda i položaj priključaka u prostoru."),

    h2("Projektni opis i principi izvedbe"),
    p("Projekt tretira postojeću kamenu kuću kao kulturni i arhitektonski resurs koji treba biti sačuvan u svojoj autentičnosti, a istodobno prilagođen suvremenom načinu života. Osnovni koncept temelji se na otvorenom, fluidnom prostoru u kojem se kuhinja, blagovaona, spavaća soba i kupaonica povezuju bez suvišnih pregrada, uz zadržavanje čvrstog karaktera tradicionalne istarske arhitekture."),
    p("1. Očuvanje i integracija \u2013 sva postojeća kamena ovojnica zadržava se i tretira kao dio arhitektonskog identiteta objekta; novi elementi ne smiju narušavati strukturu, proporcije i povijesni karakter kuće."),
    p("2. Funkcionalna i estetska logika \u2013 prostor je organiziran oko jedinstvene ravnine poda, otvorenih komunikacija i jasne funkcionalne hijerarhije."),
    p("3. Kvaliteta i kontrola \u2013 svi spojevi, čvorovi, priključci i završne obrade moraju biti kontrolirani na licu mjesta prije finalne narudžbe."),

    h2("Kontrola verzije i revizije"),
    table(
      ["Revizija","Datum","Opis promjene","Odgovorna osoba"],
      [
        ["R0","20.07.2026.","Početna verzija izvedbenog elaborata","Projektni tim"],
        ["R1","21.07.2026.","Dodani odjeljci za troškovnik, popis materijala i elektro-shemu","Projektni tim"],
      ],
      [1400, 1800, 4800, 1500]
    ),

    h2("Popis priloga"),
    bullet("Prilog A \u2013 Tlocrt i dispozicija prostora (Opcija A, HTML/PDF prikaz s dimenzijama)."),
    bullet("Prilog B \u2013 Detalji izvedbe poda, krova i pregrada."),
    bullet("Prilog C \u2013 Tehnička specifikacija materijala i opreme."),
    bullet("Prilog D \u2013 Popis troškova i kontrolni troškovnik."),
    bullet("Prilog E \u2013 Elektro i instalacijska shema s mikro-lokacijama."),

    new Paragraph({ children:[ new PageBreak() ] })
  ]
});

// ===================== 1. LEGENDA PROSTORIJA =====================
const bodyPage = (children) => ({
  properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1000, bottom: 1000, left: 1100, right: 1100 } } },
  headers: { default: new Header({ children: [ new Paragraph({ alignment: AlignmentType.RIGHT, children:[ new TextRun({text:"Boutique Apartman Ripenda \u2013 Izvedbeni elaborat", size:16, color:"999999", font:FONT}) ]}) ] }) },
  footers: { default: new Footer({ children: [ new Paragraph({ alignment: AlignmentType.CENTER, children:[
    new TextRun({text:"Stranica ", size:16, color:"999999", font:FONT}),
    new TextRun({children:[PageNumber.CURRENT], size:16, color:"999999", font:FONT})
  ]}) ] }) },
  children
});

sections.push(bodyPage([
  h1("1. LEGENDA PROSTORIJA I NETO KVADRATURA"),
  p("U skladu s izmjerama na terenu (7,00 m širine x 4,36 m dubine unutrašnjeg prostora), objekt je organiziran kao boutique studio apartman. Klasični hodnici su eliminirani \u2013 komunikacija se odvija preko otvorenog L-produžetka kuhinje prema kupaonici (Opcija A, usvojena varijanta)."),
  table(
    ["Prostorija","Dimenzije (m)","Površina","Pod / obloga"],
    [
      ["Kitchen & Dining (prednji desni dio)","3,06 x 4,36","13,34 m\u00b2","Mikrocement / pločice 60x120, bez pragova"],
      ["Master Bedroom (prednji lijevi dio)","3,94 x 2,94","11,58 m\u00b2","Jedinstveni pod, francuski prozor do poda"],
      ["Bathroom / WC (stražnji desni dio)","2,40 x 1,42","3,40 m\u00b2","Hidroizolacija, walk-in tuš kanalica, pad 2%"],
      ["Anteroom / Passage (prolaz)","1,54 x 1,42","2,18 m\u00b2","Otvoreni prolaz izravno iz kuhinje"],
    ],
    [3200, 1800, 1600, 2900]
  ),
  new Paragraph({ spacing:{before:160}, children:[ new TextRun({text:"UKUPNA NETO POVRŠINA: 30,50 m\u00b2", bold:true, size:22, color:NAVY, font:FONT}) ]}),

  h1("2. GRAĐEVINSKI I SANACIJSKI RADOVI"),
  p("Svi zahvati izvode se uz maksimalno očuvanje autentične vanjske kamene ovojnice (50\u201360 cm). Lijevi i desni bočni zidovi su potpuno slijepi i ne smiju se građevinski narušavati."),
  h2("2.1. Sanacija poda i ravnanje u \u2018nulu\u2019"),
  bullet("Kompletno uklanjanje postojeće stepenice i razlike u visini poda od 17 cm \u2013 cijela etaža u jednoj ravnini, bez pragova."),
  bullet("Hidroizolacijska bitumenska traka V4, s podizanjem uz kamene zidove min. 15 cm."),
  bullet("Toplinska izolacija poda \u2013 Styrodur (XPS) 5 cm, cementni estrih s polipropilenskim vlaknima."),
  h2("2.2. Rušenje i pregradni zidovi"),
  bullet("Rušenje, uklanjanje i odvoz na deponiju starog otvorenog kamina u desnom kutu (oslobađa prostor za blagovaonski stol)."),
  bullet("Novi pregradni zidovi kupaonice i sobe \u2013 suha gradnja, debljina 10 cm, Knauf CW75/UW75."),
  bullet("Ispuna: 50 mm polutvrde kamene vune (zvučna izolacija)."),
  bullet("Obostrano dvostruko oblaganje gips-kartonom; sa strane kupaonice i kuhinje obavezno zelene vlagootporne ploče."),

  h1("3. TEHNIČKA SPECIFIKACIJA KROVIŠTA I GREDA"),
  p("Krov ostaje potpuno otvoren do prve grede (tradicionalni istarski stil)."),
  bullet("Smjer greda: glavne masivne drvene grede (hrast/crnogorica) postavljaju se na kraći raspon kuće \u2013 prate dubinu 4,36 m radi minimalnog opterećenja i vizualno šireg prostora."),
  bullet("Čiste visine stropa (pod \u2192 donji rub grede): kuhinjski dio 2,74 m; spavaća soba 2,91 m."),
  bullet("Toplinska izolacija krova: kamena vuna visoke gustoće ili PIR ploče, isključivo IZNAD rogova."),
  bullet("Podgled: tradicionalne glinene ton-ploče ili rustikalni brodski pod. Vidljiva PUR pjena je STROGO ZABRANJENA."),
]));

sections.push(bodyPage([
  h1("4. VODOVODNE I KANALIZACIJSKE INSTALACIJE"),
  bullet("Sve vodovodne i odvodne linije (WC \u00d8110, tuš, sudoper i perilica \u00d850) izvode se duž desnog zida s minimalnim padom 2%."),
  bullet("Kupaonica se naslanja direktno desno na postojeći WC od majke \u2013 omogućuje jeftino spajanje na postojeću kanalizacijsku vertikalu bez prokopavanja kuće."),
  bullet("Topla voda: plosnati električni bojler (npr. Ariston Velis 80L), smješten u kupaonici iznad perilice rublja."),

  h1("5. ELEKTROINSTALACIJE I TABLICA MIKRO-LOKACIJA"),
  p("Razvod mora pratiti točne visinske kote. Svi prekidači i utičnice u kupaonici moraju imati IP44 zaštitu."),
  table(
    ["Uređaj / priključak","Lokacija","Visina (cm)","Napomena"],
    [
      ["Klima (unutarnja jedinica)","Desni kuhinjski zid","250","Odvod kondenzata direktno van"],
      ["TV i internet (Smart TV)","Pregradni Knauf zid u sobi","120","Kablovi skriveni unutar zida"],
      ["LED traka (kuhinjski niz)","Ispod hrastovih polica","140","Topla LED traka 3000K, cijela radna ploha"],
      ["Ugradbeni spotovi","Strop \u2013 zona anterooma","strop","Topla rasvjeta 3000K"],
      ["Viseća lampa","Iznad blagovaonskog stola","podesivo","Montaža na krovnu konstrukciju"],
      ["Utičnice/prekidači kupaonica","Desni zid kupaonice","110","Obavezna IP44 zaštita"],
      ["Utičnice radna ploča","Desni kuhinjski zid","110","Trostruki blok za male aparate"],
      ["Utičnice uz krevet","Lijevi i desni bok kreveta","40","Izmjenični prekidači iz kreveta"],
    ],
    [2600, 2600, 1200, 3100]
  ),

  h1("6. RASVJETA, PODNE OBLOGE I DIZAJN INTERIJERA"),
  h2("6.1. Ambijentalna rasvjeta i pod"),
  bullet("Toplo svjetlo 3000 K obavezno za kamen i grede \u2013 hladna/bijela sterilna svjetlost zabranjena."),
  bullet("Jedinstveni pod bez pragova kroz cijelu kuću \u2013 pločice 60x120 cm ili mikrocement, ista nijansa."),
  h2("6.2. Namještaj i sanitarije"),
  bullet("Kuhinjski niz duž cijelog desnog zida, bez gornjih elemenata \u2013 dvije masivne hrastove police s LED trakom."),
  bullet("Ugradbeni ormar (soba): 240\u2013270 cm dužine, do stropa (270 cm), push-open bez ručkica."),
  bullet("Walk-in tuš (kupaonica): 90x120 cm, fiksna staklena stijena, podna kanalica u ravnini s podom (pad 2%)."),
  bullet("Boutique terasa ispred portala: cca 2,20 x 2,50 m, istarski kamen, nivelirana u nulu s unutarnjim podom."),
]));

sections.push(bodyPage([
  h1("7. TROŠKOVNIK MATERIJALA I RADOVA"),
  p("Sljedeće tablice služe kao službena kontrolna podloga za upis količina, jediničnih cijena i obračun ukupnih troškova radova između investitora i izvođača."),
  h2("7.1. Grubi građevinski radovi i pregrade"),
  table(
    ["Stavka","Jed. mjere","Količina","Jed. cijena (€)","Ukupno (€)"],
    [
      ["Rušenje i uklanjanje starog kamina s odvozom šuta","paušal","1","",""],
      ["Uklanjanje postojeće stepenice (17 cm) i priprema podloge","paušal","1","",""],
      ["Hidroizolacijska bitumenska traka V4","m\u00b2","35","",""],
      ["Toplinska izolacija poda \u2013 Styrodur (XPS) 5 cm","m\u00b2","31","",""],
      ["Cementni estrih (glazura) d=5-6 cm s vlaknima","m\u00b2","31","",""],
      ["Pregradni zid Knauf CW75/UW75 s dvostrukim pločama","m\u00b2","12","",""],
      ["Zelene vlagootporne gips-kartonske ploče","m\u00b2","24","",""],
      ["Ispuna pregradnog zida \u2013 kamena vuna 50 mm","m\u00b2","12","",""],
      ["Toplinska izolacija krova iznad rogova","m\u00b2","68","",""],
      ["Glinene krovne ploče / brodski pod za podgled","m\u00b2","68","",""],
    ],
    [3800, 1200, 1200, 1500, 1300]
  ),
  h2("7.2. Fina oprema, keramika i stolarija"),
  table(
    ["Stavka","Jed. mjere","Količina","Jed. cijena (€)","Ukupno (€)"],
    [
      ["Podna obloga \u2013 keramika 60x120 cm / mikrocement","m\u00b2","31","",""],
      ["Klizno-podizni crni aluminijski portal (180-200 cm)","kom","1","",""],
      ["Unutarnja drvena klizna vrata (soba i WC)","kom","2","",""],
      ["Fiksna staklena stijena za walk-in tuš (90x120 cm)","kom","1","",""],
      ["Podna tuš kanalica (rešetka nehrđajući čelik)","kom","1","",""],
      ["Ugradbeni vodokotlić (npr. Geberit) s tipkom","kom","1","",""],
      ["Električni plosnati bojler 80L","kom","1","",""],
      ["Klimatizacija \u2013 unutarnja i vanjska jedinica s montažom","komplet","1","",""],
      ["Masivne hrastove police za kuhinjski niz","kom","2","",""],
      ["Ugradbeni ormar po mjeri do stropa (v=270, š=240 cm)","kom","1","",""],
    ],
    [3800, 1200, 1200, 1500, 1300]
  ),
]));

sections.push(bodyPage([
  h1("8. POPIS MATERIJALA PO STAVKAMA"),
  p("Popis materijala po stavkama omogućuje brz i profesionalan pristup narudžbi, kontrolnoj evidenciji i objašnjenju izvedbe na gradilištu."),
  h2("8.1. Građevinski materijali"),
  table(
    ["Stavka","Specifikacija","Količina","Napomena"],
    [
      ["Hidroizolacija","Bitumenska traka V4","35 m\u00b2","Podizanje uz kamene zidove min. 15 cm"],
      ["Toplinska izolacija poda","XPS 5 cm","31 m\u00b2","U ravnini s estrihom"],
      ["Pregradni sustav","Knauf CW75/UW75","12 m\u00b2","Suha gradnja, zvučna izolacija"],
      ["Ispuna zida","Kamena vuna 50 mm","12 m\u00b2","Za akustiku i toplinsku stabilnost"],
      ["Krovna izolacija","PIR ili kamena vuna","68 m\u00b2","Iznad rogova, bez PUR pjene"],
      ["Obloga poda","Keramika 60x120 / mikrocement","31 m\u00b2","Bez pragova i prijelaza"],
    ],
    [2200, 3000, 1600, 2700]
  ),
  h2("8.2. Sanitarije, stolarija i oprema"),
  table(
    ["Stavka","Specifikacija","Količina","Napomena"],
    [
      ["Walk-in tuš","Staklena stijena 90x120 cm","1 kom","S podnom kanalicom u ravnini poda"],
      ["Vodokotlić","Ugradbeni model s tipkom","1 kom","Za kupaonicu"],
      ["Bojler","Ariston Velis 80L","1 kom","U kupaonici iznad perilice"],
      ["Portal","Crni aluminijski klizni portal 180-200 cm","1 kom","Za spavaću sobu"],
      ["Kuhinja","Hrastove police + radna ploča","2 police","Bez gornjih elemenata"],
      ["Ormari","Ugradbeni ormar do stropa","1 kom","Visina 270 cm, push-open"],
    ],
    [2200, 3000, 1600, 2700]
  ),

  h1("9. DIMENZIJE NAMJEŠTAJA I OPREME"),
  p("Ključne dimenzije svih elemenata koji direktno definiraju prostor i zahtijevaju precizno usklađivanje s izvedbom."),
  h2("9.1. Kuhinja i radni elementi"),
  table(
    ["Element","Dimenzije","Materijal / napomena","Lokacija"],
    [
      ["Kuhinja s radnom plohom","300 x 60 cm","Hrast + bijela radna ploha","Desni zid"],
      ["Hrastove police","200 x 35 cm","Masivni hrast / LED rasvjeta","Ispod gornjih nivoa"],
      ["Ugradbeni ormar","270 x 240 cm","Push-open / do stropa","Prednja lijeva strana"],
      ["Radna ploča","300 x 60 cm","Kamen / kompozit","Kuhinjski radni dio"],
    ],
    [2500, 1800, 2900, 2300]
  ),
  h2("9.2. Spavaća soba, kupaonica i portal"),
  table(
    ["Element","Dimenzije","Materijal / napomena","Lokacija"],
    [
      ["Klizni aluminijski portal","180-200 cm","Crni aluminij / klizni mehanizam","Ulaz u spavaću sobu"],
      ["Walk-in tuš","90 x 120 cm","Staklena stijena, kanalica u ravnini","Kupaonica"],
      ["Bračni krevet","200 x 200 cm","Ugradbene utičnice uz krevet","Spavaća soba"],
      ["Bezpragov pod","Cijela etaža","Mikrocement / keramika 60x120","Cijeli prostor"],
    ],
    [2500, 1800, 2900, 2300]
  ),

]));

// ===================== DOPUNE PO STRUČNOJ RECENZIJI (R2) =====================
sections.push(bodyPage([
  h1("10. USKLAĐENE DIMENZIJE I NAPOMENE NA TEREN"),
  p("Vanjske/unutarnje mjere objekta jednoznačno se definiraju ovdje kako bi se izbjegla zabuna s ranijim radnim verzijama tlocrta:"),
  bullet("Unutarnje svijetle mjere objekta: 7,00 x 4,36 m (mjerodavne za sve proračune površina u ovom elaboratu)."),
  bullet("Sve pojedinačne dimenzije prostorija prikazane su u poglavlju 1. i na tlocrtu (Prilog A, zaseban ispis)."),
  bullet("Debljina kamenih bočnih zidova (50\u201360 cm) je okvirna procjena \u2013 točna debljina utvrđuje se na licu mjesta tijekom otvaranja zidova; kod starih kamenih kuća uobičajena su odstupanja od 45 do 65 cm na različitim dijelovima istog zida."),
  bullet("Prije postavljanja poda obavezno izmjeriti vlagu kamenih zidova CM metodom (karbidna metoda). Ako je vlaga povišena, prije radova na podu i pregradama provesti sanaciju vlage (npr. injektiranje ili sanacijske žbuke)."),

  h1("11. KONSTRUKTERSKI PREGLED (prije početka radova)"),
  p("Prije rušenja i sanacijskih radova obavezan je vizualni konstrukterski pregled objekta, po mogućnosti od strane ovlaštenog inženjera građevinarstva. Pregled obuhvaća:"),
  bullet("Pregled temelja \u2013 vidljiva oštećenja, slijeganje, vlaga u podnožju zidova."),
  bullet("Pregled postojećih greda \u2013 nosivost, trulež, napadnutost insektima/gljivama."),
  bullet("Pregled rogova i krovne konstrukcije \u2013 stanje drvene građe, prokišnjavanje."),
  bullet("Pregled vlage u zidovima \u2013 kapilarna vlaga, prodor oborinske vode."),
  bullet("Pregled pukotina \u2013 slijeganje, statičke pukotine vs. površinske (žbukine) pukotine."),
  p("Rezultat pregleda upisuje se kao dodatak ovom elaboratu prije narudžbe materijala za sanaciju."),

  h1("12. VENTILACIJA"),
  bullet("Kupaonica: mehanička ventilacija \u00d8100 mm s timerom (odgoda isključenja 10\u201315 min nakon paljenja svjetla)."),
  bullet("Kuhinja: napa minimalne protočnosti 500 m\u00b3/h, odvod prema van kroz kameni zid (jezgreno bušenje)."),
]));

sections.push(bodyPage([
  h1("13. ELEKTRO \u2013 RAZDJELNIK I STRUJNI KRUGOVI"),
  table(
    ["Oznaka","Krug","Napomena"],
    [
      ["R1","Rasvjeta (opća)","Kuhinja, soba, kupaonica, terasa"],
      ["R2","Utičnice (opće)","Dnevni dio, soba"],
      ["R3","Klima","Zaseban osigurač 10A, prema poglavlju 5"],
      ["R4","Kuhinjski aparati","Hladnjak, ploča, pećnica, napa"],
      ["R5","Bojler","Zaseban krug, kupaonica"],
      ["R6","Perilica rublja/posuđa","IP44 zaštićen krug, kupaonica/kuhinja"],
    ],
    [1400, 2900, 4700]
  ),
  bullet("Internet: CAT6 kabel položen od razdjelnika/ormarića do TV pozicije i do WiFi Access Pointa (centralna pozicija, npr. strop prolaza) \u2013 obavezno ugraditi u fazi grubih instalacija, prije zatvaranja Knauf pregrada."),
  bullet("Klima \u2013 odvod kondenzata: isključivo gravitacijski (bez pumpe), pad cijevi prema van min. 1,5%."),

  h1("14. STOLARIJA \u2013 TEHNIČKA SPECIFIKACIJA"),
  table(
    ["Element","Materijal","Dimenzije","Tehnički podaci"],
    [
      ["Ulazna vrata","Hrast (puno drvo/furnir)","90 x 210 cm","Toplinska propusnost U \u2248 1,3 W/m\u00b2K"],
      ["Veliki portal (soba)","Aluminij, RAL 7016 (antracit)","180\u2013200 cm, klizno-podizni","Troslojno staklo, low-E premaz"],
      ["Klizna vrata (soba, WC)","Drvo, boja hrasta","80\u2013100 x 210 cm","Šiber sustav, bez praga"],
    ],
    [2200, 2400, 2200, 2200]
  ),

  h1("15. POD \u2013 FINALNA ODLUKA"),
  p("Radi jednoznačnosti narudžbe, kao finalna podna obloga usvaja se: MIKROCEMENT u neutralnom bež-sivom tonu, u cijeloj kući bez prijelaza i pragova (uklj. kupaonicu, s odgovarajućom hidroizolacijskom podlogom ispod). Alternativa keramikom 60x120 cm napušta se radi ujednačenosti i lakše sanacije eventualnih budućih pukotina u podlozi."),
]));

sections.push(bodyPage([
  h1("16. RASVJETA \u2013 SHEMA UPRAVLJANJA"),
  table(
    ["Oznaka","Zona","Tip rasvjete"],
    [
      ["K1","Kuhinja (opća)","Ugradbeni spotovi, strop, 3000K"],
      ["K2","Stol / blagovaonica","Viseća lampa, podesiva visina"],
      ["K3","Krevet (soba)","Zidne/nočne svjetiljke, izmjenični prekidači"],
      ["K4","Kupaonica","Ugradbeni spotovi + IP44 zrcalna rasvjeta"],
      ["K5","Terasa","Vanjska zidna svjetiljka, IP65, topla bijela"],
    ],
    [1400, 2900, 4700]
  ),

  h1("17. VODOVOD \u2013 PROMJERI CIJEVI"),
  table(
    ["Instalacija","Promjer / oznaka","Primjena"],
    [
      ["Dovod hladne/tople vode","PEX 16 mm","Ogranci prema umivaoniku, perilici"],
      ["Dovod glavne vertikale","PEX 20 mm","Glavni razvod od priključka"],
      ["Odvod umivaonik/perilica","PVC \u00d850 mm","Kupaonica"],
      ["Odvod WC","PVC \u00d8110 mm","Spoj na postojeću vertikalu (majčin WC)"],
    ],
    [2800, 2400, 3800]
  ),

  h1("18. ORMAR (MASTER BEDROOM) \u2013 UNUTARNJI RASPORED"),
  bullet("Gornja zona: police za posteljinu i sezonsku odjeću (3\u20134 police)."),
  bullet("Srednja zona: vješalica (šipka) za vješanje odjeće na pleća, obostrano ili preko cijele dužine."),
  bullet("Ladice: 3\u20134 ladice s tihim zatvaranjem za rublje i sitniji sadržaj."),
  bullet("Donja zona: pletene/tekstilne košare za obuću i dodatke."),
  bullet("Sve fronte push-open, bez ručki, boja hrasta \u2013 prema poglavlju 6.2."),

  h1("19. KUHINJA \u2013 RASPORED ELEMENATA (redoslijed s lijeva na desno)"),
  bullet("FR (Hladnjak) \u2192 PEĆ (Pećnica) \u2192 PLOČA (Ploča za kuhanje) \u2192 SUDOPER \u2192 PERILICA (posuđa) \u2192 LADICE."),
  p("Redoslijed elemenata definiran je radi optimalnog radnog trokuta (hladnjak \u2013 ploča \u2013 sudoper) i uručuje se proizvođaču kuhinje kao osnova za izradu radioničkog nacrta."),
]));

sections.push(bodyPage([
  h1("20. TERASA \u2013 PROŠIRENA SPECIFIKACIJA"),
  table(
    ["Element","Specifikacija"],
    [
      ["Dimenzije","2,20 x 2,50 m, nivelirano u ravnini s unutarnjim podom"],
      ["Materijal poploče","Istarski kamen (lomljeni ili štokani), protuklizna obrada"],
      ["Nagib poda","1,5% prema van, radi odvodnje oborinske vode"],
      ["Odvodnja","Rubni kanal ili slobodan pad na teren, bez zadržavanja vode"],
      ["Rasvjeta","Vanjska zidna svjetiljka IP65 (vidi K5, poglavlje 16)"],
      ["Pergola","Drvena ili crna aluminijska konstrukcija, pomično platno ili trstika"],
      ["Priključak vode","Vanjska nepropusna slavina (za zalijevanje maslina/čišćenje)"],
      ["Električni priključak","Vanjska utičnica IP44, uz pergolu"],
      ["Hortikultura","2x maslina u terakota teglama, simetrično postavljene"],
    ],
    [2600, 6900]
  ),

  h1("21. FASADA \u2013 SPECIFIKACIJA"),
  table(
    ["Element","Specifikacija"],
    [
      ["Žbuka","Vapnena, paropropusna žbuka (prilagođena starim kamenim zidovima)"],
      ["Boja","Prirodna kamena / oker tonovi, usklađeno s postojećim pročeljem"],
      ["Fuge","Uske, ručno obrađene fuge, tradicionalni izgled"],
      ["Kamen (vidljivi dijelovi)","Očuvati i istaknuti postojeći klesani kamen oko otvora"],
      ["Čišćenje kamena","Mehaničko čišćenje (niskotlačno pranje), bez agresivnih kiselina"],
    ],
    [2600, 6900]
  ),

  h1("22. FAZE IZVOĐENJA"),
  table(
    ["Faza","Opis"],
    [
      ["1","Rušenje (stari kamin, nepotrebne pregrade)"],
      ["2","Sanacija (vlaga, konstrukterski pregled, popravak zidova)"],
      ["3","Instalacije (vodovod, kanalizacija, elektro, ventilacija \u2013 grubi razvod)"],
      ["4","Estrih (izolacija poda, hidroizolacija, cementna glazura)"],
      ["5","Knauf (pregradni zidovi kupaonice i sobe, gips-karton)"],
      ["6","Keramika / mikrocement (podna obloga, kupaonica)"],
      ["7","Bojanje (zidovi, stropovi, fasada)"],
      ["8","Kuhinja (ugradnja kuhinjskog niza i elemenata)"],
      ["9","Namještaj (ormar, krevet, stolarija, portal, terasa)"],
      ["10","Završni pregled (kontrola instalacija, čišćenje, primopredaja)"],
    ],
    [1400, 8100]
  ),

  h1("23. GRAFIČKI PRILOZI (isporučeni zasebno, izvan Worda)"),
  p("Sljedeći prilozi su izrađeni kao zasebni grafički dokumenti (HTML/PDF za ispis), kako bi se izbjegla deformacija crteža unutar Worda:"),
  bullet("Prilog A \u2013 Tlocrt (Opcija A, mjerilo 1:50, s kotama i namještajem)."),
  bullet("Prilog A2 \u2013 Tlocrt instalacija (vodovod, odvod, elektro simboli s legendom)."),
  bullet("Prilog B \u2013 Četiri fasade (prednja, stražnja, obje bočne) s označenim otvorima."),
  bullet("Prilog C \u2013 Tlocrt rušenja, tlocrt stropa/greda i detaljni nacrt terase (isporučeno u nastavku ovog kruga dorade)."),
  p("3D fotorealistični presjek interijera NIJE uključen \u2013 zahtijeva poseban 3D modelarski/render alat izvan dosega ravninskih (2D) crteža koje ovdje izrađujemo; može se naručiti kao zaseban korak kod vizualizacijskog studija ili kroz alat za generiranje slika ako ga imate na raspolaganju."),
]));

// ===================== DOPUNE R3 =====================
sections.push(bodyPage([
  h1("24. STOLARIJA \u2013 POPIS VRATA I PROZORA (TABLICA OZNAKA)"),
  p("Sve oznake odgovaraju pozicijama na tlocrtu (Prilog A). Tablica je referentna podloga za narudžbu kod stolara."),
  table(
    ["Oznaka","Naziv","Dimenzije","Materijal / napomena"],
    [
      ["V1","Ulazna vrata","90 x 210 cm","Hrast, U\u22481,3 W/m\u00b2K"],
      ["V2","Klizna vrata kupaonice","80 x 210 cm","Šiber, drvo boje hrasta"],
      ["V3","Klizna vrata sobe (prema anteroomu)","80 x 210 cm","Šiber, drvo boje hrasta"],
      ["P1","Prozor kuhinje (stražnji, postojeći)","120 x 100 cm","Zadržan postojeći otvor"],
      ["P2","Nizak prozor kupaonice (stražnji, postojeći)","cca 110 x 55 cm","Zadržan postojeći otvor, ventilacijsko krilo"],
      ["BP1","Veliki portal (spavaća soba, prednji)","180\u2013200 x 210 cm","Crni alu, RAL 7016, klizno-podizni, troslojno staklo"],
    ],
    [1000, 3400, 2200, 2500]
  ),

  h1("25. ELEKTRIČNI RAZVODNI ORMAR"),
  table(
    ["Element","Specifikacija"],
    [
      ["Lokacija ormara","Anteroom / prolaz (blizu ulaza, lako dostupno)"],
      ["Broj modula","Min. 18 modula (rezerva za proširenje)"],
      ["Glavni FID sklopnik","40A / 30mA, tip A"],
      ["Prenaponska zaštita","Kombinirani odvodnik prenapona tip 2 (obavezno za klimu/bojler)"],
      ["Osigurači (okvirno)","R1 rasvjeta 10A, R2 utičnice 16A, R3 klima 10A, R4 kuhinjski aparati 16A/20A, R5 bojler 16A, R6 perilica 16A (IP44 krug)"],
      ["Rezervni modul","Min. 2 slobodna mjesta za buduće proširenje"],
      ["CAT6 / mreža","Patch panel ili priključna kutija uz razvodni ormar, vod prema TV poziciji i WiFi AP-u"],
    ],
    [2600, 6900]
  ),

  h1("26. PODNA OBLOGA \u2013 IZVEDBENA NAPOMENA"),
  p("Sve podne površine (Kitchen & Dining, Master Bedroom, Anteroom, Bathroom) izvode se u mikrocementu, bez dilatacijskih prijelaza između prostorija \u2013 jedinstvena vizualna i funkcionalna ravnina poda kroz cijeli objekt, sukladno poglavlju 15."),

  h1("27. KLIMATIZACIJA \u2013 DETALJ IZVEDBE"),
  table(
    ["Element","Specifikacija"],
    [
      ["Vanjska jedinica \u2013 pozicija","Desni bočni zid, u visini kuhinje, na nosaču za zid (konzole)"],
      ["Unutarnja jedinica \u2013 pozicija","Master Bedroom, iznad pregradnog Knauf zida prema kuhinji, +250 cm"],
      ["Prolaz instalacije kroz zid","Jezgreno bušenje \u00d850 mm kroz Knauf pregradu (ne kroz kameni nosivi zid)"],
      ["Odvod kondenzata","Gravitacijski, bez pumpe, pad min. 1,5% prema kupaonici/vanjskom odvodu"],
      ["Napajanje","Zaseban krug R3, osigurač 10A"],
    ],
    [2600, 6900]
  ),
]));

sections.push(bodyPage([
  h1("28. ZAVRŠNI LIST \u2013 KONTROLA IZVEDBE"),
  p("Popunjava se na kraju radova, prije primopredaje objekta investitoru. Svaku stavku potvrđuje izvođač, po potrebi uz potpis investitora."),
  bullet("\u25a1 Vlaga kamenih zidova izmjerena (CM metoda)"),
  bullet("\u25a1 Estrih osušen (potvrđeno mjerenjem prije polaganja obloge)"),
  bullet("\u25a1 Hidroizolacija pregledana (pod, kupaonica, prag terase)"),
  bullet("\u25a1 Električne instalacije ispitane (mjerenje izolacije, funkcionalnost FID-a)"),
  bullet("\u25a1 Vodovod ispitan (tlačna proba)"),
  bullet("\u25a1 Kanalizacija ispitana (protočnost, brtvljenje spojeva)"),
  bullet("\u25a1 Klima testirana (hlađenje/grijanje, odvod kondenzata)"),
  bullet("\u25a1 Rasvjeta testirana (sve točke, prekidači, LED trake)"),
  bullet("\u25a1 Prozori podešeni (brtvljenje, mehanizam otvaranja)"),
  bullet("\u25a1 Vrata podešena (uklj. klizna/šiber i portal)"),
  bullet("\u25a1 Silikoni završeni (kupaonica, kuhinja, portal, terasa)"),
  bullet("\u25a1 Investitor preuzeo objekt (datum, potpis)"),

  new Paragraph({ spacing:{before:500}, children:[
    new TextRun({text:"Datum primopredaje: ______________________     Potpis izvođača: ______________________     Potpis investitora: ______________________", size:19, font:FONT})
  ]}),

  new Paragraph({ spacing:{before:600}, border:{top:{color:"999999", space:8, style:BorderStyle.SINGLE, size:6}}, children:[] }),
  new Paragraph({ spacing:{before:200}, alignment: AlignmentType.CENTER, children:[
    new TextRun({text:"KRAJ ELABORATA A (R3) \u2013 jedinstveni, finalni izvedbeni raspored \u2013 spremno za potpis i predaju izvođaču", italics:true, size:20, color:"777777", font:FONT})
  ]}),
]));

// ===================== BUILD DOC =====================
const doc = new Document({
  sections,
  styles: { default: { document: { run: { font: FONT } } } }
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync("/mnt/user-data/outputs/Izvedbeni_elaborat_Ripenda_R2.docx", buf);
  console.log("WRITTEN OK", buf.length);
}).catch(e => { console.error("ERROR", e); process.exit(1); });