const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, LevelFormat, PageBreak, ImageRun
} = require('docx');
const fs = require('fs');

// ── helpers ──────────────────────────────────────────────────────────────────

const CONTENT_W = 9026; // DXA (A4 with 2cm margins each side)
const CONTENT_PX = 642; // approx pixels for full-width image

const thickBorder  = { style: BorderStyle.SINGLE, size: 8,  color: "1F4E79" };
const thinBorder   = { style: BorderStyle.SINGLE, size: 4,  color: "CCCCCC" };
const noBorder     = { style: BorderStyle.NIL,    size: 0,  color: "FFFFFF" };
const allNoBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };
const allThin      = { top: thinBorder, bottom: thinBorder, left: thinBorder, right: thinBorder };
const allThick     = { top: thickBorder, bottom: thickBorder, left: thickBorder, right: thickBorder };

function h1(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(text)] });
}
function h2(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(text)] });
}
function h3(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun(text)] });
}
function p(text, extra = {}) {
  return new Paragraph({ spacing: { after: 120 }, ...extra, children: [new TextRun({ text })] });
}
function pItalic(text) {
  return new Paragraph({ spacing: { after: 100 }, alignment: AlignmentType.CENTER,
    children: [new TextRun({ text, italics: true, size: 18, color: "555555" })] });
}
function pNote(label, text) {
  return new Paragraph({ spacing: { after: 100, before: 60 }, indent: { left: 560 },
    children: [
      new TextRun({ text: label + ': ', bold: true, italics: true, size: 20, color: "1F4E79" }),
      new TextRun({ text, italics: true, size: 20, color: "333333" })
    ]
  });
}
function bullet(text) {
  return new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 80 },
    children: [new TextRun(text)] });
}
function space(n = 1) {
  return new Paragraph({ spacing: { after: 160 * n }, children: [] });
}
function divider() {
  return new Paragraph({
    spacing: { before: 80, after: 80 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "1F4E79", space: 1 } },
    children: []
  });
}
function pageBreak() {
  return new Paragraph({ children: [new PageBreak()] });
}

// Image helper: scale to target width preserving AR
function img(path, targetW, caption, alt = '') {
  const data = fs.readFileSync(path);
  // get dims by reading PNG header (bytes 16-24) or JPEG
  let w, h;
  if (path.endsWith('.png')) {
    w = data.readUInt32BE(16); h = data.readUInt32BE(20);
  } else {
    // JPEG: skip
    w = 800; h = 600;
  }
  const scaledH = Math.round(h / w * targetW);
  const children = [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 40 },
      children: [new ImageRun({
        type: 'png',
        data,
        transformation: { width: targetW, height: scaledH },
        altText: { title: alt || caption, description: caption, name: alt || caption }
      })]
    })
  ];
  if (caption) children.push(pItalic(caption));
  return children;
}

// Side-by-side two images in invisible table
function imgSideBySide(path1, w1, cap1, path2, w2, cap2) {
  const data1 = fs.readFileSync(path1); const data2 = fs.readFileSync(path2);
  function dims(d, p) {
    if (p.endsWith('.png')) return { w: d.readUInt32BE(16), h: d.readUInt32BE(20) };
    return { w: 800, h: 600 };
  }
  const d1 = dims(data1, path1); const d2 = dims(data2, path2);
  const h1s = Math.round(d1.h / d1.w * w1); const h2s = Math.round(d2.h / d2.w * w2);
  const cellW = Math.floor(CONTENT_W / 2);
  return [
    new Table({
      width: { size: CONTENT_W, type: WidthType.DXA },
      columnWidths: [cellW, CONTENT_W - cellW],
      rows: [new TableRow({ children: [
        new TableCell({ borders: allNoBorders, width: { size: cellW, type: WidthType.DXA },
          margins: { top: 0, bottom: 0, left: 0, right: 120 },
          children: [
            new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 },
              children: [new ImageRun({ type: 'png', data: data1,
                transformation: { width: w1, height: h1s },
                altText: { title: cap1, description: cap1, name: cap1 } })] }),
            pItalic(cap1)
          ]
        }),
        new TableCell({ borders: allNoBorders, width: { size: CONTENT_W - cellW, type: WidthType.DXA },
          margins: { top: 0, bottom: 0, left: 120, right: 0 },
          children: [
            new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 },
              children: [new ImageRun({ type: 'png', data: data2,
                transformation: { width: w2, height: h2s },
                altText: { title: cap2, description: cap2, name: cap2 } })] }),
            pItalic(cap2)
          ]
        })
      ]})]
    }),
    space()
  ];
}

function makeTable(headers, rows, colWidths) {
  const total = colWidths.reduce((a, b) => a + b, 0);
  return new Table({
    width: { size: total, type: WidthType.DXA },
    columnWidths: colWidths,
    rows: [
      new TableRow({
        tableHeader: true,
        children: headers.map((h, i) => new TableCell({
          borders: allThick,
          width: { size: colWidths[i], type: WidthType.DXA },
          shading: { fill: "1F4E79", type: ShadingType.CLEAR },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          verticalAlign: VerticalAlign.CENTER,
          children: [new Paragraph({ children: [new TextRun({ text: h, bold: true, color: "FFFFFF", size: 20 })] })]
        }))
      }),
      ...rows.map((row, ri) => new TableRow({
        children: row.map((cell, i) => new TableCell({
          borders: allThin,
          width: { size: colWidths[i], type: WidthType.DXA },
          shading: { fill: ri % 2 === 0 ? "EBF3F8" : "FFFFFF", type: ShadingType.CLEAR },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          children: [new Paragraph({ children: [new TextRun({ text: String(cell), size: 20 })] })]
        }))
      }))
    ]
  });
}

// ── caption box (gray background paragraph) ──────────────────────────────────
function captionBox(text) {
  return new Paragraph({
    spacing: { before: 60, after: 200 },
    shading: { fill: "F2F7FC", type: ShadingType.CLEAR },
    border: { left: { style: BorderStyle.SINGLE, size: 12, color: "2E75B6", space: 2 } },
    indent: { left: 360, right: 360 },
    children: [new TextRun({ text, italics: true, size: 20, color: "333333" })]
  });
}

// ── document ─────────────────────────────────────────────────────────────────
const doc = new Document({
  numbering: {
    config: [{
      reference: "bullets",
      levels: [{
        level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720, hanging: 360 } } }
      }]
    }]
  },
  styles: {
    default: { document: { run: { font: "Arial", size: 22 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 36, bold: true, font: "Arial", color: "1F4E79" },
        paragraph: { spacing: { before: 400, after: 160 }, outlineLevel: 0,
          border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: "1F4E79", space: 1 } } } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 28, bold: true, font: "Arial", color: "2E75B6" },
        paragraph: { spacing: { before: 280, after: 100 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 24, bold: true, font: "Arial", color: "2E75B6" },
        paragraph: { spacing: { before: 200, after: 60 }, outlineLevel: 2 } },
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 }, // A4
        margin: { top: 1134, right: 1134, bottom: 1134, left: 1134 }
      }
    },
    headers: {
      default: new Header({
        children: [new Paragraph({
          spacing: { after: 0 },
          border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "1F4E79", space: 1 } },
          children: [
            new TextRun({ text: "ESSA Stockholm Arlanda  –  Helikopter- och VFR-procedurer", bold: true, color: "1F4E79", size: 18 }),
            new TextRun({ text: "\tINTERNT UTKAST", color: "C00000", size: 18 })
          ],
          tabStops: [{ type: "right", position: 8826 }]
        })]
      })
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          spacing: { before: 0, after: 0 },
          border: { top: { style: BorderStyle.SINGLE, size: 6, color: "1F4E79", space: 1 } },
          children: [
            new TextRun({ text: "ESSA TWR / APP  –  Standardisering helikopter/VFR  –  v0.1  –  2026-05-21", size: 18, color: "666666" }),
            new TextRun({ text: "\tSida ", size: 18, color: "666666" }),
            new TextRun({ children: [PageNumber.CURRENT], size: 18, color: "666666" }),
          ],
          tabStops: [{ type: "right", position: 8826 }]
        })]
      })
    },
    children: [

      // ════════════════════════════════════════════════════════════════════════
      // TITLE PAGE
      // ════════════════════════════════════════════════════════════════════════
      new Paragraph({ spacing: { before: 2160, after: 320 }, children: [] }),
      new Paragraph({
        alignment: AlignmentType.CENTER, spacing: { after: 200 },
        children: [new TextRun({ text: "ESSA STOCKHOLM ARLANDA", bold: true, size: 52, font: "Arial", color: "1F4E79" })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER, spacing: { after: 80 },
        children: [new TextRun({ text: "Helikopter- och VFR-procedurer", size: 36, font: "Arial", color: "2E75B6" })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER, spacing: { after: 400 },
        children: [new TextRun({ text: "Nuläge, jämförelse och förslag till standardisering", size: 26, italics: true, color: "555555" })]
      }),
      divider(),
      space(),
      new Paragraph({
        alignment: AlignmentType.CENTER, spacing: { after: 160 },
        children: [new TextRun({ text: "INTERNT UTKAST – EJ PUBLICERAT", bold: true, size: 26, color: "C00000" })]
      }),
      space(2),
      makeTable(
        ["Fält", "Uppgift"],
        [
          ["Datum", "2026-05-21"],
          ["Status", "Internt utkast för granskning"],
          ["Avsedd mottagare", "ESSA TWR/APP chefer, LFV"],
          ["Nästa steg", "Intern granskning → SOP → AIP-publicering"],
        ],
        [3000, 6026]
      ),
      pageBreak(),

      // ════════════════════════════════════════════════════════════════════════
      // INLEDNING
      // ════════════════════════════════════════════════════════════════════════
      h1("Inledning"),

      p("Helikoptern är ett luftfartyg med unika egenskaper – den kan landa och lyfta vertikalt, hovra, och operera utan bana. Dessa egenskaper gör den ovärderlig för tidskritiska uppdrag som akutsjukvård, polisinsatser och räddningsoperationer. Samtidigt skapar de en utmaning i kontrollerat luftrum vid stora trafikflygplatser: helikoptern rör sig i lägre höjder, med lägre hastigheter och längs oregelbundna rutter jämfört med det konventionella IFR-trafikflödet."),
      space(),
      p("På en stor flygplats som Stockholm Arlanda delar helikoptern luftrum och markyta med tunga trafikflygplan i täta IFR-sekvenser. Helikoptrarna flyger VFR i de lägre skikten, korsar landningsbanor och rör sig till och från platser som saknar dedikerade taxibanor. Utan en gemensam referensram – namngivna waypoints, definierade rutter och publicerade höjdregler – uppstår en situation där varje flygning hanteras ad hoc. Det fungerar, men det är ineffektivt och introducerar onödig osäkerhet för både pilot och flygledare."),
      space(),
      p("Erfarenheter från jämförbara flygplatser visar att standardisering löser problemet. När rutter och waypoints är publicerade och kända av alla parter kan flygledaren ge en kortare klaring med tydlig förväntning om vad piloten ska göra. Piloten behöver inte be om detaljerade navigationsanvisningar för varje flygning. Hållplatser och höjdzoner är kända på förhand. Resultatet är snabbare handläggning, lägre arbetsbelastning och en högre gemensam situationsmedvetenhet – vilket direkt bidrar till flygsäkerheten."),
      space(),
      p("Vid ESSA har antalet helikopterrörelser ökat markant under senare år, drivet framför allt av polisens helikopterenhet (NFC) och ambulanshelikoptern (HEMS). Trafiken koncentreras till Apron L – känt som \"Patria\" – på sydvästra delen av flygplatsen. En ny FATO 07/25 är under uppförande på denna plats, vilket innebär att helikopteroperationerna vid ESSA snart kommer att formaliseras ytterligare."),
      space(),
      p("Ibruktagandet av FATO 07/25 är ett naturligt tillfälle att samtidigt etablera de standarder och rutter som länge saknats. Det är mot denna bakgrund som denna preliminära analys tagits fram. Dokumentet syftar till att:"),
      bullet("Beskriva det nuvarande läget vid ESSA och de trafikmönster som faktiskt används"),
      bullet("Visa hur jämförbara flygplatser i Nordeuropa hanterar helikoptertrafik genom publicerade kartor och procedurer"),
      bullet("Lägga fram ett konkret förslag på standardiserade rutter, waypoints och procedurer anpassade för ESSA"),
      space(),
      p("Dokumentet är ett internt arbetsmaterial avsett för granskning av ESSA TWR/APP och LFV, och utgör ett underlag inför eventuell publicering i AIP Sverige."),
      space(),
      divider(),
      pageBreak(),

      // ════════════════════════════════════════════════════════════════════════
      // SEKTION 1 – NULÄGE VID ESSA
      // ════════════════════════════════════════════════════════════════════════
      h1("1. Nuläge – Helikoptertrafik vid ESSA"),

      h2("1.1 Ökad helikoptertrafik"),
      p("Helikoptertrafiken vid Stockholm Arlanda har ökat markant under de senaste åren. De dominerande operatörerna är:"),
      bullet("Polisens helikopter (Nationellt Forensiskt Centrum, NFC) – opererar regelbundet med snabba, tidskritiska flygningar"),
      bullet("Ambulanshelikopter (HEMS – Helicopter Emergency Medical Services) – opererar dygnet runt med sjukvårdsuppdrag"),
      bullet("Övrig helikoptertrafik – privata och kommersiella helikoptrar"),
      space(),
      p("Trenden pekar mot fortsatt ökning i takt med att HEMS-verksamheten expanderar och helikoptern etableras som ett standardiserat transportmedel för blåljusmyndigheter."),
      space(),

      h2("1.2 Apron L – \"Patria\" och ny FATO 07/25"),
      p("Helikoptrarna opererar primärt till och från Apron L på sydvästra delen av Arlanda, känt i dagligt tal som \"Patria\". Denna plats fungerar i dag utan formellt publicerad FATO eller publicerade rutter i AIP."),
      p("En ny FATO 07/25 är under uppförande i anslutning till Apron L. FATO:n förväntas tas i drift under 2026 och kommer att formalisera helikopteroperationerna vid denna del av flygplatsen."),
      space(),

      h2("1.3 Faktiska trafikmönster"),
      p("Nedanstående bilder visar spårloggar från polisens och ambulansens helikoptrar under en representativ tidsperiod. Trots att inga rutter är publicerade framträder tydliga mönster:"),
      space(),

      // Full track images side by side
      ...imgSideBySide(
        'ESSA_Patria_POL.png', 300, 'Figur 1a: Polisens helikopter (NFC) – spårloggar inom ESSA CTR',
        'ESSA_Patria_DFL.png', 300, 'Figur 1b: Ambulanshelikopter (HEMS/DFL) – spårloggar inom ESSA CTR'
      ),

      p("Fyra huvudsakliga korridorer är urskiljbara:"),
      bullet("Öst – mot Markim/Vada (tät korridor öster om flygplatsen)"),
      bullet("Sydost – mot Vallentuna"),
      bullet("Nord – mot Skedviken/Odensala"),
      bullet("Väst – mot Vassunda, via TWR och korsning av RWY 19R"),
      space(),
      p("Inzoomning på Apron L visar hur i princip all trafik har sin origo/destination i exakt samma punkt – den nuvarande uppställningsplatsen vid Patria:"),
      space(),

      // Zoom side by side
      ...imgSideBySide(
        'img_patria_pol_zoom.png', 290, 'Figur 2a: Zoom Apron L – polisens helikopter, tydlig konvergering',
        'img_patria_dfl_zoom.png', 290, 'Figur 2b: Zoom Apron L – ambulanshelikopter, origo vid Patria'
      ),

      captionBox("Samtliga spårloggar konvergerar vid en och samma punkt – Apron L / Patria. Den nya FATO 07/25 planeras på denna position. Trots det tydliga trafikmönstret saknas i dag publicerade rutter eller waypoints."),

      h2("1.4 Nuvarande AIP-dokumentation"),
      p("ESSA:s nuvarande publicerade kartmaterial inkluderar en standard VAC-karta i skala 1:250 000. Denna karta:"),
      bullet("Visar inga helikopterrutter eller waypoints"),
      bullet("Saknar Apron L / FATO"),
      bullet("Saknar höjdanvisningar för helikoptrar"),
      bullet("Ger inga VFR-specifika landmarks (kyrkor, sjöar, samhällen) som visuella referenspunkter"),
      bullet("Innehåller inga procedurer för landningsbanakorsning specifikt för helikoptrar"),
      space(),
      p("Piloten och ATCO delar för tillfället inga gemensamma förväntningar på rutter eller höjder utöver vad som ges i den enskilda klaringen."),
      pageBreak(),

      // ════════════════════════════════════════════════════════════════════════
      // SEKTION 2 – JÄMFÖRELSE MED GRANNFLYGPLATSER
      // ════════════════════════════════════════════════════════════════════════
      h1("2. Hur jämförbara flygplatser hanterar helikoptertrafik"),

      p("Flera nordeuropeiska flygplatser med liknande helikoptertrafik och FATO har publicerat dedikerade kartor och procedurer i AIP. Gemensamt för dessa är:"),
      bullet("Kartor i stor skala (1:17 500 – 1:50 000) med VFR-relevanta landmarks synliga"),
      bullet("Namngivna waypoints och hållplatser (holding points) som pilot och ATCO kan referera till"),
      bullet("Definierade standardrutter med höjder per sträcka"),
      bullet("Skriftliga procedurer för landningsbanakorsning och kommunikationsavbrott (NORDO)"),
      space(),
      p("Nedan följer tre konkreta exempel."),
      space(),

      // ── 2.1 ENGM ────────────────────────────────────────────────────────────
      h2("2.1 ENGM Oslo Gardermoen – Helikopterkarta"),
      p("Oslo Gardermoen publicerar en dedikerad helikopterkarta i skala 1:17 500 (AIP Norge AD 2 ENGM 6-2). Kartan är explicit VFR-anpassad och visar:"),
      bullet("Namngivna waypoints: ROUNDY och GATY – tydligt markerade med namn och position"),
      bullet("Helikopter no-fly zone (markerad i rött) centralt på flygplatsen"),
      bullet("MAX altitude-cirklar: tre zoner med maximialtitud (MAX 500 ft, MAX 700 ft, MAX 1200 ft)"),
      bullet("Riktningspilar och numeriska höjdmarkeringar längs rutterna"),
      bullet("Kurvor och hållplatser inritade med namn"),
      space(),
      ...img('img_engm.png', Math.min(580, CONTENT_PX),
        'Figur 3: ENGM Oslo – Dedikerad helikopterkarta 1:17 500 (AIP Norge AD 2 ENGM 6-2). Namngivna waypoints ROUNDY och GATY, MAX altitude-zoner och heli no-fly zone.',
        'ENGM helikopterkarta'),
      p("Textuella procedurer (AIP Norge AD 2.22) specificerar bl.a.:"),
      pNote("Korsning", "\"Crossing of any runway requires a specific ATC clearance\" – ingen helikopter korsar bana utan uttrycklig klaring, oavsett om det finns en publicerad rutt eller ej."),
      space(),

      // ── 2.2 LFMN ────────────────────────────────────────────────────────────
      h2("2.2 LFMN Nice Côte d'Azur – Detaljerade helikopterprocedurer"),
      p("Nice är ett av de mest välutbyggda exemplen i Europa. AIP France innehåller totalt sex separata ATT-kartor (Approach and Traffic Track charts) för helikoptrar, plus två sidor textuella procedurer. Kartorna visar:"),
      bullet("Färgkodade rutter: E1/E2 (öst), S1/S2 (syd), N1/N2 (nord), W1/W2 (väst) – varje riktning har en dedikerad färg"),
      bullet("Höjd per sträcka explicit angiven i foten (t.ex. \"ALT 500\", \"ALT 800\")"),
      bullet("Stadsnamn och visuella landmarks synliga: Antibes, Cagnes-sur-Mer, La Trinité"),
      bullet("14 FATO-platser namngivna och koordinatsatta"),
      bullet("Ruttlinjer kurvar runt restriktioner och landningsbanor"),
      space(),
      ...img('img_lfmn_p3.png', 440,
        'Figur 4: LFMN Nice – Helikopter-VAC med alla rutter inlagda. Notera färgkodade korridorer, höjdmarkeringar och synliga städer/kuster som VFR-referens.',
        'LFMN helikopter VAC'),
      space(),
      ...img('img_lfmn2_p1.png', 440,
        'Figur 5: LFMN Nice – Specifik helikopter-inflygningskarta med detaljerade ruttlinjer, höjder och namngivna punkter.',
        'LFMN helikopter inflygning'),
      p("De textuella procedurerna (AD 3 LFMN TXT 01–02) beskriver bl.a.:"),
      bullet("Vilken frekvens som används för begäran om helikopterklar"),
      bullet("Hållprocedurer vid varje holding point"),
      bullet("NORDO-procedur specifik för helikoptrar"),
      bullet("Hur trafik koordineras mellan TWR och APP vid helikopteiankomst"),
      space(),

      // ── 2.3 EFHK ────────────────────────────────────────────────────────────
      h2("2.3 EFHK Helsingfors Vantaa – VFR Copter Routes"),
      p("Helsinki Vantaa publicerar en \"VFR Copter Routes\"-karta (EFHK AD 2.14-3). Kartan är anpassad för VFR-helikoptertrafik och visar flygplatsen i sitt omgivande landskap:"),
      bullet("9 namngivna rapporteringspunkter: TELGI, TUUSU, KEIMO, LINTU, NUPPU, VEJKA, KOIVU, POHJA, VIHTI"),
      bullet("Definerade holding-mönster inritade direkt på kartan vid varje punkt"),
      bullet("FATO-area i separat inzoomad ruta (inset) med tydlig markering"),
      bullet("Städer och orter synliga: Tuusula, Nurmijärvi, Sipoo – ger VFR-navigatörer visuella referenser"),
      bullet("Vattendrag och skogsytor i bakgrunden som naturliga landmärken"),
      bullet("Tabell med koordinater för alla rapporteringspunkter"),
      space(),
      ...img('img_efhk.png', CONTENT_PX,
        'Figur 6: EFHK Helsingfors – VFR Copter Routes (EFHK AD 2.14-3). Namngivna reporting points med holding-mönster, FATO-inset och stadsnamn som VFR-referens.',
        'EFHK VFR Copter Routes'),
      space(),

      // ── Sammanfattning jämförelse ────────────────────────────────────────────
      h2("2.4 Sammanfattning – Vad grannflygplatserna publicerar"),
      makeTable(
        ["Funktion", "ESSA (nu)", "ENGM Oslo", "LFMN Nice", "EFHK HEL"],
        [
          ["Dedikerad helikopterkarta",       "–",          "1:17 500",    "6 kartor",    "1 karta"],
          ["Namngivna waypoints",             "–",          "ROUNDY/GATY", "12+ st",      "9 st"],
          ["Standardrutter med höjder",       "–",          "Ja",          "Ja/leg",      "Ja"],
          ["VFR-landmarks synliga på karta",  "–",          "Ja",          "Ja",          "Ja"],
          ["FATO visad på karta",             "–",          "Ja",          "14 platser",  "Ja (inset)"],
          ["Banakorsning – skriftlig proc.",  "–",          "Explicit ATC","Ja",          "Ja"],
          ["NORDO-procedur för heli",         "–",          "Ja",          "Ja",          "Ja"],
          ["Hållplatser definierade",         "Informella", "Ja",          "Ja",          "Ja"],
        ],
        [3400, 1100, 1300, 1300, 1926]
      ),
      space(),
      captionBox("ESSA uppfyller i dag ingen av de åtta grundläggande funktioner som grannflygplatserna publicerar. Behovet av standardisering är tydligt."),
      pageBreak(),

      // ════════════════════════════════════════════════════════════════════════
      // SEKTION 3 – FÖRSLAG TILL STANDARDISERING
      // ════════════════════════════════════════════════════════════════════════
      h1("3. Förslag – Standardisering av helikopter/VFR vid ESSA"),

      p("Baserat på trafikmönstren som visas i avsnitt 1 och de publicerade lösningarna i avsnitt 2 föreslås en stegvis standardisering av helikopter- och VFR-trafiken vid ESSA. Arbetet bör samordnas med ibruktagandet av den nya FATO 07/25 vid Apron L."),
      space(),

      h2("3.1 Förbättrad kartografi – En dedikerad helikopterkarta"),
      p("Det viktigaste första steget är att ta fram en dedikerad helikopterkarta för ESSA, likt ENGM och EFHK. Kartan bör:"),
      bullet("Ha en skala anpassad för helikoptrar (förslag: 1:50 000 eller bättre)"),
      bullet("Visa Apron L / FATO 07/25 tydligt"),
      bullet("Inkludera VFR-landmarks: kyrkor, sjöar (Fysingen, Garnsviken, Valloxen), samhällen (Märsta, Sigtuna, Vallentuna)"),
      bullet("Rita ut definierade standardrutter som namngivna linjer"),
      bullet("Markera alla namngivna waypoints och hållplatser"),
      bullet("Visa MAX altitude-zoner (inre/yttre zon)"),
      space(),

      h2("3.2 Namngivna referenspunkter"),
      p("Följande punkter föreslås namnges och publiceras:"),
      makeTable(
        ["Namn", "Typ", "Koordinat (WGS-84)", "Anmärkning"],
        [
          ["MARKIM",    "VFR Holding",    "59°36.25'N / 18°04.09'E", "Vänstervarv. Markims kyrka = cirkelns västra punkt"],
          ["ODENSALA",  "VFR Holding",    "59°39.35'N / 17°50.85'E", "Vänstervarv. Odensala kyrka = cirkelns östra punkt"],
          ["TWR",       "VFR Holding*",   "59°39.0'N / 17°55.7'E",   "*Förslag – formalisering av informell hållplats vid TWR-byggnaden"],
          ["VADA",      "VFR Reporting",  "59°35.44'N / 18°12.88'E", "Östlig CTR-gräns"],
          ["VASSUNDA",  "VFR Reporting",  "59°42.49'N / 17°44.71'E", "Västlig CTR-gräns"],
          ["SKEDVIKEN", "VFR Reporting",  "59°46.44'N / 18°15.30'E", "Nordlig CTR-gräns"],
          ["VALLENTUNA","VFR Reporting",  "59°30.73'N / 18°00.98'E", "Sydostlig CTR-gräns"],
        ],
        [1900, 1700, 2500, 3926]
      ),
      space(),

      h2("3.3 Standardrutter – syfte och filosofi"),
      p("Föreslagna avgångs- och ankomstrutter är standarder och förväntningsbilder, inte rigida klareringar."),
      space(),
      new Paragraph({
        spacing: { before: 80, after: 200 },
        shading: { fill: "E8F0F8", type: ShadingType.CLEAR },
        border: {
          top:    { style: BorderStyle.SINGLE, size: 6, color: "1F4E79", space: 2 },
          bottom: { style: BorderStyle.SINGLE, size: 6, color: "1F4E79", space: 2 },
          left:   { style: BorderStyle.SINGLE, size: 18, color: "1F4E79", space: 4 },
          right:  { style: BorderStyle.NIL, size: 0, color: "FFFFFF" }
        },
        indent: { left: 440, right: 440 },
        children: [new TextRun({
          text: "Standardrutterna anger förväntad väg och höjd. ATC kan alltid ge avvikande klaring vid behov – t.ex. vid trafikbelastning, väder eller prioriterade flygningar. Syftet är att pilot och ATCO delar grundläggande förväntningar om vad som ska hända, vilket möjliggör snabbare och säkrare handläggning av helikoptertrafiken till och från Patria.",
          size: 22, italics: false, color: "1F4E79"
        })]
      }),
      space(),

      h3("Rutt Öst – FATO ↔ VADA (via MARKIM)"),
      makeTable(["Steg", "Beskrivning", "Höjd"],
        [
          ["Avgång", "FATO → Markims kyrka (visuell ref. för MARKIM holding) → östlig kurs → VADA", "Stiga till 1 000–1 500 ft MSL"],
          ["Ankomst", "VADA → MARKIM → sista inflygning FATO från öster. Håll vid MARKIM om ATC instruerar.", "≤1 500 ft MSL"],
        ], [1200, 5800, 2026]),
      space(),

      h3("Rutt Sydost – FATO ↔ VALLENTUNA (via MARKIM)"),
      makeTable(["Steg", "Beskrivning", "Höjd"],
        [
          ["Avgång", "FATO → Markims kyrka → sydostlig kurs → VALLENTUNA", "Stiga till 1 000–1 500 ft MSL"],
          ["Ankomst", "VALLENTUNA → MARKIM → sista inflygning FATO från sydost. Håll vid MARKIM om ATC.", "≤1 500 ft MSL"],
        ], [1200, 5800, 2026]),
      space(),

      h3("Rutt Nord – FATO ↔ SKEDVIKEN (via ODENSALA)"),
      makeTable(["Steg", "Beskrivning", "Höjd"],
        [
          ["Avgång", "FATO → nordlig kurs via Odensala kyrka (visuell ref. ODENSALA holding) → SKEDVIKEN", "Stiga till 1 000–2 500 ft MSL"],
          ["Ankomst", "SKEDVIKEN → sydlig kurs → ODENSALA → sista inflygning FATO. Håll vid ODENSALA om ATC.", "≤3 000 ft MSL"],
        ], [1200, 5800, 2026]),
      space(),

      h3("Rutt Väst – FATO ↔ VASSUNDA (via TWR, korsning RWY 19R)"),
      makeTable(["Steg", "Beskrivning", "Höjd"],
        [
          ["Avgång", "FATO → västlig kurs → TWR holding point → invänta banakorsningsklaring → korsa RWY 19R → VASSUNDA", "≤1 500 ft MSL"],
          ["Ankomst", "VASSUNDA → östlig kurs → ODENSALA (om sekvens) → TWR holding → banakorsningsklaring → FATO", "≤3 000 ft MSL, håll vid TWR"],
        ], [1200, 5800, 2026]),
      space(),

      h2("3.4 Korsning av landningsbana"),
      p("I linje med ENGM-principen föreslås att följande publiceras explicit:"),
      pNote("Regel", "Helikoptrar ska inte korsa någon landningsbana vid ESSA utan uttrycklig ATC-klaring, oavsett om publicerad rutt föreligger."),
      space(),
      p("För RWY 19R/01L (den bana som västrutten korserar) föreslås TWR holding point som standardiserad väntplats för helikoptrar i avvaktan på korsningsklaring."),
      space(),

      h2("3.5 Höjdregler"),
      makeTable(
        ["Zon", "Område", "Max höjd MSL", "Min höjd MSL"],
        [
          ["Inre zon",  "FATO ↔ MARKIM / ODENSALA",     "1 500 ft",  "335 ft (radargräns)"],
          ["Yttre zon", "MARKIM/ODENSALA ↔ CTR-gräns",  "3 000 ft",  "335 ft (radargräns)"],
        ],
        [1700, 3400, 2000, 1926]
      ),
      space(),

      h2("3.6 Kommunikationsfel (NORDO)"),
      p("Specifik NORDO-procedur för helikoptrar publiceras. Grundprincip: squawk 7600, försök nå ATC på alla frekvenser inkl. 121.5 MHz, undvik alla landningsbanor, direkt inflygning till FATO, titta efter ljussignaler från TWR."),
      space(),

      h2("3.7 Föreslagna nästa steg"),
      makeTable(
        ["Prioritet", "Åtgärd", "Tidsram"],
        [
          ["Hög",    "Intern granskning av detta dokument, ESSA TWR/APP + LFV",      "Q2 2026"],
          ["Hög",    "Formell definition av TWR holding point (namn + koordinat)",   "Q2 2026"],
          ["Hög",    "Bekräfta FATO 07/25 koordinat när bygget är klart",            "2026"],
          ["Medium", "Beställ dedikerad helikopterkarta 1:50 000 (LFV Karttjänst)",  "Q3 2026"],
          ["Medium", "AIP SUP med nya procedurer och waypoints",                      "Q3–Q4 2026"],
          ["Låg",    "Permanent AIP-sektion AD 2 ESSA (VAC Heli el. AD 2.24)",      "2027"],
        ],
        [1200, 5600, 2226]
      ),
      space(2),
      divider(),
      new Paragraph({
        alignment: AlignmentType.CENTER, spacing: { before: 240 },
        children: [new TextRun({ text: "Slut på dokument – ESSA Helikopter/VFR SOP v0.1 – Internt utkast", italics: true, size: 18, color: "888888" })]
      }),
    ]
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync("ESSA_Helikopter_SOP.docx", buf);
  console.log("OK: ESSA_Helikopter_SOP.docx");
}).catch(err => { console.error("FEL:", err.message); process.exit(1); });
