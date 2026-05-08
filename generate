const pptxgen = require("pptxgenjs");
const fs = require("fs");

// Read args
const dataFile = process.argv[2];
const outputFile = process.argv[3];

if (!dataFile || !outputFile) {
  console.error("Usage: node generate.js <data.json> <output.pptx>");
  process.exit(1);
}

const data = JSON.parse(fs.readFileSync(dataFile, "utf8"));

// ── Color Palette ────────────────────────────────────────────────
const C = {
  lightGreen:  "D9F0C0",   // Slide 1 + 4 background
  lightGray:   "F0F0F0",   // Slide 2, 3, 5 background
  darkGray:    "333333",   // Slide 6 background
  yellow:      "F5E642",   // Sichtbarkeit card
  purple:      "D8B4FE",   // Traffic card
  green:       "B9F0A0",   // Vergleich card
  darkCard:    "222222",   // Dark chart cards
  white:       "FFFFFF",
  black:       "1A1A1A",
  accent:      "A855F7",   // Purple accent text
  textGray:    "555555",
  headerSmall: "888888",
};

// ── Helper: shadow ───────────────────────────────────────────────
const mkShadow = () => ({ type: "outer", color: "000000", blur: 8, offset: 3, angle: 135, opacity: 0.12 });

// ── Helper: header bar (logo area top) ──────────────────────────
function addHeader(slide, monthYear, reportTitle = "Monthly Ad Performance Report") {
  // Company name left
  slide.addText("pmc active GmbH", {
    x: 0.4, y: 0.12, w: 3, h: 0.3,
    fontSize: 11, color: C.textGray, fontFace: "Calibri", margin: 0
  });
  // Month center
  slide.addText(monthYear, {
    x: 3.5, y: 0.12, w: 3, h: 0.3,
    fontSize: 11, color: C.textGray, fontFace: "Calibri", align: "center", margin: 0
  });
  // Report title right
  slide.addText(reportTitle, {
    x: 6.5, y: 0.12, w: 3.2, h: 0.3,
    fontSize: 11, color: C.textGray, fontFace: "Calibri", align: "right", margin: 0
  });
  // Logo placeholder (green diamond shape)
  slide.addShape("pentagon", {
    x: 0.38, y: 0.48, w: 0.35, h: 0.35,
    fill: { color: "7EE8A2" }, line: { color: "7EE8A2" }
  });
}

// ── Data extraction ──────────────────────────────────────────────
const aktionsnummer  = data.aktionsnummer  || "–";
const kampagnenname  = (data.aktuelle_daten && data.aktuelle_daten.campaign_name) || "Unbekannte Kampagne";
const platform       = (data.aktuelle_daten && data.aktuelle_daten.platform)      || "Meta";
const heute          = new Date().toLocaleDateString("de-DE", { month: "long", year: "numeric" });

// Current KPIs
const impressionen   = (data.aktuelle_daten && data.aktuelle_daten.impressions)   || 0;
const klicks         = (data.aktuelle_daten && data.aktuelle_daten.clicks)        || 0;
const ctr            = (data.aktuelle_daten && data.aktuelle_daten.ctr)           || 0;
const spend          = (data.aktuelle_daten && data.aktuelle_daten.spend)         || 0;
const cpc            = (data.aktuelle_daten && data.aktuelle_daten.cpc)           || 0;

// Historical averages
const hist = (data.durchschnitte && data.durchschnitte[0]) || {};
const avgImpr  = hist.avg_impressions || 0;
const avgKlick = hist.avg_clicks      || 0;
const avgCtr   = hist.avg_ctr         || 0;
const avgSpend = hist.avg_spend       || 0;
const avgCpc   = hist.avg_cpc         || 0;

// % change helper
function pct(current, avg) {
  if (!avg || avg === 0) return "–";
  const diff = ((current - avg) / avg) * 100;
  return (diff >= 0 ? "+" : "") + diff.toFixed(0) + "%";
}

// Historical timeline for charts
const historisch = data.historische_daten || [];
const chartLabels    = historisch.map(d => d.report_date ? d.report_date.slice(5) : "");
const chartImpr      = historisch.map(d => Number(d.impressions) || 0);
const chartKlicks    = historisch.map(d => Number(d.clicks)      || 0);

// AI analysis text
const analyseRaw = data.analyse || "";
// Extract summary and recommendations sections
function extractSection(text, keywords) {
  for (const kw of keywords) {
    const idx = text.indexOf(kw);
    if (idx !== -1) {
      const after = text.slice(idx);
      const end = after.search(/\n#{1,3} /);
      return after.slice(0, end > 0 ? end : 2000).replace(/^.*\n/, "").trim();
    }
  }
  return text.slice(0, 800).trim();
}

const summary = extractSection(analyseRaw, ["Management Summary", "Zusammenfassung", "Gesamtbewertung"]);
const empfehlungen = extractSection(analyseRaw, ["Handlungsempfehlung", "Empfehlung", "Recommendation"]);

// ════════════════════════════════════════════════════════════════
//  BUILD PRESENTATION
// ════════════════════════════════════════════════════════════════
const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";   // 10" × 5.625"
pres.author  = "pmc active GmbH";
pres.title   = "Monthly Ad Performance Report";

// ── SLIDE 1: Title ───────────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.lightGreen };

  // Logo top-left
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.4, y: 0.3, w: 0.5, h: 0.5,
    fill: { color: "7EE8A2" }, line: { color: "7EE8A2" }
  });

  // Client name top-right
  s.addText("pmc active GmbH", {
    x: 5.5, y: 0.3, w: 4.1, h: 0.4,
    fontSize: 18, color: C.black, fontFace: "Calibri", align: "right", margin: 0
  });

  // Big title
  s.addText("Monthly Ad Performance", {
    x: 2.5, y: 1.6, w: 7.1, h: 1.0,
    fontSize: 52, color: C.black, fontFace: "Calibri", bold: false, margin: 0
  });
  s.addText("Report", {
    x: 2.5, y: 2.55, w: 7.1, h: 1.0,
    fontSize: 72, color: C.accent, fontFace: "Calibri", bold: false, margin: 0
  });

  // Bottom row
  s.addText("Erstellt von", {
    x: 0.4, y: 4.9, w: 1.3, h: 0.35,
    fontSize: 11, color: C.textGray, fontFace: "Calibri", margin: 0
  });
  s.addText("Kerstin Wöldering", {
    x: 1.75, y: 4.9, w: 3, h: 0.35,
    fontSize: 11, color: C.black, fontFace: "Calibri", bold: true, margin: 0
  });
  s.addText("Director Marketing & Consulting", {
    x: 4.75, y: 4.9, w: 3, h: 0.35,
    fontSize: 11, color: C.textGray, fontFace: "Calibri", margin: 0
  });
  s.addText("www.pmc-active.de", {
    x: 7.5, y: 4.9, w: 2.1, h: 0.35,
    fontSize: 11, color: C.textGray, fontFace: "Calibri", align: "right", margin: 0
  });
}

// ── SLIDE 2: Metrics ─────────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.lightGray };
  addHeader(s, heute);

  // Logo icon
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.38, y: 0.48, w: 0.35, h: 0.35,
    fill: { color: "7EE8A2" }, line: { color: "7EE8A2" }
  });

  // Section title
  s.addText("Metrics", {
    x: 0.4, y: 0.85, w: 5, h: 0.8,
    fontSize: 44, color: C.black, fontFace: "Calibri", bold: false, margin: 0
  });

  // ── Card 1: Sichtbarkeit (yellow) ──
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.3, y: 1.75, w: 2.9, h: 3.55,
    fill: { color: C.yellow }, line: { color: C.yellow }, rectRadius: 0.15,
    shadow: mkShadow()
  });
  s.addText("Sichtbarkeit", {
    x: 0.4, y: 1.9, w: 2.7, h: 0.35,
    fontSize: 14, color: C.black, fontFace: "Calibri", bold: true, align: "center", margin: 0
  });
  // Impressions row
  s.addText(Number(impressionen).toLocaleString("de-DE"), {
    x: 0.4, y: 2.35, w: 2.7, h: 0.45,
    fontSize: 28, color: C.black, fontFace: "Calibri", bold: true, align: "center", margin: 0
  });
  s.addText("Impressionen", { x: 0.4, y: 2.78, w: 1.6, h: 0.22, fontSize: 9, color: C.textGray, fontFace: "Calibri", margin: 0 });
  s.addText("Total",        { x: 2.0, y: 2.78, w: 1.1, h: 0.22, fontSize: 9, color: C.textGray, fontFace: "Calibri", align: "right", margin: 0 });
  s.addShape(pres.shapes.LINE, { x: 0.4, y: 3.02, w: 2.7, h: 0, line: { color: C.black, width: 0.5 } });

  // Spend row
  s.addText(spend.toLocaleString("de-DE", { minimumFractionDigits: 2 }) + " €", {
    x: 0.4, y: 3.1, w: 2.7, h: 0.45,
    fontSize: 24, color: C.black, fontFace: "Calibri", bold: true, align: "center", margin: 0
  });
  s.addText("Ad Spend",  { x: 0.4, y: 3.55, w: 1.6, h: 0.22, fontSize: 9, color: C.textGray, fontFace: "Calibri", margin: 0 });
  s.addText("Total",     { x: 2.0, y: 3.55, w: 1.1, h: 0.22, fontSize: 9, color: C.textGray, fontFace: "Calibri", align: "right", margin: 0 });
  s.addShape(pres.shapes.LINE, { x: 0.4, y: 3.79, w: 2.7, h: 0, line: { color: C.black, width: 0.5 } });

  // Aktionsnummer
  s.addText(aktionsnummer, {
    x: 0.4, y: 3.87, w: 2.7, h: 0.4,
    fontSize: 18, color: C.black, fontFace: "Calibri", bold: true, align: "center", margin: 0
  });
  s.addText("Aktionsnummer", { x: 0.4, y: 4.27, w: 2.7, h: 0.22, fontSize: 9, color: C.textGray, fontFace: "Calibri", align: "center", margin: 0 });

  // ── Card 2: Traffic (purple) ──
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 3.55, y: 1.75, w: 2.9, h: 3.55,
    fill: { color: C.purple }, line: { color: C.purple }, rectRadius: 0.15,
    shadow: mkShadow()
  });
  s.addText("Traffic", {
    x: 3.65, y: 1.9, w: 2.7, h: 0.35,
    fontSize: 14, color: C.black, fontFace: "Calibri", bold: true, align: "center", margin: 0
  });
  s.addText(Number(klicks).toLocaleString("de-DE"), {
    x: 3.65, y: 2.35, w: 2.7, h: 0.45,
    fontSize: 28, color: C.black, fontFace: "Calibri", bold: true, align: "center", margin: 0
  });
  s.addText("Klicks",  { x: 3.65, y: 2.78, w: 1.6, h: 0.22, fontSize: 9, color: C.textGray, fontFace: "Calibri", margin: 0 });
  s.addText("Total",   { x: 5.25, y: 2.78, w: 1.1, h: 0.22, fontSize: 9, color: C.textGray, fontFace: "Calibri", align: "right", margin: 0 });
  s.addShape(pres.shapes.LINE, { x: 3.65, y: 3.02, w: 2.7, h: 0, line: { color: C.black, width: 0.5 } });

  s.addText(ctr.toLocaleString("de-DE", { minimumFractionDigits: 2 }) + " %", {
    x: 3.65, y: 3.1, w: 2.7, h: 0.45,
    fontSize: 24, color: C.black, fontFace: "Calibri", bold: true, align: "center", margin: 0
  });
  s.addText("CTR",   { x: 3.65, y: 3.55, w: 1.6, h: 0.22, fontSize: 9, color: C.textGray, fontFace: "Calibri", margin: 0 });
  s.addText("Total", { x: 5.25, y: 3.55, w: 1.1, h: 0.22, fontSize: 9, color: C.textGray, fontFace: "Calibri", align: "right", margin: 0 });
  s.addShape(pres.shapes.LINE, { x: 3.65, y: 3.79, w: 2.7, h: 0, line: { color: C.black, width: 0.5 } });

  s.addText(cpc.toLocaleString("de-DE", { minimumFractionDigits: 2 }) + " €", {
    x: 3.65, y: 3.87, w: 2.7, h: 0.4,
    fontSize: 18, color: C.black, fontFace: "Calibri", bold: true, align: "center", margin: 0
  });
  s.addText("CPC", { x: 3.65, y: 4.27, w: 2.7, h: 0.22, fontSize: 9, color: C.textGray, fontFace: "Calibri", align: "center", margin: 0 });

  // ── Card 3: Vergleich (green) ──
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 6.8, y: 1.75, w: 2.9, h: 3.55,
    fill: { color: C.green }, line: { color: C.green }, rectRadius: 0.15,
    shadow: mkShadow()
  });
  s.addText("Vergleich Ø-Werte", {
    x: 6.9, y: 1.9, w: 2.7, h: 0.35,
    fontSize: 13, color: C.black, fontFace: "Calibri", bold: true, align: "center", margin: 0
  });

  const vcItems = [
    ["Impressionen", pct(impressionen, avgImpr)],
    ["Klicks",       pct(klicks,       avgKlick)],
    ["CTR",          pct(ctr,          avgCtr)],
    ["Ad Spend",     pct(spend,        avgSpend)],
    ["CPC",          pct(cpc,          avgCpc)],
  ];
  vcItems.forEach(([label, val], i) => {
    const yBase = 2.35 + i * 0.58;
    const isNeg = val.startsWith("-");
    const col = isNeg ? "CC3333" : "1A7A1A";
    s.addText(label, { x: 6.9, y: yBase, w: 1.5, h: 0.28, fontSize: 10, color: C.textGray, fontFace: "Calibri", margin: 0 });
    s.addText(val,   { x: 8.1, y: yBase, w: 1.5, h: 0.28, fontSize: 16, color: col, fontFace: "Calibri", bold: true, align: "right", margin: 0 });
    if (i < vcItems.length - 1) {
      s.addShape(pres.shapes.LINE, { x: 6.9, y: yBase + 0.3, w: 2.7, h: 0, line: { color: C.black, width: 0.3 } });
    }
  });
}

// ── SLIDE 3: Key Metrics Breakdown ───────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.lightGray };
  addHeader(s, heute);

  s.addText("Key Metrics Breakdown", {
    x: 0.4, y: 0.5, w: 9.2, h: 0.75,
    fontSize: 36, color: C.black, fontFace: "Calibri", bold: false, margin: 0
  });

  // Dark card: Impressionen chart
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.3, y: 1.35, w: 4.3, h: 3.95,
    fill: { color: C.darkCard }, line: { color: C.darkCard }, rectRadius: 0.15
  });
  s.addText("Impressionen", {
    x: 0.5, y: 1.5, w: 3.9, h: 0.4,
    fontSize: 16, color: C.white, fontFace: "Calibri", bold: true, margin: 0
  });

  // Line chart: Impressionen over time
  if (chartLabels.length > 0) {
    s.addChart(pres.charts.AREA, [{
      name: "Impressionen", labels: chartLabels, values: chartImpr
    }], {
      x: 0.35, y: 1.95, w: 4.2, h: 2.1,
      chartColors: ["7B5EA7"],
      chartArea: { fill: { color: C.darkCard } },
      plotArea: { fill: { color: C.darkCard } },
      catAxisLabelColor: "AAAAAA", valAxisLabelColor: "AAAAAA",
      catGridLine: { style: "none" }, valGridLine: { color: "444444", size: 0.5 },
      showLegend: false, lineSmooth: true,
      catAxisLineShow: false, valAxisLineShow: false,
    });
  }

  s.addText(Number(impressionen).toLocaleString("de-DE"), {
    x: 0.5, y: 4.1, w: 3.9, h: 0.45,
    fontSize: 22, color: C.white, fontFace: "Calibri", bold: true, margin: 0
  });
  s.addText("Impressionen gesamt", {
    x: 0.5, y: 4.55, w: 3.9, h: 0.3,
    fontSize: 10, color: "AAAAAA", fontFace: "Calibri", margin: 0
  });

  // Dark card: Klicks chart
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 4.85, y: 1.35, w: 4.3, h: 3.95,
    fill: { color: C.darkCard }, line: { color: C.darkCard }, rectRadius: 0.15
  });
  s.addText("Klicks", {
    x: 5.05, y: 1.5, w: 3.9, h: 0.4,
    fontSize: 16, color: C.white, fontFace: "Calibri", bold: true, margin: 0
  });

  if (chartLabels.length > 0) {
    s.addChart(pres.charts.AREA, [{
      name: "Klicks", labels: chartLabels, values: chartKlicks
    }], {
      x: 4.9, y: 1.95, w: 4.2, h: 2.1,
      chartColors: ["7B5EA7"],
      chartArea: { fill: { color: C.darkCard } },
      plotArea: { fill: { color: C.darkCard } },
      catAxisLabelColor: "AAAAAA", valAxisLabelColor: "AAAAAA",
      catGridLine: { style: "none" }, valGridLine: { color: "444444", size: 0.5 },
      showLegend: false, lineSmooth: true,
      catAxisLineShow: false, valAxisLineShow: false,
    });
  }

  s.addText(Number(klicks).toLocaleString("de-DE"), {
    x: 5.05, y: 4.1, w: 3.9, h: 0.45,
    fontSize: 22, color: C.white, fontFace: "Calibri", bold: true, margin: 0
  });
  s.addText("Klicks gesamt", {
    x: 5.05, y: 4.55, w: 3.9, h: 0.3,
    fontSize: 10, color: "AAAAAA", fontFace: "Calibri", margin: 0
  });
}

// ── SLIDE 4: Management Summary ──────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.lightGreen };
  addHeader(s, heute);

  s.addText("Management Summary", {
    x: 0.4, y: 0.5, w: 9.2, h: 0.75,
    fontSize: 36, color: C.black, fontFace: "Calibri", bold: false, margin: 0
  });

  // White text box
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.4, y: 1.4, w: 9.2, h: 3.85,
    fill: { color: C.white }, line: { color: C.white }, rectRadius: 0.15,
    shadow: mkShadow()
  });

  const summaryText = summary || analyseRaw.slice(0, 900);
  s.addText(summaryText, {
    x: 0.65, y: 1.6, w: 8.7, h: 3.5,
    fontSize: 12, color: C.black, fontFace: "Calibri",
    valign: "top", wrap: true, margin: 0
  });
}

// ── SLIDE 5: Empfehlungen ────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.lightGray };
  addHeader(s, heute);

  s.addText("Empfehlungen", {
    x: 0.4, y: 0.5, w: 9.2, h: 0.75,
    fontSize: 44, color: C.black, fontFace: "Calibri", bold: false, margin: 0
  });

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.4, y: 1.4, w: 9.2, h: 3.85,
    fill: { color: C.white }, line: { color: C.white }, rectRadius: 0.15,
    shadow: mkShadow()
  });

  const empfText = empfehlungen || "Bitte AI-Analyse durchführen um Empfehlungen zu erhalten.";
  s.addText(empfText, {
    x: 0.65, y: 1.6, w: 8.7, h: 3.5,
    fontSize: 12, color: C.black, fontFace: "Calibri",
    valign: "top", wrap: true, margin: 0
  });
}

// ── SLIDE 6: Danke / Kontakt ─────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.darkGray };
  addHeader(s, heute);

  // "Danke" big
  s.addText("Danke", {
    x: 0.4, y: 0.8, w: 6, h: 1.4,
    fontSize: 80, color: C.white, fontFace: "Calibri", bold: false, margin: 0
  });

  // Logo top-right
  s.addShape(pres.shapes.RECTANGLE, {
    x: 9.0, y: 0.3, w: 0.65, h: 0.65,
    fill: { color: "7EE8A2" }, line: { color: "7EE8A2" }
  });

  // Green contact card (left)
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.4, y: 2.85, w: 4.5, h: 2.45,
    fill: { color: C.green }, line: { color: C.green }, rectRadius: 0.15
  });
  s.addText([
    { text: "Kerstin Wöldering", options: { bold: true, breakLine: true } },
    { text: "Director Marketing & Consulting, kw@pmc-active.de", options: { breakLine: true } },
  ], {
    x: 0.65, y: 3.05, w: 4.0, h: 2.0,
    fontSize: 11, color: C.black, fontFace: "Calibri", valign: "top", margin: 0
  });

  // Purple info card (right)
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 5.2, y: 2.85, w: 4.5, h: 2.45,
    fill: { color: C.purple }, line: { color: C.purple }, rectRadius: 0.15
  });
  s.addText([
    { text: "Adresse",                         options: { bold: true, breakLine: true } },
    { text: "Bretonischer Ring 10, 85630 Grasbrunn", options: { breakLine: true } },
    { text: " ",                               options: { breakLine: true } },
    { text: "Website",                         options: { bold: true, breakLine: true } },
    { text: "www.pmc-active.de",               options: { breakLine: true } },
    { text: " ",                               options: { breakLine: true } },
    { text: "E-Mail",                          options: { bold: true, breakLine: true } },
    { text: "digital@pmc-active.de",           options: {} },
  ], {
    x: 5.45, y: 3.05, w: 4.0, h: 2.0,
    fontSize: 11, color: C.black, fontFace: "Calibri", valign: "top", margin: 0
  });
}

// ── Write file ───────────────────────────────────────────────────
pres.writeFile({ fileName: outputFile })
  .then(() => {
    console.log("OK: " + outputFile);
    process.exit(0);
  })
  .catch(err => {
    console.error("ERROR: " + err.message);
    process.exit(1);
  });
