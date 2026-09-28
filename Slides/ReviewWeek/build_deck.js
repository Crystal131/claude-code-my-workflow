// build_deck.js — Review Week deck (descriptive statistics)
// Run: NODE_PATH=<dir with pptxgenjs>/node_modules node build_deck.js
// Example data = 20 students, prior_ai_use_1to5 (see ReviewWeek_SD_Excel.xlsx).
const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_16x9"; // 10 x 5.625 in
pres.title = "Review Week: Descriptive Statistics";

const C = {
  ink: "17313B", teal: "1F4E5A", mint: "E6F2EF", sea: "3C8D93",
  coral: "E4572E", gold: "F3B61F", gray: "5B6770", light: "F4F7F8", white: "FFFFFF",
};
const HF = "Cambria", BF = "Calibri";

// ---------- helpers ----------
function title(s, text, sub) {
  s.addText(text, { x: 0.5, y: 0.3, w: 9, h: 0.65, fontFace: HF, fontSize: 30, bold: true, color: C.teal, margin: 0, isTextBox: true });
  if (sub) s.addText(sub, { x: 0.5, y: 0.93, w: 9, h: 0.35, fontFace: BF, fontSize: 14, italic: true, color: C.gray, margin: 0, isTextBox: true });
}
function badge(s, x, y, label, fill) {
  s.addShape(pres.shapes.OVAL, { x, y, w: 0.5, h: 0.5, fill: { color: fill || C.teal }, line: { color: fill || C.teal } });
  s.addText(label, { x, y, w: 0.5, h: 0.5, fontFace: BF, fontSize: 16, bold: true, color: C.white, align: "center", valign: "middle", margin: 0, isTextBox: true });
}
function card(s, x, y, w, h, fill) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.08, fill: { color: fill || C.light }, line: { color: fill || C.light } });
}
function code(s, text, x, y, w, h, size) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.05, fill: { color: C.ink }, line: { color: C.ink } });
  s.addText(text, { x: x + 0.12, y, w: w - 0.24, h, fontFace: "Courier New", fontSize: size || 13, color: "E8F6F3", valign: "middle", margin: 0, isTextBox: true });
}
function darkSlide(kicker, big, sub) {
  const s = pres.addSlide();
  s.background = { color: C.teal };
  s.addText(kicker, { x: 0.7, y: 1.5, w: 8.6, h: 0.4, fontFace: BF, fontSize: 16, bold: true, color: C.gold, charSpacing: 3, margin: 0, isTextBox: true });
  s.addText(big, { x: 0.7, y: 1.95, w: 8.6, h: 1.3, fontFace: HF, fontSize: 40, bold: true, color: C.white, margin: 0, isTextBox: true });
  if (sub) s.addText(sub, { x: 0.7, y: 3.3, w: 8.6, h: 0.8, fontFace: BF, fontSize: 18, color: "CFE6E2", margin: 0, isTextBox: true });
  return s;
}
const th = (t) => ({ text: t, options: { bold: true, color: C.white, fill: { color: C.teal }, align: "center" } });
const td = (t, o) => ({ text: String(t), options: Object.assign({ align: "center" }, o || {}) });

// ---------- data ----------
const freq = [2, 3, 4, 7, 4]; // counts of ratings 1..5, n = 20
const n = 20, mean = 3.4, ss = 30.8, variance = ss / (n - 1), sd = Math.sqrt(variance);
const f3 = (v) => (Math.round(v * 1000) / 1000).toFixed(3);
const f2 = (v) => v.toFixed(2);

// 1. Title
{
  const s = darkSlide("WEEK OF REVIEW", "Describing Data:\nCenter, Spread & z-Scores", "Housekeeping  ·  Standard deviation in Excel  ·  Practice with Dataset 1");
  s.addNotes("Review week. Three parts: (1) housekeeping about the midterm, (2) doing standard deviation step-by-step in Excel, (3) practice problems using prior_ai_use_1to5 from Dataset 1.");
}

// 2. Agenda
{
  const s = pres.addSlide();
  title(s, "Today's Plan");
  const items = [
    ["1", "Housekeeping", "What to expect on the midterm (formula sheet provided)"],
    ["2", "Excel: Standard Deviation", "Variance → SD → z-score → plot, step by step"],
    ["3", "Review with Dataset 1", "Measurement scales · mean, median, mode · shape · SD · z"],
  ];
  items.forEach(([k, h, d], i) => {
    const y = 1.35 + i * 1.3;
    card(s, 0.5, y, 9, 1.05, i === 1 ? C.mint : C.light);
    badge(s, 0.8, y + 0.28, k, i === 2 ? C.coral : C.teal);
    s.addText(h, { x: 1.6, y: y + 0.15, w: 7.6, h: 0.4, fontFace: HF, fontSize: 20, bold: true, color: C.ink, margin: 0, isTextBox: true });
    s.addText(d, { x: 1.6, y: y + 0.55, w: 7.6, h: 0.35, fontFace: BF, fontSize: 14, color: C.gray, margin: 0, isTextBox: true });
  });
}

// 3. Housekeeping: midterm
{
  const s = pres.addSlide();
  title(s, "Housekeeping: The Midterm", "You do NOT need to memorize formulas; I will provide them");
  card(s, 0.5, 1.5, 4.3, 3.6, C.mint);
  s.addText("On your formula sheet", { x: 0.75, y: 1.65, w: 3.9, h: 0.4, fontFace: HF, fontSize: 18, bold: true, color: C.teal, margin: 0, isTextBox: true });
  s.addText([
    { text: "Mean:  x̄ = Σx / n", options: { bullet: true, breakLine: true } },
    { text: "Variance:  s² = Σ(x − x̄)² / (n − 1)", options: { bullet: true, breakLine: true } },
    { text: "Standard deviation:  s = √s²", options: { bullet: true, breakLine: true } },
    { text: "z-score:  z = (x − x̄) / s", options: { bullet: true } },
  ], { x: 0.75, y: 2.15, w: 3.95, h: 2.7, fontFace: BF, fontSize: 15, color: C.ink, paraSpaceAfter: 10, valign: "top", margin: 0, isTextBox: true });

  card(s, 5.2, 1.5, 4.3, 3.6, C.light);
  s.addText("What YOU need to know", { x: 5.45, y: 1.65, w: 3.9, h: 0.4, fontFace: HF, fontSize: 18, bold: true, color: C.coral, margin: 0, isTextBox: true });
  s.addText([
    { text: "Which formula fits the question", options: { bullet: true, breakLine: true } },
    { text: "How to plug in numbers and show your steps", options: { bullet: true, breakLine: true } },
    { text: "How to interpret the answer in words", options: { bullet: true, breakLine: true } },
    { text: "Identify measurement scales and distribution shape", options: { bullet: true } },
  ], { x: 5.45, y: 2.15, w: 3.95, h: 2.7, fontFace: BF, fontSize: 15, color: C.ink, paraSpaceAfter: 10, valign: "top", margin: 0, isTextBox: true });
  s.addNotes("Emphasize: formulas are provided. The skill being tested is choosing the right formula, computing correctly, and interpreting. Edit the formula list to match the actual sheet you'll hand out.");
}

// 4. Section divider: Excel
darkSlide("PART 2", "Standard Deviation in Excel", "Five steps you can follow for any variable")
  .addNotes("Open ReviewWeek_SD_Excel.xlsx (sheet 'SD Steps') and follow along live.");

// 5. Excel: set-up + overview of steps
{
  const s = pres.addSlide();
  title(s, "The Set-Up: One Column per Step", "Data in column B (rows 2–21); results in column H");
  const rows = [
    [th(""), th("A"), th("B"), th("C"), th("D"), th("E")],
    [td("1", { bold: true }), td("student_id"), td("prior_ai_use_1to5"), td("x − mean"), td("(x − mean)²"), td("z-score")],
    [td("2", { bold: true }), td("1"), td("4"), td("=B2-$H$2", { fontFace: "Courier New" }), td("=C2^2", { fontFace: "Courier New" }), td("=(B2-$H$2)/$H$8", { fontFace: "Courier New" })],
    [td("3", { bold: true }), td("2"), td("5"), td("=B3-$H$2", { fontFace: "Courier New" }), td("=C3^2", { fontFace: "Courier New" }), td("=(B3-$H$2)/$H$8", { fontFace: "Courier New" })],
    [td("…", { bold: true }), td("…"), td("…"), td("fill down ↓"), td("fill down ↓"), td("fill down ↓")],
  ];
  s.addTable(rows, { x: 0.5, y: 1.45, w: 9, colW: [0.5, 1.2, 1.9, 1.7, 1.7, 2.0], fontFace: BF, fontSize: 12, color: C.ink, border: { type: "solid", pt: 0.75, color: "C9D6D9" }, fill: { color: C.white }, rowH: 0.36 });
  card(s, 0.5, 3.55, 9, 1.55, C.mint);
  s.addText("Tip: the $ signs", { x: 0.75, y: 3.68, w: 8.5, h: 0.35, fontFace: HF, fontSize: 16, bold: true, color: C.teal, margin: 0, isTextBox: true });
  s.addText("$H$2 is an absolute reference: it always points to the mean, even when you drag the formula down to row 21. Without the $ signs, Excel would shift it to H3, H4, … and your answers would be wrong.", { x: 0.75, y: 4.05, w: 8.5, h: 0.95, fontFace: BF, fontSize: 14, color: C.ink, margin: 0, valign: "top", isTextBox: true });
}

// 6. Excel: variance
{
  const s = pres.addSlide();
  title(s, "Step A: Calculate the Variance", "s² = Σ(x − x̄)² / (n − 1)");
  const steps = [
    ["1", "Find the mean", "H2:  =AVERAGE(B2:B21)", "3.400"],
    ["2", "Deviation from the mean", "C2:  =B2-$H$2   (fill down)", "4 − 3.4 = 0.6"],
    ["3", "Square each deviation", "D2:  =C2^2   (fill down)", "0.6² = 0.36"],
    ["4", "Add up the squares (SS)", "H6:  =SUM(D2:D21)", "30.800"],
    ["5", "Divide by n − 1", "H7:  =H6/(H5-1)", "30.8 / 19 = " + f3(variance)],
  ];
  steps.forEach(([k, h, f, r], i) => {
    const y = 1.45 + i * 0.66;
    badge(s, 0.5, y + 0.02, k, i === 4 ? C.coral : C.teal);
    s.addText(h, { x: 1.15, y, w: 2.95, h: 0.55, fontFace: BF, fontSize: 15, bold: true, color: C.ink, valign: "middle", margin: 0, isTextBox: true });
    code(s, f, 4.15, y + 0.04, 3.25, 0.47, 12);
    s.addText(r, { x: 7.55, y, w: 1.95, h: 0.55, fontFace: BF, fontSize: 13, color: C.teal, bold: true, valign: "middle", margin: 0, isTextBox: true });
  });
  s.addText([
    { text: "Check your work: ", options: { bold: true, color: C.coral } },
    { text: "=VAR.S(B2:B21) should give the same " + f3(variance) + "  (VAR.S = sample; VAR.P = population)", options: { color: C.ink } },
  ], { x: 0.5, y: 4.85, w: 9, h: 0.4, fontFace: BF, fontSize: 13, margin: 0, isTextBox: true });
  s.addNotes("H5 holds n: =COUNT(B2:B21) = 20. Stress why n − 1: we're estimating from a sample. Show that doing it by hand matches VAR.S.");
}

// 7. Excel: SD and z
{
  const s = pres.addSlide();
  title(s, "Step B: Standard Deviation & z-Score");
  // SD column
  card(s, 0.5, 1.2, 4.3, 3.95, C.light);
  s.addText("Standard deviation", { x: 0.75, y: 1.35, w: 3.9, h: 0.4, fontFace: HF, fontSize: 18, bold: true, color: C.teal, margin: 0, isTextBox: true });
  s.addText("Square root of the variance", { x: 0.75, y: 1.75, w: 3.9, h: 0.3, fontFace: BF, fontSize: 13, italic: true, color: C.gray, margin: 0, isTextBox: true });
  code(s, "H8:  =SQRT(H7)", 0.75, 2.15, 3.8, 0.45, 13);
  code(s, "check: =STDEV.S(B2:B21)", 0.75, 2.7, 3.8, 0.45, 13);
  s.addText(f3(sd), { x: 0.75, y: 3.3, w: 3.8, h: 0.9, fontFace: HF, fontSize: 48, bold: true, color: C.teal, align: "center", margin: 0, isTextBox: true });
  s.addText("Ratings typically sit about 1.27 points from the mean", { x: 0.75, y: 4.25, w: 3.8, h: 0.6, fontFace: BF, fontSize: 13, color: C.ink, align: "center", margin: 0, isTextBox: true });
  // z column
  card(s, 5.2, 1.2, 4.3, 3.95, C.mint);
  s.addText("z-score", { x: 5.45, y: 1.35, w: 3.9, h: 0.4, fontFace: HF, fontSize: 18, bold: true, color: C.coral, margin: 0, isTextBox: true });
  s.addText("How many SDs a value is from the mean", { x: 5.45, y: 1.75, w: 3.9, h: 0.3, fontFace: BF, fontSize: 13, italic: true, color: C.gray, margin: 0, isTextBox: true });
  code(s, "E2:  =(B2-$H$2)/$H$8", 5.45, 2.15, 3.8, 0.45, 13);
  code(s, "or:  =STANDARDIZE(B2,$H$2,$H$8)", 5.45, 2.7, 3.8, 0.45, 12);
  s.addText([
    { text: "Student 1 rated 4:", options: { bold: true, breakLine: true } },
    { text: "z = (4 − 3.4) / 1.273 = 0.47", options: { breakLine: true } },
    { text: "→ about half an SD above the mean", options: { italic: true, color: C.teal } },
  ], { x: 5.45, y: 3.3, w: 3.8, h: 1.5, fontFace: BF, fontSize: 14, color: C.ink, valign: "top", margin: 0, isTextBox: true });
}

// 8. Excel: plot
{
  const s = pres.addSlide();
  title(s, "Step C: Plot the Distribution");
  const steps = [
    ["1", "Make a frequency table", "Ratings 1–5 in J2:J6; in K2:  =COUNTIF($B$2:$B$21,J2), fill down"],
    ["2", "Insert the chart", "Select J1:K6 → Insert → Column chart (Histogram for continuous data)"],
    ["3", "Label it", "Chart Elements (+) → Axis Titles and Chart Title; gap width ≈ 40%"],
    ["4", "Mark the mean", "Add a note or line at x̄ = 3.4, then describe the shape"],
  ];
  steps.forEach(([k, h, d], i) => {
    const y = 1.25 + i * 0.98;
    badge(s, 0.5, y, k, C.teal);
    s.addText(h, { x: 1.15, y: y - 0.02, w: 3.6, h: 0.35, fontFace: BF, fontSize: 15, bold: true, color: C.ink, margin: 0, isTextBox: true });
    s.addText(d, { x: 1.15, y: y + 0.32, w: 3.6, h: 0.6, fontFace: BF, fontSize: 12, color: C.gray, valign: "top", margin: 0, isTextBox: true });
  });
  s.addChart(pres.charts.BAR, [{ name: "Students", labels: ["1", "2", "3", "4", "5"], values: freq }], {
    x: 5.0, y: 1.2, w: 4.5, h: 3.95, barDir: "col", barGapWidthPct: 40, chartColors: [C.sea],
    showTitle: true, title: "prior_ai_use_1to5 (n = 20)", titleFontFace: BF, titleFontSize: 13, titleColor: C.ink,
    showValue: true, dataLabelPosition: "outEnd", dataLabelColor: C.ink, dataLabelFontSize: 11,
    catAxisTitle: "Rating (1 = never … 5 = very often)", showCatAxisTitle: true, catAxisTitleFontSize: 10, catAxisTitleColor: C.gray,
    valAxisTitle: "Number of students", showValAxisTitle: true, valAxisTitleFontSize: 10, valAxisTitleColor: C.gray,
    catAxisLabelColor: C.gray, valAxisLabelColor: C.gray, valAxisMaxVal: 8, valAxisMajorUnit: 2,
    valGridLine: { color: "E3E8EA", size: 0.5 }, catGridLine: { style: "none" }, showLegend: false,
  });
}

// 9. Section divider: Review
darkSlide("PART 3", "Review with Dataset 1", "Measurement scales  ·  Center & shape  ·  SD  ·  z-scores")
  .addNotes("NOTE: numbers on the following slides use a 20-student example of prior_ai_use_1to5 (see the 'Dataset 1 (example)' sheet). Replace with the real Dataset 1 values before class if they differ.");

// 10. Measurement scales
{
  const s = pres.addSlide();
  title(s, "Measurement Scales", "Find one variable in Dataset 1 for each scale");
  const scales = [
    ["N", "Nominal", "Categories, no order", "major", "Business, Nursing, CS…", C.sea],
    ["O", "Ordinal", "Ordered, gaps not equal", "class_standing", "Freshman < Sophomore < Junior < Senior", C.teal],
    ["I", "Interval", "Equal gaps, no true zero", "birth_year", "2004 − 2002 = 2 yrs, but 0 AD ≠ 'no time'", C.gold],
    ["R", "Ratio", "Equal gaps + true zero", "study_hours_week", "0 = none; 10 hrs is twice 5 hrs", C.coral],
  ];
  scales.forEach(([k, name, def, v, ex, col], i) => {
    const x = 0.5 + i * 2.3;
    card(s, x, 1.45, 2.1, 3.25, C.light);
    badge(s, x + 0.8, 1.6, k, col);
    s.addText(name, { x: x + 0.1, y: 2.2, w: 1.9, h: 0.4, fontFace: HF, fontSize: 18, bold: true, color: C.ink, align: "center", margin: 0, isTextBox: true });
    s.addText(def, { x: x + 0.1, y: 2.6, w: 1.9, h: 0.5, fontFace: BF, fontSize: 12, italic: true, color: C.gray, align: "center", valign: "top", margin: 0, isTextBox: true });
    s.addText(v, { x: x + 0.1, y: 3.15, w: 1.9, h: 0.4, fontFace: "Courier New", fontSize: 12, bold: true, color: C.teal, align: "center", margin: 0, isTextBox: true });
    s.addText(ex, { x: x + 0.12, y: 3.6, w: 1.86, h: 1.0, fontFace: BF, fontSize: 11, color: C.ink, align: "center", valign: "top", margin: 0, isTextBox: true });
  });
  s.addText([
    { text: "Where does prior_ai_use_1to5 fit? ", options: { bold: true, color: C.coral } },
    { text: "Ordinal (a 1–5 rating). We often treat it as interval to compute a mean and SD.", options: { color: C.ink } },
  ], { x: 0.5, y: 4.85, w: 9, h: 0.4, fontFace: BF, fontSize: 13, margin: 0, isTextBox: true });
  s.addNotes("Replace the example variable names with the actual Dataset 1 column names. Quick test: Is there an order? Are gaps equal? Does zero mean 'none of it'? Can you say 'twice as much'?");
}

// 11. Mean, median, mode
{
  const s = pres.addSlide();
  title(s, "Practice: Mean, Median & Mode", "Variable: prior_ai_use_1to5  (1 = never … 5 = very often), n = 20");
  const rows = [[th("Rating (x)"), th("Frequency (f)"), th("f · x")]];
  freq.forEach((f, i) => rows.push([td(i + 1), td(f), td(f * (i + 1))]));
  rows.push([td("Total", { bold: true }), td("20", { bold: true }), td("68", { bold: true })]);
  s.addTable(rows, { x: 0.5, y: 1.45, w: 3.8, colW: [1.2, 1.4, 1.2], fontFace: BF, fontSize: 13, color: C.ink, border: { type: "solid", pt: 0.75, color: "C9D6D9" }, fill: { color: C.white }, rowH: 0.42 });
  const stats = [
    ["3.40", "Mean", "68 ÷ 20 = 3.40", "=AVERAGE(B2:B21)"],
    ["4", "Median", "10th & 11th sorted values are both 4", "=MEDIAN(B2:B21)"],
    ["4", "Mode", "4 appears most often (7 students)", "=MODE.SNGL(B2:B21)"],
  ];
  stats.forEach(([big, lab, how, fx], i) => {
    const y = 1.45 + i * 1.22;
    card(s, 4.7, y, 4.8, 1.05, i === 0 ? C.mint : C.light);
    s.addText(big, { x: 4.8, y, w: 1.3, h: 1.05, fontFace: HF, fontSize: 36, bold: true, color: i === 0 ? C.coral : C.teal, align: "center", valign: "middle", margin: 0, isTextBox: true });
    s.addText(lab, { x: 6.2, y: y + 0.1, w: 3.2, h: 0.32, fontFace: BF, fontSize: 15, bold: true, color: C.ink, margin: 0, isTextBox: true });
    s.addText(how, { x: 6.2, y: y + 0.42, w: 3.2, h: 0.3, fontFace: BF, fontSize: 12, color: C.gray, margin: 0, isTextBox: true });
    s.addText(fx, { x: 6.2, y: y + 0.7, w: 3.2, h: 0.28, fontFace: "Courier New", fontSize: 11, color: C.teal, margin: 0, isTextBox: true });
  });
  s.addNotes("Have students compute by hand first, then check in Excel. For the median with n = 20 (even), average the 10th and 11th sorted values.");
}

// 12. Shape of distribution
{
  const s = pres.addSlide();
  title(s, "What Type of Distribution Is It?");
  s.addChart(pres.charts.BAR, [{ name: "Students", labels: ["1", "2", "3", "4", "5"], values: freq }], {
    x: 0.5, y: 1.15, w: 4.6, h: 4.0, barDir: "col", barGapWidthPct: 40, chartColors: [C.sea],
    showValue: true, dataLabelPosition: "outEnd", dataLabelColor: C.ink, dataLabelFontSize: 11,
    catAxisTitle: "prior_ai_use_1to5", showCatAxisTitle: true, catAxisTitleFontSize: 10, catAxisTitleColor: C.gray,
    catAxisLabelColor: C.gray, valAxisLabelColor: C.gray, valAxisMaxVal: 8, valAxisMajorUnit: 2,
    valGridLine: { color: "E3E8EA", size: 0.5 }, catGridLine: { style: "none" }, showLegend: false,
  });
  card(s, 5.4, 1.15, 4.1, 1.35, C.mint);
  s.addText([
    { text: "Mean 3.4  <  Median 4  =  Mode 4", options: { bold: true, color: C.ink, breakLine: true } },
    { text: "The mean is pulled toward the low ratings (the tail).", options: { color: C.gray } },
  ], { x: 5.6, y: 1.25, w: 3.8, h: 1.15, fontFace: BF, fontSize: 14, valign: "middle", margin: 0, isTextBox: true });
  s.addText("Negatively (left) skewed", { x: 5.4, y: 2.65, w: 4.1, h: 0.5, fontFace: HF, fontSize: 19, bold: true, color: C.coral, margin: 0, isTextBox: true });
  s.addText("Most students report frequent AI use; a few never use it.", { x: 5.4, y: 3.15, w: 4.1, h: 0.45, fontFace: BF, fontSize: 13, color: C.ink, margin: 0, isTextBox: true });
  const rules = [["Mean < Median", "Left (negative) skew"], ["Mean ≈ Median", "Symmetric / normal"], ["Mean > Median", "Right (positive) skew"]];
  const rrows = rules.map(([a, b], i) => [td(a, { bold: i === 0, fill: { color: i === 0 ? "FBE3DC" : C.white } }), td(b, { bold: i === 0, fill: { color: i === 0 ? "FBE3DC" : C.white } })]);
  s.addTable(rrows, { x: 5.4, y: 3.75, w: 4.1, colW: [1.7, 2.4], fontFace: BF, fontSize: 12, color: C.ink, border: { type: "solid", pt: 0.75, color: "C9D6D9" }, rowH: 0.42 });
}

// 13. SD by hand
{
  const s = pres.addSlide();
  title(s, "Practice: Calculate the SD", "Use the frequency table so you don't repeat the same value 20 times");
  const rows = [[th("x"), th("f"), th("x − x̄"), th("(x − x̄)²"), th("f · (x − x̄)²")]];
  freq.forEach((f, i) => {
    const x = i + 1, d = x - mean;
    rows.push([td(x), td(f), td(f2(d)), td(f2(d * d)), td(f2(f * d * d))]);
  });
  rows.push([td("Σ", { bold: true }), td("20", { bold: true }), td(""), td(""), td("30.80", { bold: true, color: C.coral })]);
  s.addTable(rows, { x: 0.5, y: 1.45, w: 5.2, colW: [0.7, 0.7, 1.1, 1.2, 1.5], fontFace: BF, fontSize: 13, color: C.ink, border: { type: "solid", pt: 0.75, color: "C9D6D9" }, fill: { color: C.white }, rowH: 0.42 });
  card(s, 6.0, 1.45, 3.5, 2.5, C.mint);
  s.addText([
    { text: "Variance", options: { bold: true, color: C.teal, breakLine: true } },
    { text: "s² = 30.80 / (20 − 1)", options: { breakLine: true } },
    { text: "    = " + f3(variance), options: { bold: true, breakLine: true } },
    { text: " ", options: { breakLine: true, fontSize: 8 } },
    { text: "Standard deviation", options: { bold: true, color: C.teal, breakLine: true } },
    { text: "s = √" + f3(variance), options: { breakLine: true } },
    { text: "   = " + f3(sd), options: { bold: true, color: C.coral } },
  ], { x: 6.25, y: 1.6, w: 3.1, h: 2.7, fontFace: BF, fontSize: 16, color: C.ink, valign: "top", margin: 0, isTextBox: true });
  code(s, "=STDEV.S(B2:B21) → " + f3(sd), 6.0, 4.1, 3.5, 0.45, 12);
}

// 14. z-scores
{
  const s = pres.addSlide();
  title(s, "Practice: Calculate the z-Score", "z = (x − x̄) / s   with  x̄ = 3.40,  s = " + f3(sd));
  const rows = [[th("Rating x"), th("x − x̄"), th("z"), th("Meaning")]];
  const meaning = ["~1.9 SD below the mean", "1.1 SD below", "slightly below", "about ½ SD above", "~1.3 SD above"];
  for (let x = 1; x <= 5; x++) {
    const z = (x - mean) / sd;
    rows.push([td(x), td(f2(x - mean)), td(f2(z), { bold: true, color: z < 0 ? C.coral : C.teal }), td(meaning[x - 1], { align: "left" })]);
  }
  s.addTable(rows, { x: 0.5, y: 1.45, w: 5.6, colW: [1.0, 1.1, 1.0, 2.5], fontFace: BF, fontSize: 13, color: C.ink, border: { type: "solid", pt: 0.75, color: "C9D6D9" }, fill: { color: C.white }, rowH: 0.45 });
  card(s, 6.4, 1.45, 3.1, 2.7, C.light);
  s.addText([
    { text: "Worked example", options: { bold: true, color: C.teal, breakLine: true } },
    { text: "A student who never uses AI (x = 1):", options: { breakLine: true } },
    { text: "z = (1 − 3.4) / 1.273", options: { breakLine: true } },
    { text: "   = −1.89", options: { bold: true, color: C.coral } },
  ], { x: 6.6, y: 1.6, w: 2.8, h: 2.4, fontFace: BF, fontSize: 14, color: C.ink, valign: "top", paraSpaceAfter: 4, margin: 0, isTextBox: true });
  s.addText([
    { text: "Read it: ", options: { bold: true, color: C.coral } },
    { text: "negative z = below the mean, positive z = above; |z| > 2 is unusual.", options: { color: C.ink } },
  ], { x: 0.5, y: 4.4, w: 9, h: 0.5, fontFace: BF, fontSize: 13, margin: 0, isTextBox: true });
}

// 15. Wrap-up
{
  const s = pres.addSlide();
  s.background = { color: C.teal };
  s.addText("BEFORE THE MIDTERM", { x: 0.7, y: 0.7, w: 8.6, h: 0.4, fontFace: BF, fontSize: 16, bold: true, color: C.gold, charSpacing: 3, margin: 0, isTextBox: true });
  s.addText("Checklist", { x: 0.7, y: 1.15, w: 8.6, h: 0.8, fontFace: HF, fontSize: 40, bold: true, color: C.white, margin: 0, isTextBox: true });
  s.addText([
    { text: "Name the scale: nominal, ordinal, interval, or ratio (and say why)", options: { bullet: true, breakLine: true } },
    { text: "Compute mean, median, and mode, then use them to describe the shape", options: { bullet: true, breakLine: true } },
    { text: "Variance → SD using n − 1, showing each step", options: { bullet: true, breakLine: true } },
    { text: "z-score: compute it AND explain it in words", options: { bullet: true, breakLine: true } },
    { text: "Formulas are provided; practice choosing and using them", options: { bullet: true } },
  ], { x: 0.7, y: 2.3, w: 8.6, h: 2.6, fontFace: BF, fontSize: 17, color: C.white, paraSpaceAfter: 10, valign: "top", margin: 0, isTextBox: true });
}

pres.writeFile({ fileName: "ReviewWeek.pptx" }).then((f) => console.log("wrote " + f));
