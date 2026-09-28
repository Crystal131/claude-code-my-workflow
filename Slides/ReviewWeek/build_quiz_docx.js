// build_quiz_docx.js — Final Review Week quiz (19 MC + 1 open-ended) as Word.
// Run: NODE_PATH=<dir with docx>/node_modules node build_quiz_docx.js
const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType,
  AlignmentType, ShadingType, BorderStyle, PageBreak, TabStopType } = require("docx");

// [section, stem, [A, B, C, D], correctIndex]
const Q = [
  ["Measurement Scales", "Temperature measured in degrees Fahrenheit is an example of which scale of measurement?", ["Nominal", "Ordinal", "Interval", "Ratio"], 2],
  [null, "The numbers on soccer players' jerseys (#7, #10, #23) are measured on which scale?", ["Nominal", "Ordinal", "Interval", "Ratio"], 0],
  [null, "For which variable does it make sense to say \"Student A has twice as much as Student B\"?", ["Student's race/ethnicity", "Temperature in °C", "Class standing (freshman … senior)", "Hours studied per week"], 3],
  ["Mean, Median, and Mode", "For the scores 2, 4, 4, 5, 10, which set of values is correct?", ["Mean = 5, Median = 4, Mode = 4", "Mean = 4, Median = 5, Mode = 4", "Mean = 5, Median = 5, Mode = 4", "Mean = 4, Median = 4, Mode = 10"], 0],
  [null, "What is the median of 3, 7, 8, 12, 15, 20?", ["8", "10", "12", "10.83"], 1],
  [null, "Which measure of central tendency is MOST affected by extreme scores (outliers)?", ["Median", "Mode", "Median and mode equally", "Mean"], 3],
  ["Positive and Negative Skew", "Consider the scores 2, 7, 8, 9, 9. The mean is 7, the median is 8, and the mode is 9. What is the shape of this distribution?", ["Positively skewed", "Negatively skewed", "Symmetrical", "Bimodal"], 1],
  [null, "A distribution has a mean of 60 and a median of 50. It is most likely:", ["Negatively skewed", "Symmetrical", "Positively skewed", "Impossible to determine without the mode"], 2],
  ["Standard Deviation", "The standard deviation describes:", ["The difference between the highest and lowest scores", "The middle score", "The most frequent score", "The typical distance of scores from the mean"], 3],
  [null, "What is the standard deviation of the scores 5, 5, 5, 5?", ["0", "1", "5", "20"], 0],
  [null, "Treating the scores 2, 4, 6 as a sample (divide by n − 1), what is the standard deviation?", ["1.63", "2", "4", "8"], 1],
  [null, "An instructor adds 5 bonus points to every student's score. What happens to the standard deviation?", ["Increases by 5", "Decreases by 5", "Multiplied by 5", "Stays the same"], 3],
  ["Normal Distribution Characteristics", "The normal distribution is:", ["Skewed right", "Bell-shaped and symmetrical", "Flat", "Skewed left"], 1],
  [null, "Approximately what percentage of scores falls within ±1 SD of the mean?", ["68%", "50%", "95%", "99.7%"], 0],
  [null, "Test scores are normally distributed with a mean of 50 and SD of 10. About 68% of students score between:", ["30 and 70", "45 and 55", "40 and 60", "50 and 60"], 2],
  [null, "IQ is normally distributed with a mean of 100 and SD of 15. About 95% of people score between:", ["55 and 145", "85 and 115", "100 and 130", "70 and 130"], 3],
  ["Standard Normal Distribution", "The standard normal distribution has:", ["Mean 0, SD 1", "Mean 1, SD 0", "Mean 100, SD 15", "Mean 0, SD 0"], 0],
  ["z-Scores", "An exam has a mean of 70 and SD of 8. A student scores 82. What is the z-score?", ["12", "0.67", "1.50", "1.20"], 2],
  [null, "A distribution has a mean of 50 and SD of 5. Which raw score corresponds to z = −2.0?", ["40", "45", "52", "60"], 0],
];
const L = ["A", "B", "C", "D"];
const FONT = "Calibri", TEAL = "1F4E5A";

const section = (t) => new Paragraph({ spacing: { before: 280, after: 120 }, keepNext: true,
  border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "9FBFC4", space: 2 } },
  children: [new TextRun({ text: t, bold: true, size: 24, color: TEAL })] });

function questionBlock(num, stem, opts) {
  const out = [new Paragraph({ spacing: { before: 160, after: 60 }, keepNext: true, keepLines: true,
    children: [new TextRun({ text: `${num}. `, bold: true }), new TextRun(stem)] })];
  opts.forEach((o, i) => out.push(new Paragraph({ indent: { left: 540 }, spacing: { after: 20 }, keepNext: i < 3,
    children: [new TextRun({ text: `${L[i]})  `, bold: true }), new TextRun(o)] })));
  return out;
}

// ----- Student quiz -----
const quiz = [
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 },
    children: [new TextRun({ text: "Review Week Quiz", bold: true, size: 36, color: TEAL })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 240 },
    children: [new TextRun({ text: "Descriptive Statistics", size: 26, color: "5B6770" })] }),
  new Paragraph({ spacing: { after: 120 }, tabStops: [{ type: TabStopType.RIGHT, position: 9360 }],
    children: [new TextRun("Name: ______________________________"), new TextRun("\tDate: ________________")] }),
  new Paragraph({ spacing: { after: 120 }, children: [new TextRun({ italics: true, color: "5B6770",
    text: "Directions: Questions 1–19: circle the one best answer. Question 20: write your response in the space provided." })] }),
];
Q.forEach(([sec, stem, opts], i) => {
  if (sec) quiz.push(section(sec));
  quiz.push(...questionBlock(i + 1, stem, opts));
});
quiz.push(section("Open-Ended"));
quiz.push(new Paragraph({ spacing: { before: 160, after: 120 }, children: [
  new TextRun({ text: "20. ", bold: true }),
  new TextRun({ text: "Course feedback: ", bold: true }),
  new TextRun("What feedback do you have for this class so far? What is helping you learn, and what could be improved? "),
  new TextRun({ text: "(Your answer will not affect your grade.)", italics: true })] }));
// Writing lines: a borderless table whose rows have only a bottom rule (adjacent bordered paragraphs merge into one box).
const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const rule = { style: BorderStyle.SINGLE, size: 4, color: "B0B8BC" };
quiz.push(new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [9360],
  rows: Array.from({ length: 6 }, () => new TableRow({ height: { value: 440, rule: "atLeast" }, children: [new TableCell({
    width: { size: 9360, type: WidthType.DXA }, borders: { top: none, left: none, right: none, bottom: rule },
    children: [new Paragraph({ children: [] })] })] })) }));

// ----- Answer key -----
const key = [
  new Paragraph({ children: [new PageBreak()] }),
  new Paragraph({ spacing: { after: 60 }, children: [new TextRun({ text: "Answer Key (Instructor Only)", bold: true, size: 32, color: TEAL })] }),
  new Paragraph({ spacing: { after: 200 }, children: [new TextRun({ text: "Remove this page before distributing the quiz.", italics: true, color: "B03A2E" })] }),
];
const border = { style: BorderStyle.SINGLE, size: 4, color: "C9D6D9" };
const borders = { top: border, bottom: border, left: border, right: border };
const colW = [700, 1000, 4460, 3200]; // sums to 9360
const cell = (t, w, opts = {}) => new TableCell({ borders, width: { size: w, type: WidthType.DXA },
  shading: opts.head ? { fill: TEAL, type: ShadingType.CLEAR, color: "auto" } : undefined,
  margins: { top: 50, bottom: 50, left: 100, right: 100 },
  children: [new Paragraph({ alignment: opts.left ? AlignmentType.LEFT : AlignmentType.CENTER,
    children: [new TextRun({ text: t, bold: !!opts.head || !!opts.bold, color: opts.head ? "FFFFFF" : undefined })] })] });
const topic = []; let cur = "";
Q.forEach(([sec]) => { if (sec) cur = sec; topic.push(cur); });
const rows = [new TableRow({ tableHeader: true, children: [cell("Q", colW[0], { head: true }), cell("Ans.", colW[1], { head: true }), cell("Correct response", colW[2], { head: true }), cell("Topic", colW[3], { head: true })] })];
Q.forEach(([, , opts, k], i) => rows.push(new TableRow({ children: [
  cell(String(i + 1), colW[0]), cell(L[k], colW[1], { bold: true }), cell(opts[k], colW[2], { left: true }), cell(topic[i], colW[3], { left: true })] })));
rows.push(new TableRow({ children: [cell("20", colW[0]), cell("—", colW[1]), cell("Open-ended (not scored)", colW[2], { left: true }), cell("Course feedback", colW[3], { left: true })] }));
key.push(new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: colW, rows }));

key.push(new Paragraph({ spacing: { before: 280, after: 100 }, children: [new TextRun({ text: "Worked Solutions", bold: true, size: 26, color: TEAL })] }));
[
  ["Q4", "Mean = (2 + 4 + 4 + 5 + 10) / 5 = 25 / 5 = 5; median = middle (3rd) score = 4; mode = 4 (appears twice)."],
  ["Q5", "n = 6 (even), so average the 3rd and 4th scores: (8 + 12) / 2 = 10. (Option D, 10.83, is the mean.)"],
  ["Q7", "Mean = 35 / 5 = 7 < median = 8 < mode = 9, so the tail is on the low side: negatively skewed."],
  ["Q8", "Mean (60) > median (50): the mean is pulled toward high scores, so positively skewed."],
  ["Q10", "Every score equals the mean (5), so every deviation is 0 and SD = 0."],
  ["Q11", "Mean = 4; deviations −2, 0, 2; SS = 4 + 0 + 4 = 8; s² = 8 / (3 − 1) = 4; s = √4 = 2. (Option A, 1.63, divides by n.)"],
  ["Q12", "Adding a constant shifts every score and the mean equally, so the distances from the mean do not change."],
  ["Q15", "50 ± 1(10) = 40 to 60."],
  ["Q16", "100 ± 2(15) = 70 to 130."],
  ["Q18", "z = (82 − 70) / 8 = 12 / 8 = 1.50."],
  ["Q19", "X = mean + z(SD) = 50 + (−2)(5) = 40."],
].forEach(([q, t]) => key.push(new Paragraph({ spacing: { after: 80 }, children: [new TextRun({ text: `${q}: `, bold: true }), new TextRun(t)] })));

const doc = new Document({
  styles: { default: { document: { run: { font: FONT, size: 22 } } } },
  sections: [{ properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } },
    children: [...quiz, ...key] }],
});
Packer.toBuffer(doc).then((b) => { fs.writeFileSync("ReviewWeek_Quiz_Final.docx", b); console.log("wrote ReviewWeek_Quiz_Final.docx"); });
