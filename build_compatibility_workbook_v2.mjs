import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = "/Users/liam/Desktop/compatibility test/outputs/compatibility_test";
await fs.mkdir(outputDir, { recursive: true });

const questions = ["God", "Politics", "Marriage", "Children", "Ethicity", "Past", "D/S", "True/kind", "Animal"];
const people = [
  ["Boy", "joshuah", "N-", "L-=I", "S^=I", "Y-=N", "A^=D", "Y-NM", "S-=D", "T-=N", "A-"],
  ["Boy", "alex", "N^=N", "R-=N", "L^=I", "Y^=I", "A^=M", "Y-", "S-DD", "T^=D", "A-"],
  ["Boy", "ethan", "Y^=M", "R^=N", "S-LN", "Y-=D", "A^=I", "Y-", "S-=D", "T^KI", "N-"],
  ["Boy", "liam", "Y-", "L-=N", "L-=I", "Y-", "A^=M", "N-=I", "S-DD", "T^=N", "A-"],
  ["Boy", "arthur", "N-", "R-LN", "S-=N", "Y^=I", "A^", "N-=D", "S-DN", "T^KM", "A-=D"],
  ["Boy", "harry", "N-=N", "R-", "S^=M", "Y^=I", "A^=N", "N^=D", "S^DN", "T-=I", "N-"],
  ["Boy", "lawrence", "N-=D", "R-", "S^=I", "Y-=N", "A^=M", "Y-", "S-=N", "T-KD", "N-"],
  ["Boy", "richard", "Y-", "R-", "S-=N", "Y-=I", "A^=D", "N-=N", "S-=N", "T^=M", "A^=D"],
  ["Boy", "tony", "Y^=N", "L^=D", "L-=I", "Y^=I", "A^=M", "Y-", "D^SD", "T^", "A^=N"],
  ["Boy", "will", "N^=D", "L-", "S-=I", "N-=D", "W^AI", "N-=N", "D-=M", "K-TN", "A^"],
  ["Boy", "jt", "N^=I", "R^=N", "L-", "Y-=D", "A^=I", "Y^ND", "S-", "T^=M", "A-=N"],
  ["Boy", "leo", "N^=D", "L^=M", "L-=N", "Y-=I", "A^=I", "N^=D", "S-DN", "T-=I", "A^"],
  ["Boy", "james", "N-=D", "L^=M", "S-=D", "Y-", "A^", "N-", "D-=I", "T-=I", "A-=N"],
  ["Boy", "ayaan", "N^=N", "L^=M", "L^=I", "Y-", "I^", "Y^NI", "D^=N", "T-=D", "A^=D"],
  ["Girl", "jasmine", "Y^=M", "R-=I", "S^=I", "N-", "A^", "N^=D", "S-DD", "K^TN", "A^=N"],
  ["Girl", "cathleen", "N^=I", "L^=D", "S-LD", "N-=M", "A^=N", "Y^", "D-SN", "T^", "A^=I"],
  ["Girl", "edith", "N-", "R-", "L-=D", "N-=D", "A^=N", "N-=N", "S-=I", "T-=M", "A^=I"],
  ["Girl", "sophiah", "Y^=M", "R^", "L-SD", "Y-=D", "A^=1", "Y-", "D^SD", "T-KN", "A^=I"],
  ["Girl", "angela", "Y^=I", "R-=D", "L-SI", "Y^=M", "A^", "N^=I", "S-DD", "K^TN", "A^=N"],
  ["Girl", "issabel", "Y-", "L^=I", "S-LN", "Y^=I", "A^=M", "Y-NN", "D^", "T^=D", "A^=D"],
];

const workbook = Workbook.create();
const calc = workbook.worksheets.add("Calculator");
const data = workbook.worksheets.add("Data");
const parsed = workbook.worksheets.add("Parsed");
const pairs = workbook.worksheets.add("All Matches");
const rules = workbook.worksheets.add("Rules");
for (const sheet of [calc, data, parsed, pairs, rules]) sheet.showGridLines = false;

data.getRange("A1:K1").values = [["Group", "Person", ...questions]];
data.getRangeByIndexes(1, 0, people.length, 11).values = people;
data.tables.add(`A1:K${people.length + 1}`, true, "PeopleTable");
data.getRange("A1:K1").format = { fill: "#1F2937", font: { bold: true, color: "#FFFFFF" } };
data.getRange("A:A").format.columnWidthPx = 70;
data.getRange("B:B").format.columnWidthPx = 120;
data.getRange("C:K").format.columnWidthPx = 86;
data.freezePanes.freezeRows(1);

const parsedRows = [];
for (const person of people) {
  for (const q of questions) parsedRows.push([`${person[1]}|${q}`, person[1], q]);
}
const parsedLast = parsedRows.length + 1;
parsed.getRange("A1:M1").values = [["Key", "Person", "Question", "Code", "Norm", "Pos", "Rest", "Actual", "Strength", "Want raw", "Desired", "Care code", "Care"]];
parsed.getRangeByIndexes(1, 0, parsedRows.length, 3).values = parsedRows;
parsed.getRange("D2").formulas = [[`=INDEX(Data!$C$2:$K$${people.length + 1},MATCH(B2,Data!$B$2:$B$${people.length + 1},0),MATCH(C2,Data!$C$1:$K$1,0))`]];
parsed.getRange(`D2:D${parsedLast}`).fillDown();
parsed.getRange("E2").formulas = [[`=UPPER(SUBSTITUTE(TRIM(D2),"1","I"))`]];
parsed.getRange(`E2:E${parsedLast}`).fillDown();
parsed.getRange("F2").formulas = [[`=MIN(IFERROR(FIND("^",E2),999),IFERROR(FIND("-",E2),999))`]];
parsed.getRange(`F2:F${parsedLast}`).fillDown();
parsed.getRange("G2").formulas = [[`=IF(OR(E2="",F2=999,LEN(E2)<=F2),"",MID(E2,F2+1,99))`]];
parsed.getRange(`G2:G${parsedLast}`).fillDown();
parsed.getRange("H2").formulas = [[`=IF(OR(E2="",F2=999),"",LEFT(E2,F2-1))`]];
parsed.getRange(`H2:H${parsedLast}`).fillDown();
parsed.getRange("I2").formulas = [[`=IF(MID(E2,F2,1)="-",0.75,IF(MID(E2,F2,1)="^",1,0))`]];
parsed.getRange(`I2:I${parsedLast}`).fillDown();
parsed.getRange("J2").formulas = [[`=IF(G2="","",IF(LEN(G2)=1,"=",LEFT(G2,LEN(G2)-1)))`]];
parsed.getRange(`J2:J${parsedLast}`).fillDown();
parsed.getRange("K2").formulas = [[`=IF(J2="","",IF(LEFT(J2,1)="=",H2,J2))`]];
parsed.getRange(`K2:K${parsedLast}`).fillDown();
parsed.getRange("L2").formulas = [[`=IF(G2="","",RIGHT(G2,1))`]];
parsed.getRange(`L2:L${parsedLast}`).fillDown();
parsed.getRange("M2").formulas = [[`=IF(L2="D",5,IF(L2="N",10,IF(L2="I",15,IF(L2="M",40,0))))`]];
parsed.getRange(`M2:M${parsedLast}`).fillDown();
parsed.getRange("A1:M1").format = { fill: "#374151", font: { bold: true, color: "#FFFFFF" } };
parsed.getRange("A:M").format.columnWidthPx = 94;
parsed.getRange("A:A").format.columnWidthPx = 150;
parsed.getRange("I:I").format.numberFormat = "0.00";
parsed.getRange("M:M").format.numberFormat = "0.00";
parsed.freezePanes.freezeRows(1);

function lookup(nameCell, qCellOrText, col) {
  const qPart = qCellOrText.startsWith("$") || qCellOrText.match(/^[A-Z]+\d+$/) ? qCellOrText : `"${qCellOrText}"`;
  return `INDEX(Parsed!$${col}$2:$${col}$${parsedLast},(MATCH(${nameCell},Data!$B$2:$B$${people.length + 1},0)-1)*${questions.length}+MATCH(${qPart},Parsed!$C$2:$C$10,0))`;
}
function scoreTerm(wanterCell, otherCell, question) {
  const desired = lookup(wanterCell, question, "K");
  const actual = lookup(otherCell, question, "H");
  const care = lookup(wanterCell, question, "M");
  const strength = lookup(otherCell, question, "I");
  return `IF(${desired}=${actual},${care}*${strength},0)`;
}
function scoreFormula(wanterCell, otherCell) {
  return `=IFERROR(${questions.map((q) => scoreTerm(wanterCell, otherCell, q)).join("+")},0)`;
}

calc.getRange("A1:J1").merge();
calc.getRange("A1").values = [["Compatibility Calculator"]];
calc.getRange("A1").format = { fill: "#111827", font: { bold: true, color: "#FFFFFF", size: 18 }, horizontalAlignment: "center" };
calc.getRange("A3:B8").values = [
  ["Input", "Name"],
  ["Person A", "tony"],
  ["Person B", "jasmine"],
  ["B-to-A score", ""],
  ["A-to-B score", ""],
  ["Final score", ""],
];
calc.getRange("A3:B3").format = { fill: "#0F766E", font: { bold: true, color: "#FFFFFF" } };
calc.getRange("B6").formulas = [["=SUM(F13:F21)"]];
calc.getRange("B7").formulas = [["=SUM(I13:I21)"]];
calc.getRange("B8").formulas = [["=SQRT(B6*B7)"]];
calc.getRange("B6:B8").format.numberFormat = `0.00"%"`;
calc.getRange("B8").format = { fill: "#ECFDF5", font: { bold: true, color: "#065F46", size: 14 }, numberFormat: `0.00"%"` };
calc.getRange("A10:J10").merge();
calc.getRange("A10").values = [[`B-to-A means Person B's wants are checked against Person A's answers first; A-to-B does the reverse. Final uses SQRT(B-to-A * A-to-B).`]];
calc.getRange("A10").format = { fill: "#F3F4F6", font: { italic: true, color: "#374151" }, wrapText: true };
calc.getRange("A12:J12").values = [["Question", "A code", "B code", "B wants", "A actual", "B-to-A pts", "A wants", "B actual", "A-to-B pts", "Note"]];
calc.getRange("A12:J12").format = { fill: "#1F2937", font: { bold: true, color: "#FFFFFF" } };
calc.getRangeByIndexes(12, 0, questions.length, 1).values = questions.map((q) => [q]);
for (let i = 0; i < questions.length; i++) {
  const row = 13 + i;
  calc.getRange(`B${row}`).formulas = [[`=IFERROR(${lookup("$B$4", `A${row}`, "D")},"")`]];
  calc.getRange(`C${row}`).formulas = [[`=IFERROR(${lookup("$B$5", `A${row}`, "D")},"")`]];
  calc.getRange(`D${row}`).formulas = [[`=IFERROR(${lookup("$B$5", `A${row}`, "K")},"")`]];
  calc.getRange(`E${row}`).formulas = [[`=IFERROR(${lookup("$B$4", `A${row}`, "H")},"")`]];
  calc.getRange(`F${row}`).formulas = [[`=IFERROR(IF(D${row}=E${row},${lookup("$B$5", `A${row}`, "M")}*${lookup("$B$4", `A${row}`, "I")},0),0)`]];
  calc.getRange(`G${row}`).formulas = [[`=IFERROR(${lookup("$B$4", `A${row}`, "K")},"")`]];
  calc.getRange(`H${row}`).formulas = [[`=IFERROR(${lookup("$B$5", `A${row}`, "H")},"")`]];
  calc.getRange(`I${row}`).formulas = [[`=IFERROR(IF(G${row}=H${row},${lookup("$B$4", `A${row}`, "M")}*${lookup("$B$5", `A${row}`, "I")},0),0)`]];
  calc.getRange(`J${row}`).formulas = [[`=IF(AND(F${row}=0,I${row}=0),"",IF(F${row}>I${row},"B cares match",IF(I${row}>F${row},"A cares match","both / tie")))`]];
}
calc.getRange("F13:F21").format.numberFormat = `0.00"%"`;
calc.getRange("I13:I21").format.numberFormat = `0.00"%"`;
calc.getRange("A:J").format.columnWidthPx = 105;
calc.getRange("J:J").format.columnWidthPx = 120;
calc.getRange("A10:J10").format.rowHeightPx = 42;
calc.getRange("B4:B5").dataValidation = { rule: { type: "list", formula1: `Data!$B$2:$B$${people.length + 1}` } };
calc.freezePanes.freezeRows(12);
calc.freezePanes.freezeColumns(1);

pairs.getRange("A1:E1").values = [["Boy", "Girl", "Girl-to-boy", "Boy-to-girl", "Final"]];
pairs.getRange("A1:E1").format = { fill: "#1F2937", font: { bold: true, color: "#FFFFFF" } };
const boys = people.filter((p) => p[0] === "Boy").map((p) => p[1]);
const girls = people.filter((p) => p[0] === "Girl").map((p) => p[1]);
const pairRows = [];
for (const boy of boys) for (const girl of girls) pairRows.push([boy, girl]);
pairs.getRangeByIndexes(1, 0, pairRows.length, 2).values = pairRows;
for (let i = 0; i < pairRows.length; i++) {
  const row = 2 + i;
  pairs.getRange(`C${row}`).formulas = [[scoreFormula(`B${row}`, `A${row}`)]];
  pairs.getRange(`D${row}`).formulas = [[scoreFormula(`A${row}`, `B${row}`)]];
  pairs.getRange(`E${row}`).formulas = [[`=SQRT(C${row}*D${row})`]];
}
pairs.getRange(`C2:E${pairRows.length + 1}`).format.numberFormat = `0.00"%"`;
pairs.getRange("A:E").format.columnWidthPx = 115;
pairs.tables.add(`A1:E${pairRows.length + 1}`, true, "AllMatchesTable");
pairs.freezePanes.freezeRows(1);

rules.getRange("A1:D1").values = [["Thing", "Code", "Meaning", "Score"]];
rules.getRange("A2:D6").values = [
  ["Care", "D", "don't really care", 5],
  ["Care", "N", "nice to have", 10],
  ["Care", "I", "important", 15],
  ["Care", "M", "mandatory / very important", 40],
  ["Strength", "-", "weak answer from the other person", 0.75],
];
rules.getRange("A8:D17").values = [
  ["Question", "Codes", "Meaning", ""],
  ["God / Children / Past", "Y / N", "yes / no", ""],
  ["Politics", "R / L", "right wing / left wing", ""],
  ["Marriage", "L / S", "love and pleasure / security and commitment", ""],
  ["Ethicity", "A / W / I / WA", "asian / white / indian / white and asian", ""],
  ["D/S", "D / S", "dominant / submissive", ""],
  ["True/kind", "T / K", "truthful / kind", ""],
  ["Animal", "A / N", "likes animals / does not like animals", ""],
  ["Want", "'=", "same answer as their own answer", ""],
  ["One-letter rest", "missing equals", "treated as wants same, care D", ""],
];
rules.getRange("A1:D1").format = { fill: "#0F766E", font: { bold: true, color: "#FFFFFF" } };
rules.getRange("A8:D8").format = { fill: "#0F766E", font: { bold: true, color: "#FFFFFF" } };
rules.getRange("A:D").format.columnWidthPx = 190;

const calcInspect = await workbook.inspect({
  kind: "table",
  range: "Calculator!A1:J21",
  include: "values,formulas",
  tableMaxRows: 25,
  tableMaxCols: 12,
});
console.log(calcInspect.ndjson);

const pairInspect = await workbook.inspect({
  kind: "table",
  range: "All Matches!A1:E8",
  include: "values,formulas",
  tableMaxRows: 10,
  tableMaxCols: 5,
});
console.log(pairInspect.ndjson);

const formulaErrors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 200 },
  summary: "formula error scan",
});
console.log(formulaErrors.ndjson);

for (const sheetName of ["Calculator", "Data", "All Matches", "Rules"]) {
  const preview = await workbook.render({ sheetName, autoCrop: "all", scale: 1, format: "png" });
  await fs.writeFile(`${outputDir}/${sheetName.replaceAll(" ", "_").toLowerCase()}_preview.png`, new Uint8Array(await preview.arrayBuffer()));
}

const xlsx = await SpreadsheetFile.exportXlsx(workbook);
await xlsx.save(`${outputDir}/compatibility_calculator.xlsx`);
console.log(`${outputDir}/compatibility_calculator.xlsx`);
