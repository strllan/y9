import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = "/Users/liam/Desktop/compatibility test/outputs/compatibility_test";
await fs.mkdir(outputDir, { recursive: true });

const questions = [
  "God",
  "Politics",
  "Marriage",
  "Children",
  "Ethicity",
  "Past",
  "D/S",
  "True/kind",
  "Animal",
];

const people = [
  ["Boy", "Alex", "N^=N", "R-=N", "L^=I", "Y-=I", "A^=M", "Y-", "S-DD", "T^=D", "A-"],
  ["Boy", "Ethan", "Y^=M", "R^=N", "S-LN", "Y-=D", "A^=I", "Y-", "S-=D", "T^KI", "N-"],
  ["Boy", "Liam", "N^=D", "L-=N", "L-=I", "Y-", "A^=M", "N-=I", "S-DD", "T^=N", "A-"],
  ["Boy", "Arthur", "N-", "R-LN", "S-=N", "Y^=I", "A^", "N-=D", "S-DN", "T^KM", "A-=D"],
  ["Boy", "Harry", "N-=N", "R-", "S^=M", "Y^=I", "A^=N", "N^=D", "S^DN", "T-=I", "N-"],
  ["Boy", "Lawrence", "N-=D", "R-", "S^=I", "Y-=N", "A^=M", "Y-", "S-=N", "T-KD", "N-"],
  ["Boy", "Richard", "Y-", "R-", "S-=N", "Y-=I", "A^=d", "S-=N", "S-=N", "T^=M", "A^D"],
  ["Boy", "Tony", "Y^=N", "L^=D", "L-=I", "Y^=I", "A^=M", "Y-", "D^SD", "T^", "A^=N"],
  ["Boy", "Will", "N^=D", "L-", "S-=I", "N-=D", "W^AI", "N-=N", "D-=M", "K-TN", "A^"],
  ["Boy", "JT", "N^=I", "R^=N", "L-", "Y-=D", "A^=I", "Y^ND", "S-", "T^=M", "A-=N"],
  ["Boy", "Leo", "N^=D", "L^=M", "L-=N", "Y-=I", "A^=I", "N^=D", "S-DN", "T-=I", "A^"],
  ["Boy", "James", "N-=D", "L^=M", "S-=D", "Y-", "A^", "N-", "D-=I", "T-=I", "A-=N"],
  ["Boy", "Ayaan", "N^=N", "L^=M", "L^=I", "Y-", "I^", "Y^NI", "D^=N", "T-=D", "A^=D"],
  ["Girl", "Jasmine", "Y^=M", "R-=I", "S^=I", "N-", "A^", "N^=D", "S-Dd", "K^TN", "A^=N"],
  ["Girl", "Cathleen", "N^=I", "L^=D", "S-LD", "N-=M", "A^=N", "Y^", "D-SN", "T^", "A^=I"],
  ["Girl", "Edith", "N-", "R-", "L-=D", "N-=D", "A^=N", "N-=N", "S-=I", "T-=M", "A^=I"],
  ["Girl", "Sophia (hu)", "Y^=M", "R^", "L-SD", "Y-=D", "A^=1", "Y-", "D^SD", "T-KN", "A^=I"],
  ["Girl", "Angela", "Y^=I", "R-=d", "L-SI", "Y^=M", "A^", "N^=I", "S-DD", "K^TN", "A^=N"],
  ["Girl", "Issabel", "Y-", "L^=I", "S-LN", "Y^=I", "A^=M", "Y-NN", "D^", "T^=D", "A^=D"],
];

const workbook = Workbook.create();
const calc = workbook.worksheets.add("Calculator");
const data = workbook.worksheets.add("Data");
const parsed = workbook.worksheets.add("Parsed");
const pairs = workbook.worksheets.add("All Matches");
const rules = workbook.worksheets.add("Rules");

for (const sheet of [calc, data, parsed, pairs, rules]) {
  sheet.showGridLines = false;
}

data.getRange("A1:K1").values = [["Group", "Person", ...questions]];
data.getRangeByIndexes(1, 0, people.length, 11).values = people;
data.tables.add(`A1:K${people.length + 1}`, true, "PeopleTable");
data.getRange("A1:K1").format = {
  fill: "#1F2937",
  font: { bold: true, color: "#FFFFFF" },
};
data.getRange(`A2:K${people.length + 1}`).format = { font: { color: "#111827" } };
data.getRange("A:A").format.columnWidthPx = 72;
data.getRange("B:B").format.columnWidthPx = 120;
data.getRange("C:K").format.columnWidthPx = 86;
data.freezePanes.freezeRows(1);

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
  ["Want", "=", "same answer as their own answer", ""],
  ["Blank want/care", "", "0 points; the person does not care", ""],
];
rules.getRange("A1:D1").format = { fill: "#0F766E", font: { bold: true, color: "#FFFFFF" } };
rules.getRange("A8:D8").format = { fill: "#0F766E", font: { bold: true, color: "#FFFFFF" } };
rules.getRange("A:D").format.columnWidthPx = 190;
rules.getRange("D:D").format.columnWidthPx = 80;

const parsedRows = [];
for (const person of people) {
  for (const q of questions) {
    parsedRows.push([`${person[1]}|${q}`, person[1], q]);
  }
}
parsed.getRange("A1:H1").values = [["Key", "Person", "Question", "Code", "Actual", "Strength", "Desired", "Care"]];
parsed.getRangeByIndexes(1, 0, parsedRows.length, 3).values = parsedRows;
const parsedLast = parsedRows.length + 1;
parsed.getRange("D2").formulas = [[`=INDEX(Data!$C$2:$K$${people.length + 1},MATCH(B2,Data!$B$2:$B$${people.length + 1},0),MATCH(C2,Data!$C$1:$K$1,0))`]];
parsed.getRange(`D2:D${parsedLast}`).fillDown();
parsed.getRange("E2").formulas = [[`=IF(D2="","",LEFT(UPPER(SUBSTITUTE(TRIM(D2),"1","I")),MIN(IFERROR(FIND("^",UPPER(SUBSTITUTE(TRIM(D2),"1","I"))),999),IFERROR(FIND("-",UPPER(SUBSTITUTE(TRIM(D2),"1","I"))),999))-1))`]];
parsed.getRange(`E2:E${parsedLast}`).fillDown();
parsed.getRange("F2").formulas = [[`=IF(MID(UPPER(SUBSTITUTE(TRIM(D2),"1","I")),MIN(IFERROR(FIND("^",UPPER(SUBSTITUTE(TRIM(D2),"1","I"))),999),IFERROR(FIND("-",UPPER(SUBSTITUTE(TRIM(D2),"1","I"))),999)),1)="-",0.75,IF(MID(UPPER(SUBSTITUTE(TRIM(D2),"1","I")),MIN(IFERROR(FIND("^",UPPER(SUBSTITUTE(TRIM(D2),"1","I"))),999),IFERROR(FIND("-",UPPER(SUBSTITUTE(TRIM(D2),"1","I"))),999)),1)="^",1,0))`]];
parsed.getRange(`F2:F${parsedLast}`).fillDown();
parsed.getRange("G2").formulas = [[`=IF(LEN(MID(UPPER(SUBSTITUTE(TRIM(D2),"1","I")),MIN(IFERROR(FIND("^",UPPER(SUBSTITUTE(TRIM(D2),"1","I"))),999),IFERROR(FIND("-",UPPER(SUBSTITUTE(TRIM(D2),"1","I"))),999))+1,99))=0,"",IF(LEFT(IF(LEN(MID(UPPER(SUBSTITUTE(TRIM(D2),"1","I")),MIN(IFERROR(FIND("^",UPPER(SUBSTITUTE(TRIM(D2),"1","I"))),999),IFERROR(FIND("-",UPPER(SUBSTITUTE(TRIM(D2),"1","I"))),999))+1,99))=1,"=",LEFT(MID(UPPER(SUBSTITUTE(TRIM(D2),"1","I")),MIN(IFERROR(FIND("^",UPPER(SUBSTITUTE(TRIM(D2),"1","I"))),999),IFERROR(FIND("-",UPPER(SUBSTITUTE(TRIM(D2),"1","I"))),999))+1,99),LEN(MID(UPPER(SUBSTITUTE(TRIM(D2),"1","I")),MIN(IFERROR(FIND("^",UPPER(SUBSTITUTE(TRIM(D2),"1","I"))),999),IFERROR(FIND("-",UPPER(SUBSTITUTE(TRIM(D2),"1","I"))),999))+1,99))-1))),1)="=",E2,IF(LEN(MID(UPPER(SUBSTITUTE(TRIM(D2),"1","I")),MIN(IFERROR(FIND("^",UPPER(SUBSTITUTE(TRIM(D2),"1","I"))),999),IFERROR(FIND("-",UPPER(SUBSTITUTE(TRIM(D2),"1","I"))),999))+1,99))=1,"=",LEFT(MID(UPPER(SUBSTITUTE(TRIM(D2),"1","I")),MIN(IFERROR(FIND("^",UPPER(SUBSTITUTE(TRIM(D2),"1","I"))),999),IFERROR(FIND("-",UPPER(SUBSTITUTE(TRIM(D2),"1","I"))),999))+1,99),LEN(MID(UPPER(SUBSTITUTE(TRIM(D2),"1","I")),MIN(IFERROR(FIND("^",UPPER(SUBSTITUTE(TRIM(D2),"1","I"))),999),IFERROR(FIND("-",UPPER(SUBSTITUTE(TRIM(D2),"1","I"))),999))+1,99))-1))))`]];
parsed.getRange(`G2:G${parsedLast}`).fillDown();
parsed.getRange("H2").formulas = [[`=IF(LEN(MID(UPPER(SUBSTITUTE(TRIM(D2),"1","I")),MIN(IFERROR(FIND("^",UPPER(SUBSTITUTE(TRIM(D2),"1","I"))),999),IFERROR(FIND("-",UPPER(SUBSTITUTE(TRIM(D2),"1","I"))),999))+1,99))=0,0,SWITCH(RIGHT(MID(UPPER(SUBSTITUTE(TRIM(D2),"1","I")),MIN(IFERROR(FIND("^",UPPER(SUBSTITUTE(TRIM(D2),"1","I"))),999),IFERROR(FIND("-",UPPER(SUBSTITUTE(TRIM(D2),"1","I"))),999))+1,99),1),"D",5,"N",10,"I",15,"M",40,0))`]];
parsed.getRange(`H2:H${parsedLast}`).fillDown();
parsed.getRange("A1:H1").format = { fill: "#374151", font: { bold: true, color: "#FFFFFF" } };
parsed.getRange("A:H").format.columnWidthPx = 110;
parsed.getRange("F:H").format.numberFormat = "0.00";
parsed.freezePanes.freezeRows(1);

calc.getRange("A1:J1").merge();
calc.getRange("A1").values = [["Compatibility Calculator"]];
calc.getRange("A1").format = {
  fill: "#111827",
  font: { bold: true, color: "#FFFFFF", size: 18 },
  horizontalAlignment: "center",
};
calc.getRange("A3:B8").values = [
  ["Input", "Name"],
  ["Person A", "Tony"],
  ["Person B", "Jasmine"],
  ["B-to-A score", ""],
  ["A-to-B score", ""],
  ["Final score", ""],
];
calc.getRange("A3:B3").format = { fill: "#0F766E", font: { bold: true, color: "#FFFFFF" } };
calc.getRange("B6").formulas = [["=SUM(G13:G21)"]];
calc.getRange("B7").formulas = [["=SUM(J13:J21)"]];
calc.getRange("B8").formulas = [["=SQRT(B6*B7)"]];
calc.getRange("B6:B8").format.numberFormat = `0.00"%"`;
calc.getRange("A6:A8").format = { font: { bold: true } };
calc.getRange("B8").format = {
  fill: "#ECFDF5",
  font: { bold: true, color: "#065F46", size: 14 },
  numberFormat: `0.00"%"`,
};
calc.getRange("A10:J10").values = [[
  "Direction",
  "Meaning",
  "",
  "",
  "",
  "",
  "",
  "",
  "",
  "",
]];
calc.getRange("A10:J10").merge(true);
calc.getRange("A10").values = [[`B-to-A means Person B's wants are checked against Person A's answers first; A-to-B does the reverse. Final uses SQRT(B-to-A * A-to-B).`]];
calc.getRange("A10").format = { fill: "#F3F4F6", font: { italic: true, color: "#374151" }, wrapText: true };
calc.getRange("A12:J12").values = [[
  "Question",
  "A code",
  "B code",
  "B wants",
  "A actual",
  "B-to-A pts",
  "A wants",
  "B actual",
  "A-to-B pts",
  "Note",
]];
calc.getRange("A12:J12").format = { fill: "#1F2937", font: { bold: true, color: "#FFFFFF" } };
calc.getRangeByIndexes(12, 0, questions.length, 1).values = questions.map((q) => [q]);
for (let i = 0; i < questions.length; i++) {
  const row = 13 + i;
  calc.getRange(`B${row}`).formulas = [[`=XLOOKUP($B$4&"|"&A${row},Parsed!$A$2:$A$${parsedLast},Parsed!$D$2:$D$${parsedLast},"")`]];
  calc.getRange(`C${row}`).formulas = [[`=XLOOKUP($B$5&"|"&A${row},Parsed!$A$2:$A$${parsedLast},Parsed!$D$2:$D$${parsedLast},"")`]];
  calc.getRange(`D${row}`).formulas = [[`=XLOOKUP($B$5&"|"&A${row},Parsed!$A$2:$A$${parsedLast},Parsed!$G$2:$G$${parsedLast},"")`]];
  calc.getRange(`E${row}`).formulas = [[`=XLOOKUP($B$4&"|"&A${row},Parsed!$A$2:$A$${parsedLast},Parsed!$E$2:$E$${parsedLast},"")`]];
  calc.getRange(`F${row}`).formulas = [[`=IFERROR(IF(OR(D${row}="",E${row}="",XLOOKUP($B$5&"|"&A${row},Parsed!$A$2:$A$${parsedLast},Parsed!$H$2:$H$${parsedLast},0)=0),0,IF(D${row}=E${row},XLOOKUP($B$5&"|"&A${row},Parsed!$A$2:$A$${parsedLast},Parsed!$H$2:$H$${parsedLast},0)*XLOOKUP($B$4&"|"&A${row},Parsed!$A$2:$A$${parsedLast},Parsed!$F$2:$F$${parsedLast},0),0)),0)`]];
  calc.getRange(`G${row}`).formulas = [[`=XLOOKUP($B$4&"|"&A${row},Parsed!$A$2:$A$${parsedLast},Parsed!$G$2:$G$${parsedLast},"")`]];
  calc.getRange(`H${row}`).formulas = [[`=XLOOKUP($B$5&"|"&A${row},Parsed!$A$2:$A$${parsedLast},Parsed!$E$2:$E$${parsedLast},"")`]];
  calc.getRange(`I${row}`).formulas = [[`=IFERROR(IF(OR(G${row}="",H${row}="",XLOOKUP($B$4&"|"&A${row},Parsed!$A$2:$A$${parsedLast},Parsed!$H$2:$H$${parsedLast},0)=0),0,IF(G${row}=H${row},XLOOKUP($B$4&"|"&A${row},Parsed!$A$2:$A$${parsedLast},Parsed!$H$2:$H$${parsedLast},0)*XLOOKUP($B$5&"|"&A${row},Parsed!$A$2:$A$${parsedLast},Parsed!$F$2:$F$${parsedLast},0),0)),0)`]];
  calc.getRange(`J${row}`).formulas = [[`=IF(AND(F${row}=0,I${row}=0),"",IF(F${row}>I${row},"B cares match",IF(I${row}>F${row},"A cares match","both / tie")))`]];
}
calc.getRange("F13:F21").format.numberFormat = `0.00"%"`;
calc.getRange("I13:I21").format.numberFormat = `0.00"%"`;
calc.getRange("A:J").format.columnWidthPx = 105;
calc.getRange("J:J").format.columnWidthPx = 120;
calc.getRange("A10:J10").format.rowHeightPx = 42;
calc.freezePanes.freezeRows(12);
calc.freezePanes.freezeColumns(1);

const namesRange = `Data!$B$2:$B$${people.length + 1}`;
calc.getRange("B4:B5").dataValidation = { rule: { type: "list", formula1: namesRange } };

function scoreFormula(wanterCell, otherCell) {
  return `=SUMPRODUCT(--(XLOOKUP(${wanterCell}&"|"&$G$2:$G$10,Parsed!$A$2:$A$${parsedLast},Parsed!$G$2:$G$${parsedLast},"")=XLOOKUP(${otherCell}&"|"&$G$2:$G$10,Parsed!$A$2:$A$${parsedLast},Parsed!$E$2:$E$${parsedLast},"")),XLOOKUP(${wanterCell}&"|"&$G$2:$G$10,Parsed!$A$2:$A$${parsedLast},Parsed!$H$2:$H$${parsedLast},0),XLOOKUP(${otherCell}&"|"&$G$2:$G$10,Parsed!$A$2:$A$${parsedLast},Parsed!$F$2:$F$${parsedLast},0))`;
}

pairs.getRange("A1:E1").values = [["Boy", "Girl", "Girl-to-boy", "Boy-to-girl", "Final"]];
pairs.getRange("A1:E1").format = { fill: "#1F2937", font: { bold: true, color: "#FFFFFF" } };
pairs.getRange("G1").values = [["Questions"]];
pairs.getRange("G2:G10").values = questions.map((q) => [q]);
const boys = people.filter((p) => p[0] === "Boy").map((p) => p[1]);
const girls = people.filter((p) => p[0] === "Girl").map((p) => p[1]);
const pairRows = [];
for (const boy of boys) {
  for (const girl of girls) {
    pairRows.push([boy, girl]);
  }
}
pairs.getRangeByIndexes(1, 0, pairRows.length, 2).values = pairRows;
for (let i = 0; i < pairRows.length; i++) {
  const row = 2 + i;
  pairs.getRange(`C${row}`).formulas = [[scoreFormula(`B${row}`, `A${row}`)]];
  pairs.getRange(`D${row}`).formulas = [[scoreFormula(`A${row}`, `B${row}`)]];
  pairs.getRange(`E${row}`).formulas = [[`=SQRT(C${row}*D${row})`]];
}
pairs.getRange(`C2:E${pairRows.length + 1}`).format.numberFormat = `0.00"%"`;
pairs.getRange("A:E").format.columnWidthPx = 115;
pairs.getRange("G:G").format.columnWidthPx = 100;
pairs.tables.add(`A1:E${pairRows.length + 1}`, true, "AllMatchesTable");
pairs.freezePanes.freezeRows(1);

const inspected = await workbook.inspect({
  kind: "table",
  range: "Calculator!A1:J21",
  include: "values,formulas",
  tableMaxRows: 25,
  tableMaxCols: 12,
});
console.log(inspected.ndjson);

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
