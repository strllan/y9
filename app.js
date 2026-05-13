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

const questionLabels = {
  God: "Religion",
  Politics: "Politics",
  Marriage: "Relationship priorities",
  Children: "Children",
  Ethicity: "Ethnicity",
  Past: "Past relationships",
  "D/S": "Dominant or submissive",
  "True/kind": "Honest or kind",
  Animal: "Animals",
};

const people = [
  ["boy", "alex", "N^=N", "R-=N", "L^=I", "Y^=I", "A^=M", "Y-", "S-DD", "T^=D", "A-"],
  ["boy", "ethan", "Y^=M", "R^=N", "S-LN", "Y-=D", "A^=I", "Y-", "S-=D", "T^KI", "N-"],
  ["boy", "liam", "N^=D", "L-=N", "L-=I", "Y-", "WA^=M", "N-=I", "S-DD", "T^=N", "A-"],
  ["boy", "arthur", "N-", "R-LN", "S-=N", "Y^=I", "A^", "N-=D", "S-DN", "T^KM", "A-=D"],
  ["boy", "harry", "N-=N", "R-", "S^=M", "Y^=I", "A^=N", "N^=D", "S^DN", "T-=I", "N-"],
  ["boy", "lawrence", "N-=D", "R-", "S^=I", "Y-=N", "A^=M", "Y-", "S-=N", "T-KD", "N-"],
  ["boy", "richard", "Y-", "R-", "S-=N", "Y-=I", "A^=D", "N-=N", "S-=N", "T^=M", "A^=D"],
  ["boy", "tony", "Y^=N", "L^=D", "L-=I", "Y^=I", "A^=M", "Y-", "D^SD", "T^", "A^=N"],
  ["boy", "will", "N^=D", "L-", "S-=I", "N-=D", "W^AI", "N-=N", "D-=M", "K-TN", "A^"],
  ["boy", "jt", "N^=I", "R^=N", "L-", "Y-=D", "A^=I", "Y^ND", "S-", "T^=M", "A-=N"],
  ["boy", "leo", "N^=D", "L^=M", "L-=N", "Y-=I", "A^=I", "N^=D", "S-DN", "T-=I", "A^"],
  ["boy", "james", "N-=D", "L^=M", "S-=D", "Y-", "A^", "N-", "D-=I", "T-=I", "A-=N"],
  ["boy", "ayaan", "N^=N", "L^=M", "L^=I", "Y-", "I^", "Y^NI", "D^=N", "T-=D", "A^=D"],
  ["girl", "jasmine", "Y^=M", "R-=I", "S^=I", "N-", "A^", "N^=D", "S-DD", "K^TN", "A^=N"],
  ["girl", "cathleen", "N^=I", "L^=D", "S-LD", "N-=M", "A^=N", "Y^", "D-SN", "T^", "A^=I"],
  ["girl", "edith", "N-", "R-", "L-=D", "N-=D", "A^=N", "N-=N", "S-=I", "T-=M", "A^=I"],
  ["girl", "sophiah", "Y^=M", "R^", "L-SD", "Y-=D", "A^=I", "Y-", "D^SD", "T-KN", "A^=I"],
  ["girl", "angela", "Y^=I", "R-=D", "L-SI", "Y^=M", "A^", "N^=I", "S-DD", "K^TN", "A^=N"],
  ["girl", "issabel", "Y-", "L^=I", "S-LN", "Y^=I", "A^=M", "Y-NN", "D^", "T^=D", "A^=D"],
];

const careScores = { D: 5, N: 10, I: 15, M: 40 };

const byName = new Map(
  people.map(([group, name, ...answers]) => [
    name,
    {
      group,
      name,
      answers: Object.fromEntries(questions.map((question, index) => [question, parseCode(answers[index])])),
    },
  ]),
);

const personA = document.querySelector("#personA");
const personB = document.querySelector("#personB");
const finalScore = document.querySelector("#finalScore");
const bToA = document.querySelector("#bToA");
const aToB = document.querySelector("#aToB");
const breakdownRows = document.querySelector("#breakdownRows");
const topMatches = document.querySelector("#topMatches");
const leastMatches = document.querySelector("#leastMatches");
const swapButton = document.querySelector("#swapButton");

function parseCode(rawCode) {
  const code = rawCode.trim().toUpperCase().replaceAll("1", "I");
  const strengthIndex = Math.min(
    code.includes("^") ? code.indexOf("^") : Number.POSITIVE_INFINITY,
    code.includes("-") ? code.indexOf("-") : Number.POSITIVE_INFINITY,
  );

  if (!Number.isFinite(strengthIndex)) {
    return { raw: rawCode, actual: code, strength: 0, desired: "", care: 0 };
  }

  const actual = code.slice(0, strengthIndex);
  const strength = code[strengthIndex] === "-" ? 0.75 : 1;
  const rest = code.slice(strengthIndex + 1);

  if (!rest) {
    return { raw: rawCode, actual, strength, desired: "", care: 0 };
  }

  const careCode = rest.slice(-1);
  const care = careScores[careCode] ?? 0;
  const wantRaw = rest.length === 1 ? "=" : rest.slice(0, -1);
  const desired = wantRaw.startsWith("=") ? actual : wantRaw;

  return { raw: rawCode, actual, strength, desired, care };
}

function scoreDirectional(wanter, other) {
  const rows = questions.map((question) => {
    const want = wanter.answers[question];
    const answer = other.answers[question];
    const points = want.desired && want.desired === answer.actual ? want.care * answer.strength : 0;
    return { question, want, answer, points };
  });

  return {
    rows,
    total: rows.reduce((sum, row) => sum + row.points, 0),
  };
}

function scorePair(nameA, nameB) {
  const a = byName.get(nameA);
  const b = byName.get(nameB);
  const bDirection = scoreDirectional(b, a);
  const aDirection = scoreDirectional(a, b);
  return {
    a,
    b,
    bDirection,
    aDirection,
    final: Math.sqrt(bDirection.total * aDirection.total),
  };
}

function percent(value) {
  return `${value.toFixed(2)}%`;
}

function preferenceLabel(value) {
  return value || "No preference";
}

function option(name) {
  const item = document.createElement("option");
  item.value = name;
  item.textContent = name;
  return item;
}

function renderCalculator() {
  const result = scorePair(personA.value, personB.value);
  bToA.textContent = percent(result.bDirection.total);
  aToB.textContent = percent(result.aDirection.total);
  finalScore.textContent = percent(result.final);

  breakdownRows.replaceChildren(
    ...questions.map((question, index) => {
      const bRow = result.bDirection.rows[index];
      const aRow = result.aDirection.rows[index];
      const tr = document.createElement("tr");
      if (bRow.points > 0 || aRow.points > 0) tr.className = "match";
      tr.innerHTML = `
        <td>${questionLabels[question]}</td>
        <td>${preferenceLabel(bRow.want.desired)}</td>
        <td>${bRow.answer.actual}</td>
        <td class="points">${percent(bRow.points)}</td>
        <td>${preferenceLabel(aRow.want.desired)}</td>
        <td>${aRow.answer.actual}</td>
        <td class="points">${percent(aRow.points)}</td>
      `;
      return tr;
    }),
  );
}

function renderTopMatches() {
  const boys = people.filter(([group]) => group === "boy").map(([, name]) => name);
  const girls = people.filter(([group]) => group === "girl").map(([, name]) => name);
  const ranked = boys
    .flatMap((boy) => girls.map((girl) => scorePair(boy, girl)))
    .sort((left, right) => right.final - left.final);

  topMatches.replaceChildren(
    ...ranked.slice(0, 15).map((match) => {
      const item = document.createElement("li");
      item.innerHTML = `<strong>${match.a.name}</strong> and <strong>${match.b.name}</strong><span class="match-score">${percent(match.final)}</span>`;
      return item;
    }),
  );

  leastMatches.replaceChildren(
    ...ranked.slice(-15).reverse().map((match) => {
      const item = document.createElement("li");
      item.innerHTML = `<strong>${match.a.name}</strong> and <strong>${match.b.name}</strong><span class="match-score">${percent(match.final)}</span>`;
      return item;
    }),
  );
}

for (const [, name] of people) {
  personA.append(option(name));
  personB.append(option(name));
}

personA.value = "tony";
personB.value = "jasmine";
personA.addEventListener("change", renderCalculator);
personB.addEventListener("change", renderCalculator);
swapButton.addEventListener("click", () => {
  [personA.value, personB.value] = [personB.value, personA.value];
  renderCalculator();
});

renderCalculator();
renderTopMatches();
