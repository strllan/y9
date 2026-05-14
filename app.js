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
  ["boy", "joshuah", "N-", "L-=I", "S^=I", "Y-=N", "A^=D", "Y-NM", "S-=D", "T-=N", "A-"],
  ["boy", "alex", "N^=N", "R-=N", "L^=I", "Y^=I", "A^=M", "Y-", "S-DD", "T^=D", "A-"],
  ["boy", "ethan", "Y^=M", "R^=N", "S-LN", "Y-=D", "A^=I", "Y-", "S-=D", "T^KI", "N-"],
  ["boy", "liam", "Y-", "L-=N", "L-=I", "Y-", "A^=M", "N-=I", "S-DD", "T^=N", "A-"],
  ["boy", "arthur", "N-", "R-LN", "S-=N", "Y^=I", "A^", "N-=D", "S-DN", "T^KM", "A-=D"],
  ["boy", "harry", "N-=N", "R-", "S^=M", "Y^=I", "A^=N", "N^=D", "S^DN", "T-=I", "N-"],
  ["boy", "lawrence", "N-=D", "R-", "S^=I", "Y-=N", "A^=M", "Y-", "S-=N", "T-KD", "N-"],
  ["boy", "richard", "Y-", "R-", "S-=N", "Y-=I", "A^=D", "N-=N", "S-=N", "T^=M", "A^=D"],
  ["boy", "tony", "Y^=N", "L^=D", "L-=I", "Y^=I", "A^=M", "Y-", "D^SD", "T^", "A^=N"],
  ["boy", "will", "N^=D", "L-", "S-=I", "N-=D", "W^AI", "N-=N", "D-=M", "K-TN", "A^"],
  ["boy", "jt", "N^=I", "R^=N", "L-", "Y-=D", "A^=I", "Y^ND", "S-", "T^=M", "A-=N"],
  ["boy", "leo", "N^=D", "L^=M", "L-=N", "Y-=I", "A^=I", "N^=D", "S-DN", "T-=I", "A^"],
  ["boy", "james", "N-=D", "L^=M", "L-=D", "Y-", "A^", "N-", "D-SI", "T-=I", "A-=N"],
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
const selfPerson = document.querySelector("#selfPerson");
const rankingType = document.querySelector("#rankingType");
const finalScore = document.querySelector("#finalScore");
const bToA = document.querySelector("#bToA");
const aToB = document.querySelector("#aToB");
const breakdownRows = document.querySelector("#breakdownRows");
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

function scoreLabel(value) {
  return value.toFixed(2);
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
  bToA.textContent = scoreLabel(result.bDirection.total);
  aToB.textContent = scoreLabel(result.aDirection.total);
  finalScore.textContent = scoreLabel(result.final);

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
        <td class="points">${scoreLabel(bRow.points)}</td>
        <td>${preferenceLabel(aRow.want.desired)}</td>
        <td>${aRow.answer.actual}</td>
        <td class="points">${scoreLabel(aRow.points)}</td>
      `;
      return tr;
    }),
  );
}

function makePairRankings(firstNames, secondNames, sameGroup = false) {
  const pairs = [];
  for (let i = 0; i < firstNames.length; i += 1) {
    const start = sameGroup ? i + 1 : 0;
    for (let j = start; j < secondNames.length; j += 1) {
      pairs.push(scorePair(firstNames[i], secondNames[j]));
    }
  }
  return pairs.sort((left, right) => right.final - left.final);
}

function renderList(targetId, entries, renderEntry) {
  const target = document.querySelector(`#${targetId}`);
  target.replaceChildren(
    ...entries.map((entry, index) => {
      const item = document.createElement("li");
      item.innerHTML = renderEntry(entry);
      item.style.setProperty("--i", index);
      return item;
    }),
  );
}

function getPairScoreForPerson(personName, otherName) {
  const person = byName.get(personName);
  return person.group === "boy" ? scorePair(personName, otherName) : scorePair(otherName, personName);
}

function getOneWayScore(personName, otherName) {
  return scoreDirectional(byName.get(personName), byName.get(otherName)).total;
}

function average(values) {
  return values.reduce((sum, value) => sum + value, 0) / values.length;
}

function renderAverages() {
  const averageRankings = people
    .map(([, name]) => {
      const others = people.map(([, otherName]) => otherName).filter((otherName) => otherName !== name);
      const matchAverage = average(others.map((otherName) => getPairScoreForPerson(name, otherName).final));
      const oneWayAverage = average(others.map((otherName) => getOneWayScore(name, otherName)));
      return { name, matchAverage, oneWayAverage };
    })
    .sort((left, right) => right.matchAverage - left.matchAverage);

  const oneWayAverageRankings = [...averageRankings].sort((left, right) => right.oneWayAverage - left.oneWayAverage);

  renderList("averageTopMatches", averageRankings, (entry) =>
    `<strong>${entry.name}</strong><span class="match-score">${scoreLabel(entry.matchAverage)} average</span>`,
  );
  renderList("averageLeastMatches", [...averageRankings].reverse(), (entry) =>
    `<strong>${entry.name}</strong><span class="match-score">${scoreLabel(entry.matchAverage)} average</span>`,
  );
  renderList("averageTopSingles", oneWayAverageRankings, (entry) =>
    `<strong>${entry.name}</strong><span class="match-score">${scoreLabel(entry.oneWayAverage)} average</span>`,
  );
  renderList("averageLeastSingles", [...oneWayAverageRankings].reverse(), (entry) =>
    `<strong>${entry.name}</strong><span class="match-score">${scoreLabel(entry.oneWayAverage)} average</span>`,
  );
}

function renderSelfCheck() {
  const selectedName = selfPerson.value;
  const selectedGroup = byName.get(selectedName).group;

  const rankMatches = (targetGroup) =>
    people
      .filter(([group, name]) => group === targetGroup && name !== selectedName)
      .map(([, name]) => ({ name, score: getPairScoreForPerson(selectedName, name).final }))
      .sort((left, right) => right.score - left.score);

  const rankOneWay = (targetGroup) =>
    people
      .filter(([group, name]) => group === targetGroup && name !== selectedName)
      .map(([, name]) => ({ name, score: getOneWayScore(selectedName, name) }))
      .sort((left, right) => right.score - left.score);

  const renderSelfEntry = (entry) => `<strong>${entry.name}</strong><span class="match-score">${scoreLabel(entry.score)}</span>`;

  renderList("selfGirlMatches", rankMatches("girl"), renderSelfEntry);
  renderList("selfBoyMatches", rankMatches("boy"), renderSelfEntry);
  renderList("selfGirlSingles", rankOneWay("girl"), renderSelfEntry);
  renderList("selfBoySingles", rankOneWay("boy"), renderSelfEntry);

  document.querySelector("#self-check-title").textContent = `${selectedName} Self-Check`;
}

function renderRankingBoard(prefix, ranked) {
  const oneWay = ranked
    .flatMap((match) => [
      {
        from: match.a.name,
        to: match.b.name,
        score: match.aDirection.total,
      },
      {
        from: match.b.name,
        to: match.a.name,
        score: match.bDirection.total,
      },
    ])
    .sort((left, right) => right.score - left.score);

  renderList(`${prefix}TopMatches`, ranked, (match) =>
    `<strong>${match.a.name}</strong> and <strong>${match.b.name}</strong><span class="match-score">${scoreLabel(match.final)}</span>`,
  );
  renderList(`${prefix}LeastMatches`, [...ranked].reverse(), (match) =>
    `<strong>${match.a.name}</strong> and <strong>${match.b.name}</strong><span class="match-score">${scoreLabel(match.final)}</span>`,
  );
  renderList(`${prefix}TopSingles`, oneWay, (single) =>
    `<strong>${single.from}</strong> → <strong>${single.to}</strong><span class="match-score">${scoreLabel(single.score)}</span>`,
  );
  renderList(`${prefix}LeastSingles`, [...oneWay].reverse(), (single) =>
    `<strong>${single.from}</strong> → <strong>${single.to}</strong><span class="match-score">${scoreLabel(single.score)}</span>`,
  );
}

function renderSelectedRankings() {
  const rankingSets = {
    boyGirl: {
      title: "Boy-Girl Matches",
      ranked: makePairRankings(boys, girls),
    },
    boyBoy: {
      title: "Boy-Boy Matches",
      ranked: makePairRankings(boys, boys, true),
    },
    girlGirl: {
      title: "Girl-Girl Matches",
      ranked: makePairRankings(girls, girls, true),
    },
  };
  const selected = rankingSets[rankingType.value];
  document.querySelector("#rankings-title").textContent = selected.title;
  renderRankingBoard("ranking", selected.ranked);
}

function showPage(pageId) {
  document.querySelectorAll(".page").forEach((page) => {
    page.classList.toggle("active", page.id === pageId);
  });
  document.querySelectorAll("[data-page-link]").forEach((control) => {
    control.classList.toggle("active", control.dataset.pageLink === pageId);
  });
  window.scrollTo({ top: 0, behavior: "smooth" });
}

for (const [, name] of people) {
  personA.append(option(name));
  personB.append(option(name));
  selfPerson.append(option(name));
}

personA.value = "tony";
personB.value = "jasmine";
selfPerson.value = "tony";
personA.addEventListener("change", renderCalculator);
personB.addEventListener("change", renderCalculator);
selfPerson.addEventListener("change", renderSelfCheck);
rankingType.addEventListener("change", renderSelectedRankings);
swapButton.addEventListener("click", () => {
  [personA.value, personB.value] = [personB.value, personA.value];
  renderCalculator();
});

document.querySelectorAll("[data-page-link]").forEach((control) => {
  control.addEventListener("click", (event) => {
    event.preventDefault();
    showPage(control.dataset.pageLink);
  });
});

const boys = people.filter(([group]) => group === "boy").map(([, name]) => name);
const girls = people.filter(([group]) => group === "girl").map(([, name]) => name);

renderCalculator();
renderAverages();
renderSelfCheck();
renderSelectedRankings();
