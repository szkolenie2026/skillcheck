/*
 * SkillCheck — bank zadań i silnik oceny (czysty JavaScript, bez zależności od Excela).
 * Ten sam plik posłuży później wersji Google Sheets (Apps Script) — platforma dostarcza
 * tylko odczyt: formuły, wartości oryginalne, wartości po perturbacji.
 *
 * Zasada: KOD liczy werdykt i typ błędu. AI dostaje gotowy wynik i decyduje tylko o ścieżce.
 */
(function (root) {
  "use strict";

  const TOL = 0.01;

  // ---------- losowość deterministyczna (ten sam kod sesji = te same dane) ----------
  function hashSeed(str) {
    let h = 1779033703 ^ str.length;
    for (let i = 0; i < str.length; i++) {
      h = Math.imul(h ^ str.charCodeAt(i), 3432918353);
      h = (h << 13) | (h >>> 19);
    }
    h = Math.imul(h ^ (h >>> 16), 2246822507);
    h = Math.imul(h ^ (h >>> 13), 3266489909);
    return (h ^= h >>> 16) >>> 0;
  }
  function mulberry32(a) {
    return function () {
      a |= 0; a = (a + 0x6d2b79f5) | 0;
      let t = Math.imul(a ^ (a >>> 15), 1 | a);
      t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
      return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
    };
  }
  const rngFrom = (str) => mulberry32(hashSeed(str));
  const randInt = (rng, a, b) => a + Math.floor(rng() * (b - a + 1));
  const pick = (rng, arr) => arr[Math.floor(rng() * arr.length)];
  function shuffle(rng, arr) {
    const a = arr.slice();
    for (let i = a.length - 1; i > 0; i--) {
      const j = Math.floor(rng() * (i + 1));
      [a[i], a[j]] = [a[j], a[i]];
    }
    return a;
  }

  // ---------- pomocnicze do porównań ----------
  const isEmpty = (v) => v === "" || v === null || v === undefined;
  const num = (v) =>
    typeof v === "number" ? v
      : typeof v === "string" && v.trim() !== "" && !isNaN(Number(v)) ? Number(v)
      : NaN;
  const eq = (v, e) => { const n = num(v); return !isNaN(n) && Math.abs(n - e) < TOL; };
  const matchMatrix = (vals, exp) => exp.map((row, i) => row.map((e, j) => eq(vals[i][j], e)));
  const all = (m) => m.every((r) => r.every(Boolean));
  const ratio = (m) => { let c = 0, n = 0; m.forEach((r) => r.forEach((x) => { n++; if (x) c++; })); return n ? c / n : 0; };
  const and = (a, b) => a.map((r, i) => r.map((x, j) => x && b[i][j]));
  const allEmpty = (vals) => vals.every((r) => r.every(isEmpty));
  const anyEmpty = (vals) => vals.some((r) => r.some(isEmpty));
  const hasFormula = (f) => typeof f === "string" && f.startsWith("=");
  const countNoFormula = (formulas) => formulas.flat().filter((f) => !hasFormula(f) && !isEmpty(f)).length;
  // odwołanie do innego skoroszytu lub adres www = próba nadużycia
  const EXTERNAL = /\[[^\]]+\.(xlsx|xlsm|xlsb|xls|csv|ods)\]|https?:\/\//i;

  // Rozmiar zakresu "B2:C13" -> {rows, cols}
  function colNum(s) { let n = 0; for (const ch of s) n = n * 26 + (ch.charCodeAt(0) - 64); return n; }
  function dims(addr) {
    const [a, b] = addr.split(":");
    const p = (x) => { const m = /^([A-Z]+)(\d+)$/.exec(x); return { c: colNum(m[1]), r: Number(m[2]) }; };
    const s = p(a), e = b ? p(b) : s;
    return { rows: e.r - s.r + 1, cols: e.c - s.c + 1 };
  }
  const filled = (addr, v) => { const d = dims(addr); return Array.from({ length: d.rows }, () => Array(d.cols).fill(v)); };

  function result(task, error_type, detail, score, extra) {
    return Object.assign({
      task_id: task.id, skill: task.skill, level: task.level,
      verdict: error_type === "correct" ? "pass" : "fail",
      error_type, detail, score,
    }, extra || {});
  }

  // =====================================================================
  // ZADANIE 1 — poziom podstawowy
  // Skill: logika JEŻELI + odwołanie bezwzględne do parametru
  // =====================================================================
  const NAMES = ["Anna Kowalska", "Piotr Nowak", "Marta Wiśniewska", "Tomasz Wójcik", "Katarzyna Kamińska",
    "Michał Lewandowski", "Agnieszka Zielińska", "Paweł Szymański", "Ewa Woźniak", "Jakub Dąbrowski",
    "Magdalena Kozłowska", "Krzysztof Jankowski", "Joanna Mazur", "Łukasz Krawczyk", "Natalia Piotrowska"];
  const RATES = [0.03, 0.04, 0.05, 0.06, 0.07];

  function t1rows(rng, names) {
    const rows = names.map((name) => {
      const target = randInt(rng, 20, 60) * 1000;
      const sales = Math.round((target * (0.7 + rng() * 0.6)) / 100) * 100;
      return { name, sales, target };
    });
    const idx = shuffle(rng, [...Array(rows.length).keys()]);
    rows[idx[0]].sales = rows[idx[0]].target;            // przypadki graniczne: sprzedaż = cel
    rows[idx[1]].sales = rows[idx[1]].target;
    for (const i of idx.slice(2, 5)) rows[i].sales = rows[i].target - randInt(rng, 1, 50) * 100; // poniżej
    for (const i of idx.slice(5, 8)) rows[i].sales = rows[i].target + randInt(rng, 1, 50) * 100; // powyżej
    return rows;
  }

  const T1 = {
    id: "T1_IF_PREMIA",
    skill: "logika_IF_odwolanie_bezwzgledne",
    level: "podstawowy",
    sheet: "Zadanie 1",
    title: "Premia handlowców",
    timeLimitSec: 360,
    instruction:
      "W zakresie D2:D13 oblicz premię każdego handlowca. Premia przysługuje, gdy sprzedaż jest równa celowi " +
      "lub wyższa od niego, i wynosi: sprzedaż × stawka premii z komórki G2. W pozostałych przypadkach premia wynosi 0. " +
      "Wynik musi się poprawnie przeliczać po zmianie danych lub stawki.",
    answer: "D2:D13",

    build(rng) {
      return { rows: t1rows(rng, shuffle(rng, NAMES).slice(0, 12)), rate: pick(rng, RATES) };
    },
    perturb(data, rng) {
      const rows = t1rows(rng, data.rows.map((r) => r.name));
      const rate = pick(rng, RATES.filter((r) => r !== data.rate));
      return { rows, rate };
    },
    layout(d) {
      return [
        { addr: "A1:D1", values: [["Handlowiec", "Sprzedaż", "Cel", "Premia"]], bold: true, fill: "#DDEBF7" },
        { addr: "A2:C13", values: d.rows.map((r) => [r.name, r.sales, r.target]) },
        { addr: "B2:C13", numberFormat: "#,##0" },
        { addr: "F2", values: [["Stawka premii"]], bold: true },
        { addr: "G2", values: [[d.rate]], numberFormat: "0%" },
        { addr: "D2:D13", fill: "#FFF2CC", numberFormat: "#,##0.00" },
      ];
    },
    // komórki, które test perturbacyjny podmienia i potem przywraca
    sourceWrites(d) {
      return [
        { addr: "B2:C13", values: d.rows.map((r) => [r.sales, r.target]) },
        { addr: "G2", values: [[d.rate]] },
      ];
    },
    expected(d, variant) {
      const cond = {
        ok: (r) => r.sales >= r.target,
        gt: (r) => r.sales > r.target,        // błąd graniczny
        rev: (r) => r.sales < r.target,       // logika odwrócona
        rev_le: (r) => r.sales <= r.target,
      }[variant || "ok"];
      return d.rows.map((r) => [cond(r) ? r.sales * d.rate : 0]);
    },
    classify({ formulas, valuesOrig, valuesPert, dataO, dataP }) {
      const mO = matchMatrix(valuesOrig, this.expected(dataO, "ok"));
      const mP = matchMatrix(valuesPert, this.expected(dataP, "ok"));
      const both = and(mO, mP);
      if (all(both)) return result(this, "correct", "Poprawnie, formuła reaguje na zmianę danych i stawki.", 1);
      if (all(mO)) {
        return countNoFormula(formulas) > 0
          ? result(this, "hardcoded", "Wyniki wpisane ręcznie zamiast formuły.", 0)
          : result(this, "hardcoded", "Formuła nie reaguje na zmianę danych — prawdopodobnie stawka wpisana na sztywno zamiast odwołania do G2.", 0.25);
      }
      if (all(matchMatrix(valuesOrig, this.expected(dataO, "rev"))) || all(matchMatrix(valuesOrig, this.expected(dataO, "rev_le"))))
        return result(this, "logic_reversed", "Warunek odwrócony — premia naliczona osobom poniżej celu.", 0);
      if (all(matchMatrix(valuesOrig, this.expected(dataO, "gt"))))
        return result(this, "partial", "Błąd graniczny: użyto > zamiast >= (sprzedaż równa celowi bez premii).", 0.5);
      if (anyEmpty(valuesOrig) && ratio(both) > 0)
        return result(this, "partial", "Nie wszystkie komórki zostały wypełnione.", 0.5 * ratio(both));
      if (ratio(both) >= 0.5)
        return result(this, "partial", "Część wyników poprawna.", 0.5 * ratio(both));
      return result(this, "wrong_function", "Wyniki niezgodne z treścią zadania.", 0);
    },
  };

  // =====================================================================
  // ZADANIE 2 — poziom średniozaawansowany
  // Skill: SUMA.WARUNKÓW (SUMIFS) z dwoma kryteriami + odwołania mieszane przy kopiowaniu
  // =====================================================================
  const REGIONS = ["Północ", "Południe", "Wschód", "Zachód"];
  const QUARTERS = ["Q1", "Q2", "Q3", "Q4"];
  const PRODUCTS = ["Laptop", "Monitor", "Drukarka", "Router", "Tablet", "Słuchawki"];
  const T2_N = 48;

  const T2 = {
    id: "T2_SUMIFS_MATRIX",
    skill: "SUMIFS_dwa_kryteria_odwolania_mieszane",
    level: "sredniozaawansowany",
    sheet: "Zadanie 2",
    title: "Sprzedaż według regionu i kwartału",
    timeLimitSec: 480,
    instruction:
      "W tabeli G2:J5 oblicz łączną wartość sprzedaży dla każdej pary: region (kolumna F) i kwartał (wiersz 1). " +
      "Dane źródłowe znajdują się w zakresie A2:D49. Zalecane: jedna formuła w G2 skopiowana na cały zakres. " +
      "Wyniki muszą się poprawnie przeliczać po zmianie danych lub nagłówków.",
    answer: "G2:J5",

    build(rng) {
      const combos = [];
      for (const r of REGIONS) for (const q of QUARTERS) combos.push([r, q]);
      const rows = [];
      for (let i = 0; i < T2_N; i++) {
        const [region, quarter] = i < combos.length ? combos[i] : [pick(rng, REGIONS), pick(rng, QUARTERS)];
        rows.push({ region, quarter, product: pick(rng, PRODUCTS), value: randInt(rng, 50, 900) * 10 });
      }
      return { rows: shuffle(rng, rows), rowLabels: REGIONS.slice(), colLabels: QUARTERS.slice() };
    },
    perturb(data, rng) {
      return {
        rows: data.rows.map((r) => Object.assign({}, r, { value: randInt(rng, 50, 900) * 10 })),
        rowLabels: data.rowLabels.slice().reverse(),                          // inna kolejność regionów
        colLabels: data.colLabels.slice(2).concat(data.colLabels.slice(0, 2)), // inna kolejność kwartałów
      };
    },
    layout(d) {
      return [
        { addr: "A1:D1", values: [["Region", "Kwartał", "Produkt", "Wartość"]], bold: true, fill: "#DDEBF7" },
        { addr: `A2:D${T2_N + 1}`, values: d.rows.map((r) => [r.region, r.quarter, r.product, r.value]) },
        { addr: `D2:D${T2_N + 1}`, numberFormat: "#,##0" },
        { addr: "F1", values: [["Region / Kwartał"]], bold: true },
        { addr: "F2:F5", values: d.rowLabels.map((x) => [x]), bold: true, fill: "#DDEBF7" },
        { addr: "G1:J1", values: [d.colLabels], bold: true, fill: "#DDEBF7" },
        { addr: "G2:J5", fill: "#FFF2CC", numberFormat: "#,##0" },
      ];
    },
    sourceWrites(d) {
      return [
        { addr: `D2:D${T2_N + 1}`, values: d.rows.map((r) => [r.value]) },
        { addr: "F2:F5", values: d.rowLabels.map((x) => [x]) },
        { addr: "G1:J1", values: [d.colLabels] },
      ];
    },
    expected(d, variant) {
      const v = variant || "ok";
      return d.rowLabels.map((rl) => d.colLabels.map((cl) => {
        let s = 0;
        for (const r of d.rows) {
          const okR = r.region === rl, okQ = r.quarter === cl;
          if (v === "ok" && okR && okQ) s += r.value;
          else if (v === "regionOnly" && okR) s += r.value;
          else if (v === "quarterOnly" && okQ) s += r.value;
          else if (v === "count" && okR && okQ) s += 1;
        }
        return s;
      }));
    },
    classify({ formulas, valuesOrig, valuesPert, dataO, dataP }) {
      const mO = matchMatrix(valuesOrig, this.expected(dataO, "ok"));
      const mP = matchMatrix(valuesPert, this.expected(dataP, "ok"));
      const both = and(mO, mP);
      if (all(both)) return result(this, "correct", "Poprawnie, wyniki reagują na zmianę danych i nagłówków.", 1);
      if (all(mO)) {
        return countNoFormula(formulas) > 0
          ? result(this, "hardcoded", "Wyniki wpisane ręcznie zamiast formuły.", 0)
          : result(this, "hardcoded", "Kryteria wpisane na sztywno w formułach (np. \"Północ\") zamiast odwołań do nagłówków.", 0.25);
      }
      if (all(matchMatrix(valuesOrig, this.expected(dataO, "regionOnly"))) || all(matchMatrix(valuesOrig, this.expected(dataO, "quarterOnly"))))
        return result(this, "partial", "Pominięte jedno z dwóch kryteriów (tylko region albo tylko kwartał).", 0.25);
      if (all(matchMatrix(valuesOrig, this.expected(dataO, "count"))))
        return result(this, "wrong_function", "Policzono liczbę transakcji zamiast sumy wartości (COUNTIFS zamiast SUMIFS).", 0);
      const r = ratio(both);
      if (mO[0][0] && ratio(mO) < 0.5)
        return result(this, "partial", "Pierwsza komórka poprawna, pozostałe błędne — przy kopiowaniu przesunęły się odwołania (brak $).", 0.25);
      if (anyEmpty(valuesOrig) && r > 0)
        return result(this, "partial", "Nie wszystkie komórki zostały wypełnione.", 0.5 * r);
      if (r >= 0.5) return result(this, "partial", "Część wyników poprawna.", 0.5 * r);
      return result(this, "wrong_function", "Wyniki niezgodne z treścią zadania.", 0);
    },
  };

  // =====================================================================
  // Wspólna ocena — reguły ważne dla każdego zadania
  // =====================================================================
  function evaluate(task, input) {
    const flags = [];
    if (input.tampered) flags.push("source_modified");
    const ext = input.formulas.flat().find((f) => typeof f === "string" && EXTERNAL.test(f));
    if (ext) return result(task, "abuse", "Odwołanie zewnętrzne w formule — sesja oznaczona.", 0, { flags: flags.concat("external_reference") });
    if (allEmpty(input.valuesOrig)) return result(task, "empty", "Brak odpowiedzi.", 0, { flags });
    const r = task.classify(input);
    r.flags = flags;
    r.score = Math.round(r.score * 100) / 100;
    return r;
  }

  // Werdykt poziomu — reguła z dokumentacji SkillCheck (liczy kod, nie AI)
  function levelVerdict(results) {
    const pct = (lvl) => {
      const rs = results.filter((r) => r.level === lvl);
      return rs.length ? rs.reduce((s, r) => s + r.score, 0) / rs.length : 0;
    };
    const basic = pct("podstawowy"), inter = pct("sredniozaawansowany");
    let level = "Podstawowy";
    if (inter >= 0.6) level = "Zaawansowany";
    else if (inter >= 0.2 || basic >= 0.6) level = "Średniozaawansowany";
    return { level, basic_pct: Math.round(basic * 100), intermediate_pct: Math.round(inter * 100) };
  }

  const api = {
    TASKS: [T1, T2],
    CONTROL_TASKS: {},          // skill -> [zadania kontrolne]; do uzupełnienia w kolejnym kroku
    rngFrom, evaluate, levelVerdict, dims, filled, isEmpty, TOL,
  };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  root.SkillCheckTasks = api;
})(typeof globalThis !== "undefined" ? globalThis : this);
