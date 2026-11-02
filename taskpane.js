/* SkillCheck — sterowanie panelem: kolejka zadań, timer, ocena, decyzja AI, raport. */
(function () {
  "use strict";
  const S = globalThis.SkillCheckTasks;
  const X = globalThis.SkillCheckExcel;
  const AI = globalThis.SkillCheckAI;
  const $ = (id) => document.getElementById(id);

  const state = {
    session: "", queue: [], idx: 0, current: null,
    results: [], decisions: [], gaps: [], probes: {},
    timer: null, startedAt: 0, edits: 0, firstEditAt: null, phase: "idle", watch: null,
  };

  function show(screen) {
    ["start", "task", "result", "summary"].forEach((s) => ($("screen-" + s).hidden = s !== screen));
  }
  const fmt = (s) => String(Math.floor(s / 60)).padStart(2, "0") + ":" + String(s % 60).padStart(2, "0");
  const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
  const allControlTasks = () => Object.values(S.CONTROL_TASKS).flat();
  const findTask = (id) => S.TASKS.concat(allControlTasks()).find((t) => t.id === id);

  // ---------- start ----------
  async function startSession() {
    $("start-error").textContent = "";
    state.session = ($("session").value || "").trim() || "S-" + Math.random().toString(36).slice(2, 8).toUpperCase();
    state.queue = S.TASKS.slice();
    state.idx = 0;
    $("btn-start").disabled = true;
    try { await runTask(state.queue[0]); }
    catch (e) { $("start-error").textContent = "Nie udało się przygotować zadania: " + e.message; $("btn-start").disabled = false; }
  }

  async function runTask(task) {
    const rng = S.rngFrom(state.session + ":" + task.id);
    const dataO = task.build(rng);
    const dataP = task.perturb(dataO, S.rngFrom(state.session + ":" + task.id + ":p"));
    state.current = { task, dataO, dataP };
    await X.setupTask(task, dataO);

    $("task-title").textContent = task.title;
    $("task-instruction").textContent = task.instruction;
    $("task-sheet").textContent = task.sheet;
    $("task-answer").textContent = task.answer;
    $("task-status").textContent = "";
    $("btn-done").disabled = false;
    $("progress").textContent = `Zadanie ${state.idx + 1} / ${state.queue.length}`;
    show("task");

    state.edits = 0; state.firstEditAt = null; state.phase = "working";
    state.watch = await X.watchEdits(task, () => {
      if (state.phase !== "working") return;
      state.edits++;
      if (!state.firstEditAt) state.firstEditAt = Date.now();
    }).catch(() => null);

    state.startedAt = Date.now();
    $("timer").textContent = fmt(task.timeLimitSec);
    $("timer").classList.remove("late");
    clearInterval(state.timer);
    state.timer = setInterval(() => {
      const el = Math.floor((Date.now() - state.startedAt) / 1000);
      const left = task.timeLimitSec - el;
      $("timer").textContent = fmt(Math.max(0, left));
      $("timer").classList.toggle("late", left <= 60);
      if (left <= 0) submit(true);
    }, 500);
  }

  // ---------- Gotowe ----------
  async function submit(timeout) {
    if (state.phase !== "working") return;
    state.phase = "checking";
    clearInterval(state.timer);
    $("btn-done").disabled = true;
    $("task-status").textContent = timeout ? "Koniec czasu — sprawdzanie…" : "Sprawdzanie…";
    const { task, dataO, dataP } = state.current;
    const timeSec = Math.round((Date.now() - state.startedAt) / 1000);
    try {
      await X.unwatch(state.watch);
      const ans = await X.readAnswer(task);
      const tampered = await X.checkTampered(task, dataO);
      const valuesPert = await X.perturbCheck(task, dataO, dataP);
      const res = S.evaluate(task, { formulas: ans.formulas, valuesOrig: ans.values, valuesPert, dataO, dataP, tampered });
      Object.assign(res, {
        time_sec: timeSec, timed_out: !!timeout, edits: state.edits,
        time_to_first_edit_sec: state.firstEditAt ? Math.round((state.firstEditAt - state.startedAt) / 1000) : null,
        formula_sample: ans.formulas.flat().find((f) => typeof f === "string" && f.startsWith("=")) || null,
      });
      state.results.push(res);

      const remaining = state.queue.slice(state.idx + 1).map((t) => t.id);
      const control = {};
      for (const [skill, list] of Object.entries(S.CONTROL_TASKS)) control[skill] = list.map((t) => t.id).filter((id) => !state.results.some((r) => r.task_id === id));
      const decision = await AI.decide({
        last_result: res, history: state.results.slice(0, -1),
        remaining_task_ids: remaining, control_tasks_available: control, probes_used: state.probes,
      });
      state.decisions.push(Object.assign({ after_task: task.id }, decision));
      if (decision.action === "probe") {
        state.probes[res.skill] = (state.probes[res.skill] || 0) + 1;
        state.queue.splice(state.idx + 1, 0, findTask(decision.task_id));
      }
      if (decision.action === "note_gap") state.gaps.push(decision.gap);
      renderResult(res, decision);
    } catch (e) {
      $("task-status").textContent = "Błąd odczytu: " + e.message;
      $("btn-done").disabled = false;
      state.phase = "working";
    }
  }

  function tagFor(r) {
    const cls = r.error_type === "correct" ? "correct" : r.error_type === "partial" ? "partial" : "fail";
    return `<span class="tag ${cls}">${esc(r.error_type)}</span>`;
  }

  function renderResult(r, d) {
    const examiner = $("examiner").checked;
    $("result-box").innerHTML = examiner
      ? `<div class="card">${tagFor(r)} <b>${esc(r.score)} pkt</b><p>${esc(r.detail)}</p>
         <table><tr><th>Czas</th><td>${fmt(r.time_sec)}${r.timed_out ? " (limit)" : ""}</td></tr>
         <tr><th>Edycje</th><td>${r.edits}</td></tr>
         <tr><th>Formuła</th><td><code>${esc(r.formula_sample || "—")}</code></td></tr>
         ${r.flags.length ? `<tr><th>Flagi</th><td>${esc(r.flags.join(", "))}</td></tr>` : ""}</table></div>`
      : `<div class="card">Odpowiedź zapisana.</div>`;
    $("decision-box").innerHTML = examiner
      ? `<div class="card"><b>Decyzja AI (${esc(d.source)}):</b> ${esc(d.action)}${d.task_id ? " → " + esc(d.task_id) : ""}<br><span class="muted">${esc(d.reason || "")}</span></div>`
      : "";
    show("result");
  }

  async function next() {
    state.idx++;
    if (state.idx < state.queue.length) {
      $("btn-next").disabled = true;
      try { await runTask(state.queue[state.idx]); } finally { $("btn-next").disabled = false; }
    } else summary();
  }

  function summary() {
    state.phase = "done";
    const verdict = S.levelVerdict(state.results);
    const report = {
      session: state.session, finished_at: new Date().toISOString(),
      level: verdict, results: state.results, ai_decisions: state.decisions, gaps: state.gaps,
    };
    $("summary-box").innerHTML =
      `<div class="card"><b>Poziom: ${esc(verdict.level)}</b><br>
       <span class="muted">Podstawowy: ${verdict.basic_pct}% · Średniozaawansowany: ${verdict.intermediate_pct}%</span></div>
       <table><tr><th>Zadanie</th><th>Wynik</th><th>Pkt</th></tr>
       ${state.results.map((r) => `<tr><td>${esc(r.task_id)}</td><td>${tagFor(r)}</td><td>${r.score}</td></tr>`).join("")}</table>
       ${state.gaps.length ? `<p><b>Luki:</b></p><ul>${state.gaps.map((g) => `<li>${esc(g)}</li>`).join("")}</ul>` : ""}`;
    $("report-json").textContent = JSON.stringify(report, null, 2);
    $("progress").textContent = "Zakończono";
    show("summary");
  }

  async function copyReport() {
    const txt = $("report-json").textContent;
    try { await navigator.clipboard.writeText(txt); $("btn-copy").textContent = "Skopiowano"; }
    catch (e) {
      const range = document.createRange(); range.selectNodeContents($("report-json"));
      const sel = window.getSelection(); sel.removeAllRanges(); sel.addRange(range);
      $("btn-copy").textContent = "Zaznaczono — Ctrl+C";
    }
  }

  Office.onReady((info) => {
    if (info.host !== Office.HostType.Excel) { $("start-error").textContent = "Ten dodatek działa tylko w Excelu."; return; }
    if (!Office.context.requirements.isSetSupported("ExcelApi", "1.7")) { $("start-error").textContent = "Ta wersja Excela jest zbyt stara (wymagane ExcelApi 1.7)."; return; }
    $("btn-start").onclick = startSession;
    $("btn-done").onclick = () => submit(false);
    $("btn-next").onclick = next;
    $("btn-copy").onclick = copyReport;
  });
})();
