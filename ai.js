/*
 * SkillCheck — warstwa AI (dyrygent przebiegu).
 * AI dostaje gotowy werdykt z kodu i wybiera WYŁĄCZNIE ścieżkę:
 *   continue  — następne zadanie z planu,
 *   probe     — zadanie kontrolne tego samego skilla (tylko z banku, max 2 na skill),
 *   note_gap  — zanotuj lukę i idź dalej.
 * Granice egzekwuje ten kod, nie prompt. Awaria / zła odpowiedź AI => "continue".
 * Klucz API nigdy nie trafia do dodatku: wywołanie idzie przez backend (worker/ai-worker.js).
 */
(function (root) {
  "use strict";

  const AI = {
    endpoint: "",          // np. "https://skillcheck-ai.<konto>.workers.dev/decide"; puste = reguła-zaślepka
    timeoutMs: 8000,
    MAX_PROBES_PER_SKILL: 2,
  };

  // Reguła-zaślepka: zachowuje się jak rozsądne AI, dopóki nie podepniemy prawdziwego
  function ruleDecide(p) {
    const last = p.last_result;
    if (last.error_type === "correct") return { action: "continue", reason: "Zadanie poprawne." };
    const controls = p.control_tasks_available[last.skill] || [];
    if (["logic_reversed", "partial", "hardcoded"].includes(last.error_type) && controls.length && (p.probes_used[last.skill] || 0) < AI.MAX_PROBES_PER_SKILL)
      return { action: "probe", task_id: controls[0], reason: "Błąd może być przypadkowy — zadanie kontrolne." };
    return { action: "note_gap", gap: `${last.skill}: ${last.detail}`, reason: "Luka potwierdzona albo brak zadań kontrolnych." };
  }

  // Twarda walidacja decyzji — AI nie może wyjść poza bank zadań ani limity
  function validate(d, p) {
    if (!d || !["continue", "probe", "note_gap"].includes(d.action)) return null;
    if (d.action === "probe") {
      const skill = p.last_result.skill;
      const allowed = p.control_tasks_available[skill] || [];
      if (!allowed.includes(d.task_id) || (p.probes_used[skill] || 0) >= AI.MAX_PROBES_PER_SKILL)
        return { action: "note_gap", gap: `${skill}: ${p.last_result.detail}`, reason: "Decyzja AI poza limitem — zamieniona na notatkę luki." };
    }
    if (d.action === "note_gap" && !d.gap) d.gap = `${p.last_result.skill}: ${p.last_result.detail}`;
    return d;
  }

  async function remoteDecide(p) {
    const ctrl = new AbortController();
    const t = setTimeout(() => ctrl.abort(), AI.timeoutMs);
    try {
      const res = await fetch(AI.endpoint, {
        method: "POST", headers: { "Content-Type": "application/json" },
        body: JSON.stringify(p), signal: ctrl.signal,
      });
      if (!res.ok) throw new Error("HTTP " + res.status);
      return await res.json();
    } finally { clearTimeout(t); }
  }

  /**
   * payload = { last_result, history, remaining_task_ids, control_tasks_available, probes_used }
   * zwraca  = { action, task_id?, gap?, reason, source: "ai" | "rule" | "fallback" }
   */
  AI.decide = async function (payload) {
    if (!AI.endpoint) return Object.assign(validate(ruleDecide(payload), payload), { source: "rule" });
    try {
      const d = validate(await remoteDecide(payload), payload);
      if (d) return Object.assign(d, { source: "ai" });
    } catch (e) { /* przejście do fallbacku */ }
    return { action: "continue", reason: "AI niedostępne — kontynuacja planu.", source: "fallback" };
  };

  AI._ruleDecide = ruleDecide;
  AI._validate = validate;
  if (typeof module !== "undefined" && module.exports) module.exports = AI;
  root.SkillCheckAI = AI;
})(typeof globalThis !== "undefined" ? globalThis : this);
