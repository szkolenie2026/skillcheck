/*
 * SkillCheck — warstwa Excela (Office.js).
 * Odpowiada tylko za: wstawienie arkusza z zadaniem, odczyt formuł i wartości,
 * test perturbacyjny (podmiana danych źródłowych -> odczyt -> przywrócenie).
 * Formuły czytane przez range.formulas są zawsze w notacji angielskiej
 * (JEŻELI -> IF), niezależnie od języka Excela kandydata.
 */
(function (root) {
  "use strict";
  const { dims, filled, TOL } = root.SkillCheckTasks;

  function applyWrite(ws, w) {
    const r = ws.getRange(w.addr);
    if (w.values) r.values = w.values;
    if (w.numberFormat) r.numberFormat = filled(w.addr, w.numberFormat);
    if (w.bold) r.format.font.bold = true;
    if (w.fill) r.format.fill.color = w.fill;
  }

  async function setupTask(task, data) {
    await Excel.run(async (ctx) => {
      const sheets = ctx.workbook.worksheets;
      const old = sheets.getItemOrNullObject(task.sheet);
      await ctx.sync();
      if (!old.isNullObject) { old.name = task.sheet + "_old_" + Date.now() % 100000; await ctx.sync(); }
      const ws = sheets.add(task.sheet);
      for (const w of task.layout(data)) applyWrite(ws, w);
      ws.getRange("A1:J1").format.columnWidth = 90;
      ws.getRange("A:A").format.columnWidth = 130;
      ws.activate();
      ws.getRange(task.answer.split(":")[0]).select();
      if (!old.isNullObject) old.delete();
      await ctx.sync();
    });
  }

  async function readAnswer(task) {
    return Excel.run(async (ctx) => {
      const r = ctx.workbook.worksheets.getItem(task.sheet).getRange(task.answer);
      r.load("formulas,values");
      await ctx.sync();
      return { formulas: r.formulas, values: r.values };
    });
  }

  // Czy kandydat zmienił dane źródłowe?
  async function checkTampered(task, data) {
    return Excel.run(async (ctx) => {
      const ws = ctx.workbook.worksheets.getItem(task.sheet);
      const writes = task.sourceWrites(data);
      const ranges = writes.map((w) => { const r = ws.getRange(w.addr); r.load("values"); return r; });
      await ctx.sync();
      return writes.some((w, k) => w.values.some((row, i) => row.some((v, j) => {
        const got = ranges[k].values[i][j];
        return typeof v === "number" ? Math.abs(Number(got) - v) > TOL : String(got) !== String(v);
      })));
    });
  }

  // Test perturbacyjny: dane B -> odczyt wyników -> przywrócenie danych A
  async function perturbCheck(task, dataO, dataP) {
    return Excel.run(async (ctx) => {
      const ws = ctx.workbook.worksheets.getItem(task.sheet);
      try {
        for (const w of task.sourceWrites(dataP)) ws.getRange(w.addr).values = w.values;
        await ctx.sync();
        try { ctx.workbook.application.calculate(Excel.CalculationType.full); await ctx.sync(); } catch (e) { /* web: przeliczanie automatyczne */ }
        const r = ws.getRange(task.answer);
        r.load("values");
        await ctx.sync();
        return r.values;
      } finally {
        for (const w of task.sourceWrites(dataO)) ws.getRange(w.addr).values = w.values;
        await ctx.sync();
      }
    });
  }

  // Telemetria: liczba edycji w arkuszu zadania
  async function watchEdits(task, onEdit) {
    return Excel.run(async (ctx) => {
      const ws = ctx.workbook.worksheets.getItem(task.sheet);
      const handle = ws.onChanged.add(async (ev) => { onEdit(ev.address); });
      await ctx.sync();
      return handle;
    });
  }
  async function unwatch(handle) {
    if (!handle) return;
    try { await Excel.run(handle.context, async (ctx) => { handle.remove(); await ctx.sync(); }); } catch (e) { /* ignoruj */ }
  }

  root.SkillCheckExcel = { setupTask, readAnswer, checkTampered, perturbCheck, watchEdits, unwatch, dims };
})(globalThis);
