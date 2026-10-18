(function () {
  "use strict";
  const $ = id => document.getElementById(id);
  const core = window.AppCore, Store = window.LocalStore, X = window.XLSX;
  if (!core || !Store || !X || !X.read || !X.CFB) { window.reportStartupError("程式元件未完整載入，請重新整理網頁。"); return; }

  const COLORS = ["#1f5fae", "#c2410c", "#15803d", "#7c3aed", "#be185d", "#0f766e"];
  const SCORES = ["0.1", "0.2", "0.3", "0.4", "0.5", "0.6", "0.7", "0.8"];

  /* ───────── 狀態 ───────── */
  let cohorts = [];                 // [{id,label,color,fileName,importedAt,sheetName,students}]
  let settings = {activeCohort: "", autoLockMinutes: 15, tab: ""};
  let pTpl = null;                  // {fileName, importedAt, bytes, options}
  let rTpl = null;                  // {fileName, importedAt, bytes, parsed}
  let pRecords = [], rRecords = [];

  const status = (id, message, kind) => { const el = $(id); el.hidden = false; el.textContent = message; el.className = "status" + (kind ? " " + kind : ""); };
  const fmtDate = t => { const d = new Date(t); return d.getFullYear() + "/" + (d.getMonth() + 1) + "/" + d.getDate(); };
  const stampDash = () => new Date().toLocaleDateString("sv-SE");
  const activeCohort = () => cohorts.find(c => c.id === settings.activeCohort) || null;
  const otherCohorts = () => cohorts.filter(c => c.id !== settings.activeCohort);
  const colorOf = label => { const c = cohorts.find(c => c.label === label); return c ? c.color : "#64748b"; };
  function cohortTag(label) {
    const span = document.createElement("span");
    span.className = "cohort-tag"; span.textContent = label || "—"; span.style.setProperty("--c", colorOf(label));
    return span;
  }
  function fillOptions(select, values, placeholder) {
    select.replaceChildren();
    if (placeholder !== undefined) select.add(new Option(placeholder, ""));
    values.forEach(value => select.add(new Option(value, value)));
  }

  // 下拉選單：值維持原本的條款目代碼，顯示文字加上事由，方便直接看內容挑選
  function fillLabeled(select, pairs, placeholder) {
    select.replaceChildren(new Option(placeholder, ""));
    pairs.forEach(([value, label]) => { const option = new Option(label, value); option.title = label; select.add(option); });
  }
  // 「條」顯示獎懲程度（學生獎懲規定第16～24條）；範本若有其他條，改顯示範本設定的額度
  const ARTICLE_LEVELS = {
    "第16條": "予以嘉獎", "第17條": "予以記功", "第18條": "予以記大功",
    "第19條": "予以申誡", "第20條": "予以記過", "第21條": "予以記大過",
    "第22條": "予以留校察看", "第23條": "予以勒令退學或退訓", "第24條": "予以開除學籍或退訓"
  };
  function articleChoices(t) {
    return core.unique(t.rules.map(rule => rule.article)).map(article => {
      const level = ARTICLE_LEVELS[article] || (t.defaultAmounts[article] ? "額度：" + t.defaultAmounts[article] : "");
      return [article, level ? article + "｜" + level : article];
    });
  }
  function clauseChoices(t, article) {
    const rows = t.rules.filter(rule => rule.article === article);
    return core.unique(rows.map(rule => rule.clause)).map(clause => {
      const own = rows.filter(rule => rule.clause === clause), items = core.unique(own.map(rule => rule.item).filter(Boolean));
      return [clause, clause + "｜" + (items.length ? "（分 " + items.length + " 目，請在下一欄選擇）" : own[0].reason)];
    });
  }
  function itemChoices(relevant) {
    const seen = new Set();
    return relevant.filter(rule => rule.item && !seen.has(rule.item) && seen.add(rule.item)).map(rule => [rule.item, rule.item + "｜" + rule.reason]);
  }

  /* ───────── 存檔 ───────── */
  let saveTimer = null;
  function saveDrafts() {
    clearTimeout(saveTimer);
    saveTimer = setTimeout(flushDrafts, 250);
    updateTabCounts();
  }
  async function flushDrafts() {
    clearTimeout(saveTimer); saveTimer = null;
    try { await Store.setJSON("drafts", {points: pRecords, reward: rRecords}); }
    catch (error) { console.warn("草稿儲存失敗", error); }
  }
  const saveSettings = () => Store.setJSON("settings", settings).catch(error => console.warn(error));
  const saveCohorts = () => Store.setJSON("cohorts", cohorts);

  function download(bytes, filename) {
    const url = URL.createObjectURL(new Blob([bytes], {type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}));
    const a = document.createElement("a");
    a.href = url; a.download = filename; a.rel = "noopener";
    a.style.display = "none";
    document.body.appendChild(a); a.click();
    setTimeout(() => { a.remove(); URL.revokeObjectURL(url); }, 30000);
  }

  function readFile(file) {
    return new Promise((resolve, reject) => {
      if (!file.size) { reject(new Error("檔案是空的；若來自 iCloud，請先下載至「檔案」後重選。")); return; }
      const reader = new FileReader();
      const timer = setTimeout(() => { reader.onabort = null; reader.abort(); reject(new Error("讀取逾時，請先將 Excel 下載至本機再重新選取。")); }, 30000);
      reader.onload = () => { clearTimeout(timer); resolve(new Uint8Array(reader.result)); };
      reader.onerror = () => { clearTimeout(timer); reject(new Error("無法讀取檔案，請先下載至本機後重選。")); };
      reader.onabort = () => { clearTimeout(timer); reject(new Error("檔案讀取已中止，請重新選取。")); };
      try { reader.readAsArrayBuffer(file); } catch (error) { clearTimeout(timer); reject(new Error("無法開啟檔案：" + error.message)); }
    });
  }
  function parseWorkbook(bytes) {
    try {
      const wb = X.read(bytes, {type: "array"});
      if (!wb.SheetNames.length) throw new Error("沒有可讀取的工作表。");
      return wb;
    } catch (error) { throw new Error("Excel 解析失敗：" + error.message); }
  }

  /* ═══════════ 鎖定與解鎖 ═══════════ */

  async function boot() {
    try {
      const state = await Store.status();
      $("bootStatus").hidden = true;
      if (state === "new") { $("setupForm").hidden = false; $("setupPw").focus(); }
      else { $("unlockForm").hidden = false; $("unlockPw").focus(); }
    } catch (error) {
      status("bootStatus", error.message, "error");
    }
    document.documentElement.dataset.appReady = "true";
  }

  $("setupForm").addEventListener("submit", async event => {
    event.preventDefault();
    const pw = $("setupPw").value, pw2 = $("setupPw2").value;
    if (pw.length < 6) { $("setupStatus").textContent = "密碼至少 6 個字元。"; return; }
    if (pw !== pw2) { $("setupStatus").textContent = "兩次輸入的密碼不一致。"; return; }
    $("setupBtn").disabled = true; $("setupStatus").textContent = "";
    try { await Store.create(pw); await enterApp(true); }
    catch (error) { $("setupStatus").textContent = error.message; $("setupBtn").disabled = false; }
  });

  $("unlockForm").addEventListener("submit", async event => {
    event.preventDefault();
    $("unlockBtn").disabled = true; $("unlockStatus").textContent = "";
    try {
      if (await Store.unlock($("unlockPw").value)) { $("unlockPw").value = ""; await enterApp(false); }
      else { $("unlockStatus").textContent = "密碼不正確。"; $("unlockPw").select(); }
    } catch (error) { $("unlockStatus").textContent = error.message; }
    $("unlockBtn").disabled = false;
  });

  async function wipeEverything() {
    if (!confirm("確定清除這台裝置上本 App 的全部資料？\n（名冊、範本、尚未匯出的紀錄、密碼設定都會刪除，無法復原）")) return;
    if (!confirm("再次確認：真的要清除嗎？")) return;
    await Store.wipe();
    location.reload();
  }
  $("forgotWipe").addEventListener("click", wipeEverything);
  $("wipeAll").addEventListener("click", wipeEverything);

  async function lockNow() {
    await flushDrafts();
    Store.lock();
    location.reload(); // 重新載入頁面，清掉記憶體中的學生資料
  }
  $("lockBtn").addEventListener("click", lockNow);

  // 閒置自動鎖定
  let lastActive = Date.now();
  ["pointerdown", "keydown", "input", "wheel", "touchstart"].forEach(type =>
    document.addEventListener(type, () => { lastActive = Date.now(); }, {passive: true, capture: true}));
  setInterval(() => {
    const minutes = Number(settings.autoLockMinutes);
    if ($("app").hidden || !minutes) return;
    if (Date.now() - lastActive > minutes * 60000) lockNow();
  }, 15000);
  document.addEventListener("visibilitychange", () => {
    const minutes = Number(settings.autoLockMinutes);
    if (document.visibilityState === "visible" && !$("app").hidden && minutes && Date.now() - lastActive > minutes * 60000) lockNow();
    if (document.visibilityState === "hidden" && !$("app").hidden) flushDrafts();
  });

  async function enterApp(firstTime) {
    cohorts = await Store.getJSON("cohorts", []);
    settings = Object.assign(settings, await Store.getJSON("settings", {}));
    const drafts = await Store.getJSON("drafts", {points: [], reward: []});
    pRecords = drafts.points || []; rRecords = drafts.reward || [];
    const loadTpl = async kind => {
      const meta = await Store.getJSON("tpl-" + kind, null), bytes = await Store.getBytes("tpl-" + kind);
      if (!meta || !bytes) return null;
      try {
        const wb = parseWorkbook(bytes);
        return kind === "points" ? Object.assign({}, meta, {bytes, options: core.extractPointsTemplate(wb, X).options})
          : Object.assign({}, meta, {bytes, parsed: core.extractRewardTemplate(wb, X)});
      } catch (error) { console.warn(error); return null; }
    };
    pTpl = await loadTpl("points");
    rTpl = await loadTpl("reward");
    if (!activeCohort()) settings.activeCohort = cohorts[0] ? cohorts[0].id : "";

    $("lockScreen").hidden = true;
    $("app").hidden = false;
    lastActive = Date.now();
    $("autoLock").value = String(settings.autoLockMinutes);

    renderCohorts();
    renderTemplates();
    pRenderCategories(); pRenderEditScores();
    pRenderRecords();
    rResetRuleControls();
    rRender();
    const ready = cohorts.length && pTpl && rTpl;
    showTab(firstTime || !cohorts.length ? "settings" : settings.tab || (ready ? "points" : "settings"));
    if (firstTime) status("rosterStatus", "密碼已設定。請先匯入名冊（每期一份），再匯入兩份學校範本；之後打開 App 輸入密碼就能直接使用。", "ok");
    if (pRecords.length || rRecords.length) {
      const parts = [];
      if (pRecords.length) parts.push("加扣分 " + pRecords.length + " 筆");
      if (rRecords.length) parts.push("獎懲 " + rRecords.length + " 筆");
      status("rResult", "已還原上次未清空的紀錄：" + parts.join("、") + "。", "ok");
      $("pError").style.color = "var(--ok)";
      $("pError").textContent = "已還原上次未清空的紀錄：" + parts.join("、") + "。";
    }
    const persisted = await Store.requestPersistence();
    $("persistStatus").textContent = persisted === true ? "✓ 瀏覽器已同意長期保存本機資料。" :
      persisted === false ? "瀏覽器未同意長期保存；長時間未使用或空間不足時，資料可能被清除（iPhone 請加入主畫面使用）。" : "";
  }

  /* ═══════════ 分頁 ═══════════ */

  function showTab(tab) {
    ["points", "reward", "settings"].forEach(name => {
      $("tab-" + name).hidden = name !== tab;
      $("tabBtn-" + name).setAttribute("aria-selected", String(name === tab));
    });
    if (settings.tab !== tab) { settings.tab = tab; saveSettings(); }
  }
  document.querySelectorAll(".tab").forEach(button => button.addEventListener("click", () => showTab(button.dataset.tab)));
  function updateTabCounts() {
    $("pTabCount").textContent = pRecords.length || "";
    $("rTabCount").textContent = rRecords.length || "";
  }

  /* ═══════════ 期別 ═══════════ */

  function renderCohorts() {
    const active = activeCohort();
    document.documentElement.style.setProperty("--cohort", active ? active.color : "#ae2530");
    const box = $("cohortButtons");
    box.replaceChildren();
    cohorts.forEach(cohort => {
      const button = document.createElement("button");
      button.type = "button"; button.className = "cohort-btn"; button.setAttribute("role", "radio");
      button.setAttribute("aria-checked", String(cohort.id === settings.activeCohort));
      button.style.setProperty("--c", cohort.color);
      button.textContent = cohort.label;
      const small = document.createElement("small"); small.textContent = cohort.students.length + " 人";
      button.append(small);
      button.addEventListener("click", () => setActiveCohort(cohort.id));
      box.append(button);
    });
    $("cohortEmpty").hidden = cohorts.length > 0;
    document.querySelectorAll(".active-cohort-tag").forEach(tag => {
      tag.textContent = active ? active.label : "尚未選擇期別";
      tag.style.setProperty("--c", active ? active.color : "#ae2530");
    });

    const list = $("cohortList");
    list.replaceChildren();
    if (!cohorts.length) {
      const p = document.createElement("p"); p.className = "err"; p.textContent = "尚未匯入任何名冊。";
      list.append(p);
    }
    cohorts.forEach(cohort => {
      const card = document.createElement("div");
      card.className = "cohort-card"; card.style.setProperty("--c", cohort.color);
      const info = document.createElement("div"); info.className = "info";
      const title = document.createElement("strong"); title.textContent = cohort.label;
      const meta = document.createElement("div"); meta.className = "note";
      meta.textContent = cohort.students.length + " 位學生｜" + cohort.fileName + "｜" + fmtDate(cohort.importedAt) + " 匯入";
      info.append(title, meta);
      const update = document.createElement("button");
      update.type = "button"; update.className = "secondary"; update.textContent = "更新名冊";
      update.addEventListener("click", () => { pendingTarget = cohort.id; $("rosterFile").click(); });
      const remove = document.createElement("button");
      remove.type = "button"; remove.className = "danger"; remove.textContent = "刪除";
      remove.addEventListener("click", async () => {
        if (!confirm("確定刪除「" + cohort.label + "」名冊？已新增的紀錄會保留。")) return;
        cohorts = cohorts.filter(c => c.id !== cohort.id);
        if (settings.activeCohort === cohort.id) settings.activeCohort = cohorts[0] ? cohorts[0].id : "";
        await saveCohorts(); saveSettings();
        renderCohorts(); refreshChecks();
        status("rosterStatus", "已刪除「" + cohort.label + "」名冊。", "ok");
      });
      card.append(info, update, remove);
      list.append(card);
    });
  }

  function setActiveCohort(id) {
    if (settings.activeCohort === id) return;
    settings.activeCohort = id;
    saveSettings();
    renderCohorts();
    refreshChecks();
  }
  function refreshChecks() { pCheckStudents(); rCheckStudents(); pUpdateAddState(); rRefreshAddState(); }

  // 匯入名冊
  let pendingRoster = null, pendingTarget = "";
  $("rosterFile").addEventListener("change", async event => {
    const input = event.target, file = input.files && input.files[0];
    if (!file) return;
    $("rosterPending").hidden = true;
    status("rosterStatus", "正在讀取：" + file.name + "…");
    try {
      const result = core.extractRoster(parseWorkbook(await readFile(file)), X);
      pendingRoster = {fileName: file.name, sheetName: result.sheetName, students: result.students, guess: core.guessCohortLabel(result.students)};
      showPendingRoster();
      $("rosterStatus").hidden = true;
    } catch (error) {
      status("rosterStatus", "名冊匯入失敗：" + error.message, "error");
      pendingTarget = "";
    } finally { input.value = ""; }
  });

  function showPendingRoster() {
    const p = pendingRoster;
    $("rosterPendingInfo").textContent = p.fileName + "｜「" + p.sheetName + "」：" + p.students.length + " 位學生" + (p.guess ? "（依建制隊號判斷為" + p.guess + "）" : "");
    const target = $("rosterTarget");
    target.replaceChildren(new Option("新增為新的期別", "new"));
    cohorts.forEach(c => target.add(new Option("更新既有名冊：" + c.label, c.id)));
    const sameLabel = cohorts.find(c => c.label === p.guess);
    target.value = pendingTarget || (sameLabel ? sameLabel.id : "new");
    $("rosterLabel").value = p.guess;
    $("rosterPending").hidden = false;
    syncPendingRoster();
    $("rosterPending").scrollIntoView({block: "nearest"});
  }
  function overlap(students, cohort) {
    const ids = new Set(cohort.students.map(s => s.school));
    return students.filter(s => ids.has(s.school)).length;
  }
  function syncPendingRoster() {
    if (!pendingRoster) return;
    const target = $("rosterTarget").value, students = pendingRoster.students;
    $("rosterLabelWrap").hidden = target !== "new";
    const warnings = [];
    if (target === "new") {
      cohorts.forEach(c => {
        const n = overlap(students, c);
        if (n > students.length / 2) warnings.push("這份名冊有 " + n + " 位學生已在「" + c.label + "」，可能是同一期；若是，請改選「更新既有名冊：" + c.label + "」。");
      });
    } else {
      const c = cohorts.find(c => c.id === target);
      const n = overlap(students, c);
      if (c && n < Math.min(students.length, c.students.length) * 0.2) warnings.push("這份名冊和原本的「" + c.label + "」幾乎沒有相同學生（" + n + " 位），請確認沒有選錯期別。");
    }
    $("rosterPendingWarn").textContent = warnings.join("\n");
  }
  $("rosterTarget").addEventListener("change", syncPendingRoster);
  $("rosterCancel").addEventListener("click", () => { pendingRoster = null; pendingTarget = ""; $("rosterPending").hidden = true; });
  $("rosterSave").addEventListener("click", async () => {
    if (!pendingRoster) return;
    const target = $("rosterTarget").value, p = pendingRoster;
    let saved;
    if (target === "new") {
      const label = core.clean($("rosterLabel").value);
      if (!label) { $("rosterPendingWarn").textContent = "請輸入期別名稱，例如：第45期"; return; }
      if (cohorts.some(c => c.label === label)) { $("rosterPendingWarn").textContent = "已有「" + label + "」，若要覆蓋請改選「更新既有名冊」。"; return; }
      const used = new Set(cohorts.map(c => c.color));
      saved = {id: "c" + Date.now().toString(36) + Math.random().toString(36).slice(2, 6), label,
        color: COLORS.find(c => !used.has(c)) || COLORS[cohorts.length % COLORS.length]};
      cohorts.push(saved);
    } else saved = cohorts.find(c => c.id === target);
    Object.assign(saved, {fileName: p.fileName, sheetName: p.sheetName, importedAt: Date.now(), students: p.students});
    if (!activeCohort()) settings.activeCohort = saved.id;
    try {
      await saveCohorts(); await saveSettings();
    } catch (error) { status("rosterStatus", "儲存失敗：" + error.message, "error"); return; }
    pendingRoster = null; pendingTarget = "";
    $("rosterPending").hidden = true;
    renderCohorts(); refreshChecks();
    status("rosterStatus", "已儲存「" + saved.label + "」名冊：" + saved.students.length + " 位學生。之後開啟 App 不需要再匯入。", "ok");
  });

  /* ═══════════ 範本 ═══════════ */

  function renderTemplates() {
    if (pTpl) status("pTplStatus", "✓ " + pTpl.fileName + "｜" + Object.keys(pTpl.options).length + " 個分類｜" + fmtDate(pTpl.importedAt) + " 匯入", "ok");
    else status("pTplStatus", "尚未匯入（加扣分功能需要此範本）", "warning");
    if (rTpl) status("rTplStatus", "✓ " + rTpl.fileName + "｜獎懲項目 " + rTpl.parsed.rules.length + " 筆｜" + fmtDate(rTpl.importedAt) + " 匯入", "ok");
    else status("rTplStatus", "尚未匯入（獎懲功能需要此範本）", "warning");
  }

  async function importTemplate(kind, input) {
    const file = input.files && input.files[0];
    if (!file) return;
    const statusId = kind === "points" ? "pTplStatus" : "rTplStatus";
    status(statusId, "正在讀取：" + file.name + "…");
    try {
      const bytes = await readFile(file), wb = parseWorkbook(bytes);
      const meta = {fileName: file.name, importedAt: Date.now()};
      let next;
      if (kind === "points") {
        next = Object.assign({}, meta, {bytes, options: core.extractPointsTemplate(wb, X).options});
        core.exportPointsWorkbook(bytes, [{school: "0", name: "x", cat: "優良x", reason: "x", detail: "", score: 0.1}], X, window); // 先試匯出，確認格式可用
      } else next = Object.assign({}, meta, {bytes, parsed: core.extractRewardTemplate(wb, X)});
      await Store.setBytes("tpl-" + kind, bytes);
      await Store.setJSON("tpl-" + kind, meta);
      if (kind === "points") { pTpl = next; pRenderCategories(); }
      else { rTpl = next; rResetRuleControls(); rRefreshAddState(); }
      renderTemplates();
    } catch (error) {
      renderTemplates();
      status(statusId, "匯入失敗：" + error.message + (kind === "points" ? (pTpl ? "\n（仍使用原本的範本）" : "") : (rTpl ? "\n（仍使用原本的範本）" : "")), "error");
    } finally { input.value = ""; }
  }
  $("pTplFile").addEventListener("change", event => importTemplate("points", event.target));
  $("rTplFile").addEventListener("change", event => importTemplate("reward", event.target));

  /* ═══════════ 設定 ═══════════ */

  $("autoLock").addEventListener("change", () => { settings.autoLockMinutes = Number($("autoLock").value); saveSettings(); });
  $("pwForm").addEventListener("submit", async event => {
    event.preventDefault();
    const out = $("pwStatus");
    out.className = "err";
    if ($("pwNew").value.length < 6) { out.textContent = "新密碼至少 6 個字元。"; return; }
    if ($("pwNew").value !== $("pwNew2").value) { out.textContent = "兩次輸入的新密碼不一致。"; return; }
    out.className = ""; out.textContent = "變更中…";
    try {
      await flushDrafts();
      await Store.changePassword($("pwOld").value, $("pwNew").value);
      $("pwForm").reset(); out.className = "ok"; out.textContent = "✓ 密碼已變更。";
    } catch (error) { out.className = "err"; out.textContent = error.message; }
  });

  /* ═══════════ 學生辨識面板（兩個分頁共用） ═══════════ */

  function renderCheck(panel, value, mode) {
    panel.replaceChildren();
    const cohort = activeCohort();
    if (!core.clean(value)) { panel.textContent = "尚未輸入學生資料"; return null; }
    if (!cohort) { panel.textContent = "請先到「名冊與設定」匯入名冊。"; panel.classList.add("err"); return null; }
    panel.classList.remove("err");
    try {
      const resolved = core.resolveStudents(value, cohort.students, otherCohorts());
      resolved.matches.forEach(({student, count, notes}) => {
        const row = document.createElement("div");
        row.className = "student-row" + (count > 1 ? " error" : "");
        const dupText = mode === "points" ? "｜本批輸入 " + count + " 次，只新增 1 筆" : "｜本批共 " + count + " 筆（重複 " + (count - 1) + " 次，全部新增）";
        row.textContent = student.name + (count > 1 ? dupText : "｜辨識成功") +
          "\n女生隊號：" + (student.female || "—") + "　建制隊號：" + (student.team || "—") + "　學籍號碼：" + student.school +
          (notes.length ? "\n註記：" + notes.join("、") : "");
        panel.appendChild(row);
      });
      resolved.errors.forEach(error => {
        const row = document.createElement("div");
        row.className = "student-row error";
        row.textContent = error.token + "：" + error.message;
        panel.appendChild(row);
      });
      const summary = document.createElement("p");
      summary.className = "summary-line";
      summary.textContent = "【" + cohort.label + "】輸入 " + resolved.tokens.length + " 次；辨識 " + resolved.matches.length + " 位；重複 " +
        resolved.matches.filter(match => match.count > 1).length + " 位；無法辨識 " + resolved.errors.length + " 筆。";
      panel.prepend(summary);
      return resolved;
    } catch (error) {
      panel.textContent = error.message; panel.classList.add("err");
      return null;
    }
  }

  /* ═══════════ 加扣分分頁 ═══════════ */

  const pOptions = () => (pTpl ? pTpl.options : {});
  const favoriteReasons = {
    "加分": [
      "24.對指派或推選擔任之工作，執行得力，成績優良者",
      "25.公差勤務工作努力，績效優良者",
      "30.一週內務經常保持整潔者",
      "31.在校精神飽滿服裝整潔足資表率者"
    ],
    "扣分": [
      "22.對指派或推選擔任之工作執行錯誤，情節尚輕者",
      "26.處理事物怠於注意，致生不良後果，情節尚輕者",
      "59.儀容不整，指甲未修者",
      "63.內務不整者",
      "69.其它違反生活規範規定，情節輕微者"
    ]
  };
  // 比對文字而不只比對編號；匯出時保留範本原文。
  const reasonKey = text => text.normalize("NFKC").replace(/^\s*\d+\s*[.、．:：)）-]?\s*/, "").replace(/[\s\p{P}]/gu, "");

  fillOptions($("pScore"), SCORES); $("pScore").value = "0.2";
  function pRenderEditScores() { fillOptions($("pEditScore"), SCORES); }

  function pCheckStudents() { renderCheck($("pCheck"), $("pIds").value, "points"); }
  function pUpdateAddState() { $("pAdd").disabled = false; }

  function pRenderCategories() {
    const prefix = core.kindPrefix($("pKind").value);
    fillOptions($("pCategory"), Object.keys(pOptions()).filter(k => k.startsWith(prefix)));
    pRenderItems();
  }
  function pRenderItems() {
    fillOptions($("pItem"), pOptions()[$("pCategory").value] || []);
    pRenderReasonBrowser(false);
  }
  $("pKind").addEventListener("change", pRenderCategories);
  $("pCategory").addEventListener("change", pRenderItems);
  $("pItem").addEventListener("change", () => pRenderReasonBrowser(false));
  $("pEditItem").addEventListener("change", () => pRenderReasonBrowser(true));
  $("pSearch").addEventListener("input", () => pRenderReasonBrowser(false));
  $("pEditSearch").addEventListener("input", () => pRenderReasonBrowser(true));
  $("pIds").addEventListener("input", pCheckStudents);

  function pRenderReasonBrowser(edit) {
    const prefix = edit ? "pEdit" : "p", kind = $(prefix + "Kind").value, options = pOptions();
    const category = $(prefix + "Category"), item = $(prefix + "Item");
    const words = $(prefix + "Search").value.trim().toLocaleLowerCase().split(/\s+/).filter(Boolean);
    const results = $(prefix + "Results"); results.replaceChildren();
    let count = 0, groups = 0;
    const categories = Object.entries(options).filter(([cat]) => cat.startsWith(core.kindPrefix(kind)));
    const makeChoice = (cat, reason, target) => {
      const button = document.createElement("button");
      button.type = "button"; button.className = "reason-choice"; button.textContent = reason;
      button.setAttribute("aria-pressed", String(category.value === cat && item.value === reason));
      button.setAttribute("aria-label", cat + "：" + reason);
      button.addEventListener("click", () => {
        category.value = cat;
        if (edit) pEditItemOptions(reason);
        else { pRenderItems(); item.value = reason; pRenderReasonBrowser(false); }
        const pressed = $(prefix + target).querySelector('[aria-pressed="true"]');
        if (pressed) pressed.focus({preventScroll: true});
      });
      return button;
    };
    const favorites = $(prefix + "Favorites"); favorites.replaceChildren();
    let missing = 0;
    favoriteReasons[kind].forEach(label => {
      const matches = categories.flatMap(([cat, items]) => Array.from(new Set(items)).filter(reason => reasonKey(reason) === reasonKey(label)).map(reason => ({cat, reason})));
      if (!matches.length) {
        missing++;
        const button = document.createElement("button");
        button.type = "button"; button.className = "reason-choice"; button.disabled = true; button.textContent = label;
        const note = document.createElement("small"); note.textContent = pTpl ? "範本中未找到此項目" : "請先匯入加扣分範本";
        button.append(note); favorites.append(button);
      } else matches.forEach(({cat, reason}) => {
        const button = makeChoice(cat, reason, "Favorites"), note = document.createElement("small");
        note.textContent = cat; button.append(note); favorites.append(button);
      });
    });
    $(prefix + "FavoriteHint").textContent = !pTpl ? "匯入加扣分範本後，即可點選常用項目。" :
      kind + "常用項目，點選即帶入歸類與事蹟。" + (missing ? " 灰色項目表示範本中未找到相符內容。" : "");
    categories.forEach(([cat, items]) => {
      const matches = items.filter(reason => words.every(word => (cat + " " + reason).toLocaleLowerCase().includes(word)));
      if (!matches.length) return;
      groups++;
      const heading = document.createElement("h3"); heading.textContent = cat; results.append(heading);
      matches.forEach(reason => { count++; results.append(makeChoice(cat, reason, "Results")); });
    });
    $(prefix + "ResultCount").textContent = !pTpl ? "請先匯入加扣分範本，即可瀏覽事蹟說明。" :
      count ? kind + "：共 " + groups + " 個分類、" + count + " 項事蹟。" : words.length ? "找不到符合的事蹟，請更換關鍵字。" : "範本中沒有「" + kind + "」事蹟。";
  }

  $("pAdd").addEventListener("click", () => {
    const out = $("pError");
    out.style.color = "var(--err)";
    try {
      const result = core.buildPointsRecords($("pIds").value, activeCohort(), otherCohorts(), pTpl, {
        kind: $("pKind").value, score: $("pScore").value, cat: $("pCategory").value, reason: $("pItem").value, detail: $("pDetail").value
      });
      pRecords.push(...result.records);
      pRenderRecords(); saveDrafts();
      out.style.color = "var(--ok)";
      out.textContent = "已新增 " + result.records.length + " 筆紀錄（" + activeCohort().label + "）" +
        (result.warnings.length ? "；以下輸入未新增或已合併：\n" + result.warnings.join("\n") : "。");
      if (result.warnings.length) out.style.color = "var(--warn)";
    } catch (error) {
      out.textContent = error.message + (pRecords.length ? "\n原有 " + pRecords.length + " 筆紀錄已保留。" : "");
    }
  });

  // ── 預覽、勾選、編輯 ──
  let pEditingIndex = null, pEditingTargets = [], pBulk = false;
  const pSelected = new Set();

  function pSyncEditFields() {
    $("pEditDetail").disabled = pBulk && !$("pApplyDetail").checked;
    $("pEditScore").disabled = pBulk && !$("pApplyScore").checked;
    $("pEditKind").disabled = pBulk && !$("pApplyKind").checked;
    $("pEditReasonFields").hidden = pBulk && !$("pApplyKind").checked;
  }
  function pSetEditMode(bulk) {
    pBulk = bulk;
    ["Detail", "Kind", "Score"].forEach(key => { const box = $("pApply" + key); box.checked = !bulk; box.parentElement.hidden = !bulk; });
    pSyncEditFields();
  }
  ["Detail", "Kind", "Score"].forEach(key => $("pApply" + key).addEventListener("change", pSyncEditFields));

  function pUpdateSelection() {
    for (const r of Array.from(pSelected)) if (!pRecords.includes(r)) pSelected.delete(r);
    const n = pSelected.size, all = pRecords.length > 0 && n === pRecords.length, some = n > 0 && n < pRecords.length;
    ["", "Bottom"].forEach(suffix => {
      $("pSelected" + suffix).textContent = "已勾選 " + n + " 筆";
      $("pBulkEdit" + suffix).disabled = n === 0;
      $("pBulkDelete" + suffix).disabled = n === 0;
      $("pSelectAll" + suffix).disabled = pRecords.length === 0;
      $("pSelectAll" + suffix).checked = all;
      $("pSelectAll" + suffix).indeterminate = some;
    });
    $("pRecords").querySelectorAll("button[data-edit-index]").forEach(button => {
      button.textContent = n > 1 && pSelected.has(pRecords[Number(button.dataset.editIndex)]) ? "編輯勾選 " + n + " 筆" : "編輯";
    });
    $("pRecords").querySelectorAll("input[data-record-index]").forEach(box => {
      box.checked = pSelected.has(pRecords[Number(box.dataset.recordIndex)]);
      box.closest("tr").classList.toggle("selected-record", box.checked);
    });
  }
  function pSelectAll(event) {
    if (pBulk) pCancelEditing();
    pSelected.clear();
    if (event.target.checked) pRecords.forEach(r => pSelected.add(r));
    pUpdateSelection();
  }
  $("pSelectAll").addEventListener("change", pSelectAll);
  $("pSelectAllBottom").addEventListener("change", pSelectAll);

  function pEditCategoryOptions(selectedCategory, selectedItem) {
    const prefix = core.kindPrefix($("pEditKind").value);
    const categories = Object.keys(pOptions()).filter(k => k.startsWith(prefix));
    fillOptions($("pEditCategory"), categories);
    if (categories.includes(selectedCategory)) $("pEditCategory").value = selectedCategory;
    pEditItemOptions(selectedItem);
  }
  function pEditItemOptions(selectedItem) {
    const items = pOptions()[$("pEditCategory").value] || [];
    fillOptions($("pEditItem"), items);
    if (items.includes(selectedItem)) $("pEditItem").value = selectedItem;
    pRenderReasonBrowser(true);
  }
  $("pEditKind").addEventListener("change", () => pEditCategoryOptions(null, null));
  $("pEditCategory").addEventListener("change", () => pEditItemOptions(null));

  function pParkEditPanel() {
    $("pEditPanel").hidden = true;
    $("pCount").after($("pEditPanel"));
    const row = $("pInlineEditRow"); if (row) row.remove();
  }
  function pCancelEditing() { pBulk = false; pEditingIndex = null; pEditingTargets = []; pParkEditPanel(); $("pEditError").textContent = ""; }
  $("pCancelEdit").addEventListener("click", pCancelEditing);

  function pStartEditing(i) {
    const r = pRecords[i]; if (!r) return;
    if (!pTpl) { alert("請先匯入加扣分範本才能編輯。"); return; }
    pParkEditPanel(); pSetEditMode(false);
    pEditingIndex = i; pEditingTargets = [r];
    $("pEditHint").textContent = "僅修改目前選取的學生，不影響其他紀錄。";
    $("pSaveEdit").textContent = "儲存修改";
    $("pEditTitle").textContent = "編輯紀錄：" + (r.female || r.team) + " " + r.name + "（" + r.cohort + "）";
    $("pEditKind").value = r.kind; $("pEditScore").value = r.score.toFixed(1);
    pEditCategoryOptions(r.cat, r.reason);
    $("pEditDetail").value = r.detail || ""; $("pEditError").textContent = "";
    const row = document.createElement("tr"); row.id = "pInlineEditRow";
    const cell = document.createElement("td"); cell.colSpan = 11; row.append(cell);
    $("pRecords").children[i].after(row); cell.append($("pEditPanel"));
    $("pEditPanel").hidden = false;
  }
  function pStartBulkEditing(anchorIndex, position) {
    const targets = pRecords.filter(r => pSelected.has(r)); if (!targets.length) return;
    pStartEditing(anchorIndex == null ? pRecords.indexOf(targets[0]) : anchorIndex);
    if (pEditingIndex === null) return;
    pEditingTargets = targets;
    const first = targets[0];
    $("pEditKind").value = first.kind; $("pEditScore").value = first.score.toFixed(1);
    pEditCategoryOptions(first.cat, first.reason); $("pEditDetail").value = first.detail || "";
    pSetEditMode(true);
    $("pEditTitle").textContent = "批次編輯 " + targets.length + " 筆紀錄";
    $("pEditHint").textContent = "先帶入第一筆勾選紀錄。只有右側勾選「一起編輯」的欄位會統一套用，未勾選的欄位保留各筆原值。編輯加扣分類型時，歸類與事蹟也會一併套用。";
    $("pSaveEdit").textContent = "套用至 " + targets.length + " 筆紀錄";
    if (anchorIndex == null) {
      $(position === "bottom" ? "pBulkBottom" : "pBulkTop").after($("pEditPanel"));
      const row = $("pInlineEditRow"); if (row) row.remove();
    }
  }
  $("pBulkEdit").addEventListener("click", () => pStartBulkEditing(null, "top"));
  $("pBulkEditBottom").addEventListener("click", () => pStartBulkEditing(null, "bottom"));

  function pDeleteSelected() {
    const targets = new Set(pRecords.filter(r => pSelected.has(r)));
    if (!targets.size) return;
    if (!confirm("確定刪除勾選的 " + targets.size + " 筆紀錄？未勾選的紀錄會保留。")) return;
    pCancelEditing();
    pRecords = pRecords.filter(r => !targets.has(r)); pSelected.clear();
    pRenderRecords(); saveDrafts();
    $("pCount").textContent += "（已刪除 " + targets.size + " 筆紀錄）";
  }
  $("pBulkDelete").addEventListener("click", pDeleteSelected);
  $("pBulkDeleteBottom").addEventListener("click", pDeleteSelected);

  $("pSaveEdit").addEventListener("click", () => {
    if (pEditingIndex === null || !pRecords[pEditingIndex]) return;
    const applyDetail = !pBulk || $("pApplyDetail").checked, applyKind = !pBulk || $("pApplyKind").checked, applyScore = !pBulk || $("pApplyScore").checked;
    const patch = {};
    if (applyDetail) patch.detail = $("pEditDetail").value.trim();
    if (applyScore) patch.score = Math.round(Number($("pEditScore").value) * 10) / 10;
    if (applyKind) Object.assign(patch, {kind: $("pEditKind").value, cat: $("pEditCategory").value, reason: $("pEditItem").value});
    const errors = core.validatePointsPatch(pTpl, patch);
    $("pEditError").textContent = errors.join("\n");
    if (errors.length) return;
    const targets = pEditingTargets.filter(r => pRecords.includes(r));
    if (targets.length !== pEditingTargets.length) return;
    targets.forEach(r => Object.assign(r, patch));
    pCancelEditing(); pRenderRecords(); saveDrafts();
    $("pCount").textContent += "（已儲存 " + targets.length + " 筆修改）";
  });

  function cohortSummary(records) {
    const counts = new Map();
    records.forEach(r => counts.set(r.cohort, (counts.get(r.cohort) || 0) + 1));
    return counts.size > 1 ? "（" + Array.from(counts).map(([label, n]) => label + " " + n + " 筆").join("、") + "）" : "";
  }

  function pRenderRecords() {
    if (pEditingIndex !== null) pCancelEditing();
    $("pCount").textContent = pRecords.length ? "本次共 " + pRecords.length + " 筆" + cohortSummary(pRecords) : "尚無紀錄";
    $("pRecords").replaceChildren();
    pRecords.forEach((r, i) => {
      const tr = document.createElement("tr");
      const selectCell = document.createElement("td"), box = document.createElement("input");
      box.type = "checkbox"; box.dataset.recordIndex = i;
      box.setAttribute("aria-label", "勾選第 " + (i + 1) + " 筆：" + (r.female || "") + " " + r.name);
      box.addEventListener("change", () => { if (pBulk) pCancelEditing(); if (box.checked) pSelected.add(r); else pSelected.delete(r); pUpdateSelection(); });
      selectCell.append(box); tr.append(selectCell);
      const cohortCell = document.createElement("td"); cohortCell.append(cohortTag(r.cohort)); tr.append(cohortCell);
      [[r.female, "nowrap"], [r.team, "nowrap"], [r.name, "nowrap"], [r.detail || "", "wrap"], [r.school, "nowrap"], [r.kind, "nowrap"], [r.score.toFixed(1), "nowrap"], [r.reason, "wrap"]].forEach(([text, cls]) => {
        const td = document.createElement("td"); td.className = cls; td.textContent = text; tr.appendChild(td);
      });
      const td = document.createElement("td"); td.className = "nowrap";
      const edit = document.createElement("button"), del = document.createElement("button");
      edit.type = "button"; edit.textContent = "編輯"; edit.dataset.editIndex = i;
      edit.addEventListener("click", () => { if (pSelected.size > 1 && pSelected.has(r)) pStartBulkEditing(i); else pStartEditing(i); });
      del.type = "button"; del.textContent = "刪除"; del.className = "danger";
      del.addEventListener("click", () => { pRecords.splice(pRecords.indexOf(r), 1); pSelected.delete(r); pRenderRecords(); saveDrafts(); });
      td.append(edit, del); tr.appendChild(td);
      $("pRecords").appendChild(tr);
    });
    $("pExport").disabled = !pRecords.length;
    $("pClear").disabled = !pRecords.length;
    pUpdateSelection(); updateTabCounts();
  }

  $("pClear").addEventListener("click", () => {
    if (!confirm("確定清空全部 " + pRecords.length + " 筆加扣分紀錄？")) return;
    pRecords = []; pSelected.clear(); pCancelEditing(); pRenderRecords(); saveDrafts();
    $("pError").textContent = "";
  });
  $("pExport").addEventListener("click", () => {
    const out = $("pError");
    try {
      if (!pTpl) throw new Error("請先到「名冊與設定」匯入加扣分範本");
      if (!pRecords.length) throw new Error("尚無紀錄");
      const bytes = core.exportPointsWorkbook(pTpl.bytes, pRecords, X, window);
      download(bytes, "加扣分APP匯入_" + stampDash() + ".xlsx");
      out.style.color = "var(--ok)"; out.textContent = "已下載 " + pRecords.length + " 筆紀錄。確認匯入學校系統成功後，可按「清空紀錄」。";
    } catch (error) { out.style.color = "var(--err)"; out.textContent = "匯出失敗：" + error.message; }
  });

  /* ═══════════ 獎懲分頁 ═══════════ */

  const rT = () => (rTpl ? rTpl.parsed : null);
  const rSelectedRows = new Set();
  let rEditingRows = [];

  function rRefreshAddState() { $("rAdd").disabled = !rTpl || !activeCohort(); }
  function rCheckStudents() { renderCheck($("rCheck"), $("rIds").value, "reward"); }
  $("rIds").addEventListener("input", rCheckStudents);

  function rResetRuleControls() {
    const t = rT();
    fillLabeled($("rArticle"), t ? articleChoices(t) : [], t ? "請選擇條" : "請先匯入獎懲範本");
    fillOptions($("rClause"), [], "請先選擇條");
    fillOptions($("rItem"), [], "請先選擇款");
    $("rAmount").value = "";
    fillOptions($("rAdjustReason"), t ? t.adjustmentOptions : [], "無（選填）");
    fillOptions($("rAdjustAmount"), t ? t.amountOptions : [], "無（選填）");
    $("rArticle").disabled = !t;
    ["rClause", "rItem", "rAmount", "rAdjustReason", "rAdjustAmount"].forEach(id => { $(id).disabled = true; });
    $("rReason").value = ""; $("rReason").disabled = true;
  }
  function rSyncRuleDetails() {
    const t = rT();
    const rule = t && t.rules.find(row => row.article === $("rArticle").value && row.clause === $("rClause").value && row.item === $("rItem").value);
    $("rReason").value = rule ? rule.reason : "";
    $("rReason").disabled = !rule; $("rAmount").disabled = !rule; $("rAdjustReason").disabled = !rule;
    $("rAmount").value = rule ? t.defaultAmounts[rule.article] || "" : "";
    $("rAdjustReason").value = ""; $("rAdjustAmount").value = ""; $("rAdjustAmount").disabled = true;
  }
  $("rArticle").addEventListener("change", () => {
    const t = rT(), article = $("rArticle").value;
    fillLabeled($("rClause"), article ? clauseChoices(t, article) : [], article ? "請選擇款" : "請先選擇條");
    $("rClause").disabled = !article;
    fillOptions($("rItem"), [], "請先選擇款"); $("rItem").disabled = true;
    rSyncRuleDetails();
  });
  $("rClause").addEventListener("change", () => {
    const relevant = rT().rules.filter(rule => rule.article === $("rArticle").value && rule.clause === $("rClause").value);
    const items = itemChoices(relevant);
    fillLabeled($("rItem"), items, items.length ? "請選擇目" : relevant.length ? "無目（不需選擇）" : "請先選擇款");
    $("rItem").disabled = !items.length;
    rSyncRuleDetails();
  });
  $("rItem").addEventListener("change", rSyncRuleDetails);
  $("rAdjustReason").addEventListener("change", () => {
    $("rAdjustAmount").disabled = !$("rAdjustReason").value;
    if (!$("rAdjustReason").value) $("rAdjustAmount").value = "";
  });

  function rReport(message, kind) { status("rResult", message, kind); }

  $("rAdd").addEventListener("click", () => {
    try {
      const selection = Object.fromEntries([["article", "rArticle"], ["clause", "rClause"], ["item", "rItem"], ["reason", "rReason"], ["adjustReason", "rAdjustReason"], ["adjustAmount", "rAdjustAmount"]].map(([key, id]) => [key, $(id).value]));
      const result = core.buildRewardRecords($("rIds").value, activeCohort(), otherCohorts(), rT(), selection);
      rRecords = rRecords.concat(result.records);
      rRender(); saveDrafts();
      const skipped = result.resolved.errors.length;
      rReport("已新增 " + result.records.length + " 筆紀錄（" + activeCohort().label + "）；目前共 " + rRecords.length + " 筆，可直接匯出。" +
        (skipped ? "\n已略過 " + skipped + " 筆無法辨識的輸入，請看上方紅字。" : ""), skipped ? "warning" : "ok");
    } catch (error) {
      rReport("新增失敗：" + error.message + (rRecords.length ? "\n原有 " + rRecords.length + " 筆紀錄已保留。" : ""), "error");
    }
  });
  $("rClearInput").addEventListener("click", () => {
    $("rIds").value = ""; rResetRuleControls(); rCheckStudents();
    rReport(rRecords.length ? "已清除上方輸入；原有 " + rRecords.length + " 筆紀錄保留。" : "已清除輸入。");
  });

  function rUpdateSelection() {
    const n = rSelectedRows.size, total = rRecords.length;
    $("rSelected").textContent = "已選取 " + n + " / " + total + " 筆";
    $("rEditSelected").disabled = !n || !rTpl;
    $("rClearRecords").disabled = !total;
    $("rExport").disabled = !total;
    $("rSelectAll").disabled = !total;
    $("rSelectAll").checked = total > 0 && n === total;
    $("rSelectAll").indeterminate = n > 0 && n < total;
  }
  function rUpdateDuplicates() {
    const duplicates = core.duplicateStudents(rRecords);
    $("rDup").hidden = !duplicates.length;
    $("rDup").textContent = duplicates.length ? "重複學生提醒（各筆均會保留及匯出）\n" +
      duplicates.map(s => s.name + "（" + s.cohort + "，學號 " + s.school + "）：共 " + s.count + " 筆，重複 " + (s.count - 1) + " 次。").join("\n") : "";
  }
  function rRender() {
    for (const i of Array.from(rSelectedRows)) if (i >= rRecords.length) rSelectedRows.delete(i);
    $("rBody").replaceChildren();
    rRecords.forEach((record, index) => {
      const row = document.createElement("tr");
      const selectCell = document.createElement("td"), label = document.createElement("label"), checkbox = document.createElement("input");
      label.className = "check-label"; checkbox.type = "checkbox";
      checkbox.checked = rSelectedRows.has(index);
      checkbox.setAttribute("aria-label", "選取 " + record.name + "（" + record.school + "）");
      checkbox.addEventListener("change", () => { if (checkbox.checked) rSelectedRows.add(index); else rSelectedRows.delete(index); row.classList.toggle("selected-record", checkbox.checked); rUpdateSelection(); });
      label.append(checkbox); selectCell.append(label); row.append(selectCell);
      row.classList.toggle("selected-record", checkbox.checked);
      const actionCell = document.createElement("td"); actionCell.className = "nowrap";
      const edit = document.createElement("button"), remove = document.createElement("button");
      edit.type = "button"; edit.className = "secondary"; edit.textContent = "編輯"; edit.disabled = !rTpl;
      edit.addEventListener("click", () => rOpenEditor([index]));
      remove.type = "button"; remove.className = "danger"; remove.textContent = "刪除";
      remove.addEventListener("click", () => rDelete(index));
      actionCell.append(edit, remove); row.append(actionCell);
      const cohortCell = document.createElement("td"); cohortCell.append(cohortTag(record.cohort)); row.append(cohortCell);
      core.rewardFields.forEach(field => {
        const cell = document.createElement("td");
        cell.className = field === "reason" || field === "adjustReason" ? "wrap" : "nowrap";
        cell.textContent = record[field] || "";
        row.append(cell);
      });
      $("rBody").append(row);
    });
    rUpdateSelection(); rUpdateDuplicates(); updateTabCounts();
  }
  function rDelete(index) {
    if (!rRecords[index]) return;
    rRecords = rRecords.filter((r, i) => i !== index);
    const next = Array.from(rSelectedRows).filter(i => i !== index).map(i => i > index ? i - 1 : i);
    rSelectedRows.clear(); next.forEach(i => rSelectedRows.add(i));
    rRender(); saveDrafts();
    rReport(rRecords.length ? "已刪除 1 筆紀錄；目前共 " + rRecords.length + " 筆。" : "已刪除最後一筆紀錄；可再次批次新增。", "ok");
  }
  $("rSelectAll").addEventListener("change", () => {
    rSelectedRows.clear();
    if ($("rSelectAll").checked) rRecords.forEach((r, i) => rSelectedRows.add(i));
    rRender();
  });
  $("rClearRecords").addEventListener("click", () => {
    if (!confirm("確定清空全部 " + rRecords.length + " 筆獎懲紀錄？")) return;
    rRecords = []; rSelectedRows.clear(); rRender(); saveDrafts();
    rReport("已清空紀錄；名冊、範本與上方輸入都保留。");
  });
  $("rExport").addEventListener("click", () => {
    try {
      if (!rTpl) throw new Error("請先到「名冊與設定」匯入獎懲範本");
      const workbook = core.exportRewardWorkbook(rRecords, rT(), X);
      const bytes = X.write(workbook, {bookType: "xlsx", type: "array"});
      download(bytes, "獎懲紀錄_" + stampDash().replace(/-/g, "") + ".xlsx");
      rReport("已下載 " + rRecords.length + " 筆紀錄。確認匯入學校系統成功後，可按「清空紀錄」。", "ok");
    } catch (error) { rReport("匯出失敗：" + error.message, "error"); }
  });

  // ── 獎懲編輯視窗 ──
  function rSyncEditor() {
    const t = rT(), applyRule = $("rEditApplyRule").checked;
    const relevant = t.rules.filter(rule => rule.article === $("rEditArticle").value && rule.clause === $("rEditClause").value);
    $("rEditArticle").disabled = !applyRule;
    $("rEditClause").disabled = !applyRule || !$("rEditArticle").value;
    $("rEditItem").disabled = !applyRule || !relevant.some(rule => rule.item);
    $("rEditReason").disabled = !$("rEditApplyReason").checked;
    $("rEditAdjustReason").disabled = !$("rEditApplyAdjust").checked;
    $("rEditAdjustAmount").disabled = !$("rEditApplyAdjust").checked || !$("rEditAdjustReason").value;
  }
  function rFillEditorReason() {
    const t = rT();
    const rule = t.rules.find(rule => rule.article === $("rEditArticle").value && rule.clause === $("rEditClause").value && rule.item === $("rEditItem").value);
    $("rEditReason").value = rule ? rule.reason : "";
    $("rEditAmount").value = rule ? t.defaultAmounts[rule.article] || "" : "";
  }
  function rFillEditItems(value) {
    const relevant = rT().rules.filter(rule => rule.article === $("rEditArticle").value && rule.clause === $("rEditClause").value);
    const items = itemChoices(relevant);
    fillLabeled($("rEditItem"), items, items.length ? "請選擇目" : relevant.length ? "無目（不需選擇）" : "請先選擇款");
    $("rEditItem").value = value || "";
    rFillEditorReason(); rSyncEditor();
  }
  function rFillEditClauses(clause, item) {
    fillLabeled($("rEditClause"), $("rEditArticle").value ? clauseChoices(rT(), $("rEditArticle").value) : [], "請選擇款");
    $("rEditClause").value = clause || "";
    rFillEditItems(item);
  }
  function rOpenEditor(indexes) {
    if (!rRecords.length || !indexes.length || !rTpl) return;
    const t = rT();
    rEditingRows = indexes.slice();
    const records = indexes.map(i => rRecords[i]);
    const common = field => records.every(r => r[field] === records[0][field]) ? records[0][field] : "";
    const single = indexes.length === 1;
    $("rEditTitle").textContent = "編輯 " + indexes.length + " 筆紀錄";
    $("rEditSummary").textContent = single ? records[0].name + "（" + records[0].cohort + "，學號 " + records[0].school + "）" :
      records.slice(0, 5).map(r => r.name).join("、") + (records.length > 5 ? "等 " + records.length + " 位學生" : "");
    ["rEditApplyRule", "rEditApplyReason", "rEditApplyAdjust"].forEach(id => { $(id).checked = single; });
    fillLabeled($("rEditArticle"), articleChoices(t), "請選擇條");
    $("rEditArticle").value = common("article");
    fillOptions($("rEditAdjustReason"), t.adjustmentOptions, "無（清除）");
    $("rEditAdjustReason").value = common("adjustReason");
    fillOptions($("rEditAdjustAmount"), t.amountOptions, "無（選填）");
    $("rEditAdjustAmount").value = common("adjustAmount");
    rFillEditClauses(common("clause"), common("item"));
    $("rEditReason").value = common("reason");
    $("rEditAmount").value = common("amount");
    $("rEditStatus").textContent = ""; $("rEditStatus").className = "";
    $("rEditor").showModal();
  }
  $("rEditSelected").addEventListener("click", () => rOpenEditor(Array.from(rSelectedRows).sort((a, b) => a - b)));
  $("rEditArticle").addEventListener("change", () => rFillEditClauses());
  $("rEditClause").addEventListener("change", () => rFillEditItems());
  $("rEditItem").addEventListener("change", () => { rFillEditorReason(); rSyncEditor(); });
  ["rEditApplyRule", "rEditApplyReason", "rEditApplyAdjust"].forEach(id => $(id).addEventListener("change", rSyncEditor));
  $("rEditAdjustReason").addEventListener("change", () => { if (!$("rEditAdjustReason").value) $("rEditAdjustAmount").value = ""; rSyncEditor(); });
  $("rCancelEdit").addEventListener("click", () => $("rEditor").close());
  $("rEditor").addEventListener("close", () => { rEditingRows = []; });
  $("rEditForm").addEventListener("submit", event => {
    event.preventDefault();
    try {
      const patch = {};
      if ($("rEditApplyRule").checked) Object.assign(patch, {article: $("rEditArticle").value, clause: $("rEditClause").value, item: $("rEditItem").value});
      if ($("rEditApplyReason").checked) patch.reason = $("rEditReason").value;
      if ($("rEditApplyAdjust").checked) {
        patch.adjustReason = $("rEditAdjustReason").value;
        patch.adjustAmount = patch.adjustReason ? $("rEditAdjustAmount").value : "";
      }
      const count = rEditingRows.length;
      rRecords = core.editRewardRecords(rRecords, rEditingRows, rT(), patch);
      $("rEditor").close();
      rRender(); saveDrafts();
      rReport("已更新 " + count + " 筆紀錄；目前共 " + rRecords.length + " 筆。", "ok");
    } catch (error) {
      $("rEditStatus").textContent = error.message; $("rEditStatus").className = "err";
    }
  });

  window.addEventListener("pagehide", () => { if (!$("app").hidden) flushDrafts(); });
  boot();
})();
