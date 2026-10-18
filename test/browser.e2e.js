// 端對端測試（全部虛構資料）
const {chromium, devices} = require("playwright");
const assert = require("node:assert/strict");
const path = require("node:path");
const fs = require("node:fs");
const FX = path.join(__dirname, "fixtures");
const OUT = path.join(__dirname, "output"); fs.mkdirSync(OUT, {recursive: true});
const URL = process.env.APP_URL || "http://localhost:8765/index.html";

let rDupBefore = 0;
(async () => {
  const browser = await chromium.launch();
  const context = await browser.newContext(process.env.MOBILE ? devices["iPhone 13"] : {viewport: {width: 1200, height: 900}, acceptDownloads: true});
  const page = await context.newPage();
  const external = [], errors = [];
  page.on("request", r => { if (!r.url().startsWith("http://localhost:8765") && !r.url().startsWith("blob:") && !r.url().startsWith("data:")) external.push(r.url()); });
  page.on("pageerror", e => errors.push(e.message));
  page.on("console", m => { if (m.type() === "error") errors.push(m.text()); });
  page.on("dialog", d => d.accept());
  const step = s => console.log("•", s);

  await page.goto(URL);
  step("首次開啟：不需密碼直接進入");
  await page.waitForSelector("#app:not([hidden])");
  assert.equal(await page.isHidden("#lockBtn"), true);
  assert.equal(await page.isVisible("#pwNoneBox"), true); // 第一次開啟停在設定分頁，顯示「未設定密碼」
  assert.equal(await page.getAttribute("#tabBtn-settings", "aria-selected"), "true");

  step("匯入兩期名冊");
  await page.setInputFiles("#rosterFile", path.join(FX, "roster44.xlsx"));
  await page.waitForSelector("#rosterPending:not([hidden])");
  assert.equal(await page.inputValue("#rosterLabel"), "第44期");
  await page.click("#rosterSave");
  await page.waitForFunction(() => document.querySelectorAll(".cohort-card").length === 1);
  await page.setInputFiles("#rosterFile", path.join(FX, "roster45.xlsx"));
  await page.waitForSelector("#rosterPending:not([hidden])");
  assert.equal(await page.inputValue("#rosterLabel"), "第45期");
  await page.click("#rosterSave");
  await page.waitForFunction(() => document.querySelectorAll(".cohort-card").length === 2);
  // 重複匯入同一期 → 預設為更新並不新增
  await page.setInputFiles("#rosterFile", path.join(FX, "roster44.xlsx"));
  await page.waitForSelector("#rosterPending:not([hidden])");
  assert.notEqual(await page.inputValue("#rosterTarget"), "new");
  await page.selectOption("#rosterTarget", "new");
  assert.match(await page.textContent("#rosterPendingWarn"), /可能是同一期/);
  await page.click("#rosterCancel");

  step("匯入兩份範本");
  await page.setInputFiles("#pTplFile", path.join(FX, "points_template.xlsx"));
  await page.waitForFunction(() => /✓/.test(document.getElementById("pTplStatus").textContent));
  await page.setInputFiles("#rTplFile", path.join(FX, "reward_template.xlsx"));
  await page.waitForFunction(() => /✓/.test(document.getElementById("rTplStatus").textContent));
  // 錯誤的檔案放到範本欄位
  await page.setInputFiles("#rTplFile", path.join(FX, "roster44.xlsx"));
  await page.waitForFunction(() => /匯入失敗/.test(document.getElementById("rTplStatus").textContent));
  assert.match(await page.textContent("#rTplStatus"), /仍使用原本的範本/);

  step("重新整理 → 直接進入，不需重新匯入");
  await page.reload();
  await page.waitForSelector("#app:not([hidden])");
  assert.equal(await page.locator(".cohort-btn").count(), 2);
  assert.match(await page.textContent("#pTplStatus"), /✓/);
  assert.match(await page.textContent("#rTplStatus"), /✓/);

  step("加扣分：第44期 002");
  await page.click("#tabBtn-points");
  await page.click(".cohort-btn:has-text('第44期')");
  await page.fill("#pIds", "內務櫃未確實上鎖 002 1\n002");
  const check44 = await page.textContent("#pCheck");
  assert.match(check44, /林小華/); assert.match(check44, /王小美/); assert.match(check44, /第44期/);
  await page.selectOption("#pKind", "扣分");
  await page.click("#pFavorites button:has-text('63.內務不整者')");
  assert.equal(await page.inputValue("#pCategory"), "不良事蹟(一)");
  await page.fill("#pDetail", "測試說明");
  await page.click("#pAdd");
  await page.waitForFunction(() => document.querySelectorAll("#pRecords tr").length === 2);
  const rows44 = await page.$$eval("#pRecords tr", trs => trs.map(tr => tr.innerText));
  assert.match(rows44[0], /第44期[\s\S]*002[\s\S]*林小華[\s\S]*測試說明 （內務櫃未確實上鎖）/);
  assert.match(await page.textContent("#pError"), /重複學生/);

  step("加扣分：切到第45期，同樣輸入 002 得到不同學生");
  await page.click(".cohort-btn:has-text('第45期')");
  await page.fill("#pIds", "002 王小美");
  const check45 = await page.textContent("#pCheck");
  assert.match(check45, /李小婷/);
  assert.match(check45, /王小美：目前期別查無此學生（第44期有符合的學生/);
  await page.selectOption("#pKind", "加分");
  await page.click("#pFavorites button:has-text('25.公差勤務')");
  await page.selectOption("#pScore", "0.5");
  await page.fill("#pDetail", "");
  await page.click("#pAdd");
  await page.waitForFunction(() => document.querySelectorAll("#pRecords tr").length === 3);
  const last = await page.$eval("#pRecords tr:last-child", tr => Array.from(tr.cells).map(td => td.textContent).join(" "));
  assert.match(last, /第45期[\s\S]*002[\s\S]*李小婷[\s\S]*加分[\s\S]*0\.5/);
  assert.match(await page.textContent("#pCount"), /第44期 2 筆、第45期 1 筆/);

  if (!process.env.MOBILE) {
    step("電腦：勾選、操作在最左邊，編輯在刪除左側並排");
    for (const width of [1280, 1024, 901]) {
      await page.setViewportSize({width, height: 900});
      const wrap = await page.locator("#tab-points .tablewrap").boundingBox();
      const row = page.locator("#pRecords > tr").first();
      const edit = await row.locator("button:has-text('編輯')").boundingBox(), del = await row.locator("button:has-text('刪除')").boundingBox();
      const box = await row.locator("input[type=checkbox]").boundingBox();
      assert.ok(box.x - wrap.x < 30 && box.x < edit.x, width + "px：勾選不在最左邊");
      assert.ok(edit.x - wrap.x < 110, width + "px：操作欄不在勾選右側");
      const overflow = await page.$eval("#tab-points .tablewrap", el => el.scrollWidth - el.clientWidth);
      assert.ok(overflow <= 0, width + "px：加扣分表格需要左右捲動（多 " + overflow + "px）");
      assert.ok(edit.x + edit.width <= del.x && Math.abs(edit.y - del.y) < 4, width + "px：編輯與刪除未左右並排");
      assert.ok(del.x + del.width <= wrap.x + wrap.width, width + "px：刪除超出可視範圍");
      assert.equal(await page.textContent("#tab-points thead th:first-child"), "勾選");
      assert.equal(await page.textContent("#tab-points thead th:nth-child(2)"), "操作");
    }
    await page.setViewportSize({width: 1200, height: 900});
  }

  step("加扣分：批次編輯分數");
  await page.check("#pRecords tr:nth-child(1) input[type=checkbox]");
  await page.check("#pRecords tr:nth-child(2) input[type=checkbox]");
  // 勾選多筆後，列上的按鈕變成一次處理全部勾選紀錄
  assert.equal(await page.locator("#pBulkEdit, #pBulkDelete, #pBulkEditBottom").count(), 0);
  if (process.env.MOBILE) await page.click("#pRecords > tr:nth-child(1) .cell-summary");
  const multiEdit = page.locator("#pRecords > tr:nth-child(1) button[data-edit-index]");
  // 按鈕文字固定；筆數顯示在全選旁與底部懸浮提示，按鈕另有滑鼠提示
  assert.equal(await multiEdit.textContent(), "編輯");
  assert.equal(await page.textContent("#pRecords > tr:nth-child(1) button[data-delete-index]"), "刪除");
  assert.match(await multiEdit.getAttribute("title"), /一次編輯勾選的 2 筆/);
  assert.equal(await page.getAttribute("#pRecords > tr:nth-child(3) button[data-edit-index]", "title"), "");
  assert.match(await page.textContent("#tab-points .bulk-hint"), /一次處理這 2 筆/);
  assert.equal(await page.isVisible("#selBar"), true);
  assert.match(await page.textContent("#selBarCount"), /已勾選 2 筆/);
  assert.equal(await page.isVisible("#selBarEdit"), true);
  assert.equal(await page.isVisible("#selBarDelete"), true);
  if (process.env.MOBILE) {
    // 手機：懸浮列不超出畫面、不壓到回頂端按鈕，並用懸浮列的「編輯」開始編輯
    const bar = await page.locator("#selBar").boundingBox();
    assert.ok(bar.x >= 0 && bar.x + bar.width <= 402 - 50, "懸浮列超出或壓到右側按鈕：" + JSON.stringify(bar));
    const order = await page.$$eval("#selBar > *", els => els.map(e => e.id || e.className));
    assert.deepEqual(order, ["selBarEdit", "selBarDelete", "sel-text", "selBarClear"]);
    await page.click("#selBarEdit");
  } else await multiEdit.click();
  // 不再跳出下方的編輯面板，而是在列上直接編輯
  assert.equal(await page.locator("#pEditPanel").count(), 0);
  const editing = page.locator("#pRecords > tr.inline-editing");
  assert.equal(await editing.count(), 1);
  assert.equal(await page.locator("#pRecords > tr.edit-target").count(), 2);
  assert.match(await editing.locator(".inline-hint").textContent(), /勾選的 2 筆/);
  // 切到獎懲分頁時，懸浮提示不顯示加扣分的勾選
  await page.click("#tabBtn-reward"); assert.equal(await page.isVisible("#selBar"), false); await page.click("#tabBtn-points");
  await editing.locator("select.inline-score").selectOption("0.3");
  await editing.locator("button.inline-save").click();
  assert.equal(await page.locator("#pRecords > tr.inline-editing").count(), 0);
  assert.match(await page.textContent("#pSelected"), /已勾選 0 筆/);
  assert.equal(await page.locator("#pRecords input[type=checkbox]:checked").count(), 0);
  // 剛修改的 2 筆維持淺藍底；查看操作不取消，下一個操作才取消
  assert.equal(await page.locator("#pRecords > tr.just-edited").count(), 2);
  await page.fill("#pFilterText", "林"); await page.click("#pFilterClear");
  assert.equal(await page.locator("#pRecords > tr.just-edited").count(), 2);
  await page.check("#pRecords > tr:nth-child(3) input[type=checkbox]");
  assert.equal(await page.locator("#pRecords > tr.just-edited").count(), 0);
  await page.uncheck("#pRecords > tr:nth-child(3) input[type=checkbox]");
  const scores = await page.$$eval("#pRecords tr td[data-label='分數']", tds => tds.map(td => td.textContent));
  assert.deepEqual(scores, ["0.3", "0.3", "0.5"]);

  step("清單工具：分組、搜尋、期別篩選");
  const visibleRows = sel => page.$$eval(sel + " > tr", trs => trs.filter(tr => !tr.hidden).length);
  assert.equal(await page.locator("#pGroups .group-item").count(), 2);
  assert.match(await page.textContent("#pGroups"), /扣分｜63\.內務不整者｜0\.3 分\s*2 筆[\s\S]*林小華、王小美/);
  await page.click("#pGroups .group-item:has-text('63.內務不整者')");
  assert.equal(await visibleRows("#pRecords"), 2);
  assert.match(await page.textContent("#pFilterInfo"), /顯示 2 \/ 3 筆/);
  // 篩選中按全選，只勾看得到的 2 筆
  await page.check("#pSelectAll");
  assert.match(await page.textContent("#pSelected"), /已勾選 2 筆/);
  await page.uncheck("#pSelectAll");
  await page.click("#pGroups .group-item:has-text('63.內務不整者')");
  assert.equal(await visibleRows("#pRecords"), 3);
  await page.fill("#pFilterText", "李小");
  assert.equal(await visibleRows("#pRecords"), 1);
  await page.click("#pFilterClear");
  await page.selectOption("#pFilterCohort", "第45期");
  assert.equal(await visibleRows("#pRecords"), 1);
  await page.click("#pFilterClear");
  assert.equal(await visibleRows("#pRecords"), 3);
  if (process.env.MOBILE) {
    step("手機：精簡列表、點開詳細、切換完整模式");
    assert.equal(await page.getAttribute(".view-seg button[data-view=compact]", "aria-pressed"), "true");
    const firstRow = page.locator("#pRecords > tr").nth(2); // 尚未展開過的一列
    assert.equal(await firstRow.locator("td[data-label='學籍號碼']").isVisible(), false);
    const compactHeight = (await firstRow.boundingBox()).height;
    assert.ok(compactHeight < 90, "精簡列高度 " + compactHeight);
    await firstRow.locator(".cell-summary").click();
    assert.equal(await firstRow.locator("td[data-label='學籍號碼']").isVisible(), true);
    assert.equal(await firstRow.locator("button:has-text('編輯')").isVisible(), true);
    await firstRow.locator(".cell-summary").click();
    await page.click("#tab-points .view-seg button[data-view=full]");
    assert.equal(await page.locator("#pRecords > tr").nth(1).locator("td[data-label='學籍號碼']").isVisible(), true);
    await page.click("#tab-points .view-seg button[data-view=compact]");
  }

  step("加扣分：在列上直接編輯單筆（說明、類型、事由）與取消");
  const row3 = "#pRecords > tr:nth-child(3)";
  const openEdit = async sel => {
    if (process.env.MOBILE && !/expanded/.test(await page.getAttribute(sel, "class") || "")) await page.click(sel + " .cell-summary");
    await page.click(sel + " button[data-edit-index]");
  };
  await openEdit(row3);
  await page.fill(row3 + " textarea.inline-detail", "改成新的說明");
  if (!process.env.MOBILE) {
    const ov = await page.$eval("#tab-points .tablewrap", el => el.scrollWidth - el.clientWidth);
    assert.ok(ov <= 0, "編輯中表格需要左右捲動（多 " + ov + "px）");
  }
  await page.click(row3 + " button.inline-save");
  assert.equal(await page.textContent(row3 + " td[data-label='說明']"), "改成新的說明");
  assert.equal(await page.textContent(row3 + " td[data-label='類型']"), "加分");
  assert.notEqual(await page.textContent("#pRecords > tr:nth-child(1) td[data-label='說明']"), "改成新的說明");
  await openEdit(row3);
  await page.selectOption(row3 + " select.inline-kind", "扣分");
  assert.equal(await page.inputValue(row3 + " select.inline-reason"), "");
  await page.click(row3 + " button.inline-save");
  assert.match(await page.textContent(row3 + " .inline-error"), /請選擇分類及事由/);
  await page.selectOption(row3 + " select.inline-reason", {label: "63.內務不整者"});
  await page.click(row3 + " button.inline-save");
  assert.equal(await page.textContent(row3 + " td[data-label='類型']"), "扣分");
  assert.equal(await page.textContent(row3 + " td[data-label='事由']"), "63.內務不整者");
  // Esc 取消，不會改動
  const row1 = "#pRecords > tr:nth-child(1)";
  await openEdit(row1);
  await page.selectOption(row1 + " select.inline-score", "0.8");
  await page.press(row1 + " select.inline-score", "Escape");
  assert.equal(await page.locator("#pRecords > tr.inline-editing").count(), 0);
  assert.equal(await page.textContent(row1 + " td[data-label='分數']"), "0.3");

  step("加扣分匯出");
  const [dl1] = await Promise.all([page.waitForEvent("download"), page.click("#pExport")]);
  const pFile = path.join(OUT, "points.xlsx"); await dl1.saveAs(pFile);
  // 無頭 Chromium 對中文檔名一律回報 "download"，實際瀏覽器會保留中文檔名

  step("獎懲：第45期 002 + 第44期 002（同號不同人）");
  await page.click("#tabBtn-reward");
  await page.fill("#rIds", "文書公差：002 002");
  await page.selectOption("#rArticle", "第20條");
  assert.match(await page.textContent("#rArticle"), /第19條｜予以申誡[\s\S]*第20條｜予以記過/);
  assert.match(await page.textContent("#rClause"), /第17款｜（分 2 目/);
  await page.selectOption("#rClause", "第17款");
  assert.match(await page.textContent("#rItem"), /第1目｜測試事由乙。[\s\S]*第2目｜測試事由丙。/);
  await page.selectOption("#rItem", "第2目");
  assert.equal(await page.inputValue("#rReason"), "測試事由丙。");
  assert.equal(await page.inputValue("#rAmount"), "小過一次");
  await page.click("#rAdd");
  await page.waitForFunction(() => document.querySelectorAll("#rBody tr").length === 2);
  await page.click(".cohort-btn:has-text('第44期')");
  await page.fill("#rIds", "002");
  await page.selectOption("#rArticle", "第19條");
  assert.match(await page.textContent("#rClause"), /第1款｜測試事由甲。/);
  await page.selectOption("#rClause", "第1款");
  await page.selectOption("#rAdjustReason", "行為之動機"); await page.selectOption("#rAdjustAmount", "申誡一次");
  await page.click("#rAdd");
  await page.waitForFunction(() => document.querySelectorAll("#rBody tr").length === 3);
  const rRows = await page.$$eval("#rBody tr", trs => trs.map(tr => tr.innerText));
  assert.match(rRows[0], /第45期[\s\S]*李小婷[\s\S]*測試事由丙。（文書公差）/);
  assert.match(rRows[2], /第44期[\s\S]*林小華[\s\S]*申誡一次/);
  assert.match(await page.textContent("#rDup"), /李小婷（第45期/);
  step("獎懲：編輯視窗");
  if (!process.env.MOBILE) {
    const wrapR = await page.locator("#tab-reward .tablewrap").boundingBox();
    const rowR = page.locator("#rBody > tr").first();
    const e = await rowR.locator("button:has-text('編輯')").boundingBox(), d = await rowR.locator("button:has-text('刪除')").boundingBox();
    const cb = await rowR.locator("input[type=checkbox]").boundingBox();
    assert.ok(cb.x < e.x && e.x - wrapR.x < 110 && e.x + e.width <= d.x && Math.abs(e.y - d.y) < 4, "獎懲：操作欄位置或排列不正確");
    assert.equal(await page.textContent("#tab-reward thead th:first-child"), "勾選");
    assert.equal(await page.textContent("#tab-reward thead th:nth-child(2)"), "操作");
  }
  if (process.env.MOBILE) await page.click("#rBody tr:nth-child(2) .cell-summary");
  await page.check("#rBody tr:nth-child(2) input[type=checkbox]");
  assert.match(await page.textContent("#rSelected"), /已勾選 1 \//);
  await page.click("#rBody tr:nth-child(2) button:has-text('編輯')");
  assert.match(await page.textContent("#rEditClause"), /第17款｜（分 2 目/);
  assert.match(await page.textContent("#rEditItem"), /第2目｜測試事由丙。/);
  await page.fill("#rEditReason", "改過的事由");
  await page.click("#rEditForm button[type=submit]");
  await page.waitForFunction(() => !document.getElementById("rEditor").open);
  assert.match(await page.textContent("#rSelected"), /已勾選 0 \//);
  assert.equal(await page.locator("#rBody > tr.just-edited").count(), 1);
  assert.equal(await page.getAttribute("#rBody > tr:nth-child(2)", "class").then(c => /just-edited/.test(c)), true);
  assert.match(await page.$eval("#rBody tr:nth-child(2)", tr => tr.innerText), /改過的事由/);
  const [dl2] = await Promise.all([page.waitForEvent("download"), page.click("#rExport")]);
  step("獎懲：勾選多筆一次刪除");
  rDupBefore = await page.locator("#rBody > tr").count();
  // 新增兩筆暫時紀錄再一次刪除，確認其餘紀錄不受影響
  await page.fill("#rIds", "001 003");
  await page.selectOption("#rArticle", "第19條"); await page.selectOption("#rClause", "第1款");
  await page.click("#rAdd");
  await page.waitForFunction(n => document.querySelectorAll("#rBody > tr").length === n + 2, rDupBefore);
  assert.equal(await page.locator("#rDeleteSelected, #rEditSelected, #rDeleteSelectedBottom").count(), 0);
  await page.check("#rBody > tr:nth-child(" + (rDupBefore + 1) + ") input[type=checkbox]");
  await page.check("#rBody > tr:nth-child(" + (rDupBefore + 2) + ") input[type=checkbox]");
  assert.match(await page.textContent("#rSelected"), /已勾選 2 \//);
  const lastRow = "#rBody > tr:nth-child(" + (rDupBefore + 2) + ")";
  if (process.env.MOBILE) await page.click(lastRow + " .cell-summary");
  assert.equal(await page.textContent(lastRow + " button[data-edit-index]"), "編輯");
  assert.match(await page.getAttribute(lastRow + " button[data-delete-index]", "title"), /一次刪除勾選的 2 筆/);
  assert.match(await page.textContent("#selBarCount"), /已勾選 2 筆/);
  if (!process.env.MOBILE) {
    const overflowR = await page.$eval("#tab-reward .tablewrap", el => el.scrollWidth - el.clientWidth);
    assert.ok(overflowR <= 0, "獎懲表格需要左右捲動（多 " + overflowR + "px）");
  }
  // 未勾選的列仍只處理自己
  assert.equal(await page.textContent("#rBody > tr:nth-child(1) button[data-delete-index]"), "刪除");
  if (process.env.MOBILE) await page.click("#selBarDelete"); // 手機用懸浮列刪除
  else await page.click(lastRow + " button[data-delete-index]");
  await page.waitForFunction(n => document.querySelectorAll("#rBody > tr").length === n, rDupBefore);
  assert.match(await page.textContent("#rResult"), /已刪除 2 筆/);
  assert.equal(await page.isVisible("#selBar"), false);
  assert.match(await page.$eval("#rBody > tr:nth-child(2)", tr => tr.textContent), /改過的事由/);
  assert.match(await page.textContent("#rSelected"), /已勾選 0 \//);
  const rFile = path.join(OUT, "reward.xlsx"); await dl2.saveAs(rFile);

  step("設定開啟密碼");
  await page.click("#tabBtn-settings");
  assert.equal(await page.isVisible("#pwNoneBox"), true);
  await page.click("details:has(#pwSetForm) summary");
  await page.fill("#pwSet1", "abc123"); await page.fill("#pwSet2", "abc124");
  await page.click("#pwSetForm button[type=submit]");
  assert.match(await page.textContent("#pwSetStatus"), /不一致/);
  await page.fill("#pwSet2", "abc123"); await page.click("#pwSetForm button[type=submit]");
  await page.waitForSelector("#lockBtn:not([hidden])");
  assert.equal(await page.isVisible("#pwSetBox"), true);

  step("鎖定 → 密碼錯誤 → 正確解鎖，紀錄仍在");
  await page.click("#lockBtn");
  await page.waitForSelector("#unlockForm:not([hidden])");
  assert.equal(await page.isHidden("#app"), true);
  assert.equal(await page.evaluate(() => document.body.innerText.includes("李小婷")), false);
  await page.fill("#unlockPw", "wrong-pw"); await page.click("#unlockBtn");
  await page.waitForFunction(() => /不正確/.test(document.getElementById("unlockStatus").textContent));
  assert.equal(await page.isHidden("#app"), true);
  await page.fill("#unlockPw", "abc123"); await page.click("#unlockBtn");
  await page.waitForSelector("#app:not([hidden])");
  assert.equal(await page.locator("#pRecords tr").count(), 3);
  assert.equal(await page.locator("#rBody tr").count(), 3);
  assert.equal(await page.textContent("#pTabCount"), "3");

  step("本機資料庫中看不到明文");
  const leaked = await page.evaluate(() => new Promise(resolve => {
    const req = indexedDB.open("student-records-local-v1");
    req.onsuccess = () => {
      const t = req.result.transaction("kv").objectStore("kv").getAll();
      t.onsuccess = () => {
        const dec = new TextDecoder();
        resolve(t.result.some(v => v.ct && /小婷|小美|1449900002/.test(dec.decode(v.ct))));
      };
    };
  }));
  assert.equal(leaked, false);

  step("變更密碼");
  await page.click("#tabBtn-settings");
  await page.click("details:has(#pwForm) summary");
  await page.fill("#pwOld", "abc123"); await page.fill("#pwNew", "newpass9"); await page.fill("#pwNew2", "newpass9");
  await page.click("#pwForm button[type=submit]");
  await page.waitForFunction(() => /已變更/.test(document.getElementById("pwStatus").textContent));
  await page.reload();
  await page.waitForSelector("#unlockForm:not([hidden])");
  await page.fill("#unlockPw", "abc123"); await page.click("#unlockBtn");
  await page.waitForFunction(() => /不正確/.test(document.getElementById("unlockStatus").textContent));
  await page.fill("#unlockPw", "newpass9"); await page.click("#unlockBtn");
  await page.waitForSelector("#app:not([hidden])");
  assert.equal(await page.locator("#rBody tr").count(), 3);
  assert.equal(await page.locator(".cohort-btn").count(), 2);
  await page.screenshot({path: path.join(OUT, "settings.png"), fullPage: true});
  await page.click("#tabBtn-points");
  await page.screenshot({path: path.join(OUT, "points.png"), fullPage: true});
  await page.click("#tabBtn-reward");
  await page.screenshot({path: path.join(OUT, "reward.png"), fullPage: true});

  step("取消密碼 → 重新開啟直接進入");
  await page.click("#tabBtn-settings");
  await page.click("details:has(#pwRemoveForm) summary");
  await page.fill("#pwRemoveOld", "wrong-pw"); await page.click("#pwRemoveForm button[type=submit]");
  await page.waitForFunction(() => /不正確/.test(document.getElementById("pwRemoveStatus").textContent));
  await page.fill("#pwRemoveOld", "newpass9"); await page.click("#pwRemoveForm button[type=submit]");
  await page.waitForSelector("#lockBtn[hidden]", {state: "attached"});
  await page.reload();
  await page.waitForSelector("#app:not([hidden])");
  assert.equal(await page.locator("#rBody tr").count(), 3);

  step("清除本機全部資料");
  await page.click("#tabBtn-settings");
  await page.click("#wipeAll");
  await page.waitForFunction(() => !document.getElementById("app").hidden && document.querySelectorAll(".cohort-card").length === 0);

  assert.deepEqual(external, [], "不應有外部連線：" + external.join(", "));
  assert.deepEqual(errors, [], "Console 錯誤：" + errors.join("\n"));
  await browser.close();
  console.log("ALL PASSED");
})().catch(error => { console.error("FAILED:", error); process.exit(1); });
