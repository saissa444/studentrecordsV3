// 端對端測試（全部虛構資料）
const {chromium, devices} = require("playwright");
const assert = require("node:assert/strict");
const path = require("node:path");
const fs = require("node:fs");
const FX = path.join(__dirname, "fixtures");
const OUT = path.join(__dirname, "output"); fs.mkdirSync(OUT, {recursive: true});
const URL = process.env.APP_URL || "http://localhost:8765/index.html";

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
  step("首次設定密碼");
  await page.waitForSelector("#setupForm:not([hidden])");
  await page.fill("#setupPw", "abc123"); await page.fill("#setupPw2", "abc124");
  await page.click("#setupBtn");
  assert.match(await page.textContent("#setupStatus"), /不一致/);
  await page.fill("#setupPw2", "abc123"); await page.click("#setupBtn");
  await page.waitForSelector("#app:not([hidden])");
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

  step("重新整理 → 密碼錯誤 → 正確解鎖，不需重新匯入");
  await page.reload();
  await page.waitForSelector("#unlockForm:not([hidden])");
  await page.fill("#unlockPw", "wrong-pw"); await page.click("#unlockBtn");
  await page.waitForFunction(() => /不正確/.test(document.getElementById("unlockStatus").textContent));
  assert.equal(await page.isHidden("#app"), true);
  await page.fill("#unlockPw", "abc123"); await page.click("#unlockBtn");
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
  const last = await page.$eval("#pRecords tr:last-child", tr => tr.innerText);
  assert.match(last, /第45期[\s\S]*002[\s\S]*李小婷[\s\S]*加分[\s\S]*0\.5/);
  assert.match(await page.textContent("#pCount"), /第44期 2 筆、第45期 1 筆/);

  step("加扣分：批次編輯分數");
  await page.check("#pRecords tr:nth-child(1) input[type=checkbox]");
  await page.check("#pRecords tr:nth-child(2) input[type=checkbox]");
  await page.click("#pBulkEdit");
  await page.check("#pApplyScore"); await page.uncheck("#pApplyDetail"); await page.uncheck("#pApplyKind");
  await page.selectOption("#pEditScore", "0.3");
  await page.click("#pSaveEdit");
  const scores = await page.$$eval("#pRecords tr td:nth-child(9)", tds => tds.map(td => td.textContent));
  assert.deepEqual(scores, ["0.3", "0.3", "0.5"]);

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
  await page.click("#rBody tr:nth-child(2) button:has-text('編輯')");
  assert.match(await page.textContent("#rEditClause"), /第17款｜（分 2 目/);
  assert.match(await page.textContent("#rEditItem"), /第2目｜測試事由丙。/);
  await page.fill("#rEditReason", "改過的事由");
  await page.click("#rEditForm button[type=submit]");
  await page.waitForFunction(() => !document.getElementById("rEditor").open);
  assert.match(await page.$eval("#rBody tr:nth-child(2)", tr => tr.innerText), /改過的事由/);
  const [dl2] = await Promise.all([page.waitForEvent("download"), page.click("#rExport")]);
  const rFile = path.join(OUT, "reward.xlsx"); await dl2.saveAs(rFile);

  step("鎖定 → 解鎖後紀錄仍在");
  await page.click("#lockBtn");
  await page.waitForSelector("#unlockForm:not([hidden])");
  assert.equal(await page.isHidden("#app"), true);
  assert.equal(await page.evaluate(() => document.body.innerText.includes("李小婷")), false);
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

  step("清除本機全部資料");
  await page.click("#tabBtn-settings");
  await page.click("#wipeAll");
  await page.waitForSelector("#setupForm:not([hidden])");

  assert.deepEqual(external, [], "不應有外部連線：" + external.join(", "));
  assert.deepEqual(errors, [], "Console 錯誤：" + errors.join("\n"));
  await browser.close();
  console.log("ALL PASSED");
})().catch(error => { console.error("FAILED:", error); process.exit(1); });
