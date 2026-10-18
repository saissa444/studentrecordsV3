/* 共用解析與驗證邏輯。所有資料只在本機瀏覽器處理，不會傳送到任何伺服器。 */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory();
  else root.AppCore = factory();
})(typeof globalThis !== "undefined" ? globalThis : this, function () {
  "use strict";

  const clean = value => String(value == null ? "" : value).trim();
  const normalize = value => clean(value).normalize("NFKC").replace(/\s+/g, "").toLowerCase();
  const identifier = value => normalize(value).replace(/^'(?=\d)/, "").replace(/\.0+$/, "");
  const unique = values => Array.from(new Set(values));
  const rows = (xlsx, sheet) => xlsx.utils.sheet_to_json(sheet, {header: 1, defval: "", raw: false});

  /* ───────────── 名冊 ───────────── */

  function column(headers, names) {
    const exact = headers.findIndex(header => names.some(name => normalize(header) === normalize(name)));
    return exact >= 0 ? exact : headers.findIndex(header => names.some(name => normalize(header).includes(normalize(name))));
  }

  function extractRoster(workbook, xlsx) {
    const preferred = workbook.SheetNames.filter(name => normalize(name) === "姓名查詢");
    const names = preferred.concat(workbook.SheetNames.filter(name => !preferred.includes(name)));
    let best = null;
    for (const sheetName of names) {
      const data = rows(xlsx, workbook.Sheets[sheetName]);
      for (let index = 0; index < Math.min(data.length, 30); index++) {
        const headers = data[index];
        const cols = {
          name: column(headers, ["姓名", "學生姓名"]),
          school: column(headers, ["學籍號碼", "完整學籍號碼", "學籍號", "學號"]),
          female: column(headers, ["女生隊號", "女生隊編號"]),
          team: column(headers, ["建制隊號", "建制隊編號"])
        };
        if (cols.name < 0 || cols.school < 0) continue;
        const students = new Map();
        for (const row of data.slice(index + 1)) {
          const student = {
            name: clean(row[cols.name]), school: identifier(row[cols.school]),
            female: cols.female < 0 ? "" : identifier(row[cols.female]),
            team: cols.team < 0 ? "" : identifier(row[cols.team])
          };
          if (!student.name || !/^\d{9,12}$/.test(student.school)) continue;
          if (/^\d{1,3}$/.test(student.female)) student.female = student.female.padStart(3, "0");
          const previous = students.get(student.school);
          if (previous && ["name", "female", "team"].some(key => normalize(previous[key]) !== normalize(student[key]))) {
            throw new Error("同一學籍號碼有不同名冊資料，請先確認 Excel。");
          }
          students.set(student.school, student);
        }
        const result = {students: Array.from(students.values()), sheetName};
        if (result.students.length && preferred.includes(sheetName)) return result;
        if (!best || result.students.length > best.students.length) best = result;
        break; // 只用第一個完整標題區塊，不混入旁邊的查詢表。
      }
    }
    if (!best || !best.students.length) {
      throw new Error("無法辨識名冊：請確認「姓名查詢」包含姓名及學籍號碼。");
    }
    const females = best.students.map(s => s.female).filter(Boolean);
    const dupFemale = females.find((value, index) => females.indexOf(value) !== index);
    if (dupFemale) throw new Error("同一份名冊中女生隊號重複：" + dupFemale + "。若檔案內含兩期學生，請分成兩個檔案各自匯入。");
    return best;
  }

  /** 由建制隊號前兩碼猜期別，例如 440012 → 第44期。猜不出來回傳空字串。 */
  function guessCohortLabel(students) {
    const prefixes = unique(students.map(s => s.team).filter(t => /^\d{6}$/.test(t)).map(t => t.slice(0, 2)));
    return prefixes.length === 1 ? "第" + Number(prefixes[0]) + "期" : "";
  }

  /* ───────────── 學生輸入解析 ───────────── */

  function parseTokens(value) {
    const tokens = [];
    for (const token of clean(value).normalize("NFKC").split(/[\s,、;]+/).filter(Boolean)) {
      const range = token.match(/^(\d{1,3})[-–~～](\d{1,3})$/);
      if (!range) { tokens.push(token); continue; }
      const first = Number(range[1]), last = Number(range[2]);
      if (first < 1 || last < first || last > 999) throw new Error("女生隊號範圍不正確：" + token);
      for (let n = first; n <= last; n++) tokens.push(String(n).padStart(3, "0"));
    }
    return tokens;
  }

  function matchStudent(token, roster) {
    const value = identifier(token);
    const candidates = roster.filter(student =>
      normalize(student.name) === normalize(token) || student.school === value ||
      (student.team && identifier(student.team) === value) ||
      (/^\d{1,3}$/.test(value) && /^\d{1,3}$/.test(student.female) &&
        student.female.padStart(3, "0") === value.padStart(3, "0"))
    );
    const matches = Array.from(new Map(candidates.map(student => [student.school, student])).values());
    if (!matches.length) return {error: "查無此學生"};
    if (matches.length > 1) return {error: "符合多位學生，請改用女生隊號或建制隊號"};
    return {student: matches[0]};
  }

  function parseStudentInputs(value, roster) {
    const inputs = [], names = new Set(roster.map(student => normalize(student.name)));
    const add = (token, note) => parseTokens(token).forEach(token => inputs.push({token, note}));
    for (const line of clean(value).split(/[\r\n]+/)) {
      let note = "";
      for (const group of line.split(/[,，、;；]+/)) {
        let pending = [];
        const emit = (token, prefix) => {
          if (pending.length || prefix !== undefined) {
            note = clean(pending.concat(prefix === undefined ? [] : [prefix]).join(" ")).replace(/[:：]+\s*$/, "").trim();
          }
          pending = [];
          add(token, note);
        };
        for (const word of group.split(/\s+/).filter(Boolean)) {
          const normalized = word.normalize("NFKC");
          // 完整姓名優先；不從姓名後綴猜學生。
          if (names.has(normalize(word)) || /^'?\d+(?:\.0+)?$|^\d+[-–~～]\d+$/.test(normalized)) {
            emit(word);
            continue;
          }
          const colon = Math.max(word.lastIndexOf(":"), word.lastIndexOf("："));
          if (colon >= 0 && colon < word.length - 1) {
            emit(word.slice(colon + 1), word.slice(0, colon));
            continue;
          }
          const attached = word.match(/^(.*?[^0-9０-９.．\-–~～])([0-9０-９]+(?:[-–~～][0-9０-９]+|[.．][0０]+)?)$/);
          if (attached && /[^\d\s:：'.\-–~～]/.test(attached[1].normalize("NFKC"))) {
            emit(attached[2], attached[1]);
            continue;
          }
          pending.push(word);
        }
        // 只有註記、後面沒有學生，也要回報，不能默默丟掉。
        pending.forEach(token => add(token, ""));
      }
    }
    return inputs;
  }

  /**
   * others：其他期別 [{label, students}]，只用來在查無時提示「這個人在別期」。
   * 絕不會自動改用別期的學生。
   */
  function resolveStudents(value, roster, others) {
    const inputs = parseStudentInputs(value, roster);
    const tokens = inputs.map(input => input.token), matches = new Map(), errors = [], students = [], entries = [];
    for (const {token, note} of inputs) {
      const match = matchStudent(token, roster);
      if (match.error) {
        let message = match.error;
        if (match.error === "查無此學生" && others && others.length) {
          const elsewhere = others.filter(cohort => !matchStudent(token, cohort.students).error).map(cohort => cohort.label);
          if (elsewhere.length) message = "目前期別查無此學生（" + elsewhere.join("、") + "有符合的學生，請確認是否要切換期別）";
        }
        errors.push({token, message});
        continue;
      }
      students.push(match.student);
      entries.push({student: match.student, note});
      const key = match.student.school;
      const previous = matches.get(key);
      if (previous) {
        previous.count++;
        if (note && !previous.notes.includes(note)) previous.notes.push(note);
      } else matches.set(key, {student: match.student, count: 1, notes: note ? [note] : []});
    }
    return {tokens, matches: Array.from(matches.values()), errors, students, entries};
  }

  function duplicateStudents(records) {
    const counts = new Map();
    records.forEach(record => {
      const previous = counts.get(record.school);
      if (previous) previous.count++;
      else counts.set(record.school, {school: record.school, name: record.name, cohort: record.cohort, count: 1});
    });
    return Array.from(counts.values()).filter(student => student.count > 1);
  }

  /* ───────────── 獎懲 ───────────── */

  const rewardFields = ["school", "name", "article", "clause", "item", "reason", "amount", "adjustReason", "adjustAmount"];

  function headerField(value) {
    const header = normalize(value);
    const mapping = {
      "學號": "school", "學籍號碼": "school", "姓名": "name",
      "獎懲依據(條)": "article", "獎懲依據(款)": "clause", "獎懲依據(目)": "item",
      "獎懲事由": "reason", "額度": "amount", "加重減輕額度": "adjustAmount"
    };
    return header.startsWith("加重減輕事由") ? "adjustReason" : mapping[header];
  }

  function extractRewardTemplate(workbook, xlsx) {
    const ruleSheetName = workbook.SheetNames.find(name => normalize(name) === "獎懲項目");
    if (!ruleSheetName) throw new Error("找不到「獎懲項目」工作表，這可能不是獎懲範本。");
    const rules = [], amounts = [], adjustments = [], defaults = {};
    for (const row of rows(xlsx, workbook.Sheets[ruleSheetName])) {
      const [article, clause, item, , reason, , key, amount, , adjustment] = Array.from({length: 10}, (_, i) => clean(row[i]));
      if (/^第\d+條$/.test(key) && amount) defaults[key] = amount;
      if (/^\d+$/.test(key) && amount && amount !== "以下空白") amounts.push(amount);
      if (adjustment && adjustment !== "加重減輕事由") adjustments.push(adjustment);
      if (/^第\d+條$/.test(article) && /^第\d+款$/.test(clause) && reason) {
        rules.push({article, clause, item, reason});
      }
    }
    if (!rules.length) throw new Error("「獎懲項目」未找到有效的條款與事由。");
    let layout = null;
    for (const sheetName of workbook.SheetNames.filter(name => name !== ruleSheetName)) {
      for (const row of rows(xlsx, workbook.Sheets[sheetName]).slice(0, 30)) {
        const mapped = row.map(headerField);
        if (rewardFields.every(field => mapped.includes(field))) {
          const last = mapped.reduce((end, field, index) => field ? index : end, 0);
          layout = {sheetName, headers: row.slice(0, last + 1).map(clean), fields: mapped.slice(0, last + 1)};
          break;
        }
      }
      if (layout) break;
    }
    if (!layout) throw new Error("找不到範本的九個獎懲匯入欄位，請確認所選檔案。");
    return Object.assign(layout, {
      rules, defaultAmounts: defaults,
      amountOptions: unique(amounts.concat(Object.values(defaults))),
      adjustmentOptions: unique(adjustments)
    });
  }

  function rewardDetails(template, selection) {
    if (!template) throw new Error("請先匯入獎懲範本。");
    if (!selection.article || !selection.clause) throw new Error("請依序選擇獎懲依據（條）、（款）。");
    const relevant = template.rules.filter(rule => rule.article === selection.article && rule.clause === selection.clause);
    if (relevant.some(rule => rule.item) && !selection.item) throw new Error("請選擇獎懲依據（目）。");
    const rule = relevant.find(row => row.item === selection.item);
    if (!rule) throw new Error("尚未帶入有效的獎懲事由。");
    const reason = Object.prototype.hasOwnProperty.call(selection, "reason") ? clean(selection.reason) : rule.reason;
    if (!reason) throw new Error("請填寫獎懲事由。");
    const amount = template.defaultAmounts[rule.article];
    if (!amount || !template.amountOptions.includes(amount)) throw new Error("範本未設定有效的固定額度，請確認「獎懲項目」工作表。");
    if (selection.adjustReason && !template.adjustmentOptions.includes(selection.adjustReason)) throw new Error("請選擇有效的加重減輕事由。");
    if (selection.adjustAmount && !selection.adjustReason) throw new Error("選擇加重減輕額度前，請先選擇加重減輕事由。");
    if (selection.adjustAmount && !template.amountOptions.includes(selection.adjustAmount)) throw new Error("請選擇有效的加重減輕額度。");
    return {
      article: rule.article, clause: rule.clause,
      item: rule.item, reason, amount,
      adjustReason: selection.adjustReason || "", adjustAmount: selection.adjustAmount || ""
    };
  }

  function buildRewardRecords(value, cohort, others, template, selection) {
    if (!cohort || !cohort.students.length) throw new Error("請先匯入並選擇期別名冊。");
    if (!template) throw new Error("請先匯入獎懲範本。");
    const resolved = resolveStudents(value, cohort.students, others);
    if (!resolved.tokens.length) throw new Error("請輸入學生。");
    if (!resolved.matches.length) throw new Error("沒有可新增的學生，請確認上方紅字提示後重新輸入。");
    const details = rewardDetails(template, selection);
    const records = resolved.entries.map(({student, note}) => Object.assign(
      {cohort: cohort.label, school: student.school, name: student.name, female: student.female, team: student.team},
      details, {reason: details.reason + (note ? "（" + note + "）" : "")}));
    return {records, resolved};
  }

  function editRewardRecords(records, indexes, template, patch) {
    const editable = ["article", "clause", "item", "reason", "adjustReason", "adjustAmount"];
    if (!Object.keys(patch).length) throw new Error("請勾選至少一組要修改的欄位。");
    if (Object.prototype.hasOwnProperty.call(patch, "amount")) throw new Error("額度依獎懲依據固定，無法單獨修改。");
    if (Object.keys(patch).some(key => !editable.includes(key))) throw new Error("學生資料依名冊顯示，請修改獎懲欄位。");
    if (!indexes.length) throw new Error("請先選取要編輯的紀錄。");
    const selected = new Set(indexes);
    if (Array.from(selected).some(index => !Number.isInteger(index) || index < 0 || index >= records.length)) throw new Error("選取的紀錄已變更，請重新選取。");
    // 全部驗證通過才回傳；任何一筆失敗就整批不變。
    return records.map((record, index) => {
      if (!selected.has(index)) return record;
      const updated = Object.assign({}, record, patch);
      const changedRule = ["article", "clause", "item"].some(key => Object.prototype.hasOwnProperty.call(patch, key) && patch[key] !== record[key]);
      if (changedRule && !Object.prototype.hasOwnProperty.call(patch, "reason")) delete updated.reason;
      return Object.assign({}, record, rewardDetails(template, updated));
    });
  }

  function exportRewardWorkbook(records, template, xlsx) {
    if (!records.length || !template) throw new Error("請先新增紀錄。");
    const data = [template.headers, ...records.map(record => template.fields.map(field => record[field] || ""))];
    const sheet = xlsx.utils.aoa_to_sheet(data);
    sheet["!cols"] = template.fields.map(field => ({wch: field === "reason" ? 55 : field === "adjustReason" ? 30 : 18}));
    const workbook = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(workbook, sheet, template.sheetName);
    return workbook;
  }

  /* ───────────── 加扣分 ───────────── */

  const POINTS_ITEM_SHEET = "加扣分項目", POINTS_OUTPUT_SHEET = "工作表1";

  function cellValue(sheet, r, c, xlsx) {
    const cell = sheet[xlsx.utils.encode_cell({r, c})];
    return cell ? cell.v : undefined;
  }

  /** 第 14 列 A–H 為分類名稱，第 15 列起為事蹟說明（與原加扣分 App 相同）。 */
  function extractPointsTemplate(workbook, xlsx) {
    const sheet = workbook.Sheets[POINTS_ITEM_SHEET];
    if (!sheet) throw new Error("找不到「加扣分項目」工作表，這可能不是加扣分範本。");
    if (!workbook.Sheets[POINTS_OUTPUT_SHEET]) throw new Error("找不到「工作表1」，無法匯出學校格式。");
    const range = xlsx.utils.decode_range(sheet["!ref"] || "A1");
    const lastRow = Math.min(299, range.e.r);
    const options = {};
    for (let c = 0; c < 8; c++) {
      const category = clean(cellValue(sheet, 13, c, xlsx));
      if (!category) continue;
      options[category] = [];
      for (let r = 14; r <= lastRow; r++) {
        const v = cellValue(sheet, r, c, xlsx);
        if (typeof v === "string" && v.trim()) options[category].push(v.trim());
      }
    }
    if (!Object.keys(options).length) throw new Error("範本中沒有分類選項（應在「加扣分項目」第 14 列）。");
    return {options};
  }

  const kindPrefix = kind => kind === "加分" ? "優良" : "不良";

  function validScore(p) {
    return Number.isFinite(p) && p >= 0.1 - 1e-9 && p <= 0.8 + 1e-9 && Math.abs(p * 10 - Math.round(p * 10)) < 1e-7;
  }

  /** 加扣分：同一批次同一學生只新增一次（與原 App 相同），註記合併附加在說明後。 */
  function buildPointsRecords(value, cohort, others, template, selection) {
    const errors = [];
    if (!cohort || !cohort.students.length) errors.push("請先匯入並選擇期別名冊");
    if (!template) errors.push("請先匯入加扣分範本");
    if (errors.length) throw new Error(errors.join("\n"));
    const resolved = resolveStudents(value, cohort.students, others);
    const p = Number(selection.score), {kind, cat, reason} = selection;
    if (!resolved.tokens.length) errors.push("請輸入至少一筆學生資料");
    else if (!resolved.matches.length) errors.push("沒有可新增的學生，請檢查輸入的隊號或姓名");
    if (!validScore(p)) errors.push("分數須為 0.1 至 0.8，且以 0.1 為級距");
    if (!reason || !cat) errors.push("請選擇分類及事蹟說明");
    else if (!cat.startsWith(kindPrefix(kind)) || !(template.options[cat] || []).includes(reason)) errors.push("請選擇符合加扣分類型的有效事蹟說明");
    if (errors.length) throw Object.assign(new Error(errors.join("\n")), {resolved});
    const base = clean(selection.detail);
    const records = resolved.matches.map(({student, notes}) => {
      const suffix = notes.map(n => "（" + n + "）").join("");
      return {
        cohort: cohort.label, school: student.school, name: student.name, female: student.female, team: student.team,
        kind, score: Math.round(p * 10) / 10, cat, reason, detail: base + (base && suffix ? " " : "") + suffix
      };
    });
    const warnings = resolved.errors.map(e => e.token + "：" + e.message)
      .concat(resolved.matches.filter(m => m.count > 1).map(m =>
        "重複學生：" + (m.student.female || m.student.team) + " " + m.student.name + "（共輸入 " + m.count + " 次），已自動略過重複輸入"));
    return {records, resolved, warnings};
  }

  function validatePointsPatch(template, patch) {
    const errors = [];
    if (!Object.keys(patch).length) errors.push("請至少勾選一項「一起編輯」。");
    if ("score" in patch && !validScore(patch.score)) errors.push("分數須為 0.1 至 0.8，且以 0.1 為級距");
    if ("kind" in patch) {
      if (!patch.cat || !patch.reason) errors.push("請選擇分類及事由");
      else {
        if (!(template.options[patch.cat] || []).includes(patch.reason)) errors.push("請選擇範本中的有效事蹟說明");
        if (!patch.cat.startsWith(kindPrefix(patch.kind))) errors.push("分類與加扣分類型不一致");
      }
    }
    return errors;
  }

  /* 加扣分匯出：直接改寫範本 zip 內「工作表1」的 XML，其餘工作表、格式、下拉選單全部保留。 */

  const NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
  const NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const NS_XML = "http://www.w3.org/XML/1998/namespace";

  function zipEntry(zip, path) {
    const index = zip.FullPaths.findIndex(p => p.replace(/^[^/]*\//, "") === path);
    return index >= 0 ? {index, entry: zip.FileIndex[index]} : null;
  }
  function zipText(zip, path) {
    const found = zipEntry(zip, path);
    if (!found) throw new Error("範本缺少 " + path);
    return new TextDecoder("utf-8").decode(found.entry.content);
  }
  function zipPut(zip, path, text) {
    const found = zipEntry(zip, path);
    const bytes = new TextEncoder().encode(text);
    found.entry.content = bytes;
    found.entry.size = bytes.length;
  }
  function zipDelete(zip, path) {
    const found = zipEntry(zip, path);
    if (!found) return;
    zip.FullPaths.splice(found.index, 1);
    zip.FileIndex.splice(found.index, 1);
  }

  const colLetter = ref => ref.replace(/\d+$/, "");
  const colNumber = letters => letters.split("").reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);

  function exportPointsWorkbook(templateBytes, records, xlsx, dom) {
    if (!records.length) throw new Error("尚無紀錄可匯出。");
    const parser = new dom.DOMParser(), serializer = new dom.XMLSerializer();
    const parse = text => {
      const doc = parser.parseFromString(text, "application/xml");
      if (doc.getElementsByTagName("parsererror").length) throw new Error("範本 XML 無法解析");
      return doc;
    };
    const zip = xlsx.CFB.read(new Uint8Array(templateBytes), {type: "array"});

    // 找出「工作表1」對應的 XML 檔
    const wbDoc = parse(zipText(zip, "xl/workbook.xml"));
    const sheetEl = Array.from(wbDoc.getElementsByTagNameNS(NS_MAIN, "sheet")).find(el => el.getAttribute("name") === POINTS_OUTPUT_SHEET);
    if (!sheetEl) throw new Error("找不到「工作表1」");
    const rid = sheetEl.getAttributeNS(NS_REL, "id") || sheetEl.getAttribute("r:id");
    const relsDoc = parse(zipText(zip, "xl/_rels/workbook.xml.rels"));
    const rels = Array.from(relsDoc.getElementsByTagName("Relationship"));
    const rel = rels.find(el => el.getAttribute("Id") === rid);
    if (!rel) throw new Error("範本工作表關聯不完整");
    let target = rel.getAttribute("Target");
    target = target.startsWith("/") ? target.slice(1) : "xl/" + target.replace(/^\.\//, "");

    const doc = parse(zipText(zip, target));
    const sheetData = doc.getElementsByTagNameNS(NS_MAIN, "sheetData")[0];
    if (!sheetData) throw new Error("範本工作表格式不正確");
    const make = name => doc.createElementNS(NS_MAIN, name);

    // 清除第 2 列以後 A–G 欄的內容（保留儲存格格式）
    const rowEls = () => Array.from(sheetData.getElementsByTagNameNS(NS_MAIN, "row"));
    for (const row of rowEls()) {
      if (Number(row.getAttribute("r")) < 2) continue;
      for (const c of Array.from(row.getElementsByTagNameNS(NS_MAIN, "c"))) {
        if (colNumber(colLetter(c.getAttribute("r") || "")) > 7) continue;
        while (c.firstChild) c.removeChild(c.firstChild);
        c.removeAttribute("t");
      }
    }

    function getRow(n) {
      let before = null;
      for (const row of rowEls()) {
        const r = Number(row.getAttribute("r"));
        if (r === n) return row;
        if (r > n) { before = row; break; }
      }
      const row = make("row");
      row.setAttribute("r", String(n));
      sheetData.insertBefore(row, before);
      return row;
    }
    function getCell(row, ref) {
      const col = colNumber(colLetter(ref));
      let before = null;
      for (const c of Array.from(row.getElementsByTagNameNS(NS_MAIN, "c"))) {
        const n = colNumber(colLetter(c.getAttribute("r") || ""));
        if (n === col) return c;
        if (n > col) { before = c; break; }
      }
      const cell = make("c");
      cell.setAttribute("r", ref);
      row.insertBefore(cell, before);
      return cell;
    }
    function setText(row, ref, text) {
      const cell = getCell(row, ref), is = make("is"), t = make("t");
      cell.setAttribute("t", "inlineStr");
      t.setAttributeNS(NS_XML, "xml:space", "preserve");
      t.textContent = String(text == null ? "" : text);
      is.appendChild(t); cell.appendChild(is);
    }
    function setNumber(row, ref, number) {
      const cell = getCell(row, ref), v = make("v");
      v.textContent = String(number);
      cell.appendChild(v);
    }
    function setFormula(row, ref, formula, cached) {
      const cell = getCell(row, ref), f = make("f"), v = make("v");
      cell.setAttribute("t", "str");
      f.textContent = formula; v.textContent = cached;
      cell.append(f, v);
    }

    records.forEach((x, i) => {
      const n = i + 2, row = getRow(n);
      row.removeAttribute("spans"); // spans 只是提示；寫入新欄位後移除，避免與實際範圍不符
      setText(row, "A" + n, x.school);
      setText(row, "B" + n, x.name);
      setText(row, "C" + n, x.cat);
      setText(row, "D" + n, x.reason);
      setText(row, "E" + n, x.detail || "");
      setFormula(row, "F" + n, 'IF(LEFT(C' + n + ',2)="優良","加分",IF(LEFT(C' + n + ',2)="不良","扣分",""))',
        x.cat.startsWith("優良") ? "加分" : x.cat.startsWith("不良") ? "扣分" : "");
      setNumber(row, "G" + n, x.score);
    });
    const dim = doc.getElementsByTagNameNS(NS_MAIN, "dimension")[0];
    if (dim) {
      const ref = dim.getAttribute("ref") || "A1";
      const end = ref.split(":")[1] || ref;
      const lastRow = Math.max(Number(end.replace(/^[A-Z]+/, "")) || 1, records.length + 1);
      const lastCol = colNumber(colLetter(end)) >= 7 ? colLetter(end) : "G";
      dim.setAttribute("ref", "A1:" + lastCol + lastRow);
    }
    zipPut(zip, target, serializer.serializeToString(doc));

    // 移除計算鏈並要求開檔時重算，避免 Excel 出現「需要修復」
    if (zipEntry(zip, "xl/calcChain.xml")) {
      zipDelete(zip, "xl/calcChain.xml");
      rels.filter(el => /calcChain$/.test(el.getAttribute("Type") || "")).forEach(el => el.parentNode.removeChild(el));
      zipPut(zip, "xl/_rels/workbook.xml.rels", serializer.serializeToString(relsDoc));
      const ctDoc = parse(zipText(zip, "[Content_Types].xml"));
      Array.from(ctDoc.getElementsByTagName("Override")).filter(el => el.getAttribute("PartName") === "/xl/calcChain.xml")
        .forEach(el => el.parentNode.removeChild(el));
      zipPut(zip, "[Content_Types].xml", serializer.serializeToString(ctDoc));
    }
    let calcPr = wbDoc.getElementsByTagNameNS(NS_MAIN, "calcPr")[0];
    if (!calcPr) {
      calcPr = wbDoc.createElementNS(NS_MAIN, "calcPr");
      const wbRoot = wbDoc.documentElement;
      const after = Array.from(wbRoot.childNodes).find(n => n.nodeType === 1 &&
        ["oleSize", "customWorkbookViews", "pivotCaches", "smartTagPr", "smartTagTypes", "webPublishing", "fileRecoveryPr", "webPublishObjects", "extLst"].includes(n.localName));
      wbRoot.insertBefore(calcPr, after || null);
    }
    calcPr.setAttribute("fullCalcOnLoad", "1");
    zipPut(zip, "xl/workbook.xml", serializer.serializeToString(wbDoc));

    return xlsx.CFB.write(zip, {fileType: "zip", type: "array", compression: true});
  }

  return {
    clean, normalize, unique,
    extractRoster, guessCohortLabel, parseTokens, matchStudent, resolveStudents, duplicateStudents,
    rewardFields, extractRewardTemplate, rewardDetails, buildRewardRecords, editRewardRecords, exportRewardWorkbook,
    extractPointsTemplate, kindPrefix, validScore, buildPointsRecords, validatePointsPatch, exportPointsWorkbook
  };
});
