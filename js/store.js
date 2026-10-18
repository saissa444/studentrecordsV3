/*
 * 本機加密儲存。
 * - 資料存在這台裝置、這個瀏覽器的 IndexedDB，不會上傳。
 * - 每一筆內容都用「開啟密碼」推導出的 AES-GCM 256 金鑰加密（PBKDF2-SHA256，31 萬次）。
 * - 密碼本身不會被儲存；忘記密碼只能清除本機資料後重新匯入。
 */
(function (root) {
  "use strict";
  const DB_NAME = "student-records-local-v1", STORE = "kv", ITERATIONS = 310000;
  const enc = new TextEncoder(), dec = new TextDecoder();
  let key = null;

  function openDB() {
    return new Promise((resolve, reject) => {
      if (!root.indexedDB) { reject(new Error("此瀏覽器不支援本機資料庫（可能是無痕模式）。")); return; }
      const request = root.indexedDB.open(DB_NAME, 1);
      request.onupgradeneeded = () => request.result.createObjectStore(STORE);
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(new Error("無法開啟本機資料庫：" + (request.error && request.error.message)));
      request.onblocked = () => reject(new Error("本機資料庫被其他分頁占用，請關閉其他分頁後重試。"));
    });
  }

  async function tx(mode, work) {
    const db = await openDB();
    try {
      return await new Promise((resolve, reject) => {
        const t = db.transaction(STORE, mode), store = t.objectStore(STORE);
        let result;
        Promise.resolve(work(store)).then(r => { result = r; }, reject);
        t.oncomplete = () => resolve(result);
        t.onerror = () => reject(t.error);
        t.onabort = () => reject(t.error || new Error("本機資料寫入中止（空間可能不足）。"));
      });
    } finally { db.close(); }
  }
  const req = request => new Promise((resolve, reject) => { request.onsuccess = () => resolve(request.result); request.onerror = () => reject(request.error); });
  const rawGet = name => tx("readonly", store => req(store.get(name)));
  const rawPut = (name, value) => tx("readwrite", store => { store.put(value, name); });

  function requireCrypto() {
    if (!root.crypto || !root.crypto.subtle) throw new Error("此瀏覽器無法使用加密功能，請改用 https 網址開啟（例如 GitHub Pages）或更新瀏覽器。");
    return root.crypto.subtle;
  }

  async function deriveKey(password, salt, iterations) {
    const subtle = requireCrypto();
    const base = await subtle.importKey("raw", enc.encode(password.normalize("NFKC")), "PBKDF2", false, ["deriveKey"]);
    return subtle.deriveKey({name: "PBKDF2", salt, iterations, hash: "SHA-256"}, base, {name: "AES-GCM", length: 256}, false, ["encrypt", "decrypt"]);
  }

  async function encryptBytes(bytes, useKey) {
    const iv = root.crypto.getRandomValues(new Uint8Array(12));
    const ct = await requireCrypto().encrypt({name: "AES-GCM", iv}, useKey || key, bytes);
    return {iv, ct: new Uint8Array(ct)};
  }
  async function decryptBytes(box, useKey) {
    const plain = await requireCrypto().decrypt({name: "AES-GCM", iv: box.iv}, useKey || key, box.ct);
    return new Uint8Array(plain);
  }

  const CHECK = "student-records-ok";

  const Store = {
    /** 'new'：尚未設定密碼；'locked'：已有加密資料，待解鎖；'open'：已解鎖 */
    async status() {
      if (key) return "open";
      const meta = await rawGet("meta");
      return meta ? "locked" : "new";
    },

    async create(password) {
      if ((await rawGet("meta"))) throw new Error("已設定過密碼。");
      const salt = root.crypto.getRandomValues(new Uint8Array(16));
      const newKey = await deriveKey(password, salt, ITERATIONS);
      const check = await encryptBytes(enc.encode(CHECK), newKey);
      await rawPut("meta", {v: 1, salt, iterations: ITERATIONS, check, createdAt: Date.now()});
      key = newKey;
    },

    async unlock(password) {
      const meta = await rawGet("meta");
      if (!meta) throw new Error("尚未設定密碼。");
      const tryKey = await deriveKey(password, meta.salt, meta.iterations);
      try {
        const plain = await decryptBytes(meta.check, tryKey);
        if (dec.decode(plain) !== CHECK) return false;
      } catch (error) { return false; }
      key = tryKey;
      return true;
    },

    lock() { key = null; },

    async getJSON(name, fallback) {
      if (!key) throw new Error("尚未解鎖");
      const box = await rawGet("d:" + name);
      if (!box) return fallback;
      return JSON.parse(dec.decode(await decryptBytes(box)));
    },
    async setJSON(name, value) {
      if (!key) throw new Error("尚未解鎖");
      await rawPut("d:" + name, await encryptBytes(enc.encode(JSON.stringify(value))));
    },
    async getBytes(name) {
      if (!key) throw new Error("尚未解鎖");
      const box = await rawGet("b:" + name);
      return box ? decryptBytes(box) : null;
    },
    async setBytes(name, bytes) {
      if (!key) throw new Error("尚未解鎖");
      await rawPut("b:" + name, await encryptBytes(new Uint8Array(bytes)));
    },
    async remove(name) {
      await tx("readwrite", store => { store.delete("d:" + name); store.delete("b:" + name); });
    },

    async changePassword(oldPassword, newPassword) {
      const meta = await rawGet("meta");
      const oldKey = await deriveKey(oldPassword, meta.salt, meta.iterations);
      try { await decryptBytes(meta.check, oldKey); } catch (error) { throw new Error("目前密碼不正確。"); }
      const salt = root.crypto.getRandomValues(new Uint8Array(16));
      const newKey = await deriveKey(newPassword, salt, ITERATIONS);
      const all = await tx("readonly", store => Promise.all([req(store.getAllKeys()), req(store.getAll())]));
      const [names, values] = all;
      const next = [];
      for (let i = 0; i < names.length; i++) {
        if (names[i] === "meta") continue;
        next.push([names[i], await encryptBytes(await decryptBytes(values[i], oldKey), newKey)]);
      }
      const check = await encryptBytes(enc.encode(CHECK), newKey);
      await tx("readwrite", store => {
        next.forEach(([name, box]) => store.put(box, name));
        store.put({v: 1, salt, iterations: ITERATIONS, check, createdAt: meta.createdAt}, "meta");
      });
      key = newKey;
    },

    /** 刪除這台裝置上本 App 的全部資料（名冊、範本、未匯出紀錄、密碼設定）。 */
    async wipe() {
      key = null;
      await new Promise((resolve, reject) => {
        const request = root.indexedDB.deleteDatabase(DB_NAME);
        request.onsuccess = () => resolve();
        request.onerror = () => reject(request.error);
        request.onblocked = () => resolve();
      });
    },

    async requestPersistence() {
      try {
        if (!navigator.storage || !navigator.storage.persist) return null;
        if (await navigator.storage.persisted()) return true;
        return await navigator.storage.persist();
      } catch (error) { return null; }
    }
  };

  root.LocalStore = Store;
})(typeof globalThis !== "undefined" ? globalThis : this);
