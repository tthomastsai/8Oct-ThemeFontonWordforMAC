/* taskpane.js — 介面與 Word 互動
 * 流程：getFileAsync 取得整份 docx → 修改 theme1.xml → insertFileFromBase64("Replace", {importTheme:true})
 */
(function () {
  "use strict";

  // Office.js 沒有列出已安裝字型的 API，所以用常用字型清單；欄位仍可直接輸入任何名稱。
  const FONTS = [
    // 中文
    { n: "PMingLiU", alt: "新細明體" }, { n: "MingLiU", alt: "細明體" },
    { n: "DFKai-SB", alt: "標楷體" }, { n: "Microsoft JhengHei", alt: "微軟正黑體" },
    { n: "Microsoft JhengHei UI", alt: "微軟正黑體 UI" }, { n: "PingFang TC", alt: "蘋方-繁" },
    { n: "PingFang SC", alt: "蘋方-簡" }, { n: "Heiti TC", alt: "黑體-繁" },
    { n: "Songti TC", alt: "宋體-繁" }, { n: "Kaiti TC", alt: "楷體-繁" },
    { n: "STSong", alt: "華文宋體" }, { n: "Noto Sans TC", alt: "思源黑體" },
    { n: "Noto Serif TC", alt: "思源宋體" }, { n: "Microsoft YaHei", alt: "微軟雅黑" },
    { n: "SimSun", alt: "宋体" }, { n: "SimHei", alt: "黑体" },
    // 英文：襯線
    { n: "Times New Roman" }, { n: "Georgia" }, { n: "Cambria" }, { n: "Garamond" },
    { n: "Palatino" }, { n: "Book Antiqua" }, { n: "Baskerville" }, { n: "Constantia" },
    // 英文：無襯線
    { n: "Calibri" }, { n: "Calibri Light" }, { n: "Arial" }, { n: "Helvetica" },
    { n: "Helvetica Neue" }, { n: "Verdana" }, { n: "Tahoma" }, { n: "Trebuchet MS" },
    { n: "Segoe UI" }, { n: "Century Gothic" }, { n: "Gill Sans MT" }, { n: "Candara" },
    { n: "Corbel" }, { n: "Optima" }, { n: "Futura" }, { n: "Avenir" },
    { n: "Aptos" }, { n: "Aptos Display" }, { n: "Aptos Narrow" },
    // 等寬
    { n: "Courier New" }, { n: "Consolas" }, { n: "Menlo" }, { n: "Monaco" },
  ];

  const KEYS = ["majorLatin", "majorEa", "minorLatin", "minorEa"];
  const DEFAULTS = {
    majorLatin: "Times New Roman", majorEa: "PMingLiU",
    minorLatin: "Times New Roman", minorEa: "PMingLiU",
  };
  const combos = {};

  const $ = (id) => document.getElementById(id);

  function setStatus(msg, cls) {
    const s = $("status");
    s.textContent = msg || "";
    s.className = cls || "";
  }

  // ---------------------------------------------------------------- 可搜尋字型選單
  function makeCombo(host, key) {
    const wrap = document.createElement("div");
    wrap.className = "combo";
    wrap.innerHTML =
      '<input type="text" autocomplete="off" spellcheck="false" aria-label="字型">' +
      '<button type="button" class="chev" tabindex="-1" aria-label="展開">▾</button>' +
      '<ul class="menu" role="listbox"></ul>';
    host.appendChild(wrap);

    const input = wrap.querySelector("input");
    const menu = wrap.querySelector("ul");
    let items = [], active = -1;

    const fontCss = (n) => '"' + n.replace(/["\\]/g, "") + '", -apple-system, sans-serif';

    function render(filter) {
      const q = (filter || "").trim().toLowerCase();
      items = FONTS.filter((f) => !q || f.n.toLowerCase().includes(q) ||
        (f.alt && f.alt.toLowerCase().includes(q)));
      menu.textContent = "";
      items.forEach((f, i) => {
        const li = document.createElement("li");
        li.setAttribute("role", "option");
        const nm = document.createElement("span");
        nm.textContent = f.n;
        nm.style.fontFamily = fontCss(f.n);          // 每個選項用自己的字型顯示
        li.appendChild(nm);
        if (f.alt) {
          const a = document.createElement("span");
          a.className = "alt"; a.textContent = f.alt; li.appendChild(a);
        }
        li.addEventListener("mousedown", (e) => { e.preventDefault(); choose(f.n); });
        li.addEventListener("mousemove", () => highlight(i));
        menu.appendChild(li);
      });
      if (!items.length) {
        const li = document.createElement("li");
        li.className = "empty";
        li.textContent = q ? "清單中沒有符合的字型，按 Enter 直接使用「" + filter.trim() + "」" : "";
        menu.appendChild(li);
      }
      active = items.length && q ? 0 : -1;
      highlight(active, true);
    }

    function highlight(i, noScroll) {
      active = i;
      [...menu.children].forEach((li, j) => li.classList.toggle("active", j === i));
      if (!noScroll && menu.children[i]) menu.children[i].scrollIntoView({ block: "nearest" });
    }

    function open() { wrap.classList.add("open"); render(input.dataset.typing === "1" ? input.value : ""); }
    function close() { wrap.classList.remove("open"); input.dataset.typing = ""; }

    function choose(name) {
      input.value = name;
      close();
      input.dispatchEvent(new Event("change", { bubbles: true }));
    }

    input.addEventListener("focus", () => { input.select(); open(); });
    input.addEventListener("input", () => {
      input.dataset.typing = "1";
      wrap.classList.add("open");
      render(input.value);
    });
    input.addEventListener("blur", () => setTimeout(close, 120));
    input.addEventListener("keydown", (e) => {
      if (e.key === "ArrowDown") { e.preventDefault(); wrap.classList.add("open"); highlight(Math.min(items.length - 1, active + 1)); }
      else if (e.key === "ArrowUp") { e.preventDefault(); highlight(Math.max(0, active - 1)); }
      else if (e.key === "Enter") {
        e.preventDefault();
        if (wrap.classList.contains("open") && items[active] && input.dataset.typing === "1") choose(items[active].n);
        else { close(); input.dispatchEvent(new Event("change", { bubbles: true })); }
      } else if (e.key === "Escape") { close(); }
    });
    wrap.querySelector(".chev").addEventListener("mousedown", (e) => {
      e.preventDefault();
      if (wrap.classList.contains("open")) close(); else { input.focus(); }
    });

    combos[key] = input;
  }

  const getVals = () => Object.fromEntries(KEYS.map((k) => [k, combos[k].value.trim()]));
  const setVals = (v) => KEYS.forEach((k) => { if (v[k] !== undefined) combos[k].value = v[k]; });

  // ---------------------------------------------------------------- 與 Word 互動
  function getDocxBytes() {
    return new Promise((resolve, reject) => {
      Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 4 * 1024 * 1024 }, (r) => {
        if (r.status !== Office.AsyncResultStatus.Succeeded) return reject(new Error(r.error.message));
        const file = r.value, parts = [];
        const finish = () => {
          file.closeAsync(() => {});
          const total = parts.reduce((s, p) => s + p.length, 0);
          const out = new Uint8Array(total);
          let off = 0;
          parts.forEach((p) => { out.set(p, off); off += p.length; });
          resolve(out);
        };
        const next = (i) => {
          if (i >= file.sliceCount) return finish();
          file.getSliceAsync(i, (s) => {
            if (s.status !== Office.AsyncResultStatus.Succeeded) {
              file.closeAsync(() => {});
              return reject(new Error(s.error.message));
            }
            parts.push(Uint8Array.from(s.value.data));
            next(i + 1);
          });
        };
        next(0);
      });
    });
  }

  function bytesToBase64(bytes) {
    let bin = "";
    for (let i = 0; i < bytes.length; i += 0x8000) {
      bin += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
    }
    return btoa(bin);
  }

  async function readCurrent() {
    return ThemeFontsCore.readThemeFromDocx(await getDocxBytes());
  }

  async function onRead() {
    setStatus("讀取中…", "busy");
    try {
      const v = await readCurrent();
      setVals(v);
      setStatus("目前主題字型\n標題：" + v.majorLatin + " / " + v.majorEa +
        "\n內文：" + v.minorLatin + " / " + v.minorEa, "ok");
    } catch (e) { setStatus("讀取失敗：" + e.message, "err"); }
  }

  async function onApply() {
    const want = getVals();
    if (!KEYS.some((k) => want[k])) return setStatus("請至少設定一個字型。", "err");
    const btn = $("apply"); btn.disabled = true;
    try {
      setStatus("讀取文件…", "busy");
      const bytes = await getDocxBytes();

      if ($("backup").checked) {
        setStatus("開啟備份…", "busy");
        await Word.run(async (ctx) => { ctx.application.createDocument(bytesToBase64(bytes)).open(); await ctx.sync(); });
      }

      setStatus("修改主題字型…", "busy");
      const b64 = await ThemeFontsCore.patchDocx(bytes, {
        majorLatin: want.majorLatin, majorEa: want.majorEa,
        minorLatin: want.minorLatin, minorEa: want.minorEa,
      }, "base64");

      setStatus("套用到文件…", "busy");
      await Word.run(async (ctx) => {
        ctx.document.insertFileFromBase64(b64, "Replace", { importTheme: true, importStyles: false });
        await ctx.sync();
      });

      // 驗證：重新讀回文件的主題字型，確認真的有生效
      const got = await readCurrent();
      const diff = KEYS.filter((k) => want[k] && got[k] !== want[k]);
      if (diff.length) {
        setStatus("已執行，但文件的主題字型似乎沒有改變（" + diff.join("、") + "）。\n" +
          "這個 Word 版本可能不支援匯入主題，請改用 Python 版本。", "err");
      } else {
        setStatus("完成，已驗證主題字型已更新。\n標題：" + got.majorLatin + " / " + got.majorEa +
          "\n內文：" + got.minorLatin + " / " + got.minorEa, "ok");
      }
    } catch (e) {
      setStatus("失敗：" + e.message, "err");
    } finally { btn.disabled = false; }
  }

  // ---------------------------------------------------------------- 啟動
  function init() {
    document.querySelectorAll(".field").forEach((f) => makeCombo(f, f.dataset.key));
    setVals(DEFAULTS);

    // 內文跟隨標題
    ["majorLatin", "majorEa"].forEach((k) => {
      combos[k].addEventListener("change", () => {
        if ($("sync").checked) combos[k.replace("major", "minor")].value = combos[k].value;
      });
    });

    $("apply").addEventListener("click", onApply);
    $("read").addEventListener("click", onRead);

    if (window.Office && Office.context && Office.context.requirements &&
        !Office.context.requirements.isSetSupported("WordApi", "1.5")) {
      setStatus("此 Word 版本不支援 WordApi 1.5，無法使用本增益集。", "err");
      $("apply").disabled = true;
    } else {
      onRead();   // 開啟時自動帶入目前文件的主題字型
    }
  }

  Office.onReady(init);
})();
