/* core.js — 主題字型修改核心（瀏覽器與 Node 共用，邏輯與 theme_fonts.py 一致）
 * .docx 是 zip，主題字型存在 word/theme/theme1.xml。
 */
(function (root, factory) {
  if (typeof module === "object" && module.exports) {
    module.exports = factory(require("jszip"));
  } else {
    root.ThemeFontsCore = factory(root.JSZip);
  }
})(typeof self !== "undefined" ? self : this, function (JSZip) {
  "use strict";

  const THEME = "word/theme/theme1.xml";

  const esc = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;")
    .replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  const unesc = (s) => String(s).replace(/&quot;/g, '"').replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">").replace(/&amp;/g, "&");

  // 把第一個符合 tagRe 的標籤的 typeface 換掉，並移除會與新字型衝突的 panose 等屬性
  function setFace(body, tagRe, face) {
    return body.replace(tagRe, (tag) =>
      tag.replace(/typeface="[^"]*"/, 'typeface="' + esc(face) + '"')
         .replace(/\s+(panose|pitchFamily|charset)="[^"]*"/g, ""));
  }

  function blockRe(block) {
    return new RegExp("(<a:" + block + ">)([\\s\\S]*?)(</a:" + block + ">)");
  }

  function patchBlock(xml, block, latin, ea) {
    const m = xml.match(blockRe(block));
    if (!m) throw new Error("找不到 <a:" + block + ">，這份文件的主題格式不尋常。");
    let body = m[2];
    if (latin) body = setFace(body, /<a:latin\b[^>]*\/>/, latin);
    if (ea) {
      body = setFace(body, /<a:ea\b[^>]*\/>/, ea);
      // 繁體中文實際走的是 script="Hant" 這一項
      body = setFace(body, /<a:font\b[^>]*script="Hant"[^>]*\/>/, ea);
    }
    return xml.slice(0, m.index) + m[1] + body + m[3] + xml.slice(m.index + m[0].length);
  }

  /** opts: { majorLatin, majorEa, minorLatin, minorEa }（空值表示不改） */
  function patchThemeXml(xml, opts) {
    xml = patchBlock(xml, "majorFont", opts.majorLatin, opts.majorEa); // Headings
    xml = patchBlock(xml, "minorFont", opts.minorLatin, opts.minorEa); // Body
    return xml;
  }

  function faceOf(body, tagRe) {
    const tag = body.match(tagRe);
    if (!tag) return "";
    const f = tag[0].match(/typeface="([^"]*)"/);
    return f ? unesc(f[1]) : "";
  }

  function readBlock(xml, block) {
    const m = xml.match(blockRe(block));
    if (!m) return { latin: "", ea: "" };
    const body = m[2];
    const latin = faceOf(body, /<a:latin\b[^>]*\/>/);
    const ea = faceOf(body, /<a:font\b[^>]*script="Hant"[^>]*\/>/) ||
               faceOf(body, /<a:ea\b[^>]*\/>/);
    return { latin, ea };
  }

  function readThemeFonts(xml) {
    const major = readBlock(xml, "majorFont");
    const minor = readBlock(xml, "minorFont");
    return { majorLatin: major.latin, majorEa: major.ea, minorLatin: minor.latin, minorEa: minor.ea };
  }

  async function loadTheme(bytes) {
    const zip = await JSZip.loadAsync(bytes);
    const f = zip.file(THEME);
    if (!f) throw new Error("此文件沒有主題檔 (theme1.xml)。");
    return { zip, xml: await f.async("string") };
  }

  async function readThemeFromDocx(bytes) {
    const { xml } = await loadTheme(bytes);
    return readThemeFonts(xml);
  }

  async function patchDocx(bytes, opts, outType) {
    const { zip, xml } = await loadTheme(bytes);
    zip.file(THEME, patchThemeXml(xml, opts));
    return zip.generateAsync({ type: outType || "base64", compression: "DEFLATE" });
  }

  return { THEME, patchThemeXml, readThemeFonts, readThemeFromDocx, patchDocx };
});
