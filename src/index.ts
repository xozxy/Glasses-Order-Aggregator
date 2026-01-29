import Papa from "papaparse";
import { PDFDocument, StandardFonts, rgb } from "pdf-lib";
import fontkit from "@pdf-lib/fontkit";

/** =========================
 * CSV 必要字段（订单明细）
 * ========================= */
const ORDER_ID = "Order ID";
const BUNDLE_ID = "Bundle ID";
const LINE_ITEM = "Line Item";
const QTY = "Quantity";

/** =========================
 * 你的字体文件（已放到 public/fonts/）
 * 多字体兼容：JP/SC/TC/KR
 * ========================= */
const FONT_FILES = {
  jp: "/fonts/NotoSansJP.ttf",
  sc: "/fonts/NotoSansSC.ttf",
  tc: "/fonts/NotoSansTC.ttf",
  kr: "/fonts/NotoSansKR.ttf",
};

/** =========================
 * Label 尺寸（5x4cm）
 * ========================= */
const CM_TO_PT = 28.3464566929;
const PAGE_W = 5 * CM_TO_PT;
const PAGE_H = 4 * CM_TO_PT;
const DEFAULT_LABEL_LIMIT = 200;
const MAX_LABEL_LIMIT = 300;

/** =========================
 * 工具函数
 * ========================= */
function uniqPreserveOrder(arr: any[]) {
  const seen = new Set<string>();
  const out: string[] = [];
  for (const x of arr) {
    const v = String(x ?? "").trim();
    if (!v) continue;
    if (!seen.has(v)) {
      seen.add(v);
      out.push(v);
    }
  }
  return out;
}

function categorize(item: any) {
  const s = String(item ?? "").toLowerCase();
  if (s.includes("coating")) return "Coating";
  if (s.includes("index lens") || s.includes("prescription lens") || /\blens\b/i.test(s)) return "Lens";
  if (s.includes("frame") || s.includes("glasses style")) return "Frame";
  return "Other";
}

function parseIndexLensFromText(val: any): string {
  const s = String(val ?? "");
  const m = s.match(/1\.\d{1,2}|\d\.\d{1,2}/);
  if (!m) return "";
  const n = Number(m[0]);
  if (!Number.isFinite(n)) return "";
  return n.toFixed(2);
}

function fmtTwoDec(x: any): string {
  const n = Number(x);
  return Number.isFinite(n) ? n.toFixed(2) : "";
}
function fmtAxis(x: any): string {
  const n = Number(x);
  return Number.isFinite(n) ? String(Math.round(n)) : "";
}

function pick(row: Record<string, any>, keys: string[]) {
  for (const k of keys) {
    const v = row[k];
    if (v !== null && v !== undefined && String(v).trim() !== "") return v;
  }
  return "";
}

function getPD(row: Record<string, any>): [string, string] {
  const pd_od = fmtTwoDec(pick(row, ["PD_OD", "PD OD", "OD PD", "Pupillary Distance OD"]));
  const pd_os = fmtTwoDec(pick(row, ["PD_OS", "PD OS", "OS PD", "Pupillary Distance OS"]));
  if (pd_od || pd_os) return [pd_od, pd_os];

  const single = pick(row, ["Single PD", "Single_PD", "PD", "Pupillary Distance"]);
  const n = Number(single);
  if (Number.isFinite(n) && n > 0) {
    const half = n / 2;
    return [half.toFixed(2), half.toFixed(2)];
  }
  return ["", ""];
}

/** =========================
 * 读取字体 bytes（从 env.ASSETS）
 * ========================= */
async function loadFontBytes(env: any, path: string): Promise<Uint8Array> {
  if (!env?.ASSETS?.fetch) throw new Error("ASSETS binding missing. Check wrangler.jsonc assets.binding is 'ASSETS'.");
  const res = await env.ASSETS.fetch(new Request("http://local" + path));
  if (!res.ok) throw new Error(`Font not found: ${path} (status ${res.status})`);
  return new Uint8Array(await res.arrayBuffer());
}

/** =========================
 * 多字体 fallback：按字符选择字体
 * ========================= */
function isHiraganaKatakana(ch: string) {
  const cp = ch.codePointAt(0) || 0;
  return (
    (cp >= 0x3040 && cp <= 0x30ff) || // Hiragana + Katakana
    (cp >= 0x31f0 && cp <= 0x31ff) || // Katakana Phonetic Extensions
    (cp >= 0xff66 && cp <= 0xff9d) // Halfwidth Katakana
  );
}
function isHangul(ch: string) {
  const cp = ch.codePointAt(0) || 0;
  return (
    (cp >= 0xac00 && cp <= 0xd7af) || // Hangul Syllables
    (cp >= 0x1100 && cp <= 0x11ff) || // Hangul Jamo
    (cp >= 0x3130 && cp <= 0x318f) || // Hangul Compatibility Jamo
    (cp >= 0xa960 && cp <= 0xa97f) || // Hangul Jamo Extended-A
    (cp >= 0xd7b0 && cp <= 0xd7ff) // Hangul Jamo Extended-B
  );
}
function isCJKUnified(ch: string) {
  const cp = ch.codePointAt(0) || 0;
  return (
    (cp >= 0x4e00 && cp <= 0x9fff) || // CJK Unified Ideographs
    (cp >= 0x3400 && cp <= 0x4dbf) || // Extension A
    (cp >= 0x20000 && cp <= 0x2a6df) || // Extension B
    (cp >= 0x2a700 && cp <= 0x2b73f) || // Extension C
    (cp >= 0x2b740 && cp <= 0x2b81f) || // Extension D
    (cp >= 0x2b820 && cp <= 0x2ceaf) || // Extension E
    (cp >= 0x2ceb0 && cp <= 0x2ebef) || // Extension F
    (cp >= 0x30000 && cp <= 0x3134f) || // Extension G
    (cp >= 0xf900 && cp <= 0xfaff) || // Compatibility Ideographs
    (cp >= 0x2f800 && cp <= 0x2fa1f) // Compatibility Supplement
  );
}
function isAsciiPrintable(ch: string) {
  const cp = ch.codePointAt(0) || 0;
  return cp >= 0x20 && cp <= 0x7e;
}
function isAsciiString(text: string) {
  for (const ch of text) {
    if (!isAsciiPrintable(ch)) return false;
  }
  return true;
}
type FontNeeds = { jp: boolean; kr: boolean; sc: boolean; other: boolean; nonAscii: boolean };
function scanTextForFonts(text: any, needs: FontNeeds) {
  if (text === null || text === undefined) return;
  const s = String(text);
  for (const ch of s) {
    if (isAsciiPrintable(ch)) continue;
    needs.nonAscii = true;
    if (isHiraganaKatakana(ch)) {
      needs.jp = true;
    } else if (isHangul(ch)) {
      needs.kr = true;
    } else if (isCJKUnified(ch)) {
      needs.sc = true;
    } else {
      needs.other = true;
    }
  }
}

/**
 * 规则：
 * - ASCII (0x20..0x7E) → Helvetica
 * - 假名 → JP
 * - 韩文 → KR
 * - 汉字 → 默认 SC（简中优先；如果你希望繁中优先，把 sc 改 tc）
 * - 其他 → SC（或 JP），避免 WinAnsi 编码错误
 */
function pickFontForChar(fonts: any, ch: string) {
  if (isAsciiPrintable(ch)) return fonts.latin;
  if (isHiraganaKatakana(ch) && fonts.jp) return fonts.jp;
  if (isHangul(ch) && fonts.kr) return fonts.kr;
  if (isCJKUnified(ch) && fonts.sc) return fonts.sc; // 想繁中优先就改 fonts.tc
  const fallback = fonts.fallback || fonts.sc || fonts.jp || fonts.kr;
  if (!fallback) throw new Error("Missing embedded font for non-ASCII text");
  return fallback;
}

/**
 * 按 run 绘制：同一字体连续字符合并成一段绘制
 * x,y 为 baseline 坐标（pdf-lib 原生坐标）
 */
function drawTextRuns(page: any, fonts: any, text: string, x: number, y: number, size: number, color: any) {
  if (!text) return;
  if (isAsciiString(text)) {
    page.drawText(text, { x, y, size, font: fonts.latin, color });
    return;
  }
  let cursorX = x;
  let buf = "";
  let curFont: any = null;

  const flush = () => {
    if (!buf || !curFont) return;
    page.drawText(buf, { x: cursorX, y, size, font: curFont, color });
    cursorX += curFont.widthOfTextAtSize(buf, size);
    buf = "";
  };

  for (const ch of text) {
    const f = pickFontForChar(fonts, ch);
    if (!curFont) curFont = f;
    if (f !== curFont) {
      flush();
      curFont = f;
    }
    buf += ch;
  }
  flush();
}

/** =========================
 * 画 Label（top-based 坐标，接近你 Python 模板）
 * ========================= */
function normalizeText(val: any): string {
  if (val === null || val === undefined) return "";
  return String(val)
    .replace(/[\r\n\t]+/g, " ")
    .replace(/\s{2,}/g, " ")
    .trim();
}

function measureTextRuns(fonts: any, text: string, size: number) {
  if (!text) return 0;
  if (isAsciiString(text)) return fonts.latin.widthOfTextAtSize(text, size);
  let width = 0;
  let buf = "";
  let curFont: any = null;

  const flush = () => {
    if (!buf || !curFont) return;
    width += curFont.widthOfTextAtSize(buf, size);
    buf = "";
  };

  for (const ch of text) {
    const f = pickFontForChar(fonts, ch);
    if (!curFont) curFont = f;
    if (f !== curFont) {
      flush();
      curFont = f;
    }
    buf += ch;
  }
  flush();
  return width;
}

function drawLabel(page: any, fonts: any, data: {
  backer: string;
  name: string;
  presType: string;
  thickness: string;
  coating: string;
  od: { sph: string; cyl: string; axis: string; add: string; pd: string };
  os: { sph: string; cyl: string; axis: string; add: string; pd: string };
  dateStr: string;
}) {
  const black = rgb(0, 0, 0);
  const W = PAGE_W;
  const H = PAGE_H;

  const x = (frac: number) => W * frac;
  const top = (yFromBottom: number) => (1 - yFromBottom) * H;

  const drawTopLeft = (xPos: number, topY: number, text: string, size: number) => {
    const t = normalizeText(text);
    if (!t) return;
    const baselineY = H - topY - size;
    drawTextRuns(page, fonts, t, xPos, baselineY, size, black);
  };

  const drawTopCenter = (xCenter: number, topY: number, text: string, size: number) => {
    const t = normalizeText(text);
    if (!t) return;
    const baselineY = H - topY - size;
    const w = measureTextRuns(fonts, t, size);
    drawTextRuns(page, fonts, t, xCenter - w / 2, baselineY, size, black);
  };

  const hlineTop = (topY: number, thickness = 1) => {
    const y = H - topY;
    page.drawLine({ start: { x: x(0.03), y }, end: { x: x(0.97), y }, thickness, color: black });
  };

  // 顶部两行
  drawTopLeft(x(0.03), top(0.92), `Backer Number: ${data.backer || ""}`, 7);
  drawTopLeft(x(0.03), top(0.80), `Name: ${data.name || ""}`, 7);

  hlineTop(top(0.74), 1);

  // Prescription / Thickness / Coating
  drawTopLeft(x(0.03), top(0.66), "Prescription:", 6);
  drawTopLeft(x(0.45), top(0.66), data.presType || "-", 6);

  drawTopLeft(x(0.03), top(0.58), "Thickness:", 6);
  drawTopLeft(x(0.45), top(0.58), data.thickness || "index lens", 6);

  drawTopLeft(x(0.03), top(0.50), "Coating:", 6);
  drawTopLeft(x(0.45), top(0.50), data.coating || "-", 6);

  hlineTop(top(0.44), 0.8);

  // 表格区域
  const colX = [0.16, 0.32, 0.48, 0.64, 0.8].map(x);
  ["sph", "cyl", "axis", "add", "pd"].forEach((h, i) => drawTopCenter(colX[i], top(0.37), h, 6));
  drawTopCenter(x(0.06), top(0.29), "od", 6);
  drawTopCenter(x(0.06), top(0.21), "os", 6);

  const vOD = [data.od.sph, data.od.cyl, data.od.axis, data.od.add, data.od.pd].map((v) => v || "-");
  const vOS = [data.os.sph, data.os.cyl, data.os.axis, data.os.add, data.os.pd].map((v) => v || "-");

  vOD.forEach((v, i) => drawTopCenter(colX[i], top(0.29), v, 6));
  vOS.forEach((v, i) => drawTopCenter(colX[i], top(0.21), v, 6));

  hlineTop(top(0.15), 0.8);
  drawTopLeft(x(0.03), top(0.08), data.dateStr || "", 6);
}

/** =========================
 * /api/aggregate：订单聚合（一单一行 CSV）
 * ========================= */
async function handleAggregate(request: Request) {
  const url = new URL(request.url);
  const format = (url.searchParams.get("format") || "csv").toLowerCase();

  const form = await request.formData();
  const fileAny = form.get("file");
  if (!fileAny || typeof fileAny === "string") return new Response("Missing/invalid file field 'file'", { status: 400 });
  const text = await (fileAny as File).text();

  const parsed = Papa.parse<Record<string, any>>(text, { header: true, skipEmptyLines: true });
  if (parsed.errors?.length) return new Response("CSV parse error: " + JSON.stringify(parsed.errors.slice(0, 3)), { status: 400 });
  const rows = parsed.data || [];
  if (!rows.length) return new Response("Empty CSV", { status: 400 });

  for (const col of [ORDER_ID, BUNDLE_ID, LINE_ITEM, QTY]) {
    if (!(col in rows[0])) return new Response(`Missing required column: ${col}`, { status: 400 });
  }

  const allColumns = Object.keys(rows[0]);
  const rxRe = /\b(OD|OS|PD|Prism|ADD|Axis|Cylinder|Sphere|Pupillary|base)\b/i;
  const rxColumns = allColumns.filter((c) => rxRe.test(c) || c === "Lens Notes");

  const lineLevelCols = new Set([LINE_ITEM, QTY, "Line Item Price"]);
  const idCols = new Set([ORDER_ID, BUNDLE_ID]);
  const rxSet = new Set(rxColumns);
  const orderLevelCols = allColumns.filter((c) => !lineLevelCols.has(c) && !rxSet.has(c) && !idCols.has(c));

  const orderMap = new Map<string, any>();

  for (const r of rows) {
    const orderId = r[ORDER_ID];
    const bundleId = r[BUNDLE_ID];

    if (!orderMap.has(orderId)) orderMap.set(orderId, { orderId, orderFields: {}, bundles: new Map() });
    const o = orderMap.get(orderId);

    for (const c of orderLevelCols) {
      if (!o.orderFields[c] && String(r[c] ?? "").trim() !== "") o.orderFields[c] = r[c];
    }

    if (!o.bundles.has(bundleId)) o.bundles.set(bundleId, { bundleId, items: { Frame: [], Lens: [], Coating: [], Other: [] }, rx: {} });
    const b = o.bundles.get(bundleId);

    const item = r[LINE_ITEM];
    const qty = Number(r[QTY]) || 1;
    const display = qty !== 1 ? `${item} x${qty}` : String(item);
    b.items[categorize(item)].push(display);

    for (const c of rxColumns) {
      if (!b.rx[c] && String(r[c] ?? "").trim() !== "") b.rx[c] = r[c];
    }
  }

  const out: any[] = [];
  let maxBundleCount = 1;

  for (const [, o] of orderMap) {
    const bundles = Array.from(o.bundles.values()).sort((a: any, b: any) => String(a.bundleId).localeCompare(String(b.bundleId)));
    maxBundleCount = Math.max(maxBundleCount, bundles.length);

    const rowOut: any = { [ORDER_ID]: o.orderId, "Bundle Count": bundles.length, ...o.orderFields };
    bundles.forEach((b: any, idx: number) => {
      const i = idx + 1;
      rowOut[`Bundle ID (Bundle ${i})`] = b.bundleId;
      rowOut[`Frame Items (Bundle ${i})`] = uniqPreserveOrder(b.items.Frame).join("; ");
      rowOut[`Lens Items (Bundle ${i})`] = uniqPreserveOrder(b.items.Lens).join("; ");
      rowOut[`Coating Items (Bundle ${i})`] = uniqPreserveOrder(b.items.Coating).join("; ");
      rowOut[`Other Items (Bundle ${i})`] = uniqPreserveOrder(b.items.Other).join("; ");
      for (const c of rxColumns) rowOut[`${c} (Bundle ${i})`] = b.rx[c] || "";
    });

    out.push(rowOut);
  }

  const baseCols = [ORDER_ID, "Bundle Count", ...orderLevelCols];
  const bundleCols: string[] = [];
  for (let i = 1; i <= maxBundleCount; i++) {
    bundleCols.push(
      `Bundle ID (Bundle ${i})`,
      `Frame Items (Bundle ${i})`,
      `Lens Items (Bundle ${i})`,
      `Coating Items (Bundle ${i})`,
      `Other Items (Bundle ${i})`,
      ...rxColumns.map((c) => `${c} (Bundle ${i})`)
    );
  }
  const columns = [...baseCols, ...bundleCols];

  if (format === "xlsx") return new Response("xlsx 暂未启用（当前仅 CSV）", { status: 400 });

  const csv = Papa.unparse(out, { columns, skipEmptyLines: true });
  return new Response(csv, {
    headers: {
      "Content-Type": "text/csv; charset=utf-8",
      "Content-Disposition": 'attachment; filename="orders_one_row.csv"',
    },
  });
}

/** =========================
 * /api/labels：生成 PDF Label（5x4cm）
 * mode=bundle  每个 bundle 一张
 * mode=order   每个订单一张（取第一个 bundle 的处方信息）
 * ========================= */
async function handleLabels(request: Request, env: any) {
  const url = new URL(request.url);
  const mode = (url.searchParams.get("mode") || "bundle").toLowerCase(); // bundle | order
  const startRaw = url.searchParams.get("start");
  const limitRaw = url.searchParams.get("limit");
  const paging = startRaw !== null || limitRaw !== null;
  let start = startRaw ? Number.parseInt(startRaw, 10) : 0;
  let limit = limitRaw ? Number.parseInt(limitRaw, 10) : DEFAULT_LABEL_LIMIT;
  if (!Number.isFinite(start) || start < 0) start = 0;
  if (!Number.isFinite(limit) || limit <= 0) limit = DEFAULT_LABEL_LIMIT;
  if (limit > MAX_LABEL_LIMIT) {
    return new Response(`limit too large (max ${MAX_LABEL_LIMIT})`, { status: 400 });
  }

  const t0 = Date.now();
  const log = (msg: string) => console.log(`[labels] ${msg} (+${Date.now() - t0}ms)`);
  let stage = "init";

  try {
    stage = "parse-form";
    const form = await request.formData();
    const fileAny = form.get("file");
    if (!fileAny || typeof fileAny === "string") return new Response("Missing/invalid file field 'file'", { status: 400 });
    const csvText = await (fileAny as File).text();
    log("form parsed");

    stage = "parse-csv";
    const parsed = Papa.parse<Record<string, any>>(csvText, { header: true, skipEmptyLines: true });
    if (parsed.errors?.length) return new Response("CSV parse error: " + JSON.stringify(parsed.errors.slice(0, 3)), { status: 400 });
    const rows = parsed.data || [];
    if (!rows.length) return new Response("Empty CSV", { status: 400 });

    for (const col of [ORDER_ID, BUNDLE_ID, LINE_ITEM, QTY]) {
      if (!(col in rows[0])) return new Response(`Missing required column: ${col}`, { status: 400 });
    }
    log(`csv parsed rows=${rows.length}`);

    stage = "scan-text";
    const needs: FontNeeds = { jp: false, kr: false, sc: false, other: false, nonAscii: false };

    // === 组织 order/bundle ===
    const orderMap = new Map<string, any>();

    for (const r of rows) {
      const orderId = r[ORDER_ID];
      const bundleId = r[BUNDLE_ID];
      scanTextForFonts(orderId, needs);
      scanTextForFonts(bundleId, needs);

      if (!orderMap.has(orderId)) orderMap.set(orderId, { orderId, name: "", bundles: new Map<string, any>() });
      const o = orderMap.get(orderId);

      if (!o.name && String(r["Name"] ?? "").trim() !== "") {
        o.name = String(r["Name"]);
        scanTextForFonts(o.name, needs);
      }

      if (!o.bundles.has(bundleId)) {
        o.bundles.set(bundleId, { bundleId, items: { Frame: [], Lens: [], Coating: [], Other: [] }, rx: {} });
      }
      const b = o.bundles.get(bundleId);

      const item = r[LINE_ITEM];
      const qty = Number(r[QTY]) || 1;
      const display = qty !== 1 ? `${item} x${qty}` : String(item);
      b.items[categorize(item)].push(display);
      scanTextForFonts(item, needs);

      // 尽量收集处方字段（兼容不同列名）
      const keysToKeep = [
        "OD_SPH","OD_CYL","OD_AXIS","OD_ADD",
        "OS_SPH","OS_CYL","OS_AXIS","OS_ADD",
        "Sphere OD","Cylinder OD","Axis OD","ADD OD","Add OD",
        "Sphere OS","Cylinder OS","Axis OS","ADD OS","Add OS",
        "PD_OD","PD_OS","Single PD","PD",
        "Prescription Type","Prescription","Lens Type"
      ];
      for (const k of keysToKeep) {
        if (!b.rx[k] && String(r[k] ?? "").trim() !== "") {
          b.rx[k] = r[k];
          scanTextForFonts(r[k], needs);
        }
      }
    }

    const needSC = needs.sc || needs.other;
    const needJP = needs.jp;
    const needKR = needs.kr;
    const needCJK = needs.nonAscii;
    log(`scan done: nonAscii=${needs.nonAscii} jp=${needJP} sc=${needSC} kr=${needKR}`);

    stage = "count-labels";
    let totalLabels = 0;
    for (const [, o] of orderMap) {
      const bundles = Array.from(o.bundles.values());
      totalLabels += mode === "order" ? (bundles.length ? 1 : 0) : bundles.length;
    }
    if (!paging && totalLabels > DEFAULT_LABEL_LIMIT) {
      return new Response(
        `Too many labels (${totalLabels}). Use ?start=0&limit=${DEFAULT_LABEL_LIMIT} to generate in batches (max ${MAX_LABEL_LIMIT}).`,
        { status: 400 }
      );
    }
    if (start >= totalLabels) {
      return new Response(`start out of range (start=${start}, total=${totalLabels})`, { status: 400 });
    }
    log(`labels total=${totalLabels} start=${start} limit=${limit}`);

    stage = "create-pdf";
    const pdf = await PDFDocument.create();

    // 英文/数字：Helvetica（最稳）
    const fontLatin = await pdf.embedFont(StandardFonts.Helvetica);

    let fontJP: any = null;
    let fontSC: any = null;
    let fontKR: any = null;

    if (needCJK) {
      stage = "embed-fonts";
      pdf.registerFontkit(fontkit);
      const fontTasks: Promise<void>[] = [];

      if (needJP) {
        fontTasks.push(
          loadFontBytes(env, FONT_FILES.jp).then(async (bytes) => {
            fontJP = await pdf.embedFont(bytes, { subset: true });
          })
        );
      }
      if (needSC) {
        fontTasks.push(
          loadFontBytes(env, FONT_FILES.sc).then(async (bytes) => {
            fontSC = await pdf.embedFont(bytes, { subset: true });
          })
        );
      }
      if (needKR) {
        fontTasks.push(
          loadFontBytes(env, FONT_FILES.kr).then(async (bytes) => {
            fontKR = await pdf.embedFont(bytes, { subset: true });
          })
        );
      }

      await Promise.all(fontTasks);
      log("fonts embedded");
    }

    const fallback = fontSC || fontJP || fontKR;
    if (needCJK && !fallback) throw new Error("Non-ASCII detected but no CJK font embedded");

    const fonts = { latin: fontLatin, jp: fontJP, sc: fontSC, kr: fontKR, fallback };

    const today = new Date();
    const dateStr = today.toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" });

    stage = "generate-pages";
    let generated = 0;
    let idx = 0;
    outer: for (const [, o] of orderMap) {
      const bundles = Array.from(o.bundles.values()).sort((a: any, b: any) => String(a.bundleId).localeCompare(String(b.bundleId)));
      const targets = mode === "order" ? (bundles.length ? [bundles[0]] : []) : bundles;

      for (const b of targets) {
        if (idx < start) {
          idx++;
          continue;
        }
        if (generated >= limit) break outer;

        const lensText = uniqPreserveOrder(b.items.Lens).join("; ");
        const coatingText = uniqPreserveOrder(b.items.Coating).join("; ");

        const idxLens = parseIndexLensFromText(lensText);
        const thickness = idxLens ? `${idxLens} index lens` : "index lens";

        let coating = "Blue Light Blocking";
        if (coatingText) coating = coatingText.toLowerCase().includes("blue") ? "Blue Light Blocking" : coatingText;

        const presType = String(pick(b.rx, ["Prescription Type", "Prescription", "Lens Type"])) || "Single Vision";

        const od_sph  = fmtTwoDec(pick(b.rx, ["OD_SPH", "Sphere OD", "OD Sphere"]));
        const od_cyl  = fmtTwoDec(pick(b.rx, ["OD_CYL", "Cylinder OD", "OD Cylinder"]));
        const od_axis = fmtAxis(pick(b.rx, ["OD_AXIS", "Axis OD", "OD Axis"]));
        const od_add  = fmtTwoDec(pick(b.rx, ["OD_ADD", "ADD OD", "Add OD"]));

        const os_sph  = fmtTwoDec(pick(b.rx, ["OS_SPH", "Sphere OS", "OS Sphere"]));
        const os_cyl  = fmtTwoDec(pick(b.rx, ["OS_CYL", "Cylinder OS", "OS Cylinder"]));
        const os_axis = fmtAxis(pick(b.rx, ["OS_AXIS", "Axis OS", "OS Axis"]));
        const os_add  = fmtTwoDec(pick(b.rx, ["OS_ADD", "ADD OS", "Add OS"]));

        const [pd_od, pd_os] = getPD(b.rx);

        const page = pdf.addPage([PAGE_W, PAGE_H]);
        drawLabel(page, fonts, {
          backer: mode === "order" ? String(o.orderId) : String(b.bundleId),
          name: o.name || "-",
          presType,
          thickness,
          coating,
          od: { sph: od_sph, cyl: od_cyl, axis: od_axis, add: od_add, pd: pd_od },
          os: { sph: os_sph, cyl: os_cyl, axis: os_axis, add: os_add, pd: pd_os },
          dateStr,
        });

        generated++;
        idx++;
      }
    }
    log(`pages generated=${generated}`);

    stage = "save-pdf";
    const bytes = await pdf.save();
    log(`pdf saved bytes=${bytes.length}`);

    return new Response(bytes, {
      headers: {
        "Content-Type": "application/pdf",
        "Content-Disposition": 'attachment; filename="labels_5x4cm.pdf"',
      },
    });
  } catch (e: any) {
    const msg = e?.stack || e?.message || String(e);
    return new Response(`Labels error (${stage}):
${msg}`, {
      status: 500,
      headers: { "Content-Type": "text/plain; charset=utf-8" },
    });
  }
}

/** =========================
 * Worker 入口
 * ========================= */
export default {
  async fetch(request: Request, env: any) {
    const url = new URL(request.url);

    if (url.pathname === "/health") return new Response("ok");

    if (url.pathname === "/api/aggregate") {
      if (request.method !== "POST") return new Response("Method Not Allowed", { status: 405 });
      try {
        return await handleAggregate(request);
      } catch (e: any) {
        return new Response("Aggregate error:\n" + (e?.stack || e?.message || String(e)), {
          status: 500,
          headers: { "Content-Type": "text/plain; charset=utf-8" },
        });
      }
    }

    if (url.pathname === "/api/labels") {
      if (request.method !== "POST") return new Response("Method Not Allowed", { status: 405 });
      try {
        return await handleLabels(request, env);
      } catch (e: any) {
        return new Response("Labels error:\n" + (e?.stack || e?.message || String(e)), {
          status: 500,
          headers: { "Content-Type": "text/plain; charset=utf-8" },
        });
      }
    }

    // 其他路径走静态站
    if (env.ASSETS) return env.ASSETS.fetch(request);

    return new Response("Not Found", { status: 404 });
  },
};
