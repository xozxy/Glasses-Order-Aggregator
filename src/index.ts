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
const AGGREGATE_DROP_DEFAULT = new Set([
  "Shopify Order ID",
  "Email",
  "Fulfilled at",
  "Currency",
  "Subtotal",
  "Shipping",
  "Taxes",
  "Total",
  "Discount amount",
  "Shipping method",
  "Payment Method",
  "Source",
  "Billing Address",
  "Refund amount",
  "Line Item Vendor",
  "Line Item Fulfillment Status",
  "Line Item Discount",
  "Store",
  "Upload URL",
  "Commission charge",
  "Segment Height OD",
  "Segment Height OS",
  "Ocular Height OD",
  "Ocular Height OS",
  "Prism OD Horizontal",
  "Prism OD Horizontal Base Direction",
  "Prism OD Vertical",
  "Prism OD Vertical Base Direction",
  "Prism OS Horizontal",
  "Prism OS Horizontal Base Direction",
  "Prism OS Vertical",
  "Prism OS Vertical Base Direction",
  "Lens Notes",
]);

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

function parseIndexLens(value: any): string {
  const s = String(value ?? "").trim();
  if (!s) return "";

  // Prefer refractive-index-like values (1.x / 1.xx) with numeric boundaries.
  const m = s.match(/(?:^|[^0-9])1\.(\d{1,2})(?!\d)/i);
  if (!m) return "";

  const n = Number(`1.${m[1]}`);
  if (!Number.isFinite(n)) return "";
  return n.toFixed(2);
}

function fmtTwoDec(x: any): string {
  if (x === null || x === undefined) return "";
  if (typeof x === "string") {
    const s = x.trim();
    if (s === "" || s === "/" || s.toLowerCase() === "nan") return "";
    const n = Number(s);
    return Number.isFinite(n) ? n.toFixed(2) : "";
  }
  if (typeof x === "number") {
    return Number.isFinite(x) ? x.toFixed(2) : "";
  }
  const n = Number(x);
  return Number.isFinite(n) ? n.toFixed(2) : "";
}
function fmtAxis(x: any): string {
  if (x === null || x === undefined) return "";
  if (typeof x === "string") {
    const s = x.trim();
    if (s === "" || s === "/" || s.toLowerCase() === "nan") return "";
    const n = Number(s);
    return Number.isFinite(n) ? String(Math.round(n)) : "";
  }
  if (typeof x === "number") {
    return Number.isFinite(x) ? String(Math.round(x)) : "";
  }
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

function splitCoatingText(text: string) {
  const parts = text
    .split(/[;,/]+/)
    .map((s) => s.trim())
    .filter(Boolean);
  return parts;
}

function truncateToFit(fonts: any, text: string, size: number, maxWidth: number) {
  const ellipsis = "...";
  if (measureTextRuns(fonts, text, size) <= maxWidth) return text;
  const ellWidth = measureTextRuns(fonts, ellipsis, size);
  if (ellWidth > maxWidth) return ellipsis;
  const chars = Array.from(text);
  let lo = 0;
  let hi = chars.length;
  while (lo < hi) {
    const mid = Math.floor((lo + hi) / 2);
    const candidate = chars.slice(0, mid).join("") + ellipsis;
    if (measureTextRuns(fonts, candidate, size) <= maxWidth) {
      lo = mid + 1;
    } else {
      hi = mid;
    }
  }
  const cut = Math.max(0, lo - 1);
  return chars.slice(0, cut).join("") + ellipsis;
}

function fitLine(fonts: any, text: string, size: number, maxWidth: number, minSize = 4) {
  let curSize = size;
  while (curSize >= minSize) {
    if (measureTextRuns(fonts, text, curSize) <= maxWidth) return { text, size: curSize };
    curSize -= 1;
  }
  return { text: truncateToFit(fonts, text, minSize, maxWidth), size: minSize };
}

function buildCoatingLines(fonts: any, raw: string, size: number, maxWidth: number, maxLines = 2) {
  const text = normalizeText(raw);
  if (!text) return [{ text: "-", size }];
  const parts = splitCoatingText(text);
  if (!parts.length) return [{ text: "-", size }];

  const lines: string[] = [];
  let current = "";
  let overflow = false;

  for (let i = 0; i < parts.length; i++) {
    const part = parts[i];
    const candidate = current ? `${current}; ${part}` : part;
    if (measureTextRuns(fonts, candidate, size) <= maxWidth || current === "") {
      current = candidate;
    } else {
      lines.push(current);
      current = part;
      if (lines.length === maxLines - 1) {
        if (i + 1 < parts.length) {
          current = `${current}; ${parts.slice(i + 1).join("; ")}`;
          overflow = true;
          break;
        }
      }
    }
  }
  if (current) lines.push(current);
  if (lines.length > maxLines) {
    lines.length = maxLines;
    overflow = true;
  }
  if (overflow && lines.length) {
    lines[lines.length - 1] = `${lines[lines.length - 1]}...`;
  }
  return lines.map((line) => fitLine(fonts, line, size, maxWidth));
}

function buildWrappedLines(fonts: any, raw: string, size: number, maxWidth: number, maxLines = 2) {
  const text = normalizeText(raw);
  if (!text) return [{ text: "-", size }];
  if (measureTextRuns(fonts, text, size) <= maxWidth) return [{ text, size }];

  const words = text.split(/\s+/).filter(Boolean);
  if (!words.length) return [{ text: "-", size }];

  const lines: string[] = [];
  let current = "";
  let overflow = false;

  for (let i = 0; i < words.length; i++) {
    const w = words[i];
    const candidate = current ? `${current} ${w}` : w;
    if (!current || measureTextRuns(fonts, candidate, size) <= maxWidth) {
      current = candidate;
      continue;
    }
    lines.push(current);
    current = w;
    if (lines.length >= maxLines - 1) {
      if (i + 1 < words.length) {
        current = `${current} ${words.slice(i + 1).join(" ")}`;
        overflow = true;
      }
      break;
    }
  }
  if (current) lines.push(current);
  if (lines.length > maxLines) {
    lines.length = maxLines;
    overflow = true;
  }
  if (overflow && lines.length) {
    lines[lines.length - 1] = `${lines[lines.length - 1]}...`;
  }
  return lines.map((line) => fitLine(fonts, line, size, maxWidth));
}

function buildThicknessText(indexLensValue: any, fallbackText: any) {
  const indexLensRaw = normalizeText(indexLensValue);
  if (indexLensRaw) return indexLensRaw;
  const idx = parseIndexLens(fallbackText);
  return idx ? `${idx} index lens` : "index lens";
}

const RX_KEYS = {
  od_sph: ["OD SPH", "OD_SPH", "Sphere OD", "OD Sphere"],
  od_cyl: ["OD CYL", "OD_CYL", "Cylinder OD", "OD Cylinder"],
  od_axis: ["OD Axis", "OD_AXIS", "Axis OD", "OD Axis"],
  od_add: ["OD ADD", "OD_ADD", "ADD OD", "Add OD"],
  os_sph: ["OS SPH", "OS_SPH", "Sphere OS", "OS Sphere"],
  os_cyl: ["OS CYL", "OS_CYL", "Cylinder OS", "OS Cylinder"],
  os_axis: ["OS Axis", "OS_AXIS", "Axis OS", "OS Axis"],
  os_add: ["OS ADD", "OS_ADD", "ADD OS", "Add OS"],
  pd_od: ["PD OD", "PD_OD", "OD PD", "Pupillary Distance OD"],
  pd_os: ["PD OS", "PD_OS", "OS PD", "Pupillary Distance OS"],
  pres_type: ["Prescription Type", "Prescription", "Lens Type"],
};

function pickFromBundle(row: Record<string, any>, keys: string[], bundleIndex: number) {
  for (const key of keys) {
    const bundleKey = `${key} (Bundle ${bundleIndex})`;
    const v = row[bundleKey];
    if (v !== null && v !== undefined && String(v).trim() !== "") return v;
  }
  return pick(row, keys);
}

function isAggregatedHeaders(headers: string[]) {
  return headers.includes("Bundle Count") || headers.some((h) => /\(Bundle\s+\d+\)/i.test(h));
}

function getBundleIndices(headers: string[]) {
  const set = new Set<number>();
  for (const h of headers) {
    const m = h.match(/\(Bundle\s+(\d+)\)/i);
    if (m) set.add(Number(m[1]));
  }
  return Array.from(set).filter((n) => Number.isFinite(n) && n > 0).sort((a, b) => a - b);
}

function buildBundleColumnsMap(headers: string[]) {
  const map = new Map<number, string[]>();
  for (const h of headers) {
    const m = h.match(/\(Bundle\s+(\d+)\)/i);
    if (!m) continue;
    const idx = Number(m[1]);
    if (!Number.isFinite(idx) || idx <= 0) continue;
    if (!map.has(idx)) map.set(idx, []);
    map.get(idx)!.push(h);
  }
  return map;
}

function inferBundleCount(row: Record<string, any>, bundleIndices: number[], bundleColumnsMap: Map<number, string[]>) {
  const bc = Number(row["Bundle Count"]);
  if (Number.isFinite(bc) && bc > 0) return Math.floor(bc);
  let maxIdx = 0;
  for (const idx of bundleIndices) {
    const cols = bundleColumnsMap.get(idx) || [];
    for (const c of cols) {
      const v = row[c];
      if (v !== null && v !== undefined && String(v).trim() !== "") {
        maxIdx = Math.max(maxIdx, idx);
        break;
      }
    }
  }
  return maxIdx;
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
  lensGroup: string;
  coating: string;
  od: { sph: string; cyl: string; axis: string; add: string; pd: string };
  os: { sph: string; cyl: string; axis: string; add: string; pd: string };
  dateStr: string;
}) {
  const black = rgb(0, 0, 0);
  const W = PAGE_W;
  const H = PAGE_H;
  const titleSize = 6;
  const bodySize = 5;

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
  drawTopLeft(x(0.03), top(0.92), `Order Number: ${data.backer || ""}`, titleSize);
  drawTopLeft(x(0.03), top(0.85), `Name: ${data.name || ""}`, titleSize);

  hlineTop(top(0.72), 1);

  // Prescription / Thickness / Coating
  const rightValueX = x(0.45);
  const rightMaxWidth = x(0.97) - rightValueX;

  drawTopLeft(x(0.03), top(0.68), "Prescription:", bodySize);
  drawTopLeft(rightValueX, top(0.68), data.presType || "-", bodySize);

  drawTopLeft(x(0.03), top(0.60), "Thickness:", bodySize);
  const thicknessLines = buildWrappedLines(fonts, data.thickness || "index lens", bodySize, rightMaxWidth, 2);
  const thicknessTop = top(0.60);
  const thicknessGap = bodySize + 1;
  thicknessLines.forEach((line, i) => {
    drawTopLeft(rightValueX, thicknessTop + i * thicknessGap, line.text, line.size);
  });

  drawTopLeft(x(0.03), top(0.50), "Lens Group:", bodySize);
  const lensGroupLines = buildWrappedLines(fonts, data.lensGroup || "-", bodySize, rightMaxWidth, 2);
  const lensGroupTop = top(0.50);
  const coatingGap = bodySize + 1;
  lensGroupLines.forEach((line, i) => {
    drawTopLeft(rightValueX, lensGroupTop + i * coatingGap, line.text, line.size);
  });

  drawTopLeft(x(0.03), top(0.40), "Coating:", bodySize);
  const coatingLines = buildCoatingLines(fonts, data.coating || "-", bodySize, rightMaxWidth, 2);
  const coatingTop = top(0.40);
  coatingLines.forEach((line, i) => {
    drawTopLeft(rightValueX, coatingTop + i * coatingGap, line.text, line.size);
  });

  hlineTop(top(0.30), 0.8);

  // 表格区域
  const colX = [0.16, 0.32, 0.48, 0.64, 0.8].map(x);
  ["sph", "cyl", "axis", "add", "pd"].forEach((h, i) => drawTopCenter(colX[i], top(0.26), h, bodySize));
  drawTopCenter(x(0.06), top(0.20), "od", bodySize);
  drawTopCenter(x(0.06), top(0.14), "os", bodySize);

  const vOD = [data.od.sph, data.od.cyl, data.od.axis, data.od.add, data.od.pd].map((v) => v || "-");
  const vOS = [data.os.sph, data.os.cyl, data.os.axis, data.os.add, data.os.pd].map((v) => v || "-");

  vOD.forEach((v, i) => drawTopCenter(colX[i], top(0.20), v, bodySize));
  vOS.forEach((v, i) => drawTopCenter(colX[i], top(0.14), v, bodySize));

  hlineTop(top(0.09), 0.8);
  drawTopLeft(x(0.03), top(0.05), data.dateStr || "", bodySize);
}

/** =========================
 * /api/aggregate：订单聚合（一单一行 CSV）
 * ========================= */
async function handleAggregate(request: Request) {
  const url = new URL(request.url);
  const format = (url.searchParams.get("format") || "csv").toLowerCase();
  const dropMode = (url.searchParams.get("drop") || "default").toLowerCase();
  const dropColsParam = url.searchParams.get("drop_cols") || "";

  const dropSet = new Set<string>();
  const applyDefault = dropMode !== "none";
  if (applyDefault) {
    for (const c of AGGREGATE_DROP_DEFAULT) dropSet.add(c);
  }
  if (dropMode === "custom") {
    dropColsParam
      .split(",")
      .map((s) => s.trim())
      .filter(Boolean)
      .forEach((c) => dropSet.add(c));
  }

  const shouldDropColumn = (name: string) => {
    if (dropSet.has(name)) return true;
    const m = name.match(/^(.*) \\(Bundle \\d+\\)$/);
    if (m && dropSet.has(m[1])) return true;
    return false;
  };

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
  const rxColumns = allColumns.filter((c) => (rxRe.test(c) || c === "Lens Notes") && !shouldDropColumn(c));

  const lineLevelCols = new Set([LINE_ITEM, QTY, "Line Item Price"]);
  const idCols = new Set([ORDER_ID, BUNDLE_ID]);
  const rxSet = new Set(rxColumns);
  const orderLevelCols = allColumns.filter((c) => !lineLevelCols.has(c) && !rxSet.has(c) && !idCols.has(c) && !shouldDropColumn(c));

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
      const bundleKeys = {
        bundleId: `Bundle ID (Bundle ${i})`,
        frame: `Frame Items (Bundle ${i})`,
        lens: `Lens Items (Bundle ${i})`,
        coating: `Coating Items (Bundle ${i})`,
        other: `Other Items (Bundle ${i})`,
      };
      if (!shouldDropColumn(bundleKeys.bundleId)) rowOut[bundleKeys.bundleId] = b.bundleId;
      if (!shouldDropColumn(bundleKeys.frame)) rowOut[bundleKeys.frame] = uniqPreserveOrder(b.items.Frame).join("; ");
      if (!shouldDropColumn(bundleKeys.lens)) rowOut[bundleKeys.lens] = uniqPreserveOrder(b.items.Lens).join("; ");
      if (!shouldDropColumn(bundleKeys.coating)) rowOut[bundleKeys.coating] = uniqPreserveOrder(b.items.Coating).join("; ");
      if (!shouldDropColumn(bundleKeys.other)) rowOut[bundleKeys.other] = uniqPreserveOrder(b.items.Other).join("; ");
      for (const c of rxColumns) {
        const key = `${c} (Bundle ${i})`;
        if (!shouldDropColumn(key)) rowOut[key] = b.rx[c] || "";
      }
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
  const columns = [...baseCols, ...bundleCols].filter((c) => !shouldDropColumn(c));

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

    const headers = Object.keys(rows[0] || {});
    const aggregated = isAggregatedHeaders(headers);
    let bundleIndices: number[] = [];
    let bundleColumnsMap = new Map<number, string[]>();

    if (aggregated) {
      if (!(ORDER_ID in rows[0])) return new Response(`Missing required column: ${ORDER_ID}`, { status: 400 });
      bundleIndices = getBundleIndices(headers);
      bundleColumnsMap = buildBundleColumnsMap(headers);
      if (!bundleIndices.length) {
        return new Response("Aggregated CSV detected but no Bundle columns found. Please upload raw CSV or orders_one_row.csv.", { status: 400 });
      }
      if (mode === "order" && !headers.some((h) => /\(Bundle\s+1\)/i.test(h))) {
        return new Response("Aggregated CSV missing Bundle 1 columns for order mode.", { status: 400 });
      }
    } else {
      for (const col of [ORDER_ID, BUNDLE_ID, LINE_ITEM, QTY]) {
        if (!(col in rows[0])) return new Response(`Missing required column: ${col}`, { status: 400 });
      }
    }
    log(`csv parsed rows=${rows.length} aggregated=${aggregated}`);

    stage = "scan-text";
    const needs: FontNeeds = { jp: false, kr: false, sc: false, other: false, nonAscii: false };

    // === 组织 order/bundle ===
    const orderMap = new Map<string, any>();

    if (!aggregated) {
      const keysToKeep = [
        ...RX_KEYS.od_sph,
        ...RX_KEYS.od_cyl,
        ...RX_KEYS.od_axis,
        ...RX_KEYS.od_add,
        ...RX_KEYS.os_sph,
        ...RX_KEYS.os_cyl,
        ...RX_KEYS.os_axis,
        ...RX_KEYS.os_add,
        ...RX_KEYS.pd_od,
        ...RX_KEYS.pd_os,
        ...RX_KEYS.pres_type,
        "Index Lens",
        "Lens Group",
        "Single PD",
        "PD",
      ];

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

        for (const k of keysToKeep) {
          if (!b.rx[k] && String(r[k] ?? "").trim() !== "") {
            b.rx[k] = r[k];
            scanTextForFonts(r[k], needs);
          }
        }
      }
    } else {
      for (const r of rows) {
        const orderId = r[ORDER_ID];
        const name = String(r["Name"] ?? "");
        scanTextForFonts(orderId, needs);
        if (name.trim()) scanTextForFonts(name, needs);

        const bundleCount = inferBundleCount(r, bundleIndices, bundleColumnsMap);
        if (mode === "order" && bundleCount < 1) continue;
        const bundleList = mode === "order" ? [1] : Array.from({ length: bundleCount }, (_, i) => i + 1);

        for (const i of bundleList) {
          const bundleId = pickFromBundle(r, ["Bundle ID"], i);
          if (bundleId) scanTextForFonts(bundleId, needs);

          const lensText = String(pickFromBundle(r, ["Lens Items"], i) ?? "");
          const coatingText = String(pickFromBundle(r, ["Coating Items"], i) ?? "");
          const indexLensText = String(pickFromBundle(r, ["Index Lens"], i) ?? "");
          const lensGroup = String(pickFromBundle(r, ["Lens Group"], i) ?? "");
          scanTextForFonts(lensText, needs);
          scanTextForFonts(coatingText, needs);
          scanTextForFonts(indexLensText, needs);
          scanTextForFonts(lensGroup, needs);

          const presType = String(pickFromBundle(r, RX_KEYS.pres_type, i) ?? "");
          scanTextForFonts(presType, needs);

          scanTextForFonts(pickFromBundle(r, RX_KEYS.od_sph, i), needs);
          scanTextForFonts(pickFromBundle(r, RX_KEYS.od_cyl, i), needs);
          scanTextForFonts(pickFromBundle(r, RX_KEYS.od_axis, i), needs);
          scanTextForFonts(pickFromBundle(r, RX_KEYS.od_add, i), needs);
          scanTextForFonts(pickFromBundle(r, RX_KEYS.os_sph, i), needs);
          scanTextForFonts(pickFromBundle(r, RX_KEYS.os_cyl, i), needs);
          scanTextForFonts(pickFromBundle(r, RX_KEYS.os_axis, i), needs);
          scanTextForFonts(pickFromBundle(r, RX_KEYS.os_add, i), needs);
          scanTextForFonts(pickFromBundle(r, RX_KEYS.pd_od, i), needs);
          scanTextForFonts(pickFromBundle(r, RX_KEYS.pd_os, i), needs);
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
    if (!aggregated) {
      for (const [, o] of orderMap) {
        const bundles = Array.from(o.bundles.values());
        totalLabels += mode === "order" ? (bundles.length ? 1 : 0) : bundles.length;
      }
    } else {
      for (const r of rows) {
        const bundleCount = inferBundleCount(r, bundleIndices, bundleColumnsMap);
        if (mode === "order") {
          if (bundleCount >= 1) totalLabels += 1;
        } else {
          if (bundleCount >= 1) totalLabels += bundleCount;
        }
      }
      if (totalLabels === 0) {
        return new Response("No bundles found in aggregated CSV for the selected mode.", { status: 400 });
      }
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

    if (!aggregated) {
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
          const indexLensText = pick(b.rx, ["Index Lens"]);
          const thickness = buildThicknessText(indexLensText, lensText);
          const lensGroup = normalizeText(pick(b.rx, ["Lens Group"])) || "-";

          let coating = "Blue Light Blocking";
          if (coatingText) coating = coatingText.toLowerCase().includes("blue") ? "Blue Light Blocking" : coatingText;

          const presType = String(pick(b.rx, RX_KEYS.pres_type)) || "Single Vision";

          const od_sph  = fmtTwoDec(pick(b.rx, RX_KEYS.od_sph));
          const od_cyl  = fmtTwoDec(pick(b.rx, RX_KEYS.od_cyl));
          const od_axis = fmtAxis(pick(b.rx, RX_KEYS.od_axis));
          const od_add  = fmtTwoDec(pick(b.rx, RX_KEYS.od_add));

          const os_sph  = fmtTwoDec(pick(b.rx, RX_KEYS.os_sph));
          const os_cyl  = fmtTwoDec(pick(b.rx, RX_KEYS.os_cyl));
          const os_axis = fmtAxis(pick(b.rx, RX_KEYS.os_axis));
          const os_add  = fmtTwoDec(pick(b.rx, RX_KEYS.os_add));

          const [pd_od, pd_os] = getPD(b.rx);

          const page = pdf.addPage([PAGE_W, PAGE_H]);
          drawLabel(page, fonts, {
            backer: mode === "order" ? String(o.orderId) : String(b.bundleId),
            name: o.name || "-",
            presType,
            thickness,
            lensGroup,
            coating,
            od: { sph: od_sph, cyl: od_cyl, axis: od_axis, add: od_add, pd: pd_od },
            os: { sph: os_sph, cyl: os_cyl, axis: os_axis, add: os_add, pd: pd_os },
            dateStr,
          });

          generated++;
          idx++;
        }
      }
    } else {
      outer: for (const r of rows) {
        const orderId = r[ORDER_ID];
        const name = String(r["Name"] ?? "");
        const bundleCount = inferBundleCount(r, bundleIndices, bundleColumnsMap);
        if (mode === "order" && bundleCount < 1) continue;
        const bundleList = mode === "order" ? [1] : Array.from({ length: bundleCount }, (_, i) => i + 1);

        for (const i of bundleList) {
          if (idx < start) {
            idx++;
            continue;
          }
          if (generated >= limit) break outer;

          const lensText = String(pickFromBundle(r, ["Lens Items"], i) ?? "");
          const coatingText = String(pickFromBundle(r, ["Coating Items"], i) ?? "");
          const indexLensText = pickFromBundle(r, ["Index Lens"], i);
          const thickness = buildThicknessText(indexLensText, lensText);
          const lensGroup = normalizeText(pickFromBundle(r, ["Lens Group"], i)) || "-";

          let coating = "Blue Light Blocking";
          if (coatingText) coating = coatingText.toLowerCase().includes("blue") ? "Blue Light Blocking" : coatingText;

          const presType = String(pickFromBundle(r, RX_KEYS.pres_type, i)) || "Single Vision";

          const od_sph  = fmtTwoDec(pickFromBundle(r, RX_KEYS.od_sph, i));
          const od_cyl  = fmtTwoDec(pickFromBundle(r, RX_KEYS.od_cyl, i));
          const od_axis = fmtAxis(pickFromBundle(r, RX_KEYS.od_axis, i));
          const od_add  = fmtTwoDec(pickFromBundle(r, RX_KEYS.od_add, i));

          const os_sph  = fmtTwoDec(pickFromBundle(r, RX_KEYS.os_sph, i));
          const os_cyl  = fmtTwoDec(pickFromBundle(r, RX_KEYS.os_cyl, i));
          const os_axis = fmtAxis(pickFromBundle(r, RX_KEYS.os_axis, i));
          const os_add  = fmtTwoDec(pickFromBundle(r, RX_KEYS.os_add, i));

          const pd_od = fmtTwoDec(pickFromBundle(r, RX_KEYS.pd_od, i));
          const pd_os = fmtTwoDec(pickFromBundle(r, RX_KEYS.pd_os, i));

          const bundleId = pickFromBundle(r, ["Bundle ID"], i);
          const page = pdf.addPage([PAGE_W, PAGE_H]);
          drawLabel(page, fonts, {
            backer: mode === "order" ? String(orderId) : String(bundleId || `${orderId}-${i}`),
            name: name || "-",
            presType,
            thickness,
            lensGroup,
            coating,
            od: { sph: od_sph, cyl: od_cyl, axis: od_axis, add: od_add, pd: pd_od },
            os: { sph: os_sph, cyl: os_cyl, axis: os_axis, add: os_add, pd: pd_os },
            dateStr,
          });

          generated++;
          idx++;
        }
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
