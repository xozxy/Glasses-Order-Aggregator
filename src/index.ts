import Papa from "papaparse";
import { PDFDocument, rgb } from "pdf-lib";
import fontkit from "@pdf-lib/fontkit";

/** =========================
 * 配置：字体文件路径（在 public/ 下面）
 * ========================= */
const FONT_PATH = "/fonts/NotoSansSC.ttf"; // 可改：/fonts/NotoSansJP.ttf 等

/** =========================
 * 通用字段
 * ========================= */
const ORDER_ID = "Order ID";
const BUNDLE_ID = "Bundle ID";
const LINE_ITEM = "Line Item";
const QTY = "Quantity";

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
 * 读取字体（关键）
 * ========================= */
async function loadFontBytes(env: any): Promise<Uint8Array> {
  if (!env?.ASSETS?.fetch) {
    throw new Error("ASSETS binding missing. Check wrangler assets binding name is 'ASSETS'.");
  }
  const res = await env.ASSETS.fetch(new Request("http://local" + FONT_PATH));
  if (!res.ok) {
    throw new Error(`Font not found: ${FONT_PATH} (status ${res.status})`);
  }
  const ab = await res.arrayBuffer();
  return new Uint8Array(ab);
}

/** =========================
 * /api/aggregate（保持你现有逻辑：订单一行）
 * ========================= */
async function handleAggregate(request: Request) {
  const url = new URL(request.url);
  const format = (url.searchParams.get("format") || "csv").toLowerCase();

  const form = await request.formData();
  const fileAny = form.get("file");
  if (!fileAny || typeof fileAny === "string") {
    return new Response("Missing/invalid file field 'file'", { status: 400 });
  }
  const text = await (fileAny as File).text();

  const parsed = Papa.parse<Record<string, any>>(text, { header: true, skipEmptyLines: true });
  if (parsed.errors?.length) {
    return new Response("CSV parse error: " + JSON.stringify(parsed.errors.slice(0, 3)), { status: 400 });
  }
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

    if (!o.bundles.has(bundleId)) {
      o.bundles.set(bundleId, { bundleId, items: { Frame: [], Lens: [], Coating: [], Other: [] }, rx: {} });
    }
    const b = o.bundles.get(bundleId);

    const item = r[LINE_ITEM];
    const qty = Number(r[QTY]) || 1;
    const display = qty !== 1 ? `${item} x${qty}` : String(item);
    const cat = categorize(item);
    b.items[cat].push(display);

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
 * /api/labels：PDF Label（关键：Unicode 字体）
 * ========================= */
const CM_TO_PT = 28.3464566929;
const PAGE_W = 5 * CM_TO_PT;
const PAGE_H = 4 * CM_TO_PT;

function drawLabel(page: any, font: any, data: {
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

  // ==== 采用“从上往下”的坐标系（更接近 matplotlib） ====
  const W = PAGE_W;
  const H = PAGE_H;

  const marginX = 6;   // 左右边距（pt）
  const lineX1 = marginX;
  const lineX2 = W - marginX;

  // top-based 文本绘制：x为左边，topY为距离顶部的像素
  const drawTextTop = (x: number, topY: number, text: string, size: number, bold = false) => {
    const t = text ?? "";
    const baselineY = H - topY - size; // baseline 修正
    page.drawText(t, { x, y: baselineY, size, font, color: black });
  };

  // 横线：topY 为距离顶部
  const hlineTop = (topY: number, thickness = 1) => {
    const y = H - topY;
    page.drawLine({
      start: { x: lineX1, y },
      end: { x: lineX2, y },
      thickness,
      color: black,
    });
  };

  // ======== 布局参数（对齐你模板图） ========
  // 顶部两行
  drawTextTop(marginX, 10, `Backer Number: ${data.backer || ""}`, 8);
  drawTextTop(marginX, 24, `Name: ${data.name || ""}`, 8);

  // 第一条线
  hlineTop(34, 1);

  // Prescription/Thickness/Coating 三行（左右两列）
  const leftLabelX = marginX;
  const rightValueX = marginX + 92;

  drawTextTop(leftLabelX, 48, "Prescription:", 7);
  drawTextTop(rightValueX, 48, data.presType || "-", 7);

  drawTextTop(leftLabelX, 62, "Thickness:", 7);
  drawTextTop(rightValueX, 62, data.thickness || "index lens", 7);

  drawTextTop(leftLabelX, 76, "Coating:", 7);
  drawTextTop(rightValueX, 76, data.coating || "-", 7);

  // 第二条线
  hlineTop(86, 0.8);

  // 表头
  const headerTopY = 98;
  const odTopY = 112;
  const osTopY = 126;

  const colX = [42, 70, 98, 126, 154];
  const headers = ["sph", "cyl", "axis", "add", "pd"];

  headers.forEach((h, i) => drawTextTop(colX[i] - 6, headerTopY, h, 7));

  // od/os 标记
  drawTextTop(marginX + 10, odTopY, "od", 7);
  drawTextTop(marginX + 10, osTopY, "os", 7);

  const vOD = [data.od.sph, data.od.cyl, data.od.axis, data.od.add, data.od.pd].map(v => v || "-");
  const vOS = [data.os.sph, data.os.cyl, data.os.axis, data.os.add, data.os.pd].map(v => v || "-");

  vOD.forEach((v, i) => drawTextTop(colX[i] - 6, odTopY, v, 7));
  vOS.forEach((v, i) => drawTextTop(colX[i] - 6, osTopY, v, 7));

  // 第三条线
  hlineTop(140, 0.8);

  // 日期
  drawTextTop(marginX, 154, data.dateStr || "", 7);
}


async function handleLabels(request: Request, env: any) {
  const url = new URL(request.url);
  const mode = (url.searchParams.get("mode") || "bundle").toLowerCase(); // bundle | order

  const form = await request.formData();
  const fileAny = form.get("file");
  if (!fileAny || typeof fileAny === "string") {
    return new Response("Missing/invalid file field 'file'", { status: 400 });
  }
  const csvText = await (fileAny as File).text();

  const parsed = Papa.parse<Record<string, any>>(csvText, { header: true, skipEmptyLines: true });
  if (parsed.errors?.length) {
    return new Response("CSV parse error: " + JSON.stringify(parsed.errors.slice(0, 3)), { status: 400 });
  }
  const rows = parsed.data || [];
  if (!rows.length) return new Response("Empty CSV", { status: 400 });

  for (const col of [ORDER_ID, BUNDLE_ID, LINE_ITEM, QTY]) {
    if (!(col in rows[0])) return new Response(`Missing required column: ${col}`, { status: 400 });
  }

  // ✅ 读取 Unicode 字体
  const fontBytes = await loadFontBytes(env);

  const pdf = await PDFDocument.create();
  pdf.registerFontkit(fontkit);
  const font = await pdf.embedFont(fontBytes, { subset: true });

  const today = new Date();
  const dateStr = today.toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" });

  // 按 order/bundle 组织数据
  const orderMap = new Map<string, any>();

  for (const r of rows) {
    const orderId = r[ORDER_ID];
    const bundleId = r[BUNDLE_ID];

    if (!orderMap.has(orderId)) {
      orderMap.set(orderId, { orderId, name: "", bundles: new Map<string, any>() });
    }
    const o = orderMap.get(orderId);

    if (!o.name && String(r["Name"] ?? "").trim() !== "") o.name = String(r["Name"]);

    if (!o.bundles.has(bundleId)) {
      o.bundles.set(bundleId, {
        bundleId,
        items: { Frame: [], Lens: [], Coating: [], Other: [] },
        rx: {},
      });
    }
    const b = o.bundles.get(bundleId);

    const item = r[LINE_ITEM];
    const qty = Number(r[QTY]) || 1;
    const display = qty !== 1 ? `${item} x${qty}` : String(item);
    const cat = categorize(item);
    b.items[cat].push(display);

    // rx 字段尽可能收集
    const keysToKeep = [
      "OD_SPH","OD_CYL","OD_AXIS","OD_ADD",
      "OS_SPH","OS_CYL","OS_AXIS","OS_ADD",
      "Sphere OD","Cylinder OD","Axis OD","ADD OD","Add OD",
      "Sphere OS","Cylinder OS","Axis OS","ADD OS","Add OS",
      "PD_OD","PD_OS","Single PD","PD",
      "Prescription Type","Prescription","Lens Type"
    ];
    for (const k of keysToKeep) {
      if (!b.rx[k] && String(r[k] ?? "").trim() !== "") b.rx[k] = r[k];
    }
  }

  // 生成 label 页
  for (const [, o] of orderMap) {
    const bundles = Array.from(o.bundles.values()).sort((a: any, b: any) => String(a.bundleId).localeCompare(String(b.bundleId)));
    const targets = (mode === "order") ? (bundles.length ? [bundles[0]] : []) : bundles;

    for (const b of targets) {
      const lensText = uniqPreserveOrder(b.items.Lens).join("; ");
      const coatingText = uniqPreserveOrder(b.items.Coating).join("; ");

      const idx = parseIndexLensFromText(lensText);
      const thickness = idx ? `${idx} index lens` : "index lens";

      let coating = "Blue Light Blocking";
      if (coatingText) coating = coatingText.toLowerCase().includes("blue") ? "Blue Light Blocking" : coatingText;

      const presType = String(pick(b.rx, ["Prescription Type","Prescription","Lens Type"])) || "Single Vision";

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
      drawLabel(page, font, {
        backer: mode === "order" ? String(o.orderId) : String(b.bundleId),
        name: o.name || "-",
        presType,
        thickness,
        coating,
        od: { sph: od_sph, cyl: od_cyl, axis: od_axis, add: od_add, pd: pd_od },
        os: { sph: os_sph, cyl: os_cyl, axis: os_axis, add: os_add, pd: pd_os },
        dateStr,
      });
    }
  }

  const bytes = await pdf.save();
  return new Response(bytes, {
    headers: {
      "Content-Type": "application/pdf",
      "Content-Disposition": 'attachment; filename="labels_5x4cm.pdf"',
    },
  });
}

/** =========================
 * 主入口
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

    // 静态页面
    if (env.ASSETS) return env.ASSETS.fetch(request);

    return new Response("Not Found", { status: 404 });
  },
};
