import Papa from "papaparse";
import { PDFDocument, StandardFonts, rgb } from "pdf-lib";

/** =========================
 *  通用工具
 *  ========================= */
const ORDER_ID = "Order ID";
const BUNDLE_ID = "Bundle ID";
const LINE_ITEM = "Line Item";
const QTY = "Quantity";

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

function firstNonEmpty(values: any[]) {
  for (const v of values) {
    if (v !== null && v !== undefined && String(v).trim() !== "") return v;
  }
  return "";
}

function toNumberMaybe(x: any) {
  const n = Number(x);
  return Number.isFinite(n) ? n : null;
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

/** 从列名候选里取第一个非空 */
function pick(row: Record<string, any>, keys: string[]) {
  for (const k of keys) {
    const v = row[k];
    if (v !== null && v !== undefined && String(v).trim() !== "") return v;
  }
  return "";
}

/** PD 优先 OD/OS；否则用 Single PD/2 */
function getPD(row: Record<string, any>): [string, string] {
  const pd_od = fmtTwoDec(pick(row, ["PD_OD", "PD OD", "Pupillary Distance OD", "Pupillary distance OD", "OD PD"]));
  const pd_os = fmtTwoDec(pick(row, ["PD_OS", "PD OS", "Pupillary Distance OS", "Pupillary distance OS", "OS PD"]));
  if (pd_od || pd_os) return [pd_od, pd_os];

  const single = pick(row, ["Single PD", "Single_PD", "Pupillary Distance", "PD"]);
  const n = Number(single);
  if (Number.isFinite(n) && n > 0) {
    const half = n / 2;
    return [half.toFixed(2), half.toFixed(2)];
  }
  return ["", ""];
}

/** =========================
 *  /api/aggregate：订单一行
 *  ========================= */
async function handleAggregate(request: Request) {
  const url = new URL(request.url);
  const format = (url.searchParams.get("format") || "csv").toLowerCase();

  const form = await request.formData();
  const file = form.get("file") as File | null;
  if (!file) return new Response("Missing file field 'file'", { status: 400 });

  const text = await file.text();
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

  // RX 字段（尽量泛化）
  const rxRe = /\b(OD|OS|PD|Prism|ADD|Axis|Cylinder|Sphere|Pupillary|base)\b/i;
  const rxColumns = allColumns.filter((c) => rxRe.test(c) || c === "Lens Notes");

  const lineLevelCols = new Set([LINE_ITEM, QTY, "Line Item Price"]);
  const idCols = new Set([ORDER_ID, BUNDLE_ID]);
  const rxSet = new Set(rxColumns);
  const orderLevelCols = allColumns.filter((c) => !lineLevelCols.has(c) && !rxSet.has(c) && !idCols.has(c));

  // orderId -> { orderFields, bundles }
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

  // 输出宽表
  const out: any[] = [];
  let maxBundleCount = 1;

  for (const [, o] of orderMap) {
    const bundles = Array.from(o.bundles.values()).sort((a: any, b: any) => String(a.bundleId).localeCompare(String(b.bundleId)));
    maxBundleCount = Math.max(maxBundleCount, bundles.length);

    const rowOut: any = {
      [ORDER_ID]: o.orderId,
      "Bundle Count": bundles.length,
      ...o.orderFields,
    };

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

  // 列顺序
  const baseCols = [ORDER_ID, "Bundle Count", ...orderLevelCols];
  const bundleCols: string[] = [];
  for (let i = 1; i <= maxBundleCount; i++) {
    bundleCols.push(
      `Bundle ID (Bundle ${i})`,
      `Frame Items (Bundle ${i})`,
      `Lens Items (Bundle ${i})`,
      `Coating Items (Bundle ${i})`,
      `Other Items (Bundle ${i})`
    );
    for (const c of rxColumns) bundleCols.push(`${c} (Bundle ${i})`);
  }
  const columns = [...baseCols, ...bundleCols];

  if (format === "xlsx") {
    // 你想要的话后续可以加 SheetJS，这里先保持和你现在一致
    return new Response("xlsx 暂未启用（目前仅支持 CSV）", { status: 400 });
  }

  const csv = Papa.unparse(out, { columns, skipEmptyLines: true });
  return new Response(csv, {
    headers: {
      "Content-Type": "text/csv; charset=utf-8",
      "Content-Disposition": 'attachment; filename="orders_one_row.csv"',
    },
  });
}

/** =========================
 *  /api/labels：生成 PDF 标签
 *  - mode=bundle：每个 bundle 一张
 *  - mode=order：每个订单一张（取 bundle1 的信息做代表）
 *  ========================= */
const CM_TO_PT = 28.3464566929;
const PAGE_W = 5 * CM_TO_PT; // 5cm
const PAGE_H = 4 * CM_TO_PT; // 4cm

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
  const marginL = 6;
  const marginR = 6;
  const topY = PAGE_H - 6;

  const text = (x:number, y:number, t:string, size=8) => {
    page.drawText(t ?? "", { x, y, size, font, color: black });
  };
  const hline = (y:number, lw=1) => {
    page.drawLine({ start: { x: marginL, y }, end: { x: PAGE_W - marginR, y }, thickness: lw, color: black });
  };

  text(marginL, topY - 10, `Backer Number: ${data.backer}`, 8);
  text(marginL, topY - 24, `Name: ${data.name}`, 8);

  hline(topY - 34, 1);

  text(marginL, topY - 48, "Prescription:", 7);
  text(marginL + 92, topY - 48, data.presType || "-", 7);

  text(marginL, topY - 62, "Thickness:", 7);
  text(marginL + 92, topY - 62, data.thickness || "index lens", 7);

  text(marginL, topY - 76, "Coating:", 7);
  text(marginL + 92, topY - 76, data.coating || "-", 7);

  hline(topY - 86, 0.8);

  const headerY = topY - 98;
  const rowOdY  = topY - 112;
  const rowOsY  = topY - 126;

  const colX = [42, 70, 98, 126, 154];
  ["sph","cyl","axis","add","pd"].forEach((h, i) => text(colX[i] - 6, headerY, h, 7));
  text(marginL + 10, rowOdY, "od", 7);
  text(marginL + 10, rowOsY, "os", 7);

  const vOD = [data.od.sph, data.od.cyl, data.od.axis, data.od.add, data.od.pd].map(v => v || "-");
  const vOS = [data.os.sph, data.os.cyl, data.os.axis, data.os.add, data.os.pd].map(v => v || "-");
  vOD.forEach((v, i) => text(colX[i] - 6, rowOdY, v, 7));
  vOS.forEach((v, i) => text(colX[i] - 6, rowOsY, v, 7));

  hline(topY - 140, 0.8);
  text(marginL, topY - 154, data.dateStr, 7);
}

async function handleLabels(request: Request) {
  const url = new URL(request.url);
  const mode = (url.searchParams.get("mode") || "bundle").toLowerCase(); // bundle | order

  const form = await request.formData();
  const file = form.get("file") as File | null;
  if (!file) return new Response("Missing file field 'file'", { status: 400 });

  const text = await file.text();
  const parsed = Papa.parse<Record<string, any>>(text, { header: true, skipEmptyLines: true });
  if (parsed.errors?.length) {
    return new Response("CSV parse error: " + JSON.stringify(parsed.errors.slice(0, 3)), { status: 400 });
  }
  const rows = parsed.data || [];
  if (!rows.length) return new Response("Empty CSV", { status: 400 });

  // 必要列（为了能按 bundle/order 组织）
  for (const col of [ORDER_ID, BUNDLE_ID, LINE_ITEM, QTY]) {
    if (!(col in rows[0])) return new Response(`Missing required column: ${col}`, { status: 400 });
  }

  // 日期字符串
  const today = new Date();
  const dateStr = today.toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" });

  // 组装 order/bundle 数据（跟聚合类似，但我们还要拿 rx）
  const orderMap = new Map<string, any>();

  for (const r of rows) {
    const orderId = r[ORDER_ID];
    const bundleId = r[BUNDLE_ID];

    if (!orderMap.has(orderId)) {
      orderMap.set(orderId, {
        orderId,
        name: "",
        bundles: new Map<string, any>(),
      });
    }
    const o = orderMap.get(orderId);

    // name 尝试取表里的 Name（你的订单 CSV 有）
    if (!o.name && String(r["Name"] ?? "").trim() !== "") o.name = String(r["Name"]);

    if (!o.bundles.has(bundleId)) {
      o.bundles.set(bundleId, {
        bundleId,
        items: { Frame: [], Lens: [], Coating: [], Other: [] },
        rx: {}, // 存储一堆可能的字段
      });
    }
    const b = o.bundles.get(bundleId);

    const item = r[LINE_ITEM];
    const qty = Number(r[QTY]) || 1;
    const display = qty !== 1 ? `${item} x${qty}` : String(item);
    const cat = categorize(item);
    b.items[cat].push(display);

    // 保存原始行，后面 pick rx
    // 用 firstNonEmpty 思路：同字段有值就保留
    const keysToKeep = [
      "Sphere OD","Sphere OS","Cylinder OD","Cylinder OS","Axis OD","Axis OS",
      "ADD OD","ADD OS","Add OD","Add OS",
      "PD_OD","PD_OS","Single PD","PD",
      "Prescription Type","Prescription","Lens Type"
    ];
    for (const k of keysToKeep) {
      if (!b.rx[k] && String(r[k] ?? "").trim() !== "") b.rx[k] = r[k];
    }
  }

  const pdf = await PDFDocument.create();
  const font = await pdf.embedFont(StandardFonts.Helvetica);

  // 生成页面
  for (const [, o] of orderMap) {
    const bundles = Array.from(o.bundles.values()).sort((a: any, b: any) => String(a.bundleId).localeCompare(String(b.bundleId)));

    const targets = (mode === "order")
      ? (bundles.length ? [bundles[0]] : [])
      : bundles;

    for (const b of targets) {
      const lensText = uniqPreserveOrder(b.items.Lens).join("; ");
      const coatingText = uniqPreserveOrder(b.items.Coating).join("; ");
      const frameText = uniqPreserveOrder(b.items.Frame).join("; ");

      // thickness 从镜片文本里解析 1.60/1.67/1.74...
      const idx = parseIndexLensFromText(lensText);
      const thickness = idx ? `${idx} index lens` : "index lens";

      // coating 优先显示 Blue Light Blocking（如果存在），否则显示第一个或拼接
      let coating = "Blue Light Blocking";
      if (coatingText) {
        coating = coatingText.toLowerCase().includes("blue") ? "Blue Light Blocking" : coatingText;
      }

      const presType = String(pick(b.rx, ["Prescription Type","Prescription","Lens Type"])) || "Single Vision";

      // RX 字段尽量从多种列名取
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

      // 如果你也想把 frame/lens 显示到 label 上（更像工厂用），可以告诉我，我给你加一行
      // 例如：page.drawText(frameText.slice(0, 30), ...)
      void frameText; // 防止 TS 未使用告警（你不需要可以删）
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
 *  主入口
 *  ========================= */
export default {
  async fetch(request: Request, env: any) {
    const url = new URL(request.url);

    if (url.pathname === "/health") return new Response("ok");

    // /api/aggregate
    if (url.pathname === "/api/aggregate") {
      if (request.method !== "POST") return new Response("Method Not Allowed", { status: 405 });
      return handleAggregate(request);
    }

    // /api/labels
    if (url.pathname === "/api/labels") {
      if (request.method !== "POST") return new Response("Method Not Allowed", { status: 405 });
      return handleLabels(request);
    }

    // 静态页面
    if (env.ASSETS) return env.ASSETS.fetch(request);

    return new Response("Not Found", { status: 404 });
  },
};
