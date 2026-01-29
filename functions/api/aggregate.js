import Papa from "papaparse";

// ====== 可选：如果你需要导出 xlsx，取消下一行注释并 npm i xlsx ======
// import * as XLSX from "xlsx";

// 你表格里常见的列名（按你提供的 csv）
// Order ID, Bundle ID, Line Item, Quantity 等
const ORDER_ID = "Order ID";
const BUNDLE_ID = "Bundle ID";
const LINE_ITEM = "Line Item";
const QTY = "Quantity";

function categorize(item) {
  if (!item) return "Other";
  const s = String(item).toLowerCase();
  if (s.includes("coating")) return "Coating";
  if (s.includes("index lens") || s.includes("prescription lens") || /\blens\b/i.test(s)) return "Lens";
  if (s.includes("frame") || s.includes("glasses style")) return "Frame";
  return "Other";
}

function firstNonEmpty(values) {
  for (const v of values) {
    if (v !== null && v !== undefined && String(v).trim() !== "") return v;
  }
  return "";
}

function toNumberMaybe(x) {
  const n = Number(x);
  return Number.isFinite(n) ? n : null;
}

function uniqPreserveOrder(arr) {
  const seen = new Set();
  const out = [];
  for (const x of arr) {
    if (!x) continue;
    const key = String(x);
    if (!seen.has(key)) {
      seen.add(key);
      out.push(key);
    }
  }
  return out;
}

function detectRxColumns(columns) {
  const rxRe = /\b(OD|OS|PD|Prism|ADD|Axis|Cylinder|Sphere|Pupillary|base)\b/i;
  return columns.filter((c) => rxRe.test(c) || c === "Lens Notes");
}

function buildCsv(rows, columns) {
  return Papa.unparse(rows, { columns, quotes: false, skipEmptyLines: true });
}

export async function onRequest(context) {
  const { request } = context;

  if (request.method !== "POST") {
    return new Response("Method Not Allowed", { status: 405 });
  }

  const url = new URL(request.url);
  const format = (url.searchParams.get("format") || "csv").toLowerCase();

  const form = await request.formData();
  const file = form.get("file");
  if (!file) return new Response("Missing file field 'file'", { status: 400 });

  const text = await file.text();

  const parsed = Papa.parse(text, { header: true, skipEmptyLines: true });
  if (parsed.errors?.length) {
    return new Response("CSV parse error: " + JSON.stringify(parsed.errors.slice(0, 3)), { status: 400 });
  }

  const rows = parsed.data;
  if (!rows.length) return new Response("Empty CSV", { status: 400 });

  // 必要列校验
  for (const col of [ORDER_ID, BUNDLE_ID, LINE_ITEM, QTY]) {
    if (!(col in rows[0])) return new Response(`Missing required column: ${col}`, { status: 400 });
  }

  const allColumns = Object.keys(rows[0]);
  const rxColumns = detectRxColumns(allColumns);

  // 订单级字段：排除行级 + rx + 关键 id
  const lineLevelCols = new Set([LINE_ITEM, QTY, "Line Item Price"]);
  const idCols = new Set([ORDER_ID, BUNDLE_ID]);
  const rxSet = new Set(rxColumns);

  const orderLevelCols = allColumns.filter((c) => !lineLevelCols.has(c) && !rxSet.has(c) && !idCols.has(c));

  // ========== 分组：Order -> Bundle -> Items ==========
  const orderMap = new Map(); // orderId -> { orderFields, bundles: Map(bundleId -> bundleObj) }

  for (const r of rows) {
    const orderId = r[ORDER_ID];
    const bundleId = r[BUNDLE_ID];

    if (!orderMap.has(orderId)) {
      orderMap.set(orderId, { orderId, orderFields: {}, bundles: new Map() });
    }
    const orderObj = orderMap.get(orderId);

    // 订单级字段：首个非空
    for (const c of orderLevelCols) {
      if (!orderObj.orderFields[c] && String(r[c] ?? "").trim() !== "") {
        orderObj.orderFields[c] = r[c];
      }
    }

    if (!orderObj.bundles.has(bundleId)) {
      orderObj.bundles.set(bundleId, {
        bundleId,
        items: { Frame: [], Lens: [], Coating: [], Other: [] },
        rx: {},
      });
    }
    const b = orderObj.bundles.get(bundleId);

    // item 分类
    const item = r[LINE_ITEM];
    const qty = toNumberMaybe(r[QTY]) ?? 1;
    const display = qty && qty !== 1 ? `${item} x${qty}` : String(item);
    const cat = categorize(item);
    b.items[cat].push(display);

    // rx 字段：bundle 内首个非空
    for (const c of rxColumns) {
      if (!b.rx[c] && String(r[c] ?? "").trim() !== "") b.rx[c] = r[c];
    }
  }

  // ========== 输出：宽表（Bundle1/2/3...） ==========
  // 每个订单可能有多个 bundle，按 bundleId 排序并编号
  const out = [];
  let maxBundleCount = 1;

  for (const [, o] of orderMap) {
    const bundleEntries = Array.from(o.bundles.values()).sort((a, b) => String(a.bundleId).localeCompare(String(b.bundleId)));
    maxBundleCount = Math.max(maxBundleCount, bundleEntries.length);

    const rowOut = {
      [ORDER_ID]: o.orderId,
      "Bundle Count": bundleEntries.length,
      ...o.orderFields,
    };

    bundleEntries.forEach((b, idx) => {
      const i = idx + 1;
      rowOut[`Bundle ID (Bundle ${i})`] = b.bundleId;
      rowOut[`Frame Items (Bundle ${i})`] = uniqPreserveOrder(b.items.Frame).join("; ");
      rowOut[`Lens Items (Bundle ${i})`] = uniqPreserveOrder(b.items.Lens).join("; ");
      rowOut[`Coating Items (Bundle ${i})`] = uniqPreserveOrder(b.items.Coating).join("; ");
      rowOut[`Other Items (Bundle ${i})`] = uniqPreserveOrder(b.items.Other).join("; ");

      for (const c of rxColumns) {
        rowOut[`${c} (Bundle ${i})`] = b.rx[c] || "";
      }
    });

    out.push(rowOut);
  }

  // 统一列顺序（更好看）
  const baseCols = [ORDER_ID, "Bundle Count", ...orderLevelCols];
  const bundleCols = [];
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

  // ========== 返回 CSV / XLSX ==========
  if (format === "xlsx") {
    // ✅ 可选：xlsx 导出（需要 npm i xlsx）
    // SheetJS 在 Cloudflare Workers/Pages 有官方 demo 参考。 :contentReference[oaicite:7]{index=7}
    // 取消顶部 XLSX import 后启用以下代码：
    //
    // const ws = XLSX.utils.json_to_sheet(out, { header: columns });
    // const wb = XLSX.utils.book_new();
    // XLSX.utils.book_append_sheet(wb, ws, "Orders");
    // const data = XLSX.write(wb, { bookType: "xlsx", type: "array" }); // ArrayBuffer
    // return new Response(data, {
    //   headers: {
    //     "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    //     "Content-Disposition": 'attachment; filename="orders_one_row.xlsx"',
    //   },
    // });

    return new Response(
      "xlsx 导出未启用：请 npm i xlsx 并按文件注释开启。",
      { status: 400 }
    );
  }

  const csv = buildCsv(out, columns);
  return new Response(csv, {
    headers: {
      "Content-Type": "text/csv; charset=utf-8",
      "Content-Disposition": 'attachment; filename="orders_one_row.csv"',
    },
  });
}
