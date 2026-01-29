import Papa from "papaparse";

const ORDER_ID = "Order ID";
const BUNDLE_ID = "Bundle ID";
const LINE_ITEM = "Line Item";
const QTY = "Quantity";

function categorize(item: any) {
  const s = String(item ?? "").toLowerCase();
  if (s.includes("coating")) return "Coating";
  if (s.includes("index lens") || s.includes("prescription lens") || /\blens\b/i.test(s)) return "Lens";
  if (s.includes("frame") || s.includes("glasses style")) return "Frame";
  return "Other";
}

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

function detectRxColumns(columns: string[]) {
  const rxRe = /\b(OD|OS|PD|Prism|ADD|Axis|Cylinder|Sphere|Pupillary|base)\b/i;
  return columns.filter((c) => rxRe.test(c) || c === "Lens Notes");
}

function firstNonEmpty(values: any[]) {
  for (const v of values) {
    if (v !== null && v !== undefined && String(v).trim() !== "") return v;
  }
  return "";
}

export default {
  async fetch(request: Request, env: any, ctx: ExecutionContext) {
    const url = new URL(request.url);

    // ✅ API：/api/aggregate
    if (url.pathname === "/api/aggregate") {
      if (request.method !== "POST") return new Response("Method Not Allowed", { status: 405 });

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
      const rxColumns = detectRxColumns(allColumns);

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

      if (format === "xlsx") {
        return new Response("xlsx 暂未启用（先跑通 CSV 再加）", { status: 400 });
      }

      const csv = Papa.unparse(out, { columns, skipEmptyLines: true });

      return new Response(csv, {
        headers: {
          "Content-Type": "text/csv; charset=utf-8",
          "Content-Disposition": 'attachment; filename="orders_one_row.csv"',
        },
      });
    }

    // ✅ 非 /api 请求：交给静态资源（你的 index.html）
    // 需要 wrangler.toml 配了 [assets] binding = "ASSETS"
    if (env.ASSETS) return env.ASSETS.fetch(request);

    // 没有 assets 绑定时的兜底
    return new Response("Not Found", { status: 404 });
  },
};
