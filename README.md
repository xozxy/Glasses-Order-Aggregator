# Glasses Order Aggregator & Label Generator

用于将眼镜订单明细 CSV 聚合为“一单一行”并生成 5×4cm PDF 标签的 Cloudflare Workers + Pages 项目。

## 功能

- `/api/aggregate`：上传订单明细 CSV，输出“一单一行”聚合 CSV
- `/api/labels`：上传 CSV（原始明细或聚合表），生成 5×4cm PDF label
- 前端页面：一键上传并下载 CSV/PDF，自动分批生成并合并 PDF

## 技术栈

- Cloudflare Workers + Assets
- pdf-lib + fontkit（多语言字体）
- PapaParse（CSV 解析）

## 本地开发

```bash
npm i
npx wrangler dev
```

打开浏览器访问 `http://127.0.0.1:8787`。

## 接口说明

### 1) POST `/api/aggregate`

上传原始订单明细 CSV（每个产品一行）并输出聚合 CSV。

请求：

- `multipart/form-data`，字段名：`file`
- Query：`format=csv`（当前仅支持 CSV）

响应：

- `text/csv` 文件下载

### 2) POST `/api/labels`

根据 CSV 生成 5×4cm PDF label，每条记录一页。

请求：

- `multipart/form-data`，字段名：`file`
- Query：
  - `mode=bundle`（默认）：每个 Bundle 一张标签
  - `mode=order`：每个订单一张（聚合表读取 Bundle 1）
  - `start` / `limit`：分页输出，避免 Worker 资源超限（默认 200，最大 300）

响应：

- `application/pdf` 文件下载

## CSV 兼容与字段映射

### A) 原始明细表（每行一个产品）

必须包含：

- `Order ID`
- `Bundle ID`
- `Line Item`
- `Quantity`

处方字段（优先级从高到低）：

- OD：`OD SPH` / `OD CYL` / `OD Axis` / `OD ADD`
- OS：`OS SPH` / `OS CYL` / `OS Axis` / `OS ADD`
- PD：`PD OD` / `PD OS`
- 处方类型：`Prescription Type`
- 兼容 fallback：`OD_SPH`/`OS_SPH` 等下划线版本

### B) 聚合表（orders_one_row.csv）

处方字段格式：

- `OD SPH (Bundle 1)` / `OD CYL (Bundle 1)` / `OD Axis (Bundle 1)` / `OD ADD (Bundle 1)`
- `OS SPH (Bundle 1)` / `OS CYL (Bundle 1)` / `OS Axis (Bundle 1)` / `OS ADD (Bundle 1)`
- `PD OD (Bundle 1)` / `PD OS (Bundle 1)`
- `Prescription Type (Bundle 1)`

**模式规则：**

- `mode=order`：仅读取 Bundle 1
- `mode=bundle`：循环所有 Bundle（根据 `Bundle Count` 或列名推断）

如果聚合表缺少 Bundle 列会返回 400 并提示。

## PDF 生成逻辑（关键点）

- Helvetica 仅用于 ASCII；非 ASCII 必须用嵌入字体
- 按需加载字体（JP/SC/KR），减少资源消耗
- `subset:true` 缩小字体体积
- label 超过 200 页会建议分批生成（前端已自动处理）

## 前端批量合并

页面按钮“生成并下载 PDF”支持：

1. 先尝试直接生成
2. 超过上限自动分批请求 `/api/labels?start=&limit=`
3. 浏览器端使用 `pdf-lib` 合并并下载单个 PDF

## 字体文件

放在 `public/fonts/`：

- `NotoSansJP.ttf`
- `NotoSansSC.ttf`
- `NotoSansTC.ttf`（备用）
- `NotoSansKR.ttf`

Worker 通过 `env.ASSETS.fetch("http://local/fonts/xxx")` 读取。

## 最小自测

1) 原始明细 CSV  
```
curl -F "file=@/path/to/lensadvizor-orders.csv" \
  "http://127.0.0.1:8787/api/labels?mode=bundle" \
  -o labels_5x4cm.pdf
```

2) 聚合 CSV  
```
curl -F "file=@/path/to/orders_one_row.csv" \
  "http://127.0.0.1:8787/api/labels?mode=order" \
  -o labels_5x4cm.pdf
```

3) 生成聚合 CSV  
```
curl -F "file=@/path/to/lensadvizor-orders.csv" \
  "http://127.0.0.1:8787/api/aggregate" \
  -o orders_one_row.csv
```

## 常见问题

**Q: PDF 里全部是 0.00 / 0？**  
A: 字段名不匹配导致读空值，现在已按真实字段名优先读取并避免空值变 0。

**Q: 标签太多导致 503？**  
A: 使用分页生成（`start`/`limit`），前端已自动分批合并。
