# Codex 功能模块部署文档

## 概览

| 模块 | 状态 | 运行时 | 关键库 |
|------|------|--------|--------|
| 🌐 Browser | ✅ 已部署 | Node.js | playwright 1.60.0 |
| 🖥️ Computer Use | ✅ 已修复 | Node.js | @oai/sky (Named Pipe) |
| 📄 Documents | ✅ 已部署 | Python | python-docx 1.2.0 |
| 📊 Spreadsheets | ✅ 已部署 | Python | openpyxl 3.1.5 + pandas 3.0.1 |
| 📽️ Presentations | ✅ 已部署 | Node.js | pptxgenjs |

## 快速使用

### 1. 浏览器 (Browser)
```javascript
const { chromium } = await import("playwright");
const browser = await chromium.launch({ headless: false });
const page = await browser.newPage();
await page.goto("https://example.com");
```

### 2. 桌面控制 (Computer Use)
```javascript
const { setupComputerUseRuntime } = await import(
  "C:\\Users\\Lenovo\\.codex\\plugins\\cache\\openai-bundled\\computer-use\\26.602.40724\\scripts\\computer-use-client.mjs"
);
const sky = await setupComputerUseRuntime({ globals: globalThis });
const apps = await sky.list_apps();
```

### 3. Word 文档
```bash
python -c "from docx import Document; d = Document(); d.add_heading('Hello', 0); d.save('test.docx')"
```

### 4. Excel 表格
```bash
python -c "import openpyxl; wb = openpyxl.Workbook(); wb.active['A1'] = 'Hello'; wb.save('test.xlsx')"
```

### 5. PPT 演示
```javascript
const pptxgen = await import("pptxgenjs");
const pres = new pptxgen.default();
pres.addSlide().addText("Hello", { x: 1, y: 1, w: 8, h: 2, fontSize: 36 });
const buf = await pres.writeFile({ outputType: "nodebuffer" });
```

## 修复说明

**Computer Use** 的根因是 Codex Desktop Rust server 未正确注入 `nativePipe` API。修复方案：在 `computer-use-client.mjs` 中增加了 `node:net` + `node:child_process` 的故障转移路径，自动发现并直连 named pipe。

**Browser & Presentations**: 通过创建 pnpm `junction` 链接，将 `playwright-core`、`jszip` 等缺失的传递依赖暴露到 bundled `node_modules` 中。
