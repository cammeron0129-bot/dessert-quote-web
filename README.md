# 甜品菜单 → 自动报价单（离线网页版）

这个小工具可以在网页里勾选/搜索单品，自动生成报价单（可直接打印/导出 PDF）。

## 使用方式（最简单）

1. 打开 `dessert-quote-web/index.html`
2. 在「菜单」里选择产品、填写数量
3. 切到「报价单」预览，点击「打印 / 导出 PDF」

> 说明：菜单数据来自你提供的 `2024当夏茶歇单品(2).pdf`（包含分类与参考图片），报价单模板来自 `中秋国庆主题甜品台.pdf`。

## 修改菜单/价格

- 单品数据：`dessert-quote-web/data/menu.json`
- 单品图片：`dessert-quote-web/assets/menu/`（菜单 JSON 的 `image` 字段会引用这些图片）
- 报价单示例行（可选）：`dessert-quote-web/data/template_quote.json`

修改完 JSON 后，如果你也想让 `index.html` 在 `file://` 直接打开时仍能读到数据，请同时更新：

- `dessert-quote-web/data/menu.js`
- `dessert-quote-web/data/template_quote.js`

（它们只是把 JSON 变成 `window.MENU_DATA` / `window.TEMPLATE_QUOTE_DATA` 的脚本。）

## 重新从 PDF 抽取菜单（分类+图片）

在项目根目录运行：

- `source .venv/bin/activate`
- `python3 dessert-quote-web/scripts/extract_menu_from_pdf.py --pdf "/path/to/2024当夏茶歇单品(2).pdf" --out-json dessert-quote-web/data/menu.json --assets-dir dessert-quote-web/assets/menu`

