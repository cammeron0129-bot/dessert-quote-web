/* global MENU_DATA, TEMPLATE_QUOTE_DATA */

const STORAGE_KEY = "dessert_quote_web:v1";
const SIDEBAR_W_KEY = "dessert_quote_web:sidebarW";
const SERVICE_FEE_NAME =
  "摆台服务费（来回交通运输费，甜品师摆台服务，专人撤场服务，包含提供桌布，摆台器皿，仿真花装饰，餐具纸杯）";

const SERVICE_FEE_MENU_ITEM = {
  category: "服务",
  name: SERVICE_FEE_NAME,
  unitPrice: 0,
  minOrder: "1项",
  minOrderNum: 1,
  image: null,
};

function el(sel) {
  const node = document.querySelector(sel);
  if (!node) throw new Error(`Missing element: ${sel}`);
  return node;
}

function formatMoney(n) {
  const num = Number.isFinite(n) ? n : 0;
  return `${num.toFixed(2)}`;
}

function parseMinOrder(minOrder) {
  if (!minOrder) return null;
  const m = String(minOrder).match(/(\d+(?:\.\d+)?)/);
  if (!m) return null;
  return Number(m[1]);
}

function defaultMeta() {
  const now = new Date();
  const y = now.getFullYear();
  const m = String(now.getMonth() + 1).padStart(2, "0");
  const d = String(now.getDate()).padStart(2, "0");
  return {
    date: `${y}-${m}-${d}`,
    location: "",
    customer: "",
    contact: "",
    discountPercent: 100,
    taxPercent: 0,
    finalPrice: "",
    note: "不含税",
    orderNotes:
      "1. 甜品台为预约项目，建议提前预定档期。\n2. 动物奶油易融，空调环境建议摆放 2–3 小时，请合理预约时间。\n3. 交通/摆台/撤场等费用请按实际情况另行确认。",
    orderNotesTitle: "订购说明（可编辑）",
  };
}

function loadState() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return null;
    return JSON.parse(raw);
  } catch {
    return null;
  }
}

function saveState(state) {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
}

function buildInitialState() {
  return {
    meta: defaultMeta(),
    selected: {}, // { [name]: { qty, unitPrice, minOrder, category } }
    quoteLines: [], // { id, name, qty, unitPrice, source? }
    nextId: 1,
  };
}

function computeQuoteLinesFromSelected(selected) {
  const lines = Object.values(selected)
    .filter((x) => Number(x.qty) > 0)
    .map((x) => ({
      name: x.name,
      qty: Number(x.qty),
      unitPrice: Number(x.unitPrice ?? 0),
    }));
  return lines;
}

function toNumber(v) {
  const n = Number(v);
  return Number.isFinite(n) ? n : 0;
}

function clamp(n, min, max) {
  return Math.max(min, Math.min(max, n));
}

function debounce(fn, waitMs) {
  let t = null;
  return (...args) => {
    if (t) clearTimeout(t);
    t = setTimeout(() => fn(...args), waitMs);
  };
}

function init() {
  const menu = [
    SERVICE_FEE_MENU_ITEM,
    ...(Array.isArray(window.MENU_DATA) ? window.MENU_DATA : []),
  ];
  const template = Array.isArray(window.TEMPLATE_QUOTE_DATA) ? window.TEMPLATE_QUOTE_DATA : [];

  let state = loadState() || buildInitialState();
  let activeCategory = "全部";

  // Ensure meta keys exist
  state.meta = { ...defaultMeta(), ...(state.meta || {}) };
  state.selected = state.selected || {};
  state.quoteLines = state.quoteLines || [];
  state.nextId = state.nextId || 1;

  // UI refs
  const tabs = Array.from(document.querySelectorAll(".tab"));
  const panes = Array.from(document.querySelectorAll(".tabpane"));
  const menuList = el("#menuList");
  const menuSearch = el("#menuSearch");
  const onlySelected = el("#onlySelected");
  const menuCount = el("#menuCount");
  const categoryBar = el("#categoryBar");

  const metaDate = el("#metaDate");
  const metaLocation = el("#metaLocation");
  const metaCustomer = el("#metaCustomer");
  const metaContact = el("#metaContact");
  const metaDiscount = el("#metaDiscount");
  const metaTaxPercent = el("#metaTaxPercent");
  const metaFinalPrice = el("#metaFinalPrice");
  const metaNote = el("#metaNote");
  const metaOrderNotes = el("#metaOrderNotes");
  const metaOrderNotesTitle = el("#metaOrderNotesTitle");

  const qDate = el("#qDate");
  const qLocation = el("#qLocation");
  const qCustomer = el("#qCustomer");
  const qContact = el("#qContact");
  const qNote = el("#qNote");
  const qOrderNotesTitle = el("#qOrderNotesTitle");
  const qOrderNotes = el("#qOrderNotes");

  const quoteTbody = el("#quoteTbody");
  const subtotalEl = el("#subtotal");
  const taxIncludedEl = el("#taxIncluded");
  const totalAfterDiscountEl = el("#totalAfterDiscount");

  const sidebarResizer = document.querySelector("#sidebarResizer");

  const imageModal = el("#imageModal");
  const modalClose = el("#modalClose");
  const modalImg = el("#modalImg");
  const modalCaption = el("#modalCaption");

  const btnExportImage = el("#btnExportImage");
  const btnExportCsv = el("#btnExportCsv");
  const btnScreenshot = el("#btnScreenshot");
  const btnReset = el("#btnReset");
  const btnLoadTemplate = el("#btnLoadTemplate");

  function buildExportName() {
    const date = (state.meta.date || "").trim() || "日期未填";
    const customer = (state.meta.customer || "").trim() || "客户未填";
    return `当夏报价单_${date}_${customer}`;
  }

  function downloadText(filename, mime, text) {
    const blob = new Blob([text], { type: mime });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  }

  async function exportTable() {
    const exportName = buildExportName();
    const oldText = btnExportCsv.textContent;
    btnExportCsv.disabled = true;
    btnExportCsv.textContent = "生成中...";

    try {
      const isMobile = document.body.classList.contains("mobile");
      const lines = state.quoteLines.map((l, idx) => ({ ...l, seq: idx + 1 }));
      // 手机端导出给 WPS/Excel 打开时，更需要“文件自包含”（内嵌缩略图）才能避免图片/内容空白。
      // 为避免文件过大：仅使用缩略图，并限制在合理数量（与长图一致最多 50）。
      const MAX_TABLE_LINES = 50;
      const tableLines = lines.slice(0, MAX_TABLE_LINES);
      const inlineImages = (isMobile || lines.length <= 30) && tableLines.length <= 50;
      if (lines.length > MAX_TABLE_LINES) {
        alert(`已选择 ${lines.length} 项，导出表格最多支持 ${MAX_TABLE_LINES} 项，本次仅导出前 ${MAX_TABLE_LINES} 项。`);
      }

      const subtotal = tableLines.reduce((sum, l) => sum + toNumber(l.qty) * toNumber(l.unitPrice), 0);
      const taxPercent = clamp(toNumber(state.meta.taxPercent || 0), 0, 100);
      const taxIncluded = subtotal * (1 + taxPercent / 100);
      const discountPercent = clamp(toNumber(state.meta.discountPercent || 100), 0, 100);
      const computedAfterDiscount = subtotal * (discountPercent / 100);
      const manualFinal = toNumber(state.meta.finalPrice);
      const totalAfterDiscount = manualFinal > 0 ? manualFinal : computedAfterDiscount;

      async function loadDataUrlsWithLimit(urls, limit) {
        const results = new Array(urls.length).fill("");
        const cache = new Map();
        let nextIndex = 0;
        const n = Math.max(1, Math.min(limit, urls.length));

        async function worker() {
          while (nextIndex < urls.length) {
            const i = nextIndex++;
            const u = urls[i];
            if (!u) continue;
            const abs = new URL(u, location.href).toString();
            if (cache.has(abs)) {
              results[i] = cache.get(abs);
              continue;
            }
            let data = "";
            try {
              data = await toDataUrlFromUrl(u);
            } catch {
              // retry once
              try {
                await new Promise((r) => setTimeout(r, 160));
                data = await toDataUrlFromUrl(u);
              } catch {
                data = "";
              }
            }
            cache.set(abs, data);
            results[i] = data;
          }
        }

        await Promise.all(Array.from({ length: n }, () => worker()));
        return results;
      }

      const imgUrls = tableLines.map((l) => {
        if (!l.source?.startsWith("menu:")) return "";
        const name = l.source.slice("menu:".length);
        return menuThumbForName(name) || menuImageForLine(l) || "";
      });
      const imgDataUrls = inlineImages ? await loadDataUrlsWithLimit(imgUrls, 8) : [];

      let logoDataUrl = "";
      let logoUrl = "";
      try {
        logoUrl = new URL("./assets/brand/logo.png", location.href).toString();
      } catch {
        logoUrl = "./assets/brand/logo.png";
      }
      if (inlineImages) {
        try {
          logoDataUrl = await toDataUrlFromUrl("./assets/brand/logo.png");
        } catch {}
      }

      const esc = (s) =>
        String(s ?? "")
          .replace(/&/g, "&amp;")
          .replace(/</g, "&lt;")
          .replace(/>/g, "&gt;")
          .replace(/\"/g, "&quot;");

      const metaRows = [
        ["时间", state.meta.date || ""],
        ["地点", state.meta.location || ""],
        ["客户", state.meta.customer || ""],
        ["联系人", state.meta.contact || ""],
        ["备注", state.meta.note || ""],
      ];

      const rowsHtml = tableLines
        .map((l, idx) => {
          const qty = toNumber(l.qty);
          const unit = toNumber(l.unitPrice);
          const total = qty * unit;
          let src = "";
          if (inlineImages) src = imgDataUrls[idx] || "";
          else if (imgUrls[idx]) {
            try {
              src = new URL(imgUrls[idx], location.href).toString();
            } catch {
              src = imgUrls[idx];
            }
          }
          const img = src
            ? `<img src="${src}" style="width:48px;height:48px;object-fit:cover;border-radius:6px;border:1px solid #ddd" />`
            : "";
          return `
            <tr>
              <td class="num">${l.seq}</td>
              <td class="img">${img}</td>
              <td>${esc(l.name || "")}</td>
              <td class="num">${esc(qty)}</td>
              <td class="num">${esc(formatMoney(unit))}</td>
              <td class="num">${esc(formatMoney(total))}</td>
            </tr>
          `;
        })
        .join("");

      // Excel/WPS 兼容性：使用“Excel HTML”包装 + UTF-8 BOM
      const excelHtml = `
<!doctype html>
<html xmlns:o="urn:schemas-microsoft-com:office:office"
      xmlns:x="urn:schemas-microsoft-com:office:excel"
      xmlns="http://www.w3.org/TR/REC-html40">
  <head>
    <meta charset="utf-8" />
    <title>${esc(exportName)}</title>
    <!--[if gte mso 9]><xml>
      <x:ExcelWorkbook>
        <x:ExcelWorksheets>
          <x:ExcelWorksheet>
            <x:Name>报价单</x:Name>
            <x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions>
          </x:ExcelWorksheet>
        </x:ExcelWorksheets>
      </x:ExcelWorkbook>
    </xml><![endif]-->
    <style>
      body{font-family: -apple-system,BlinkMacSystemFont,"PingFang SC","Microsoft YaHei",Arial,sans-serif; padding:16px; color:#111827; background:#f6f7fb;}
      .paper{background:#fff; border:1px solid rgba(0,0,0,.10); border-radius:12px; padding:16px;}
      .head{width:100%; border-collapse:collapse; margin-bottom:8px;}
      .head td{border:0; padding:0; vertical-align:middle;}
      .logo{width:64px; height:48px; object-fit:contain;}
      .title{font-size:18px; font-weight:800; text-align:center;}
      .sep{border-top:1px dashed rgba(0,0,0,.20); margin:10px 0;}
      table{border-collapse:collapse; width:100%;}
      td,th{border-bottom:1px solid rgba(0,0,0,.10); padding:8px 6px; vertical-align:top;}
      thead th{border-bottom:1px solid rgba(0,0,0,.18); background:#f8fafc; text-align:left; font-size:12px;}
      .meta{width:100%; margin:0;}
      .meta td{border:0; padding:2px 0; font-size:12px;}
      .meta .k{color:rgba(17,24,39,.70); font-weight:700; width:84px;}
      .num{text-align:right; white-space:nowrap; font-variant-numeric: tabular-nums;}
      .img{width:74px;}
      .thumb{width:48px;height:48px;object-fit:cover;border-radius:8px;border:1px solid rgba(0,0,0,.12); background:#fff;}
      .totals{margin-top:10px; width:100%; border-collapse:collapse;}
      .totals td{border:0; padding:4px 0; font-size:13px;}
      .totals .k{color:rgba(17,24,39,.70); font-weight:800; width:84px;}
      .totals .v{font-weight:900; text-align:right; font-variant-numeric: tabular-nums;}
      .totals .hot{color:#dc2626;}
      .notesTitle{margin-top:12px; font-weight:800; font-size:12px;}
      .notes{white-space:pre-wrap; color:rgba(17,24,39,.78); margin-top:6px; font-size:11px; border:1px solid rgba(0,0,0,.10); border-radius:10px; padding:10px; background:rgba(0,0,0,.03);}
    </style>
  </head>
  <body>
    <div class="paper">
      <table class="head">
        <tr>
          <td style="width:72px">
            ${
              logoDataUrl
                ? `<img class="logo" src="${logoDataUrl}" alt="当夏烘焙" />`
                : logoUrl
                  ? `<img class="logo" src="${logoUrl}" alt="当夏烘焙" />`
                  : ""
            }
          </td>
          <td class="title">【当夏烘焙】甜品台服务 报价单</td>
          <td style="width:72px"></td>
        </tr>
      </table>

      <div class="sep"></div>

      <table class="meta">
        <tbody>
          ${metaRows
            .map(([k, v]) => `<tr><td class="k">${esc(k)}：</td><td>${esc(v)}</td></tr>`)
            .join("")}
        </tbody>
      </table>

      <div class="sep"></div>

      <table>
        <thead>
          <tr>
            <th style="width:52px">序号</th>
            <th style="width:74px">图片</th>
            <th>内容</th>
            <th style="width:70px">数量</th>
            <th style="width:90px">单价</th>
            <th style="width:90px">总价</th>
          </tr>
        </thead>
        <tbody>
          ${rowsHtml.replaceAll('style="width:48px;height:48px;object-fit:cover;border-radius:6px;border:1px solid #ddd"', 'class="thumb"')}
        </tbody>
      </table>

      <table class="totals">
        <tbody>
          <tr><td class="k">合计：</td><td class="v">${esc(formatMoney(subtotal))}</td></tr>
          <tr><td class="k">含税价：</td><td class="v">${esc(formatMoney(taxIncluded))}</td></tr>
          <tr><td class="k hot">优惠价：</td><td class="v hot">${esc(formatMoney(totalAfterDiscount))}</td></tr>
        </tbody>
      </table>

      <div class="sep"></div>
      <div class="notesTitle">${esc(state.meta.orderNotesTitle || "订购说明")}</div>
      <div class="notes">${esc(state.meta.orderNotes || "")}</div>
    </div>
  </body>
</html>`;
      const html = `\ufeff${excelHtml}`;

      // 手机端（iOS + WPS）：WPS/Excel 往往不显示 HTML-xls 里的图片（data: / VML 都不稳定）。
      // 所以手机端策略：
      // - 导出“可编辑表格”：.xls（保证能编辑，但图片可能被 WPS 忽略）
      // - 同时提供“含图片预览”：.html（浏览器打开一定有图；如需发给客户可用这个）
      if (isMobile) {
        const xlsFile = new File([html], `${exportName}.xls`, {
          type: "application/vnd.ms-excel;charset=utf-8",
        });
        const previewFile = new File([excelHtml], `${exportName}_含图预览.html`, {
          type: "text/html;charset=utf-8",
        });

        // 1) 优先直接调用系统分享（不依赖弹窗），保存到“文件”里后用 WPS/Excel 打开
        try {
          if (navigator?.share) {
            await navigator.share({
              files: [xlsFile, previewFile],
              title: exportName,
              text: "说明：.xls 可编辑但 iOS WPS 可能不显示图片；_含图预览.html 用浏览器打开一定带图片。",
            });
            return;
          }
        } catch {
          // ignore, fallback below
        }

        // 2) 无法分享时：同页打开预览（不弹窗），用户可用浏览器“分享”菜单保存
        try {
          const blob = new Blob([excelHtml], { type: "text/html;charset=utf-8" });
          const url = URL.createObjectURL(blob);
          alert("将打开“含图预览”页（一定带图片）。如需可编辑表格，请用上方“系统分享”导出 .xls。返回本页面请点浏览器“返回”。");
          window.location.assign(url);
          setTimeout(() => {
            try {
              URL.revokeObjectURL(url);
            } catch {}
          }, 120_000);
          return;
        } catch (e) {
          alert(`导出表格失败：${e?.message || e}`);
          return;
        }
        return;
      }

      // 电脑端：用 .xls 扩展名让 Excel/WPS 直接以表格打开（本质是 HTML）
      downloadText(`${exportName}.xls`, "application/vnd.ms-excel;charset=utf-8", html);
    } catch (e) {
      // eslint-disable-next-line no-console
      console.error(e);
      alert(`导出表格失败：${e?.message || e}`);
    } finally {
      btnExportCsv.disabled = false;
      btnExportCsv.textContent = oldText;
    }
  }

  async function toDataUrlFromUrl(url) {
    const res = await fetch(url, { cache: "force-cache" });
    const blob = await res.blob();
    return await new Promise((resolve) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result);
      reader.readAsDataURL(blob);
    });
  }

  async function ensureHtml2Canvas() {
    if (window.html2canvas) return window.html2canvas;
    await new Promise((resolve, reject) => {
      const s = document.createElement("script");
      s.src = "https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js";
      s.async = true;
      s.onload = resolve;
      s.onerror = () => reject(new Error("加载导出组件失败（html2canvas）"));
      document.head.appendChild(s);
    });
    return window.html2canvas;
  }

  function wrapTextCJK(ctx, text, maxWidth) {
    const s = String(text || "");
    if (!s) return [""];
    const lines = [];
    let line = "";
    for (const ch of s) {
      const test = line + ch;
      if (ctx.measureText(test).width <= maxWidth || !line) line = test;
      else {
        lines.push(line);
        line = ch;
      }
    }
    if (line) lines.push(line);
    return lines;
  }

  async function loadImageBitmap(url) {
    const abs = new URL(url, location.href).toString();
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), 30_000);
    try {
      const res = await fetch(abs, { cache: "force-cache", signal: controller.signal });
      const blob = await res.blob();
      if (!blob || blob.size === 0) throw new Error("图片为空");

      // iOS/Safari 对 createImageBitmap 支持不稳定，优先走 Image 解码，避免“生成中卡住”。
      const isIOS = /iP(hone|ad|od)/.test(navigator.userAgent);
      if (!isIOS && "createImageBitmap" in window) return await createImageBitmap(blob);

      return await new Promise((resolve, reject) => {
        const objUrl = URL.createObjectURL(blob);
        const img = new Image();
        img.crossOrigin = "anonymous";
        const t = setTimeout(() => {
          try {
            URL.revokeObjectURL(objUrl);
          } catch {}
          reject(new Error("图片解码超时"));
        }, 30_000);
        img.onload = () => {
          clearTimeout(t);
          resolve(img);
          setTimeout(() => {
            try {
              URL.revokeObjectURL(objUrl);
            } catch {}
          }, 30_000);
        };
        img.onerror = () => {
          clearTimeout(t);
          reject(new Error("图片加载失败"));
          setTimeout(() => {
            try {
              URL.revokeObjectURL(objUrl);
            } catch {}
          }, 30_000);
        };
        img.src = objUrl;
      });
    } finally {
      clearTimeout(timeout);
    }
  }

  async function loadBitmapsWithLimit(urls, limit) {
    const results = new Array(urls.length).fill(null);
    const cache = new Map(); // absUrl -> bitmap/img|null
    let nextIndex = 0;

    async function worker() {
      while (nextIndex < urls.length) {
        const i = nextIndex++;
        const u = urls[i];
        if (!u) continue;
        const abs = new URL(u, location.href).toString();
        if (cache.has(abs)) {
          results[i] = cache.get(abs);
          continue;
        }
        let bm = null;
        try {
          bm = await loadImageBitmap(u);
        } catch {
          // retry once (弱网/移动端偶发失败)
          try {
            await new Promise((r) => setTimeout(r, 180));
            bm = await loadImageBitmap(u);
          } catch {
            bm = null;
          }
        }
        cache.set(abs, bm);
        results[i] = bm;
      }
    }

    const workers = [];
    const n = Math.max(1, Math.min(limit, urls.length));
    for (let i = 0; i < n; i += 1) workers.push(worker());
    await Promise.all(workers);
    return results;
  }

  function prepareExportWindow(exportName) {
    const isMobile = document.body.classList.contains("mobile");
    if (!isMobile) return null;
    try {
      const w = window.open("about:blank", "_blank", "noopener,noreferrer");
      if (!w) return null;
      w.document.title = `${exportName}.png`;
      w.document.body.style.margin = "0";
      w.document.body.style.fontFamily =
        "ui-sans-serif, system-ui, -apple-system, PingFang SC, Microsoft YaHei";
      w.document.body.innerHTML = `<div style="padding:16px">正在生成图片，请稍等…</div>`;
      return w;
    } catch {
      return null;
    }
  }

  function renderExportWindow(preparedWindow, exportName, url) {
    if (!preparedWindow) return false;
    try {
      preparedWindow.document.title = `${exportName}.png`;
      preparedWindow.document.body.style.margin = "0";
      preparedWindow.document.body.innerHTML = `
        <div style="padding:12px; font-size:14px; color:#111827">
          <div style="margin-bottom:8px">已生成图片（iPhone：长按图片→存储到照片 或 点“分享”）</div>
          <div style="display:flex; gap:8px; flex-wrap:wrap; margin-bottom:10px">
            <button id="shareBtn" style="appearance:none; border:1px solid rgba(0,0,0,.15); background:#fff; border-radius:10px; padding:10px 12px; font-size:14px">分享/保存</button>
            <a id="openBtn" href="${url}" target="_blank" rel="noopener noreferrer" style="display:inline-flex; align-items:center; justify-content:center; border:1px solid rgba(0,0,0,.15); background:#fff; border-radius:10px; padding:10px 12px; font-size:14px; text-decoration:none; color:#111827">在新页打开</a>
          </div>
          <img id="img" src="${url}" alt="${exportName}" style="width:100%; height:auto; display:block; background:#fff" />
        </div>
      `;
      const shareBtn = preparedWindow.document.getElementById("shareBtn");
      if (shareBtn) {
        shareBtn.addEventListener("click", async () => {
          try {
            if (!preparedWindow.navigator?.share) {
              preparedWindow.alert("当前浏览器不支持“分享”。请长按图片保存。");
              return;
            }
            const res = await preparedWindow.fetch(url);
            const blob = await res.blob();
            const file = new File([blob], `${exportName}.png`, { type: "image/png" });
            await preparedWindow.navigator.share({ files: [file], title: exportName });
          } catch (e) {
            preparedWindow.alert(`分享失败：${e?.message || e}`);
          }
        });
      }
      return true;
    } catch {
      return false;
    }
  }

  async function finishDownloadPng(exportName, blobOrUrl, preparedWindow) {
    const isIOS = /iP(hone|ad|od)/.test(navigator.userAgent);
    const isMobile = document.body.classList.contains("mobile");

    const url = typeof blobOrUrl === "string" ? blobOrUrl : URL.createObjectURL(blobOrUrl);

    // 移动端/IOS：优先使用“预先打开的窗口”承接 Blob URL，避免弹窗/下载被浏览器拦截。
    if (preparedWindow) {
      // iOS 某些情况下跨窗口加载 blob: 会出现空白；优先转成 dataURL 再展示
      let displayUrl = url;
      if (typeof blobOrUrl !== "string") {
        try {
          // 小尺寸转 dataURL（避免内存炸），超大图仍用 blob
          if (blobOrUrl.size < 6_000_000) {
            displayUrl = await new Promise((resolve) => {
              const reader = new FileReader();
              reader.onload = () => resolve(reader.result);
              reader.readAsDataURL(blobOrUrl);
            });
          }
        } catch {}
      }
      const ok = renderExportWindow(preparedWindow, exportName, displayUrl);
      if (ok) {
        if (typeof blobOrUrl !== "string") setTimeout(() => URL.revokeObjectURL(url), 120_000);
        return;
      }
    }

    const a = document.createElement("a");
    a.href = url;
    a.download = `${exportName}.png`;
    a.rel = "noopener";
    try {
      document.body.appendChild(a);
      a.click();
      a.remove();
      if (isIOS || isMobile) window.open(url, "_blank", "noopener,noreferrer");
    } catch {
      window.open(url, "_blank", "noopener,noreferrer");
    }
    if (typeof blobOrUrl !== "string") setTimeout(() => URL.revokeObjectURL(url), 30_000);
  }

  async function renderQuoteToPng({ width, exportName, preparedWindow }) {
    const isMobile = document.body.classList.contains("mobile");
    const isNarrow = width < 560;

    const pad = 16;
    const fontSmall = "12px ui-sans-serif, system-ui, -apple-system, PingFang SC, Microsoft YaHei";
    const brand = "#111827";
    const muted = "rgba(17,24,39,0.72)";
    const border = "rgba(0,0,0,0.12)";
    const imgSize = 48;
    const tableHeadH = 30;
    const contentW = width - pad * 2;
    const nameW = Math.max(140, contentW - 52 - imgSize - 12 - 210);

    // 移动端和电脑端都需要稳定支持较多产品（最多 50 个带图导出）
    const MAX_EXPORT_LINES = 50;
    const allLines = state.quoteLines.map((l, idx) => ({ ...l, seq: idx + 1 }));
    const lines = allLines.slice(0, MAX_EXPORT_LINES);
    if (allLines.length > MAX_EXPORT_LINES) {
      alert(`已选择 ${allLines.length} 项，导出长图最多支持 ${MAX_EXPORT_LINES} 项，本次仅导出前 ${MAX_EXPORT_LINES} 项。`);
    }
    const rowImgUrls = lines.map((l) => {
      if (!l.source?.startsWith("menu:")) return null;
      const name = l.source.slice("menu:".length);
      const item = menu.find((m) => m.name === name);
      return item?.imageThumb || item?.image || null;
    });

    // 产品很多时，若一次性并发加载所有图片，移动端（尤其 iOS）容易丢图/解码失败。
    // 这里做“限并发 + 重试”，并且手机/电脑一致配置，保证最多 50 张也能稳定加载出来。
    const bitmaps = await loadBitmapsWithLimit(rowImgUrls, 8);

    const tmp = document.createElement("canvas");
    const tctx = tmp.getContext("2d");
    tctx.font = fontSmall;
    const rowHeights = lines.map((l) => {
      if (isNarrow) {
        const nameLines = wrapTextCJK(tctx, l.name || "", contentW - imgSize - 24);
        const infoLines = 2;
        return Math.max(imgSize + 22, nameLines.length * 16 + infoLines * 16 + 22);
      }
      const nameLines = wrapTextCJK(tctx, l.name || "", nameW);
      return Math.max(imgSize + 18, nameLines.length * 16 + 16);
    });

    const notes = state.meta.orderNotes || "";
    const notesLines = notes.split("\n").flatMap((ln) => wrapTextCJK(tctx, ln, contentW));
    const notesH = Math.max(60, notesLines.length * 14 + 18);

    const metaItems = [
      ["时间：", state.meta.date || ""],
      ["地点：", state.meta.location || ""],
      ["客户：", state.meta.customer || ""],
      ["联系人：", state.meta.contact || ""],
      ["备注：", state.meta.note || ""],
    ];
    const metaTopY = 74;
    const metaLineH = 18;
    // 头部高度自适应，避免“联系人/备注”被表头覆盖
    const headerH = Math.max(118, metaTopY + metaItems.length * metaLineH + 16);

    const tableH = (isNarrow ? 14 : tableHeadH) + rowHeights.reduce((a, b) => a + b, 0) + 10;
    const totalsH = 60;
    const fullH = headerH + tableH + totalsH + 18 + notesH + 20;

    const scale = Math.min(2, window.devicePixelRatio || 2);
    const drawAll = async (ctx) => {
      ctx.fillStyle = "#ffffff";
      ctx.fillRect(0, 0, width, fullH);

      // Header
      ctx.font = "18px ui-sans-serif, system-ui, -apple-system, PingFang SC, Microsoft YaHei";
      ctx.fillStyle = brand;
      ctx.fillText("【当夏烘焙】甜品台服务 报价单", pad + 72, 30);
      try {
        const logo = await loadImageBitmap("./assets/brand/logo.png");
        ctx.drawImage(logo, pad, 6, 64, 48);
      } catch {}

      ctx.strokeStyle = border;
      ctx.setLineDash([3, 3]);
      ctx.beginPath();
      ctx.moveTo(pad, 52);
      ctx.lineTo(width - pad, 52);
      ctx.stroke();
      ctx.setLineDash([]);

      ctx.font = fontSmall;
      let my = metaTopY;
      for (const [k, v] of metaItems) {
        ctx.fillStyle = muted;
        ctx.fillText(k, pad, my);
        ctx.fillStyle = brand;
        ctx.fillText(v, pad + 44, my);
        my += metaLineH;
      }

      // Table head / Narrow head
      let y = headerH;
      ctx.strokeStyle = border;
      ctx.beginPath();
      ctx.moveTo(pad, y);
      ctx.lineTo(width - pad, y);
      ctx.stroke();
      y += 8;
      ctx.fillStyle = muted;
      if (!isNarrow) {
        ctx.fillText("序号", pad, y + 12);
        ctx.fillText("图片", pad + 52, y + 12);
        ctx.fillText("内容", pad + 52 + imgSize + 10, y + 12);
        ctx.fillText("数量", width - pad - 190, y + 12);
        ctx.fillText("单价", width - pad - 130, y + 12);
        ctx.fillText("总价", width - pad - 70, y + 12);
        y += 22;
      } else {
        ctx.fillText("明细", pad, y + 12);
        y += 18;
      }

      for (let i = 0; i < lines.length; i += 1) {
        const l = lines[i];
        const h = rowHeights[i];
        ctx.strokeStyle = border;
        ctx.beginPath();
        ctx.moveTo(pad, y);
        ctx.lineTo(width - pad, y);
        ctx.stroke();

        const cy = y + 10;
        const qty = toNumber(l.qty);
        const unit = toNumber(l.unitPrice);
        const total = qty * unit;

        ctx.fillStyle = brand;
        if (!isNarrow) {
          ctx.fillText(String(l.seq), pad, cy + 12);
          ctx.strokeRect(pad + 46, cy, imgSize, imgSize);
          const bm = bitmaps[i];
          if (bm) ctx.drawImage(bm, pad + 46, cy, imgSize, imgSize);

          const nx = pad + 46 + imgSize + 10;
          const nameLines = wrapTextCJK(ctx, l.name || "", nameW);
          let ty = cy + 12;
          for (const ln of nameLines.slice(0, 6)) {
            ctx.fillText(ln, nx, ty);
            ty += 16;
          }

          ctx.textAlign = "right";
          ctx.fillText(String(qty), width - pad - 168, cy + 12);
          ctx.fillText(formatMoney(unit), width - pad - 100, cy + 12);
          ctx.fillText(formatMoney(total), width - pad, cy + 12);
          ctx.textAlign = "left";
        } else {
          // Narrow card-like row: image + name, then qty/unit/total below.
          ctx.strokeRect(pad, cy, imgSize, imgSize);
          const bm = bitmaps[i];
          if (bm) ctx.drawImage(bm, pad, cy, imgSize, imgSize);

          const nx = pad + imgSize + 10;
          ctx.fillStyle = muted;
          ctx.fillText(`#${l.seq}`, nx, cy + 12);
          ctx.fillStyle = brand;
          const nameLines = wrapTextCJK(ctx, l.name || "", contentW - imgSize - 24);
          let ty = cy + 30;
          for (const ln of nameLines.slice(0, 6)) {
            ctx.fillText(ln, nx, ty);
            ty += 16;
          }

          ctx.fillStyle = muted;
          ctx.fillText(`数量：${qty}   单价：${formatMoney(unit)}   总价：${formatMoney(total)}`, nx, ty + 2);
        }

        y += h;
      }
      ctx.strokeStyle = border;
      ctx.beginPath();
      ctx.moveTo(pad, y);
      ctx.lineTo(width - pad, y);
      ctx.stroke();

      const subtotal = state.quoteLines.reduce(
        (sum, l) => sum + toNumber(l.qty) * toNumber(l.unitPrice),
        0
      );
      const discountPercent = clamp(toNumber(state.meta.discountPercent || 100), 0, 100);
      const computedAfterDiscount = subtotal * (discountPercent / 100);
      const manualFinal = toNumber(state.meta.finalPrice);
      const totalAfterDiscount = manualFinal > 0 ? manualFinal : computedAfterDiscount;

    const taxPercent = clamp(toNumber(state.meta.taxPercent || 0), 0, 100);
    const taxIncluded = subtotal * (1 + taxPercent / 100);

    y += 22;
    ctx.font = "14px ui-sans-serif, system-ui, -apple-system, PingFang SC, Microsoft YaHei";
    ctx.fillStyle = muted;
    ctx.textAlign = "right";
    ctx.fillText("合计：", width - pad - 120, y);
    ctx.fillStyle = brand;
    ctx.fillText(formatMoney(subtotal), width - pad, y);
    y += 20;
    ctx.fillStyle = muted;
    ctx.fillText("含税价：", width - pad - 120, y);
    ctx.fillStyle = brand;
    ctx.fillText(formatMoney(taxIncluded), width - pad, y);
    y += 20;
    ctx.fillStyle = muted;
    ctx.fillText("优惠价：", width - pad - 120, y);
    ctx.fillStyle = "#dc2626";
    ctx.fillText(formatMoney(totalAfterDiscount), width - pad, y);
    ctx.textAlign = "left";

      y += 18;
      ctx.strokeStyle = border;
      ctx.setLineDash([3, 3]);
      ctx.beginPath();
      ctx.moveTo(pad, y);
      ctx.lineTo(width - pad, y);
      ctx.stroke();
      ctx.setLineDash([]);

      y += 18;
      ctx.font = fontSmall;
      ctx.fillStyle = brand;
      ctx.fillText(state.meta.orderNotesTitle || "订购说明", pad, y);
      y += 18;
      ctx.fillStyle = muted;
      for (const ln of notesLines) {
        ctx.fillText(ln, pad, y);
        y += 14;
        if (y > fullH - 10) break;
      }
    };

    const maxCanvasPx = isMobile ? 12_000 : 32_000;
    const targetCanvasHeightPx = Math.floor(fullH * scale);
    const tooTall = targetCanvasHeightPx > maxCanvasPx;

    if (!tooTall) {
      const canvas = document.createElement("canvas");
      canvas.width = Math.floor(width * scale);
      canvas.height = targetCanvasHeightPx;
      const ctx = canvas.getContext("2d");
      ctx.scale(scale, scale);
      await drawAll(ctx);
      const blob = await new Promise((resolve) => canvas.toBlob(resolve, "image/png"));
      if (!blob) throw new Error("图片生成失败（blob 为空）");
      await finishDownloadPng(exportName, blob, preparedWindow);
      return;
    }

    // 超过移动端 Canvas 高度限制：按固定宽度分段导出多张图
    const sliceH = Math.floor(maxCanvasPx / scale);
    const totalSlices = Math.ceil(fullH / sliceH);
    alert(`内容较长，手机浏览器限制将分 ${totalSlices} 张导出（宽度不变）。`);

    for (let s = 0; s < totalSlices; s += 1) {
      const offsetY = s * sliceH;
      const curH = Math.min(sliceH, fullH - offsetY);
      const canvas = document.createElement("canvas");
      canvas.width = Math.floor(width * scale);
      canvas.height = Math.floor(curH * scale);
      const ctx = canvas.getContext("2d");
      ctx.scale(scale, scale);
      ctx.save();
      ctx.beginPath();
      ctx.rect(0, 0, width, curH);
      ctx.clip();
      ctx.translate(0, -offsetY);
      await drawAll(ctx);
      ctx.restore();
      const blob = await new Promise((resolve) => canvas.toBlob(resolve, "image/png"));
      if (!blob) throw new Error("图片生成失败（blob 为空）");
      // 第 1 张用 preparedWindow，其余按普通下载/新开页
      await finishDownloadPng(`${exportName}_${s + 1}`, blob, s === 0 ? preparedWindow : null);
      // 避免连续触发被浏览器拦截
      await new Promise((r) => setTimeout(r, 500));
    }
  }

  async function exportScreenshotPng() {
    const exportName = `${buildExportName()}_截图`;
    const node = document.querySelector("#quotePaper");
    if (!node) return;

    const oldBtnText = btnScreenshot.textContent;
    btnScreenshot.disabled = true;
    btnScreenshot.textContent = "截图中...";
    const preparedWindow = prepareExportWindow(exportName);

    try {
      const rect = node.getBoundingClientRect();
      const width = Math.floor(rect.width);
      await renderQuoteToPng({ width, exportName, preparedWindow });
    } catch (e) {
      // eslint-disable-next-line no-console
      console.error(e);
      alert(`截图导出失败：${e?.message || e}`);
      try {
        if (preparedWindow) preparedWindow.close();
      } catch {}
    } finally {
      btnScreenshot.disabled = false;
      btnScreenshot.textContent = oldBtnText;
    }
  }

  function computeExportWidth() {
    // Desktop: A4-ish width at 96dpi (794px). Mobile: fit screen width.
    const isMobile = document.body.classList.contains("mobile");
    if (isMobile) return Math.max(360, Math.floor(window.innerWidth - 24));
    return 794;
  }

  function buildExportNode(width) {
    const wrapper = document.createElement("div");
    wrapper.className = "paper";
    wrapper.style.width = `${width}px`;
    wrapper.style.borderRadius = "0";
    wrapper.style.boxShadow = "none";
    wrapper.style.minHeight = "auto";

    const head = document.createElement("div");
    head.className = "paper__head";

    const headTop = document.createElement("div");
    headTop.className = "paper__headTop";

    const logo = document.createElement("img");
    logo.className = "paper__logo";
    logo.alt = "当夏烘焙";
    logo.src = "./assets/brand/logo.png";

    const title = document.createElement("div");
    title.className = "paper__title";
    title.textContent = "【当夏烘焙】甜品台服务 报价单";

    headTop.appendChild(logo);
    headTop.appendChild(title);

    const meta = document.createElement("div");
    meta.className = "paper__meta";
    const metaCol = document.createElement("div");
    metaCol.className = "paper__metaCol";

    const metaItems = [
      ["时间：", state.meta.date || ""],
      ["地点：", state.meta.location || ""],
      ["客户：", state.meta.customer || ""],
      ["联系人：", state.meta.contact || ""],
      ["备注：", state.meta.note || ""],
    ];
    for (const [k, v] of metaItems) {
      const item = document.createElement("div");
      item.className = "paper__metaItem";
      const ks = document.createElement("span");
      ks.className = "k";
      ks.textContent = k;
      const vs = document.createElement("span");
      vs.textContent = v;
      item.appendChild(ks);
      item.appendChild(vs);
      metaCol.appendChild(item);
    }
    meta.appendChild(metaCol);

    head.appendChild(headTop);
    head.appendChild(meta);
    wrapper.appendChild(head);

    const table = document.createElement("table");
    table.className = "table";
    const thead = document.createElement("thead");
    thead.innerHTML = `
      <tr>
        <th class="col--seq">序号</th>
        <th class="col--img">图片</th>
        <th>内容</th>
        <th class="col--qty">数量</th>
        <th class="col--price">单价</th>
        <th class="col--price">总价</th>
      </tr>
    `;
    table.appendChild(thead);
    const tbody = document.createElement("tbody");

    state.quoteLines.forEach((l, idx) => {
      const tr = document.createElement("tr");
      const qty = toNumber(l.qty);
      const unit = toNumber(l.unitPrice);
      const total = qty * unit;

      const tdSeq = document.createElement("td");
      tdSeq.className = "num";
      tdSeq.textContent = String(idx + 1);

      const tdImg = document.createElement("td");
      const imgWrap = document.createElement("div");
      const imgSrc = menuImageForLine(l);
      const thumbSrc = l.source?.startsWith("menu:")
        ? menuThumbForName(l.source.slice("menu:".length))
        : null;
      imgWrap.className = `quoteThumb ${imgSrc ? "" : "quoteThumb--empty"}`.trim();
      if (imgSrc) {
        const img = document.createElement("img");
        img.src = thumbSrc || imgSrc;
        img.alt = l.name || "图片";
        tdImg.appendChild(imgWrap);
        imgWrap.appendChild(img);
      } else {
        imgWrap.textContent = "-";
        tdImg.appendChild(imgWrap);
      }

      const tdName = document.createElement("td");
      tdName.textContent = l.name || "";

      const tdQty = document.createElement("td");
      tdQty.className = "num";
      tdQty.textContent = String(qty);

      const tdUnit = document.createElement("td");
      tdUnit.className = "num";
      tdUnit.textContent = formatMoney(unit);

      const tdTotal = document.createElement("td");
      tdTotal.className = "num";
      tdTotal.textContent = formatMoney(total);

      tr.appendChild(tdSeq);
      tr.appendChild(tdImg);
      tr.appendChild(tdName);
      tr.appendChild(tdQty);
      tr.appendChild(tdUnit);
      tr.appendChild(tdTotal);
      tbody.appendChild(tr);
    });

    table.appendChild(tbody);
    wrapper.appendChild(table);

    const subtotal = state.quoteLines.reduce((sum, l) => sum + toNumber(l.qty) * toNumber(l.unitPrice), 0);
      const taxPercent = clamp(toNumber(state.meta.taxPercent || 0), 0, 100);
      const taxIncluded = subtotal * (1 + taxPercent / 100);
    const discountPercent = clamp(toNumber(state.meta.discountPercent || 100), 0, 100);
    const computedAfterDiscount = subtotal * (discountPercent / 100);
    const manualFinal = toNumber(state.meta.finalPrice);
    const totalAfterDiscount = manualFinal > 0 ? manualFinal : computedAfterDiscount;

    const totals = document.createElement("div");
    totals.className = "totals";
    totals.innerHTML = `
      <div class="totals__row"><div class="totals__label">合计</div><div class="totals__value">${formatMoney(subtotal)}</div></div>
      <div class="totals__row"><div class="totals__label">含税价</div><div class="totals__value">${formatMoney(taxIncluded)}</div></div>
      <div class="totals__row totals__row--highlight"><div class="totals__label">优惠价</div><div class="totals__value">${formatMoney(totalAfterDiscount)}</div></div>
    `;
    wrapper.appendChild(totals);

    const foot = document.createElement("div");
    foot.className = "paper__foot";

    const footTitle = document.createElement("div");
    footTitle.className = "paper__footTitle";
    footTitle.textContent = state.meta.orderNotesTitle || "订购说明";
    foot.appendChild(footTitle);

    const notes = document.createElement("div");
    notes.className = "orderNotes";
    notes.textContent = state.meta.orderNotes || "";
    foot.appendChild(notes);

    wrapper.appendChild(foot);

    return wrapper;
  }

  async function exportLongPng() {
    const exportName = buildExportName();
    const node = document.querySelector("#quotePaper");
    if (!node) return;

    const oldBtnText = btnExportImage.textContent;
    btnExportImage.disabled = true;
    btnExportImage.textContent = "生成中...";
    const preparedWindow = prepareExportWindow(exportName);

    try {
      const width = computeExportWidth();
      await renderQuoteToPng({ width, exportName, preparedWindow });
    } catch (e) {
      // eslint-disable-next-line no-console
      console.error(e);
      alert(`导出失败：${e?.message || e}`);
      try {
        if (preparedWindow) preparedWindow.close();
      } catch {}
    } finally {
      btnExportImage.disabled = false;
      btnExportImage.textContent = oldBtnText;
    }
  }

  function openImageModal(src, caption) {
    modalImg.src = src;
    modalImg.alt = caption || "图片";
    modalCaption.textContent = caption || "";
    imageModal.setAttribute("aria-hidden", "false");
  }

  function closeImageModal() {
    imageModal.setAttribute("aria-hidden", "true");
    modalImg.removeAttribute("src");
    modalCaption.textContent = "";
  }

  function menuImageForLine(line) {
    if (!line?.source?.startsWith("menu:")) return null;
    const name = line.source.slice("menu:".length);
    const item = menu.find((m) => m.name === name);
    return item?.image || null;
  }

  function menuThumbForName(name) {
    const item = menu.find((m) => m.name === name);
    return item?.imageThumb || item?.image || null;
  }

  function setActiveTab(tabKey) {
    for (const t of tabs) t.classList.toggle("tab--active", t.dataset.tab === tabKey);
    for (const p of panes) p.classList.toggle("tabpane--active", p.dataset.pane === tabKey);
  }

  function ensureSelectedEntry(item) {
    const key = item.name;
    if (!state.selected[key]) {
      const min = parseMinOrder(item.minOrder);
      state.selected[key] = {
        name: item.name,
        qty: min ?? 0,
        unitPrice: item.unitPrice ?? 0,
        minOrder: item.minOrder ?? "",
        category: item.category ?? "",
      };
    }
    return state.selected[key];
  }

  function removeSelected(name) {
    delete state.selected[name];
  }

  function upsertQuoteLineFromMenu(name) {
    const entry = state.selected[name];
    if (!entry) return;
    const existing = state.quoteLines.find((l) => l.source === `menu:${name}`);
    const next = {
      id: existing?.id ?? state.nextId++,
      source: `menu:${name}`,
      name: entry.name,
      qty: toNumber(entry.qty),
      unitPrice: toNumber(entry.unitPrice),
    };
    if (existing) Object.assign(existing, next);
    else state.quoteLines.push(next);
  }

  function deleteQuoteLine(id) {
    state.quoteLines = state.quoteLines.filter((l) => l.id !== id);
  }

  function syncQuoteLinesFromSelected() {
    const selectedNames = Object.keys(state.selected);
    for (const name of selectedNames) upsertQuoteLineFromMenu(name);
    // Remove menu-sourced lines that are no longer selected
    state.quoteLines = state.quoteLines.filter((l) => {
      if (!l.source?.startsWith("menu:")) return true;
      const n = l.source.slice("menu:".length);
      return Boolean(state.selected[n]);
    });
  }

  function renderMenu() {
    const q = menuSearch.value.trim().toLowerCase();
    const selectedOnly = onlySelected.checked;
    const cat = activeCategory;

    const list = menu
      .filter((it) => (cat === "全部" ? true : String(it.category || "") === cat))
      .filter((it) => (q ? String(it.name).toLowerCase().includes(q) : true))
      .filter((it) => (selectedOnly ? Boolean(state.selected[it.name]) : true));

    menuCount.textContent = `共 ${list.length} 项（已选 ${Object.keys(state.selected).length}）`;

    menuList.innerHTML = "";
    for (const item of list) {
      const selected = state.selected[item.name];
      const card = document.createElement("div");
      card.className = "card";

      const grid = document.createElement("div");
      grid.className = "card__grid";

      const thumb = document.createElement("div");
      thumb.className = "thumb";
      if (item.image) {
        const img = document.createElement("img");
        img.alt = item.name;
        img.loading = "lazy";
        img.decoding = "async";
        img.fetchPriority = "low";
        img.src = item.imageThumb || item.image;
        img.addEventListener("click", () => openImageModal(item.image, item.name));
        thumb.appendChild(img);
      } else {
        thumb.textContent = "无图";
      }

      const top = document.createElement("div");
      top.className = "card__top";

      const left = document.createElement("div");
      const name = document.createElement("div");
      name.className = "card__name";
      name.textContent = item.name;
      left.appendChild(name);

      const right = document.createElement("div");
      const pill = document.createElement("span");
      pill.className = "pill";
      pill.textContent = selected ? "已选" : "未选";
      right.appendChild(pill);

      top.appendChild(left);
      top.appendChild(right);

      const meta = document.createElement("div");
      meta.className = "card__meta";
      meta.innerHTML = `
        <span>单价：<b>${formatMoney(toNumber(item.unitPrice))}</b> 元</span>
        <span>起订量：<b>${item.minOrder || "-"}</b></span>
        <span>分类：<b>${item.category || "-"}</b></span>
      `;

      const actions = document.createElement("div");
      actions.className = "card__actions";

      const qty = document.createElement("input");
      qty.className = "qty";
      qty.type = "number";
      qty.min = "0";
      qty.step = "1";
      qty.value = selected ? String(selected.qty ?? 0) : String(parseMinOrder(item.minOrder) ?? 0);
      qty.disabled = !selected;

      const btn = document.createElement("button");
      btn.className = "btn btn--ghost";
      btn.type = "button";
      btn.textContent = selected ? "移除" : "加入";

      const price = document.createElement("input");
      price.className = "qty";
      price.type = "number";
      price.min = "0";
      price.step = "0.5";
      price.value = selected ? String(selected.unitPrice ?? 0) : String(item.unitPrice ?? 0);
      price.disabled = !selected;

      const priceLabel = document.createElement("span");
      priceLabel.className = "muted";
      priceLabel.textContent = "单价";

      const qtyLabel = document.createElement("span");
      qtyLabel.className = "muted";
      qtyLabel.textContent = "数量";

      btn.addEventListener("click", () => {
        if (state.selected[item.name]) {
          removeSelected(item.name);
        } else {
          ensureSelectedEntry(item);
        }
        syncQuoteLinesFromSelected();
        saveState(state);
        renderAll();
      });

      qty.addEventListener("change", () => {
        const entry = state.selected[item.name];
        if (!entry) return;
        entry.qty = clamp(toNumber(qty.value), 0, 999999);
        upsertQuoteLineFromMenu(item.name);
        saveState(state);
        renderQuote();
      });

      price.addEventListener("change", () => {
        const entry = state.selected[item.name];
        if (!entry) return;
        entry.unitPrice = clamp(toNumber(price.value), 0, 999999);
        upsertQuoteLineFromMenu(item.name);
        saveState(state);
        renderQuote();
      });

      actions.appendChild(qtyLabel);
      actions.appendChild(qty);
      actions.appendChild(priceLabel);
      actions.appendChild(price);
      actions.appendChild(btn);

      const content = document.createElement("div");
      content.appendChild(top);
      content.appendChild(meta);
      content.appendChild(actions);

      grid.appendChild(thumb);
      grid.appendChild(content);

      card.appendChild(grid);
      menuList.appendChild(card);
    }
  }

  function renderCategories() {
    const categories = Array.from(
      new Set(menu.map((m) => String(m.category || "").trim()).filter(Boolean))
    ).sort((a, b) => a.localeCompare(b, "zh-Hans-CN"));
    const all = ["全部", ...categories];

    categoryBar.innerHTML = "";
    for (const c of all) {
      const chip = document.createElement("button");
      chip.type = "button";
      chip.className = `chip ${c === activeCategory ? "chip--active" : ""}`;
      chip.textContent = c;
      chip.addEventListener("click", () => {
        activeCategory = c;
        renderMenu();
        renderCategories();
      });
      categoryBar.appendChild(chip);
    }
  }

  function renderMeta() {
    metaDate.value = state.meta.date || "";
    metaLocation.value = state.meta.location || "";
    metaCustomer.value = state.meta.customer || "";
    metaContact.value = state.meta.contact || "";
    metaDiscount.value = String(toNumber(state.meta.discountPercent || 100));
    metaTaxPercent.value = String(toNumber(state.meta.taxPercent || 0));
    metaFinalPrice.value = state.meta.finalPrice || "";
    metaNote.value = state.meta.note || "";
    metaOrderNotes.value = state.meta.orderNotes || "";
    metaOrderNotesTitle.value = state.meta.orderNotesTitle || "";

    qDate.textContent = state.meta.date || "";
    qLocation.textContent = state.meta.location || "";
    qCustomer.textContent = state.meta.customer || "";
    qContact.textContent = state.meta.contact || "";
    qNote.textContent = state.meta.note || "";
    qOrderNotes.textContent = state.meta.orderNotes || "";
    qOrderNotesTitle.textContent = state.meta.orderNotesTitle || "订购说明（可编辑）";
  }

  function renderQuote() {
    // Keep menu lines in sync
    syncQuoteLinesFromSelected();

    quoteTbody.innerHTML = "";
    const lines = state.quoteLines.map((l, idx) => ({ ...l, seq: idx + 1 }));
    const isMobile = document.body.classList.contains("mobile");

    for (const line of lines) {
      if (isMobile) {
        const tr = document.createElement("tr");
        const td = document.createElement("td");
        td.colSpan = 7;

        const card = document.createElement("div");
        card.className = "quoteCard";

        const top = document.createElement("div");
        top.className = "quoteCard__top";

        const left = document.createElement("div");
        left.className = "quoteCard__left";

        const seq = document.createElement("div");
        seq.className = "quoteCard__seq";
        seq.textContent = `#${line.seq}`;

        const imgWrap = document.createElement("div");
        const imgSrc = menuImageForLine(line);
        const thumbSrc = line.source?.startsWith("menu:")
          ? menuThumbForName(line.source.slice("menu:".length))
          : null;
        imgWrap.className = `quoteThumb ${imgSrc ? "" : "quoteThumb--empty"}`.trim();
        if (imgSrc) {
          const img = document.createElement("img");
          img.src = thumbSrc || imgSrc;
          img.alt = line.name || "图片";
          img.loading = "lazy";
          img.decoding = "async";
          img.fetchPriority = "low";
          img.addEventListener("click", () => openImageModal(imgSrc, line.name || ""));
          imgWrap.appendChild(img);
        } else {
          imgWrap.textContent = "-";
        }

        left.appendChild(seq);
        left.appendChild(imgWrap);

        const right = document.createElement("div");
        right.className = "quoteCard__right";

        const nameInput = document.createElement("input");
        nameInput.className = "cellInput";
        nameInput.value = line.name || "";
        nameInput.addEventListener("input", debounce(() => {
          const target = state.quoteLines.find((x) => x.id === line.id);
          if (!target) return;
          target.name = nameInput.value.trim();
          saveState(state);
          renderQuote();
        }, 120));

        const grid = document.createElement("div");
        grid.className = "quoteCard__grid";

        const qtyInput = document.createElement("input");
        qtyInput.className = "cellInput cellInput--num";
        qtyInput.type = "number";
        qtyInput.min = "0";
        qtyInput.step = "1";
        qtyInput.value = String(toNumber(line.qty));
        qtyInput.addEventListener("input", debounce(() => {
          const target = state.quoteLines.find((x) => x.id === line.id);
          if (!target) return;
          target.qty = clamp(toNumber(qtyInput.value), 0, 999999);
          saveState(state);
          renderQuote();
        }, 120));

        const unitInput = document.createElement("input");
        unitInput.className = "cellInput cellInput--num";
        unitInput.type = "number";
        unitInput.min = "0";
        unitInput.step = "0.5";
        unitInput.value = String(toNumber(line.unitPrice));
        unitInput.addEventListener("input", debounce(() => {
          const target = state.quoteLines.find((x) => x.id === line.id);
          if (!target) return;
          target.unitPrice = clamp(toNumber(unitInput.value), 0, 999999);
          saveState(state);
          renderQuote();
        }, 120));

        const total = toNumber(line.qty) * toNumber(line.unitPrice);
        const totalBox = document.createElement("div");
        totalBox.className = "quoteCard__total";
        totalBox.textContent = `总价：${formatMoney(total)}`;

        grid.appendChild(labelWrap("数量", qtyInput));
        grid.appendChild(labelWrap("单价", unitInput));
        grid.appendChild(totalBox);

        const del = document.createElement("button");
        del.className = "iconBtn";
        del.type = "button";
        del.title = "删除此行";
        del.textContent = "×";
        del.addEventListener("click", () => {
          deleteQuoteLine(line.id);
          saveState(state);
          renderAll();
        });

        right.appendChild(nameInput);
        right.appendChild(grid);

        top.appendChild(left);
        top.appendChild(right);
        top.appendChild(del);

        card.appendChild(top);
        td.appendChild(card);
        tr.appendChild(td);
        quoteTbody.appendChild(tr);
        continue;
      }

      const tr = document.createElement("tr");

      const tdSeq = document.createElement("td");
      tdSeq.className = "num";
      tdSeq.textContent = String(line.seq);

      const tdImg = document.createElement("td");
      const imgWrap = document.createElement("div");
      const imgSrc = menuImageForLine(line);
      const thumbSrc = line.source?.startsWith("menu:") ? menuThumbForName(line.source.slice("menu:".length)) : null;
      imgWrap.className = `quoteThumb ${imgSrc ? "" : "quoteThumb--empty"}`.trim();
      if (imgSrc) {
        const img = document.createElement("img");
        img.src = thumbSrc || imgSrc;
        img.alt = line.name || "图片";
        img.loading = "lazy";
        img.decoding = "async";
        img.fetchPriority = "low";
        img.addEventListener("click", () => openImageModal(imgSrc, line.name || ""));
        imgWrap.appendChild(img);
      } else {
        imgWrap.textContent = "-";
      }
      tdImg.appendChild(imgWrap);

      const tdName = document.createElement("td");
      const nameInput = document.createElement("input");
      nameInput.className = "cellInput";
      nameInput.value = line.name || "";
      nameInput.addEventListener("change", () => {
        const target = state.quoteLines.find((x) => x.id === line.id);
        if (!target) return;
        target.name = nameInput.value.trim();
        saveState(state);
        renderQuote();
      });
      tdName.appendChild(nameInput);

      const tdQty = document.createElement("td");
      const qtyInput = document.createElement("input");
      qtyInput.className = "cellInput cellInput--num";
      qtyInput.type = "number";
      qtyInput.min = "0";
      qtyInput.step = "1";
      qtyInput.value = String(toNumber(line.qty));
      qtyInput.addEventListener("change", () => {
        const target = state.quoteLines.find((x) => x.id === line.id);
        if (!target) return;
        target.qty = clamp(toNumber(qtyInput.value), 0, 999999);
        saveState(state);
        renderQuote();
      });
      tdQty.className = "num";
      tdQty.appendChild(qtyInput);

      const tdUnit = document.createElement("td");
      const unitInput = document.createElement("input");
      unitInput.className = "cellInput cellInput--num";
      unitInput.type = "number";
      unitInput.min = "0";
      unitInput.step = "0.5";
      unitInput.value = String(toNumber(line.unitPrice));
      unitInput.addEventListener("change", () => {
        const target = state.quoteLines.find((x) => x.id === line.id);
        if (!target) return;
        target.unitPrice = clamp(toNumber(unitInput.value), 0, 999999);
        saveState(state);
        renderQuote();
      });
      tdUnit.className = "num";
      tdUnit.appendChild(unitInput);

      const tdTotal = document.createElement("td");
      tdTotal.className = "num";
      const total = toNumber(line.qty) * toNumber(line.unitPrice);
      tdTotal.textContent = formatMoney(total);

      const tdAct = document.createElement("td");
      tdAct.className = "no-print";
      const del = document.createElement("button");
      del.className = "iconBtn";
      del.type = "button";
      del.title = "删除此行";
      del.textContent = "×";
      del.addEventListener("click", () => {
        deleteQuoteLine(line.id);
        saveState(state);
        renderAll();
      });
      tdAct.appendChild(del);

      tr.appendChild(tdSeq);
      tr.appendChild(tdImg);
      tr.appendChild(tdName);
      tr.appendChild(tdQty);
      tr.appendChild(tdUnit);
      tr.appendChild(tdTotal);
      tr.appendChild(tdAct);
      quoteTbody.appendChild(tr);
    }

    const subtotal = state.quoteLines.reduce((sum, l) => sum + toNumber(l.qty) * toNumber(l.unitPrice), 0);
    const taxPercent = clamp(toNumber(state.meta.taxPercent || 0), 0, 100);
    const taxIncluded = subtotal * (1 + taxPercent / 100);
    const discountPercent = clamp(toNumber(state.meta.discountPercent || 100), 0, 100);
    const computedAfterDiscount = subtotal * (discountPercent / 100);
    const manualFinal = toNumber(state.meta.finalPrice);
    const totalAfterDiscount = manualFinal > 0 ? manualFinal : computedAfterDiscount;

    subtotalEl.textContent = formatMoney(subtotal);
    taxIncludedEl.textContent = formatMoney(taxIncluded);
    totalAfterDiscountEl.textContent = formatMoney(totalAfterDiscount);
  }

  function labelWrap(label, input) {
    const wrap = document.createElement("label");
    wrap.className = "quoteCard__field";
    const l = document.createElement("div");
    l.className = "quoteCard__label";
    l.textContent = label;
    wrap.appendChild(l);
    wrap.appendChild(input);
    return wrap;
  }

  function renderAll() {
    renderMeta();
    renderCategories();
    renderMenu();
    renderQuote();
  }

  function applySidebarWidth(px) {
    const w = clamp(toNumber(px), 320, 720);
    document.documentElement.style.setProperty("--sidebarW", `${w}px`);
    localStorage.setItem(SIDEBAR_W_KEY, String(w));
  }

  // init sidebar width from storage
  try {
    const stored = localStorage.getItem(SIDEBAR_W_KEY);
    if (stored) applySidebarWidth(stored);
  } catch {
    // ignore
  }

  // drag to resize sidebar (desktop only)
  if (sidebarResizer) {
    sidebarResizer.addEventListener("pointerdown", (e) => {
      sidebarResizer.setPointerCapture(e.pointerId);
      const startX = e.clientX;
      const startW =
        parseFloat(getComputedStyle(document.documentElement).getPropertyValue("--sidebarW")) || 420;

      const onMove = (ev) => {
        const dx = ev.clientX - startX;
        applySidebarWidth(startW + dx);
      };
      const onUp = () => {
        sidebarResizer.removeEventListener("pointermove", onMove);
        sidebarResizer.removeEventListener("pointerup", onUp);
        sidebarResizer.removeEventListener("pointercancel", onUp);
      };
      sidebarResizer.addEventListener("pointermove", onMove);
      sidebarResizer.addEventListener("pointerup", onUp);
      sidebarResizer.addEventListener("pointercancel", onUp);
    });
  }

  function onMetaChange() {
    state.meta.date = metaDate.value.trim();
    state.meta.location = metaLocation.value.trim();
    state.meta.customer = metaCustomer.value.trim();
    state.meta.contact = metaContact.value.trim();
    state.meta.discountPercent = clamp(toNumber(metaDiscount.value), 0, 100);
    state.meta.taxPercent = clamp(toNumber(metaTaxPercent.value), 0, 100);
    state.meta.finalPrice = metaFinalPrice.value.trim();
    state.meta.note = metaNote.value.trim();
    state.meta.orderNotes = metaOrderNotes.value || "";
    state.meta.orderNotesTitle = metaOrderNotesTitle.value.trim() || "";
    saveState(state);
    renderMeta();
    renderQuote();
  }

  // iOS 上很多输入场景不会触发 change（不失焦），用 input + debounce 确保“最新版”自动保存
  const onMetaInput = debounce(onMetaChange, 120);

  for (const t of tabs) {
    t.addEventListener("click", () => setActiveTab(t.dataset.tab));
  }

  menuSearch.addEventListener("input", renderMenu);
  onlySelected.addEventListener("change", renderMenu);

  for (const node of [
    metaDate,
    metaLocation,
    metaCustomer,
    metaContact,
    metaDiscount,
    metaTaxPercent,
    metaFinalPrice,
    metaNote,
    metaOrderNotes,
    metaOrderNotesTitle,
  ]) {
    node.addEventListener("change", onMetaChange);
    node.addEventListener("input", onMetaInput);
  }

  qOrderNotes.addEventListener("input", () => {
    state.meta.orderNotes = qOrderNotes.textContent || "";
    metaOrderNotes.value = state.meta.orderNotes;
    saveState(state);
  });

  qOrderNotesTitle.addEventListener("input", () => {
    state.meta.orderNotesTitle = qOrderNotesTitle.textContent || "";
    metaOrderNotesTitle.value = state.meta.orderNotesTitle;
    saveState(state);
  });

  btnExportCsv.addEventListener("click", exportTable);
  btnScreenshot.addEventListener("click", exportScreenshotPng);

  btnExportImage.addEventListener("click", exportLongPng);

  modalClose.addEventListener("click", closeImageModal);
  imageModal.addEventListener("click", (e) => {
    const target = e.target;
    if (target && target.dataset && target.dataset.close === "1") closeImageModal();
  });
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape" && imageModal.getAttribute("aria-hidden") === "false") closeImageModal();
  });

  btnReset.addEventListener("click", () => {
    if (!confirm("确定要清空已选产品、报价行和表头信息吗？")) return;
    state = buildInitialState();
    saveState(state);
    renderAll();
  });

  btnLoadTemplate.addEventListener("click", () => {
    // Load rows from 表2 as editable quote lines (does not overwrite selected menu)
    const rows = template
      .slice()
      .sort((a, b) => toNumber(a.seq) - toNumber(b.seq))
      .map((r) => ({
        id: state.nextId++,
        source: "template",
        name: String(r.name || "").trim(),
        qty: toNumber(r.qty),
        unitPrice: toNumber(r.unitPrice),
      }));
    const keep = state.quoteLines.filter((l) => l.source?.startsWith("menu:"));
    state.quoteLines = rows.concat(keep);
    saveState(state);
    setActiveTab("quote");
    renderAll();
  });

  // Migration: if older state used a dedicated service_fee line, convert to menu-backed selection.
  const legacy = state.quoteLines.find((l) => l.source === "service_fee");
  if (legacy) {
    state.quoteLines = state.quoteLines.filter((l) => l.source !== "service_fee");
    state.selected = state.selected || {};
    state.selected[SERVICE_FEE_NAME] = {
      name: SERVICE_FEE_NAME,
      qty: toNumber(legacy.qty) || 1,
      unitPrice: toNumber(legacy.unitPrice) || 0,
      minOrder: "1项",
      category: "服务",
    };
    syncQuoteLinesFromSelected();
    saveState(state);
  }

  renderAll();
}

document.addEventListener("DOMContentLoaded", init);
