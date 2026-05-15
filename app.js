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
    selectedStyles: [], // [{ id, category, image, imageThumb }]
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

function escapeHtml(s) {
  return String(s || "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function init() {
  const menu = [
    SERVICE_FEE_MENU_ITEM,
    ...(Array.isArray(window.MENU_DATA) ? window.MENU_DATA : []),
  ];
  const template = Array.isArray(window.TEMPLATE_QUOTE_DATA) ? window.TEMPLATE_QUOTE_DATA : [];

  let state = loadState() || buildInitialState();
  let activeCategory = "全部";

  const CUSTOM_MENU_STORAGE_KEY = "dangxia_custom_menu_v1";
  const customMenu = (() => {
    try {
      const raw = localStorage.getItem(CUSTOM_MENU_STORAGE_KEY);
      const parsed = raw ? JSON.parse(raw) : null;
      return Array.isArray(parsed) ? parsed : [];
    } catch {
      return [];
    }
  })();

  const STYLE_STORAGE_KEY = "dangxia_styles_v1";
  const styleLibrary = (() => {
    try {
      const raw = localStorage.getItem(STYLE_STORAGE_KEY);
      const parsed = raw ? JSON.parse(raw) : null;
      return Array.isArray(parsed) ? parsed : [];
    } catch {
      return [];
    }
  })();

  const saveStyleLibrary = () => {
    try {
      localStorage.setItem(STYLE_STORAGE_KEY, JSON.stringify(styleLibrary));
    } catch {}
  };

  // Ensure meta keys exist
  state.meta = { ...defaultMeta(), ...(state.meta || {}) };
  state.selected = state.selected || {};
  state.quoteLines = state.quoteLines || [];
  state.selectedStyles = Array.isArray(state.selectedStyles) ? state.selectedStyles : [];
  state.nextId = state.nextId || 1;

  // UI refs
  const tabs = Array.from(document.querySelectorAll(".tab"));
  const panes = Array.from(document.querySelectorAll(".tabpane"));
  const menuList = el("#menuList");
  const menuSearch = el("#menuSearch");
  const onlySelected = el("#onlySelected");
  const menuCount = el("#menuCount");
  const categoryBar = el("#categoryBar");
  const btnAddCustom = el("#btnAddCustom");

  // Style tab refs
  const styleList = el("#styleList");
  const styleCount = el("#styleCount");
  const btnAddStyle = el("#btnAddStyle");

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
  const quotePaper = el("#quotePaper");
  const qStyles = el("#qStyles");

  const quoteTbody = el("#quoteTbody");
  const subtotalEl = el("#subtotal");
  const taxIncludedEl = el("#taxIncluded");
  const totalAfterDiscountEl = el("#totalAfterDiscount");

  const sidebarResizer = document.querySelector("#sidebarResizer");

  const imageModal = el("#imageModal");
  const modalClose = el("#modalClose");
  const modalImg = el("#modalImg");
  const modalCaption = el("#modalCaption");

  // Custom item modal refs (optional)
  const customModal = el("#customModal");
  const customName = el("#customName");
  const customCategory = el("#customCategory");
  const customPrice = el("#customPrice");
  const customMinOrder = el("#customMinOrder");
  const customFile = el("#customFile");
  const customCategoryList = el("#customCategoryList");
  const cropWrap = el("#cropWrap");
  const cropStage = el("#cropStage");
  const cropImg = el("#cropImg");
  const cropBox = el("#cropBox");
  const cropHandle = el("#cropHandle");
  const btnCustomCancel = el("#btnCustomCancel");
  const btnCustomSave = el("#btnCustomSave");

  // Style upload modal refs (optional)
  const styleModal = el("#styleModal");
  const styleFile = el("#styleFile");
  const styleBatchList = el("#styleBatchList");
  // (batch editor renders per-item controls dynamically)
  const btnStyleCancel = el("#btnStyleCancel");
  const btnStyleSave = el("#btnStyleSave");
  const btnStyleAutoCropAll = el("#btnStyleAutoCropAll");

  const closeStyleModal = () => styleModal?.setAttribute("aria-hidden", "true");
  const openStyleModal = () => styleModal?.setAttribute("aria-hidden", "false");

  const closeCustomModal = () => customModal?.setAttribute("aria-hidden", "true");
  const openCustomModal = () => customModal?.setAttribute("aria-hidden", "false");

  const saveCustomMenu = () => {
    try {
      localStorage.setItem(CUSTOM_MENU_STORAGE_KEY, JSON.stringify(customMenu));
    } catch {}
  };

  const findMenuItemByName = (name) => [...menu, ...customMenu].find((m) => m.name === name) || null;

  // Crop state (square, cover)
  const cropState = {
    ready: false,
    imgNaturalW: 0,
    imgNaturalH: 0,
    tx: 0,
    ty: 0,
    scale: 1,
    dragging: false,
    lastX: 0,
    lastY: 0,
    boxSize: 0,
    boxCx: 0,
    boxCy: 0,
    boxDragging: false,
    boxResizing: false,
  };

  const createCropper = ({ stageEl, imgEl, boxEl, handleEl, aspect = 1 }) => {
    const st = {
      ready: false,
      imgNaturalW: 0,
      imgNaturalH: 0,
      tx: 0,
      ty: 0,
      scale: 1,
      lastX: 0,
      lastY: 0,
      draggingImg: false,
      boxSize: 0,
      boxCx: 0,
      boxCy: 0,
      boxDragging: false,
      boxResizing: false,
      aspect,
    };

    const applyBoxLayout = () => {
      if (!stageEl || !boxEl) return;
      const stageRect = stageEl.getBoundingClientRect();
      const maxW = stageRect.width;
      const maxH = stageRect.height;
      const minW = 120;
      const maxBoxW = Math.min(maxW, maxH * st.aspect);
      const w = clamp(st.boxSize || Math.floor(maxBoxW * 0.78), minW, maxBoxW);
      st.boxSize = w;
      const h = w / st.aspect;
      const halfW = w / 2;
      const halfH = h / 2;
      st.boxCx = clamp(st.boxCx || maxW / 2, halfW, maxW - halfW);
      st.boxCy = clamp(st.boxCy || maxH / 2, halfH, maxH - halfH);
      boxEl.style.width = `${w}px`;
      boxEl.style.height = `${h}px`;
      boxEl.style.left = `${st.boxCx}px`;
      boxEl.style.top = `${st.boxCy}px`;
      boxEl.style.transform = "translate(-50%, -50%)";
    };

    const setImgTransform = () => {
      if (!imgEl) return;
      imgEl.style.transform = `translate(-50%, -50%) translate(${st.tx}px, ${st.ty}px) scale(${st.scale})`;
    };

    const resetToCover = () => {
      if (!stageEl || !boxEl || !imgEl) return;
      applyBoxLayout();
      const w = st.boxSize;
      const h = w / st.aspect;
      const iw = st.imgNaturalW || 1;
      const ih = st.imgNaturalH || 1;
      st.scale = Math.max(w / iw, h / ih);
      st.tx = 0;
      st.ty = 0;
      setImgTransform();
    };

    const clampImgToCoverBox = () => {
      if (!boxEl) return;
      const w = st.boxSize || boxEl.getBoundingClientRect().width;
      const h = w / st.aspect;
      const iw = st.imgNaturalW * st.scale;
      const ih = st.imgNaturalH * st.scale;
      const maxX = Math.max(0, iw / 2 - w / 2);
      const maxY = Math.max(0, ih / 2 - h / 2);
      st.tx = clamp(st.tx, -maxX, maxX);
      st.ty = clamp(st.ty, -maxY, maxY);
    };

    const cropToDataUrl = (outW, outH) => {
      if (!stageEl || !boxEl || !imgEl) return null;
      if (!st.ready) return null;
      const stageRect = stageEl.getBoundingClientRect();
      const w = st.boxSize || boxEl.getBoundingClientRect().width;
      const h = w / st.aspect;
      const cx = st.boxCx || stageRect.width / 2;
      const cy = st.boxCy || stageRect.height / 2;
      const stageCenterX = stageRect.width / 2;
      const stageCenterY = stageRect.height / 2;

      const iw = st.imgNaturalW;
      const ih = st.imgNaturalH;
      const s = st.scale;

      const srcLeft = (cx - w / 2 - stageCenterX - st.tx) / s + iw / 2;
      const srcTop = (cy - h / 2 - stageCenterY - st.ty) / s + ih / 2;
      const srcW = w / s;
      const srcH = h / s;

      const canvas = document.createElement("canvas");
      canvas.width = outW;
      canvas.height = outH;
      const ctx = canvas.getContext("2d");
      ctx.fillStyle = "#fff";
      ctx.fillRect(0, 0, outW, outH);
      ctx.imageSmoothingEnabled = true;
      ctx.imageSmoothingQuality = "high";
      ctx.drawImage(imgEl, srcLeft, srcTop, srcW, srcH, 0, 0, outW, outH);
      return canvas.toDataURL("image/jpeg", 0.9);
    };

    const clear = () => {
      st.ready = false;
      st.tx = 0;
      st.ty = 0;
      st.scale = 1;
      st.boxSize = 0;
      st.boxCx = 0;
      st.boxCy = 0;
      st.boxDragging = false;
      st.boxResizing = false;
      st.draggingImg = false;
    };

    const bind = () => {
      if (!stageEl || !imgEl || !boxEl) return;

      const imgDown = (ev) => {
        if (!st.ready) return;
        st.draggingImg = true;
        const p = ev.touches?.[0] || ev;
        st.lastX = p.clientX;
        st.lastY = p.clientY;
      };
      stageEl.addEventListener("mousedown", imgDown);
      stageEl.addEventListener("touchstart", imgDown, { passive: true });

      const boxDown = (ev) => {
        if (!st.ready) return;
        st.boxDragging = true;
        const p = ev.touches?.[0] || ev;
        st.lastX = p.clientX;
        st.lastY = p.clientY;
        ev.preventDefault?.();
        ev.stopPropagation?.();
      };
      boxEl.addEventListener("mousedown", boxDown);
      boxEl.addEventListener("touchstart", boxDown, { passive: false });

      const handleDown = (ev) => {
        if (!st.ready) return;
        st.boxResizing = true;
        const p = ev.touches?.[0] || ev;
        st.lastX = p.clientX;
        st.lastY = p.clientY;
        ev.preventDefault?.();
        ev.stopPropagation?.();
      };
      handleEl?.addEventListener("mousedown", handleDown);
      handleEl?.addEventListener("touchstart", handleDown, { passive: false });

      const move = (ev) => {
        if (!st.draggingImg && !st.boxDragging && !st.boxResizing) return;
        const p = ev.touches?.[0] || ev;
        const dx = p.clientX - st.lastX;
        const dy = p.clientY - st.lastY;
        st.lastX = p.clientX;
        st.lastY = p.clientY;

        if (st.draggingImg) {
          st.tx += dx;
          st.ty += dy;
          clampImgToCoverBox();
          setImgTransform();
        } else if (st.boxDragging) {
          st.boxCx += dx;
          st.boxCy += dy;
          applyBoxLayout();
          clampImgToCoverBox();
          setImgTransform();
        } else if (st.boxResizing) {
          const d = (dx + dy) / 2;
          st.boxSize = (st.boxSize || 0) + d * 2;
          applyBoxLayout();
          const w = st.boxSize;
          const h = w / st.aspect;
          const iw = st.imgNaturalW || 1;
          const ih = st.imgNaturalH || 1;
          const minScale = Math.max(w / iw, h / ih);
          if (st.scale < minScale) st.scale = minScale;
          clampImgToCoverBox();
          setImgTransform();
        }
        ev.preventDefault?.();
      };
      const up = () => {
        st.draggingImg = false;
        st.boxDragging = false;
        st.boxResizing = false;
      };

      window.addEventListener("mousemove", move, { passive: false });
      window.addEventListener("touchmove", move, { passive: false });
      window.addEventListener("mouseup", up);
      window.addEventListener("touchend", up);

      stageEl.addEventListener(
        "wheel",
        (ev) => {
          if (!st.ready) return;
          const delta = ev.deltaY || 0;
          const factor = delta > 0 ? 0.92 : 1.08;
          st.scale = clamp(st.scale * factor, 0.2, 10);
          clampImgToCoverBox();
          setImgTransform();
          ev.preventDefault();
        },
        { passive: false }
      );
    };

    bind();

    return {
      state: st,
      applyBoxLayout,
      setImgTransform,
      resetToCover,
      clampImgToCoverBox,
      cropToDataUrl,
      setAspect: (nextAspect) => {
        const a = Number(nextAspect);
        if (!Number.isFinite(a) || a <= 0) return;
        st.aspect = a;
        // Re-layout box and ensure image still covers it.
        applyBoxLayout();
        resetToCover();
        clampImgToCoverBox();
        setImgTransform();
      },
      clear,
    };
  };

  const applyCropBoxLayout = () => {
    if (!cropStage || !cropBox) return;
    const stageRect = cropStage.getBoundingClientRect();
    const minSize = 80;
    const maxSize = Math.min(stageRect.width, stageRect.height);
    const size = clamp(cropState.boxSize || Math.floor(maxSize * 0.56), minSize, maxSize);
    cropState.boxSize = size;
    const half = size / 2;
    cropState.boxCx = clamp(cropState.boxCx || stageRect.width / 2, half, stageRect.width - half);
    cropState.boxCy = clamp(cropState.boxCy || stageRect.height / 2, half, stageRect.height - half);
    cropBox.style.width = `${size}px`;
    cropBox.style.height = `${size}px`;
    cropBox.style.left = `${cropState.boxCx}px`;
    cropBox.style.top = `${cropState.boxCy}px`;
    cropBox.style.transform = "translate(-50%, -50%)";
  };

  const setCropTransform = () => {
    if (!cropImg) return;
    cropImg.style.transform = `translate(-50%, -50%) translate(${cropState.tx}px, ${cropState.ty}px) scale(${cropState.scale})`;
  };

  const resetCropToCover = () => {
    if (!cropStage || !cropBox || !cropImg) return;
    applyCropBoxLayout();
    const boxSize = cropState.boxSize || Math.min(cropBox.getBoundingClientRect().width, cropBox.getBoundingClientRect().height);
    const iw = cropState.imgNaturalW || 1;
    const ih = cropState.imgNaturalH || 1;
    cropState.scale = Math.max(boxSize / iw, boxSize / ih);
    cropState.tx = 0;
    cropState.ty = 0;
    setCropTransform();
  };

  const clampCrop = () => {
    if (!cropBox) return;
    const boxSize = cropState.boxSize || Math.min(cropBox.getBoundingClientRect().width, cropBox.getBoundingClientRect().height);
    const iw = cropState.imgNaturalW * cropState.scale;
    const ih = cropState.imgNaturalH * cropState.scale;
    const halfBox = boxSize / 2;
    const maxX = Math.max(0, iw / 2 - halfBox);
    const maxY = Math.max(0, ih / 2 - halfBox);
    cropState.tx = clamp(cropState.tx, -maxX, maxX);
    cropState.ty = clamp(cropState.ty, -maxY, maxY);
  };

  const cropToThumbDataUrl = () => {
    if (!cropStage || !cropBox || !cropImg) return null;
    if (!cropState.ready) return null;

    const stageRect = cropStage.getBoundingClientRect();
    const boxSize = cropState.boxSize || Math.min(cropBox.getBoundingClientRect().width, cropBox.getBoundingClientRect().height);
    const cx = cropState.boxCx || stageRect.width / 2;
    const cy = cropState.boxCy || stageRect.height / 2;
    const stageCenterX = stageRect.width / 2;
    const stageCenterY = stageRect.height / 2;

    const iw = cropState.imgNaturalW;
    const ih = cropState.imgNaturalH;
    const s = cropState.scale;

    const srcLeft = (cx - boxSize / 2 - stageCenterX - cropState.tx) / s + iw / 2;
    const srcTop = (cy - boxSize / 2 - stageCenterY - cropState.ty) / s + ih / 2;
    const srcSize = boxSize / s;

    const canvas = document.createElement("canvas");
    const outSize = 224; // 56px * 4
    canvas.width = outSize;
    canvas.height = outSize;
    const ctx = canvas.getContext("2d");
    ctx.fillStyle = "#fff";
    ctx.fillRect(0, 0, outSize, outSize);
    ctx.imageSmoothingEnabled = true;
    ctx.imageSmoothingQuality = "high";
    ctx.drawImage(cropImg, srcLeft, srcTop, srcSize, srcSize, 0, 0, outSize, outSize);
    return canvas.toDataURL("image/jpeg", 0.88);
  };

  const btnExportImage = el("#btnExportImage");
  const btnExportCsv = el("#btnExportCsv");
  const btnExportPdf = el("#btnExportPdf");
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

      function dataUrlToBase64(dataUrl) {
        const s = String(dataUrl || "");
        const idx = s.indexOf("base64,");
        if (idx < 0) return "";
        return s.slice(idx + "base64,".length);
      }

      async function toArrayBufferFromUrl(url) {
        const abs = new URL(url, location.href).toString();
        const res = await fetch(abs, { cache: "force-cache" });
        const buf = await res.arrayBuffer();
        return buf;
      }

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
            ? `<div style="width:48px;height:48px;overflow:hidden;margin:0 auto;">
                 <img src="${src}" width="48" height="48" style="display:block;border:1px solid rgba(0,0,0,.12);background:#fff;mso-width-source:userset;mso-height-source:userset;" />
               </div>`
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
      .img{width:74px; text-align:center; line-height:0; padding-top:6px; padding-bottom:6px;}
      .img div{display:block;}
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
          ${rowsHtml}
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

      // 真正的 .xlsx（带图片嵌入单元格）— iOS/电脑都一致可用
      const ExcelJS = window.ExcelJS;
      if (!ExcelJS) {
        alert("导出表格组件未加载（ExcelJS）。将改用兼容模式导出。");
        downloadText(`${exportName}.xls`, "application/vnd.ms-excel;charset=utf-8", html);
        return;
      }

      const wb = new ExcelJS.Workbook();
      wb.creator = "Dangxia Quote Web";
      wb.created = new Date();
      const ws = wb.addWorksheet("报价单", {
        properties: { defaultRowHeight: 18 },
        pageSetup: { paperSize: 9, orientation: "portrait" },
      });
      const BRAND_GREEN = "FFA4CC9A";
      const WHITE = "FFFFFFFF";

      // Columns
      ws.columns = [
        { header: "序号", key: "seq", width: 6 },
        { header: "图片", key: "img", width: 12 },
        { header: "内容", key: "name", width: 34 },
        { header: "数量", key: "qty", width: 8 },
        { header: "单价", key: "unit", width: 10 },
        { header: "总价", key: "total", width: 12 },
      ];

      // Header row (logo + title)
      ws.mergeCells("A1:F1");
      const titleCell = ws.getCell("A1");
      titleCell.value = "【当夏烘焙】甜品台服务 报价单";
      titleCell.font = { name: "PingFang SC", size: 16, bold: true, color: { argb: "FFFFFFFF" } };
      titleCell.alignment = { vertical: "middle", horizontal: "center" };
      titleCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: BRAND_GREEN } };
      ws.getRow(1).height = 34;
      // 标题行底色需覆盖整行
      for (let c = 1; c <= 6; c += 1) {
        const cell = ws.getCell(1, c);
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: BRAND_GREEN } };
      }

      // Add logo at title row (right)
      try {
        const logoBuf = await toArrayBufferFromUrl("./assets/brand/logo.png");
        const logoId = wb.addImage({ buffer: logoBuf, extension: "png" });
        // place near column F, row 1 right side
        ws.addImage(logoId, { tl: { col: 5.15, row: 0.15 }, ext: { width: 54, height: 40 }, editAs: "oneCell" });
      } catch {
        // ignore
      }

      // Meta
      let r = 3;
      const metaPairs = [
        ["时间", state.meta.date || ""],
        ["地点", state.meta.location || ""],
        ["客户", state.meta.customer || ""],
        ["联系人", state.meta.contact || ""],
        ["备注", state.meta.note || ""],
      ];
      for (const [k, v] of metaPairs) {
        ws.mergeCells(`A${r}:B${r}`);
        ws.mergeCells(`C${r}:F${r}`);
        ws.getCell(`A${r}`).value = `${k}：`;
        ws.getCell(`A${r}`).font = { bold: true, color: { argb: "FF6B7280" } };
        ws.getCell(`A${r}`).alignment = { vertical: "middle", horizontal: "left" };
        ws.getCell(`C${r}`).value = String(v || "");
        ws.getCell(`C${r}`).alignment = { vertical: "middle", horizontal: "left" };
        // ensure white background
        for (let c = 1; c <= 6; c += 1) ws.getCell(r, c).fill = { type: "pattern", pattern: "solid", fgColor: { argb: WHITE } };
        r += 1;
      }

      r += 1;

      // Table header
      const headerRowIndex = r;
      const headerRow = ws.getRow(r);
      headerRow.values = ["序号", "图片", "内容", "数量", "单价", "总价"];
      headerRow.font = { bold: true, color: { argb: "FFFFFFFF" } };
      headerRow.alignment = { vertical: "middle", horizontal: "left" };
      headerRow.height = 20;
      headerRow.eachCell((cell) => {
        // 表头底色品牌绿
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: BRAND_GREEN } };
        cell.border = { bottom: { style: "thin", color: { argb: "FFCBD5E1" } } };
      });
      r += 1;

      // Load thumbs as buffers
      const imgBuffers = await Promise.all(
        imgUrls.map(async (u, idx) => {
          if (!u) return null;
          try {
            if (inlineImages && imgDataUrls[idx]) {
              const b64 = dataUrlToBase64(imgDataUrls[idx]);
              return { base64: b64, extension: "png" };
            }
            const buf = await toArrayBufferFromUrl(u);
            // guess png/jpg
            const ext = String(u).toLowerCase().includes(".jpg") || String(u).toLowerCase().includes(".jpeg")
              ? "jpeg"
              : "png";
            return { buffer: buf, extension: ext };
          } catch {
            return null;
          }
        })
      );

      // Rows with images anchored to cell B
      for (let i = 0; i < tableLines.length; i += 1) {
        const l = tableLines[i];
        const qty = toNumber(l.qty);
        const unit = toNumber(l.unitPrice);
        const total = qty * unit;
        const row = ws.getRow(r);
        row.getCell(1).value = l.seq;
        row.getCell(3).value = l.name || "";
        row.getCell(4).value = qty;
        row.getCell(5).value = unit;
        row.getCell(6).value = total;
        row.height = 44;
        row.alignment = { vertical: "middle" };
        // borders
        for (let c = 1; c <= 6; c += 1) {
          const cell = row.getCell(c);
          cell.border = { bottom: { style: "thin", color: { argb: "FFE5E7EB" } } };
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: WHITE } };
          if (c === 4 || c === 5 || c === 6) cell.alignment = { vertical: "middle", horizontal: "left" };
          if (c === 1) cell.alignment = { vertical: "middle", horizontal: "left" };
        }
        // image in column B (cell 2)
        const imgSpec = imgBuffers[i];
        if (imgSpec) {
          const imgId = wb.addImage(imgSpec);
          ws.addImage(imgId, {
            tl: { col: 1 + 0.2, row: r - 1 + 0.15 },
            ext: { width: 40, height: 40 },
            editAs: "oneCell",
          });
        }
        r += 1;
      }

      r += 1;
      // Totals
      ws.mergeCells(`A${r}:E${r}`);
      ws.getCell(`A${r}`).value = "合计：";
      ws.getCell(`A${r}`).font = { bold: true, color: { argb: "FF6B7280" } };
      ws.getCell(`A${r}`).alignment = { horizontal: "left", vertical: "middle" };
      ws.getCell(`F${r}`).value = subtotal;
      ws.getCell(`F${r}`).numFmt = "0.00";
      ws.getCell(`F${r}`).font = { bold: true };
      ws.getCell(`F${r}`).alignment = { horizontal: "left", vertical: "middle" };
      for (let c = 1; c <= 6; c += 1) ws.getCell(r, c).fill = { type: "pattern", pattern: "solid", fgColor: { argb: WHITE } };
      r += 1;

      ws.mergeCells(`A${r}:E${r}`);
      ws.getCell(`A${r}`).value = "含税价：";
      ws.getCell(`A${r}`).font = { bold: true, color: { argb: "FF6B7280" } };
      ws.getCell(`A${r}`).alignment = { horizontal: "left", vertical: "middle" };
      ws.getCell(`F${r}`).value = taxIncluded;
      ws.getCell(`F${r}`).numFmt = "0.00";
      ws.getCell(`F${r}`).font = { bold: true };
      ws.getCell(`F${r}`).alignment = { horizontal: "left", vertical: "middle" };
      for (let c = 1; c <= 6; c += 1) ws.getCell(r, c).fill = { type: "pattern", pattern: "solid", fgColor: { argb: WHITE } };
      r += 1;

      ws.mergeCells(`A${r}:E${r}`);
      ws.getCell(`A${r}`).value = "优惠价：";
      ws.getCell(`A${r}`).font = { bold: true, color: { argb: "FFDC2626" } };
      ws.getCell(`A${r}`).alignment = { horizontal: "left", vertical: "middle" };
      ws.getCell(`F${r}`).value = totalAfterDiscount;
      ws.getCell(`F${r}`).numFmt = "0.00";
      ws.getCell(`F${r}`).font = { bold: true, color: { argb: "FFDC2626" } };
      ws.getCell(`F${r}`).alignment = { horizontal: "left", vertical: "middle" };
      for (let c = 1; c <= 6; c += 1) ws.getCell(r, c).fill = { type: "pattern", pattern: "solid", fgColor: { argb: WHITE } };
      r += 2;

      // Notes
      ws.mergeCells(`A${r}:F${r}`);
      ws.getCell(`A${r}`).value = state.meta.orderNotesTitle || "订购说明";
      ws.getCell(`A${r}`).font = { bold: true };
      ws.getCell(`A${r}`).fill = { type: "pattern", pattern: "solid", fgColor: { argb: WHITE } };
      r += 1;
      const noteLines = String(state.meta.orderNotes || "")
        .split("\n")
        .map((x) => x.trimEnd())
        .filter((x) => x.length > 0);
      if (noteLines.length === 0) noteLines.push("");
      for (const line of noteLines) {
        ws.mergeCells(`A${r}:F${r}`);
        ws.getCell(`A${r}`).value = line;
        ws.getCell(`A${r}`).alignment = { wrapText: true, vertical: "top" };
        ws.getCell(`A${r}`).fill = { type: "pattern", pattern: "solid", fgColor: { argb: WHITE } };
        ws.getRow(r).height = 18;
        r += 1;
      }

      // 统一底色：除标题行/表头行外，其他所有单元格强制白色（避免透明导致显示成其他底色）
      const lastRowIndex = Math.max(1, r - 1);
      for (let rowIndex = 1; rowIndex <= lastRowIndex; rowIndex += 1) {
        for (let colIndex = 1; colIndex <= 6; colIndex += 1) {
          const cell = ws.getCell(rowIndex, colIndex);
          const isGreenRow = rowIndex === 1 || rowIndex === headerRowIndex;
          cell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: isGreenRow ? BRAND_GREEN : WHITE },
          };
        }
      }

      const buf = await wb.xlsx.writeBuffer();
      const blob = new Blob([buf], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });

      if (isMobile && navigator?.share) {
        const file = new File([blob], `${exportName}.xlsx`, { type: blob.type });
        await navigator.share({ files: [file], title: exportName });
        return;
      }

      // Desktop / fallback download
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `${exportName}.xlsx`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 30_000);
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
      const titleText = "【当夏烘焙】甜品台服务 报价单";
      const titleMetrics = ctx.measureText(titleText);
      const titleX = Math.max(pad, Math.round((width - titleMetrics.width) / 2));
      ctx.fillText(titleText, titleX, 30);
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

          // 数字列改为左对齐
          ctx.textAlign = "left";
          ctx.fillText(String(qty), width - pad - 190, cy + 12);
          ctx.fillText(formatMoney(unit), width - pad - 130, cy + 12);
          ctx.fillText(formatMoney(total), width - pad - 70, cy + 12);
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
    ctx.textAlign = "left";
    ctx.fillText("合计：", width - pad - 190, y);
    ctx.fillStyle = brand;
    ctx.fillText(formatMoney(subtotal), width - pad - 130, y);
    y += 20;
    ctx.fillStyle = muted;
    ctx.fillText("含税价：", width - pad - 190, y);
    ctx.fillStyle = brand;
    ctx.fillText(formatMoney(taxIncluded), width - pad - 130, y);
    y += 20;
    ctx.fillStyle = muted;
    ctx.fillText("优惠价：", width - pad - 190, y);
    ctx.fillStyle = "#dc2626";
    ctx.fillText(formatMoney(totalAfterDiscount), width - pad - 130, y);
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
    const rawNotes = state.meta.orderNotes || "";
    // 分行渲染，方便 PDF 分页断点对齐，避免整段被切割
    notes.innerHTML = "";
    for (const line of String(rawNotes).split("\n")) {
      const div = document.createElement("div");
      div.textContent = line;
      notes.appendChild(div);
    }
    foot.appendChild(notes);

    wrapper.appendChild(foot);

    // Styles appended below order notes (same width as quote paper)
    if (Array.isArray(state.selectedStyles) && state.selectedStyles.length > 0) {
      const stylesWrap = document.createElement("div");
      stylesWrap.className = "quoteStyles";
      const ids = state.selectedStyles.map((x) => x.id);
      const selected = ids.map((id) => styleLibrary.find((s) => s.id === id)).filter(Boolean);
      for (const s of selected) {
        const item = document.createElement("div");
        item.className = "quoteStyles__item";
        const fullSrc = s.imageFull || s.image;
        if (fullSrc) {
          const img = document.createElement("img");
          img.className = "quoteStyles__img";
          img.src = fullSrc;
          img.alt = s.category || "摆台风格";
          item.appendChild(img);
        }
        const cap = document.createElement("div");
        cap.className = "quoteStyles__cap";
        cap.innerHTML = `<span class="quoteStyles__tag">摆台风格</span><span>${escapeHtml(s.category || "")}</span>`;
        item.appendChild(cap);
        stylesWrap.appendChild(item);
      }
      wrapper.appendChild(stylesWrap);
    }

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

  async function exportPdfA4OnePage(opts = {}) {
    const exportName = buildExportName();
    const oldText = btnExportPdf.textContent;
    btnExportPdf.disabled = true;
    btnExportPdf.textContent = "生成中...";

    try {
      const watermarkUrl = await (async () => {
        // Watermark is implemented as CSS background so it only appears in print/PDF
        // when the user enables browser print option “背景图形 / Background graphics”.
        const toDataUrl = (blob) =>
          new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = () => resolve(String(reader.result || ""));
            reader.onerror = () => reject(new Error("read watermark failed"));
            reader.readAsDataURL(blob);
          });

        const candidates = [
          "./assets/brand/watermark.png",
          "assets/brand/watermark.png",
          // fallback to absolute based on current origin/path
          new URL("./assets/brand/watermark.png", location.href).toString(),
        ];

        for (const url of candidates) {
          try {
            const res = await fetch(url, { cache: "force-cache" });
            if (!res.ok) continue;
            const blob = await res.blob();
            const dataUrl = await toDataUrl(blob);
            if (dataUrl.startsWith("data:image/")) return dataUrl;
          } catch {}
        }
        // If fetch/dataURL fails, fall back to relative URL (may still work).
        return "./assets/brand/watermark.png";
      })();

      // A4 at 96dpi (approx). Fit width, then paginate vertically if needed.
      const A4_W = 794;
      const A4_H = 1123;
      const MARGIN_X = 48; // left/right
      const MARGIN_TOP = 24; // reduce top whitespace by half
      const MARGIN_BOTTOM = 48;
      const contentW = A4_W - MARGIN_X * 2;

      const node = buildExportNode(contentW);
      // Optional: hide top-left logo in PDF export
      if (!opts.showLogo) {
        const logo = node.querySelector(".paper__logo");
        if (logo) logo.style.display = "none";
      }
      // Ensure latest render
      // (buildExportNode uses current state)

      // Use an iframe instead of window.open to avoid popup blockers.
      const frame = document.createElement("iframe");
      frame.style.position = "fixed";
      frame.style.right = "0";
      frame.style.bottom = "0";
      frame.style.width = "0";
      frame.style.height = "0";
      frame.style.border = "0";
      frame.setAttribute("aria-hidden", "true");
      document.body.appendChild(frame);

      const doc = frame.contentDocument;
      doc.open();
      doc.write(`<!doctype html>
<html lang="zh-CN">
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <base href="${location.href}">
    <title>${exportName}</title>
    <link rel="stylesheet" href="./styles.css?v=highlight2" />
    <style>
      @page { size: A4; margin: 0; }
      html, body { margin: 0; padding: 0; background: #fff !important; }
      .page { width: ${A4_W}px; height: ${A4_H}px; background: #fff; overflow: hidden; box-sizing: border-box; page-break-after: always; position: relative; }
      .page:last-child { page-break-after: auto; }
      .stage { width: ${A4_W - MARGIN_X * 2}px; height: ${A4_H - MARGIN_TOP - MARGIN_BOTTOM}px; margin: ${MARGIN_TOP}px ${MARGIN_X}px ${MARGIN_BOTTOM}px ${MARGIN_X}px; overflow: hidden; box-sizing: border-box; position: relative; z-index: 0; background: transparent; }
      .scale {
        transform-origin: top left;
        position: relative;
        z-index: 1;
      }
      /* Watermark: CSS background so it appears only when enabling browser print option “背景图形”. */
      .stage::after{
        content:"";
        position:absolute;
        /* cover stage area */
        top:-45%;
        left:-45%;
        width:190%;
        height:190%;
        background-image: url("${watermarkUrl}");
        background-repeat: repeat;
        background-size: 220px 220px;
        background-position: 0 0;
        transform: rotate(-30deg);
        opacity: 0.1;
        pointer-events:none;
        z-index: 0;
      }
      /* remove shadows/rounding for print */
      .paper { box-shadow: none !important; border-radius: 0 !important; min-height: auto !important; background: transparent !important; }
      .topbar, .panel--left, .resizer, .modal, .no-print { display: none !important; }
    </style>
  </head>
  <body>
    <div id="pages"></div>
  </body>
</html>`);
      doc.close();

      const pagesEl = doc.getElementById("pages");
      doc.title = exportName;

      // Mount once for measurement
      const measurePage = doc.createElement("div");
      measurePage.className = "page";
      const measureStage = doc.createElement("div");
      measureStage.className = "stage";
      const measureScale = doc.createElement("div");
      measureScale.className = "scale";
      measureStage.appendChild(measureScale);
      measurePage.appendChild(measureStage);
      pagesEl.appendChild(measurePage);

      const imported = doc.importNode(node, true);
      measureScale.appendChild(imported);

      const waitImages = () => {
        const imgs = Array.from(doc.images || []);
        if (imgs.length === 0) return Promise.resolve();
        return Promise.all(
          imgs.map((img) => {
            if (img.complete && img.naturalWidth > 0) return Promise.resolve();
            return new Promise((resolve) => {
              const done = () => resolve();
              img.addEventListener("load", done, { once: true });
              img.addEventListener("error", done, { once: true });
            });
          })
        );
      };

      const waitWatermark = async () => {
        // Background images aren't part of doc.images; preload so print isn't blank.
        if (!watermarkUrl) return;
        try {
          const img = doc.createElement("img");
          img.style.position = "fixed";
          img.style.width = "1px";
          img.style.height = "1px";
          img.style.opacity = "0";
          img.style.pointerEvents = "none";
          img.setAttribute("aria-hidden", "true");
          img.src = watermarkUrl;
          doc.body.appendChild(img);
          await new Promise((resolve) => {
            const done = () => resolve();
            img.addEventListener("load", done, { once: true });
            img.addEventListener("error", done, { once: true });
          });
          try {
            img.remove();
          } catch {}
        } catch {}
      };

      await waitWatermark();
      await waitImages();
      // Fit width, then paginate height
      const stage = measureStage;
      const rectStage = stage.getBoundingClientRect();
      const rectContent = imported.getBoundingClientRect();
      const contentTop = rectContent.top;
      const sx = rectStage.width / rectContent.width;
      const s = Math.min(1, sx);

      const contentHeight = Math.ceil(rectContent.height);
      const stageHeight = rectStage.height;
      // 留一点安全空间，避免因为四舍五入/字体渲染导致底部切到一行
      const SAFE_PAD = 16;
      const pageHeightUnscaled = Math.max(1, (stageHeight - SAFE_PAD) / s);

      // 构建“不可切割块”列表：每一行、合计块、订购说明块与行、摆台风格块
      const blocks = [];
      const pushBlock = (el, kind = "generic") => {
        if (!el || !el.getBoundingClientRect) return;
        const rEl = el.getBoundingClientRect();
        const top = Math.max(0, rEl.top - contentTop);
        const bottom = Math.max(0, rEl.bottom - contentTop);
        if (Number.isFinite(top) && Number.isFinite(bottom) && bottom > top + 1) blocks.push({ top, bottom, kind });
      };

      for (const tr of Array.from(imported.querySelectorAll("tbody tr"))) pushBlock(tr, "row");
      pushBlock(imported.querySelector(".totals"), "totals");
      pushBlock(imported.querySelector(".paper__foot"), "foot");
      for (const div of Array.from(imported.querySelectorAll(".orderNotes > div"))) pushBlock(div, "noteLine");
      // style blocks should not be cut; additionally support "shrink a bit" if only small split would happen
      for (const el of Array.from(imported.querySelectorAll(".quoteStyles__item"))) pushBlock(el, "style");

      blocks.sort((a, b) => a.top - b.top);

      const effectiveContentHeight = (() => {
        const maxBottom = blocks.reduce((m, b) => Math.max(m, b.bottom), 0);
        // Prefer last “real content” bottom to avoid extra blank pages caused by
        // min-height/padding or off-by-one rounding in rectContent.height.
        const h = Math.max(1, Math.ceil(maxBottom || contentHeight));
        return h;
      })();

      const segments = [];
      let start = 0;
      const EPS = 6;
      while (start < effectiveContentHeight - EPS) {
        const max = start + pageHeightUnscaled;
        // 找到第一个会在本页被切割的块（top < max 且 bottom > max）
        const overflow = blocks.find((b) => b.top > start + EPS && b.top < max && b.bottom > max);

        if (overflow) {
          if (overflow.kind === "style") {
            const blockH = Math.max(1, overflow.bottom - overflow.top);
            const cut = overflow.bottom - max; // how much would be cut off
            const cutRatio = cut / blockH;
            // If cut is small (<=20%), shrink ONLY THIS PAGE slightly so the whole style block fits.
            if (cutRatio <= 0.2) {
              const need = overflow.bottom - start + 2; // include tiny safety
              const sNeeded = (stageHeight - SAFE_PAD) / Math.max(1, need);
              const segScale = clamp(sNeeded, 0.7, s);
              if (segScale < s - 0.001) {
                segments.push({ start, heightUnscaled: Math.max(1, need), scale: segScale });
                start = Math.max(start + 1, overflow.bottom);
                continue;
              }
            }
            // cut would be large: move the whole block to next page
          }
          // 本页只显示到 overflow.top（不显示半行），下一页从 overflow.top 开始
          const nextStart = Math.max(overflow.top, start + EPS);
          // 再往上收一点点，避免因像素取整导致“下一行露出一条边/一点图片”
          const end = Math.max(start + 1, nextStart - 2);
          const heightUnscaled = Math.max(1, end - start);
          segments.push({ start, heightUnscaled, scale: s });
          start = nextStart;
          continue;
        }

        // 没有切割：正常推进一页
        const heightUnscaled = Math.min(pageHeightUnscaled, effectiveContentHeight - start);
        segments.push({ start, heightUnscaled, scale: s });
        start = start + pageHeightUnscaled;
      }
      // Remove trailing zero-content segments (can happen due to rounding)
      while (segments.length > 1) {
        const last = segments[segments.length - 1];
        if (!last) break;
        if ((last.start || 0) >= effectiveContentHeight - EPS) segments.pop();
        else break;
      }
      const totalPages = Math.max(1, segments.length);

      // Clear and rebuild pages with clipped offsets
      pagesEl.innerHTML = "";
      for (let i = 0; i < totalPages; i += 1) {
        const seg = segments[i] || { start: 0, heightUnscaled: pageHeightUnscaled, scale: s };
        const page = doc.createElement("div");
        page.className = "page";
        page.style.position = "relative";
        const st = doc.createElement("div");
        st.className = "stage";
        // 本页裁剪高度（避免显示半行），其余留白
        st.style.height = `${Math.max(1, Math.ceil(seg.heightUnscaled * (seg.scale || s)))}px`;
        const sc = doc.createElement("div");
        sc.className = "scale";
        sc.style.transform = `scale(${seg.scale || s})`;
        const clone = doc.importNode(node, true);
        clone.style.position = "relative";
        clone.style.top = `-${Math.floor(seg.start || 0)}px`;
        sc.appendChild(clone);
        st.appendChild(sc);
        page.appendChild(st);
        pagesEl.appendChild(page);
      }

      // Print (user chooses Save as PDF)
      setTimeout(() => {
        try {
          frame.contentWindow.focus();
        } catch {}
        try {
          frame.contentWindow.print();
        } catch {}
      }, 250);

      // Cleanup later (after print dialog opens)
      setTimeout(() => {
        try {
          frame.remove();
        } catch {}
      }, 30_000);
    } catch (e) {
      // eslint-disable-next-line no-console
      console.error(e);
      alert(`导出 PDF 失败：${e?.message || e}`);
    } finally {
      btnExportPdf.disabled = false;
      btnExportPdf.textContent = oldText;
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

  // PDF export options (logo on/off)
  const pdfOptionsModal = el("#pdfOptionsModal");
  const pdfOptShowLogo = el("#pdfOptShowLogo");
  const btnPdfOptCancel = el("#btnPdfOptCancel");
  const btnPdfOptGo = el("#btnPdfOptGo");

  const openPdfOptions = () => {
    if (!pdfOptionsModal) return exportPdfA4OnePage({ showLogo: true });
    // default: show logo (matches current behavior)
    if (pdfOptShowLogo) pdfOptShowLogo.checked = true;
    pdfOptionsModal.setAttribute("aria-hidden", "false");
  };
  const closePdfOptions = () => {
    pdfOptionsModal?.setAttribute("aria-hidden", "true");
  };

  function menuImageForLine(line) {
    if (!line?.source?.startsWith("menu:")) return null;
    const name = line.source.slice("menu:".length);
    const item = findMenuItemByName(name);
    return item?.image || null;
  }

  function menuThumbForName(name) {
    const item = findMenuItemByName(name);
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

    const allMenu = [...menu, ...customMenu];
    const list = allMenu
      .filter((it) => (cat === "全部" ? true : String(it.category || "") === cat))
      .filter((it) => (q ? String(it.name).toLowerCase().includes(q) : true))
      .filter((it) => (selectedOnly ? Boolean(state.selected[it.name]) : true));

    menuCount.textContent = `共 ${list.length} 项（已选 ${Object.keys(state.selected).length}，自定义 ${customMenu.length}）`;

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

      const btnRemoveCustom = document.createElement("button");
      btnRemoveCustom.className = "btn btn--ghost";
      btnRemoveCustom.type = "button";
      btnRemoveCustom.textContent = "删除";
      btnRemoveCustom.style.display = item.custom ? "" : "none";

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

      btnRemoveCustom.addEventListener("click", () => {
        if (!item.custom) return;
        if (!confirm(`删除自定义产品：${item.name}？`)) return;
        // remove selection and quote lines if any
        removeSelected(item.name);
        state.quoteLines = state.quoteLines.filter((l) => l.source !== `menu:${item.name}`);
        const idx = customMenu.findIndex((m) => m.name === item.name);
        if (idx >= 0) customMenu.splice(idx, 1);
        saveCustomMenu();
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
      actions.appendChild(btnRemoveCustom);

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
      new Set([...menu, ...customMenu].map((m) => String(m.category || "").trim()).filter(Boolean))
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
    renderStyles();
    renderQuote();
    renderStylesInQuote();
    renderWatermarkPreview();
  }

  function initCustomUpload() {
    if (!btnAddCustom || !customModal) return;

    const clearForm = () => {
      if (customName) customName.value = "";
      if (customCategory) customCategory.value = "";
      if (customPrice) customPrice.value = "";
      if (customMinOrder) customMinOrder.value = "";
      if (customFile) customFile.value = "";
      if (cropWrap) cropWrap.hidden = true;
      cropState.ready = false;
      cropState.tx = 0;
      cropState.ty = 0;
      cropState.scale = 1;
      cropState.boxSize = 0;
      cropState.boxCx = 0;
      cropState.boxCy = 0;
    };

    btnAddCustom.addEventListener("click", () => {
      clearForm();
      // populate category dropdown from existing categories
      if (customCategoryList) {
        const cats = Array.from(
          new Set([...menu, ...customMenu].map((m) => String(m.category || "").trim()).filter(Boolean))
        ).sort((a, b) => a.localeCompare(b, "zh-Hans-CN"));
        customCategoryList.innerHTML = "";
        for (const c of cats) {
          const opt = document.createElement("option");
          opt.value = c;
          customCategoryList.appendChild(opt);
        }
      }
      openCustomModal();
      try {
        customName?.focus();
      } catch {}
    });

    for (const elClose of Array.from(document.querySelectorAll("[data-close-custom=\"1\"]"))) {
      elClose.addEventListener("click", closeCustomModal);
    }
    btnCustomCancel?.addEventListener("click", closeCustomModal);

    customFile?.addEventListener("change", async () => {
      const f = customFile.files?.[0];
      if (!f) return;
      if (!String(f.type || "").startsWith("image/")) {
        alert("请选择图片文件。");
        return;
      }
      const url = URL.createObjectURL(f);
      cropImg.src = url;
      if (cropWrap) cropWrap.hidden = false;

      await new Promise((resolve) => {
        const done = () => resolve();
        cropImg.addEventListener("load", done, { once: true });
        cropImg.addEventListener("error", done, { once: true });
      });

      cropState.imgNaturalW = cropImg.naturalWidth || 0;
      cropState.imgNaturalH = cropImg.naturalHeight || 0;
      cropState.ready = cropState.imgNaturalW > 0 && cropState.imgNaturalH > 0;
      applyCropBoxLayout();
      resetCropToCover();
      clampCrop();
      setCropTransform();

      try {
        URL.revokeObjectURL(url);
      } catch {}
    });

    const onDown = (ev) => {
      if (!cropState.ready) return;
      cropState.dragging = true;
      const p = ev.touches?.[0] || ev;
      cropState.lastX = p.clientX;
      cropState.lastY = p.clientY;
    };
    const onMove = (ev) => {
      if (!cropState.dragging) return;
      const p = ev.touches?.[0] || ev;
      const dx = p.clientX - cropState.lastX;
      const dy = p.clientY - cropState.lastY;
      cropState.lastX = p.clientX;
      cropState.lastY = p.clientY;
      cropState.tx += dx;
      cropState.ty += dy;
      clampCrop();
      setCropTransform();
      ev.preventDefault?.();
    };
    const onUp = () => {
      cropState.dragging = false;
    };
    cropStage?.addEventListener("mousedown", onDown);
    cropStage?.addEventListener("touchstart", onDown, { passive: true });
    window.addEventListener("mousemove", onMove, { passive: false });
    window.addEventListener("touchmove", onMove, { passive: false });
    window.addEventListener("mouseup", onUp);
    window.addEventListener("touchend", onUp);

    cropStage?.addEventListener(
      "wheel",
      (ev) => {
        if (!cropState.ready) return;
        const delta = ev.deltaY || 0;
        const factor = delta > 0 ? 0.92 : 1.08;
        cropState.scale = clamp(cropState.scale * factor, 0.2, 10);
        clampCrop();
        setCropTransform();
        ev.preventDefault();
      },
      { passive: false }
    );

    // Manual crop box: drag to move, handle to resize
    const boxDown = (ev) => {
      if (!cropState.ready) return;
      cropState.boxDragging = true;
      const p = ev.touches?.[0] || ev;
      cropState.lastX = p.clientX;
      cropState.lastY = p.clientY;
      ev.preventDefault?.();
      ev.stopPropagation?.();
    };
    cropBox?.addEventListener("mousedown", boxDown);
    cropBox?.addEventListener("touchstart", boxDown, { passive: false });

    const handleDown = (ev) => {
      if (!cropState.ready) return;
      cropState.boxResizing = true;
      const p = ev.touches?.[0] || ev;
      cropState.lastX = p.clientX;
      cropState.lastY = p.clientY;
      ev.preventDefault?.();
      ev.stopPropagation?.();
    };
    cropHandle?.addEventListener("mousedown", handleDown);
    cropHandle?.addEventListener("touchstart", handleDown, { passive: false });

    const onBoxMove = (ev) => {
      if (!cropState.boxDragging && !cropState.boxResizing) return;
      const p = ev.touches?.[0] || ev;
      const dx = p.clientX - cropState.lastX;
      const dy = p.clientY - cropState.lastY;
      cropState.lastX = p.clientX;
      cropState.lastY = p.clientY;

      if (cropState.boxDragging) {
        cropState.boxCx += dx;
        cropState.boxCy += dy;
        applyCropBoxLayout();
        clampCrop();
        setCropTransform();
      } else if (cropState.boxResizing) {
        const d = (dx + dy) / 2;
        cropState.boxSize = (cropState.boxSize || 0) + d * 2;
        applyCropBoxLayout();
        const iw = cropState.imgNaturalW || 1;
        const ih = cropState.imgNaturalH || 1;
        const minScale = Math.max(cropState.boxSize / iw, cropState.boxSize / ih);
        if (cropState.scale < minScale) cropState.scale = minScale;
        clampCrop();
        setCropTransform();
      }
      ev.preventDefault?.();
    };
    const onBoxUp = () => {
      cropState.boxDragging = false;
      cropState.boxResizing = false;
    };
    window.addEventListener("mousemove", onBoxMove, { passive: false });
    window.addEventListener("touchmove", onBoxMove, { passive: false });
    window.addEventListener("mouseup", onBoxUp);
    window.addEventListener("touchend", onBoxUp);

    btnCustomSave?.addEventListener("click", () => {
      const name = String(customName?.value || "").trim();
      const category = String(customCategory?.value || "").trim() || "自定义";
      const unitPrice = toNumber(customPrice?.value || 0);
      const minOrder = String(customMinOrder?.value || "").trim();
      if (!name) return alert("请填写产品名称。");
      if (!Number.isFinite(unitPrice) || unitPrice < 0) return alert("请填写正确的单价。");
      if (findMenuItemByName(name)) return alert("该产品名称已存在，请换一个名称（或在菜单里直接编辑数量/价格）。");

      const thumb = cropToThumbDataUrl();

      customMenu.unshift({
        name,
        category,
        unitPrice,
        minOrder,
        ...(thumb ? { image: thumb, imageThumb: thumb } : {}),
        custom: true,
      });
      saveCustomMenu();
      closeCustomModal();
      renderAll();
    });
  }

  function renderStyles() {
    if (!styleList || !styleCount) return;
    styleCount.textContent = `共 ${styleLibrary.length} 项（已加入 ${state.selectedStyles.length}）`;
    styleList.innerHTML = "";

    for (const s of styleLibrary) {
      const card = document.createElement("div");
      card.className = "card";

      const grid = document.createElement("div");
      grid.className = "card__grid";

      const thumb = document.createElement("div");
      thumb.className = "thumb";
      if (s.image) {
        const img = document.createElement("img");
        img.alt = s.category || "摆台风格";
        img.loading = "lazy";
        img.decoding = "async";
        img.fetchPriority = "low";
        img.src = s.imageThumb || s.image;
        img.addEventListener("click", () => openImageModal(s.image, s.category || "摆台风格"));
        thumb.appendChild(img);
      } else {
        thumb.textContent = "无图";
      }

      const content = document.createElement("div");
      const top = document.createElement("div");
      top.className = "card__top";
      const left = document.createElement("div");
      const nameInput = document.createElement("input");
      nameInput.className = "input";
      nameInput.style.padding = "8px 10px";
      nameInput.style.borderRadius = "10px";
      nameInput.value = s.category || "";
      nameInput.placeholder = "分类（可编辑）";
      nameInput.addEventListener("input", debounce(() => {
        s.category = nameInput.value.trim();
        saveStyleLibrary();
        renderStylesInQuote();
      }, 120));
      left.appendChild(nameInput);
      top.appendChild(left);

      const meta = document.createElement("div");
      meta.className = "card__meta";
      meta.innerHTML = `<span>比例：<b>${s.aspect || "-"}</b></span>`;

      const actions = document.createElement("div");
      actions.className = "card__actions";

      const joined = state.selectedStyles.some((x) => x.id === s.id);
      const btnJoin = document.createElement("button");
      btnJoin.className = "btn btn--ghost";
      btnJoin.type = "button";
      btnJoin.textContent = joined ? "移除" : "加入";

      const btnDel = document.createElement("button");
      btnDel.className = "btn btn--ghost";
      btnDel.type = "button";
      btnDel.textContent = "删除";

      btnJoin.addEventListener("click", () => {
        if (joined) state.selectedStyles = state.selectedStyles.filter((x) => x.id !== s.id);
        else state.selectedStyles = [{ id: s.id }, ...state.selectedStyles];
        saveState(state);
        renderAll();
      });

      btnDel.addEventListener("click", () => {
        if (!confirm(`删除摆台风格：${s.category || ""}？`)) return;
        const idx = styleLibrary.findIndex((x) => x.id === s.id);
        if (idx >= 0) styleLibrary.splice(idx, 1);
        state.selectedStyles = state.selectedStyles.filter((x) => x.id !== s.id);
        saveStyleLibrary();
        saveState(state);
        renderAll();
      });

      actions.appendChild(btnJoin);
      actions.appendChild(btnDel);

      content.appendChild(top);
      content.appendChild(meta);
      content.appendChild(actions);

      grid.appendChild(thumb);
      grid.appendChild(content);
      card.appendChild(grid);
      styleList.appendChild(card);
    }
  }

  function renderStylesInQuote() {
    if (!qStyles) return;
    qStyles.innerHTML = "";
    const ids = state.selectedStyles.map((x) => x.id);
    const selected = ids.map((id) => styleLibrary.find((s) => s.id === id)).filter(Boolean);
    if (selected.length === 0) return;

    for (const s of selected) {
      const wrap = document.createElement("div");
      wrap.className = "quoteStyles__item";
      const fullSrc = s.imageFull || s.image;
      if (fullSrc) {
        const img = document.createElement("img");
        img.className = "quoteStyles__img";
        img.src = fullSrc;
        img.alt = s.category || "摆台风格";
        img.loading = "lazy";
        img.decoding = "async";
        img.fetchPriority = "low";
        img.addEventListener("click", () => openImageModal(fullSrc, s.category || "摆台风格"));
        wrap.appendChild(img);
      }
      const cap = document.createElement("div");
      cap.className = "quoteStyles__cap";
      cap.innerHTML = `<span class="quoteStyles__tag">摆台风格</span><span>${escapeHtml(s.category || "")}</span>`;
      wrap.appendChild(cap);
      qStyles.appendChild(wrap);
    }
  }

  function initStyleUpload() {
    if (!btnAddStyle || !styleModal) return;

    let pending = [];
    let activeId = null;

    const aspectMap = { "16:9": 16 / 9, "9:16": 9 / 16, "4:3": 4 / 3, "3:4": 3 / 4 };
    const aspectToDims = (key) => {
      switch (key) {
        case "9:16":
          return { thumbW: 270, thumbH: 480, fullW: 900, fullH: 1600 };
        case "4:3":
          return { thumbW: 480, thumbH: 360, fullW: 1600, fullH: 1200 };
        case "3:4":
          return { thumbW: 360, thumbH: 480, fullW: 1200, fullH: 1600 };
        case "16:9":
        default:
          return { thumbW: 480, thumbH: 270, fullW: 1600, fullH: 900 };
      }
    };

    const loadImage = (src) =>
      new Promise((resolve) => {
        const img = new Image();
        img.onload = () => resolve(img);
        img.onerror = () => resolve(img);
        img.src = src;
      });

    const autoCropToDataUrl = async ({ objectUrl, aspectKey, thumbW, thumbH, fullW, fullH }) => {
      const img = await loadImage(objectUrl);
      const iw = img.naturalWidth || img.width || 0;
      const ih = img.naturalHeight || img.height || 0;
      if (!iw || !ih) return { thumb: null, full: null };
      const aspect = aspectMap[aspectKey] || 16 / 9;
      let sw;
      let sh;
      if (iw / ih > aspect) {
        sh = ih;
        sw = ih * aspect;
      } else {
        sw = iw;
        sh = iw / aspect;
      }
      const sx = (iw - sw) / 2;
      const sy = (ih - sh) / 2;
      const draw = (ow, oh) => {
        const c = document.createElement("canvas");
        c.width = ow;
        c.height = oh;
        const ctx = c.getContext("2d");
        ctx.fillStyle = "#fff";
        ctx.fillRect(0, 0, ow, oh);
        ctx.imageSmoothingEnabled = true;
        ctx.imageSmoothingQuality = "high";
        ctx.drawImage(img, sx, sy, sw, sh, 0, 0, ow, oh);
        return c.toDataURL("image/jpeg", 0.9);
      };
      return { thumb: draw(thumbW, thumbH), full: draw(fullW, fullH) };
    };

    const clearForm = () => {
      if (styleFile) styleFile.value = "";
      pending = [];
      activeId = null;
      if (styleBatchList) styleBatchList.innerHTML = "";
    };

    const renderPending = () => {
      if (!styleBatchList) return;
      styleBatchList.innerHTML = "";
      for (const item of pending) {
        const wrap = document.createElement("div");
        wrap.className = "batchItem";

        const thumb = document.createElement("div");
        thumb.className = "batchItem__thumb";
        const a = aspectMap[item.aspect || "16:9"] || 16 / 9;
        thumb.style.aspectRatio = String(a);
        const img = document.createElement("img");
        img.alt = item.fileName || "图片";
        img.src = item.thumbDataUrl || item.objectUrl;
        thumb.appendChild(img);

        const meta = document.createElement("div");
        meta.className = "batchItem__meta";

        const row1 = document.createElement("div");
        row1.className = "batchItem__row";
        const cat = document.createElement("input");
        cat.className = "input";
        cat.type = "text";
        cat.placeholder = "分类（必填）";
        cat.value = item.category || "";
        cat.addEventListener("input", () => {
          item.category = cat.value;
        });

        const sel = document.createElement("select");
        sel.className = "input";
        sel.innerHTML = `
          <option value="4:3">4:3</option>
          <option value="3:4">3:4</option>
          <option value="16:9">16:9</option>
          <option value="9:16">9:16</option>
        `;
        sel.value = item.aspect || "16:9";
        sel.addEventListener("change", () => {
          item.aspect = sel.value;
          renderPending();
        });

        row1.appendChild(cat);
        row1.appendChild(sel);

        const actions = document.createElement("div");
        actions.className = "batchItem__actions";
        const btnEdit = document.createElement("button");
        btnEdit.className = "btn btn--ghost";
        btnEdit.type = "button";
        btnEdit.textContent = activeId === item.id ? "收起" : "编辑";
        btnEdit.addEventListener("click", () => {
          activeId = activeId === item.id ? null : item.id;
          renderPending();
        });

        const btnRemove = document.createElement("button");
        btnRemove.className = "btn btn--ghost";
        btnRemove.type = "button";
        btnRemove.textContent = "移除";
        btnRemove.addEventListener("click", () => {
          if (!confirm("移除该图片？")) return;
          const idx = pending.findIndex((x) => x.id === item.id);
          if (idx >= 0) {
            try {
              if (pending[idx].objectUrl) URL.revokeObjectURL(pending[idx].objectUrl);
            } catch {}
            pending.splice(idx, 1);
          }
          if (activeId === item.id) activeId = null;
          renderPending();
        });

        const status = document.createElement("div");
        status.className = "batchItem__status";
        status.textContent = item.fullDataUrl ? "已裁剪" : "未裁剪";

        actions.appendChild(btnEdit);
        actions.appendChild(btnRemove);
        actions.appendChild(status);

        meta.appendChild(row1);
        meta.appendChild(actions);

        wrap.appendChild(thumb);
        wrap.appendChild(meta);

        if (activeId === item.id) {
          const editor = document.createElement("div");
          editor.className = "batchItem__editor";
          editor.innerHTML = `
            <div class="cropStage">
              <img alt="裁剪预览" />
              <div class="cropMask" aria-hidden="true"></div>
              <div class="cropBox" aria-hidden="true"><div class="cropHandle" aria-hidden="true"></div></div>
            </div>
            <div class="controls__row" style="justify-content:flex-end">
              <button class="btn btn--ghost btn--sm" type="button" data-save-one="1">保存本张</button>
            </div>
          `;
          wrap.appendChild(editor);

          queueMicrotask(async () => {
            if (activeId !== item.id) return;
            const stageEl = editor.querySelector(".cropStage");
            const imgEl = editor.querySelector("img");
            const boxEl = editor.querySelector(".cropBox");
            const handleEl = editor.querySelector(".cropHandle");
            const btnSaveOne = editor.querySelector("button[data-save-one=\"1\"]");
            if (!stageEl || !imgEl || !boxEl) return;
            const aspect = aspectMap[item.aspect || "16:9"] || 16 / 9;
            stageEl.style.aspectRatio = String(aspect);
            const cropper = createCropper({ stageEl, imgEl, boxEl, handleEl, aspect });
            imgEl.src = item.objectUrl;
            await new Promise((resolve) => {
              const done = () => resolve();
              imgEl.addEventListener("load", done, { once: true });
              imgEl.addEventListener("error", done, { once: true });
            });
            cropper.state.imgNaturalW = imgEl.naturalWidth || 0;
            cropper.state.imgNaturalH = imgEl.naturalHeight || 0;
            cropper.state.ready = cropper.state.imgNaturalW > 0 && cropper.state.imgNaturalH > 0;
            cropper.setAspect(aspect);
            btnSaveOne?.addEventListener("click", () => {
              const category = String(item.category || "").trim();
              if (!category) return alert("请填写分类。");
              const dims = aspectToDims(item.aspect || "16:9");
              const thumb2 = cropper.cropToDataUrl(dims.thumbW, dims.thumbH);
              const full2 = cropper.cropToDataUrl(dims.fullW, dims.fullH);
              if (!thumb2 || !full2) return alert("请先完成裁剪。");
              item.thumbDataUrl = thumb2;
              item.fullDataUrl = full2;
              item.fullW = dims.fullW;
              item.fullH = dims.fullH;
              renderPending();
            });
          });
        }

        styleBatchList.appendChild(wrap);
      }
    };

    btnAddStyle.addEventListener("click", () => {
      clearForm();
      openStyleModal();
    });

    for (const elClose of Array.from(document.querySelectorAll("[data-close-style=\"1\"]"))) {
      elClose.addEventListener("click", closeStyleModal);
    }
    btnStyleCancel?.addEventListener("click", closeStyleModal);

    styleFile?.addEventListener("change", async () => {
      const files = Array.from(styleFile.files || []);
      if (files.length === 0) return;
      pending = [];
      activeId = null;
      for (const f of files) {
        if (!String(f.type || "").startsWith("image/")) continue;
        const objectUrl = URL.createObjectURL(f);
        let guessedCategory = "";
        try {
          const rel = String(f.webkitRelativePath || "");
          if (rel.includes("/")) guessedCategory = rel.split("/")[0] || "";
        } catch {}
        pending.push({
          id: `tmp_${Date.now()}_${Math.random().toString(16).slice(2)}`,
          fileName: f.name,
          objectUrl,
          category: guessedCategory,
          aspect: "16:9",
          thumbDataUrl: null,
          fullDataUrl: null,
          fullW: 0,
          fullH: 0,
        });
      }
      renderPending();
    });

    btnStyleAutoCropAll?.addEventListener("click", async () => {
      if (!pending.length) return alert("请先选择图片。");
      const missing = pending.find((x) => !String(x.category || "").trim());
      if (missing) return alert("请先为每张图片填写分类（可在列表中直接输入）。");

      btnStyleAutoCropAll.disabled = true;
      const old = btnStyleAutoCropAll.textContent;
      btnStyleAutoCropAll.textContent = "裁剪中...";
      try {
        for (const item of pending) {
          if (item.fullDataUrl) continue;
          const dims = aspectToDims(item.aspect || "16:9");
          const { thumb, full } = await autoCropToDataUrl({
            objectUrl: item.objectUrl,
            aspectKey: item.aspect || "16:9",
            ...dims,
          });
          item.thumbDataUrl = thumb;
          item.fullDataUrl = full;
          item.fullW = dims.fullW;
          item.fullH = dims.fullH;
          renderPending();
          await new Promise((r) => setTimeout(r, 10));
        }
      } finally {
        btnStyleAutoCropAll.disabled = false;
        btnStyleAutoCropAll.textContent = old;
      }
    });

    btnStyleSave?.addEventListener("click", () => {
      if (!pending.length) return alert("请先选择图片。");
      const ready = pending.filter((x) => String(x.category || "").trim() && x.thumbDataUrl && x.fullDataUrl);
      if (ready.length !== pending.length) return alert("请先完成裁剪（可逐张编辑保存，或点“一键裁剪全部”）。");

      for (const x of ready.reverse()) {
        const id = `st_${Date.now()}_${Math.random().toString(16).slice(2)}`;
        styleLibrary.unshift({
          id,
          category: String(x.category || "").trim(),
          aspect: x.aspect || "16:9",
          image: x.thumbDataUrl,
          imageThumb: x.thumbDataUrl,
          imageFull: x.fullDataUrl,
          fullW: x.fullW,
          fullH: x.fullH,
        });
      }
      saveStyleLibrary();
      closeStyleModal();
      renderAll();
    });
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
  btnExportPdf.addEventListener("click", openPdfOptions);
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

  // PDF options modal events
  for (const elClose of Array.from(document.querySelectorAll("[data-close-pdfopt=\"1\"]"))) {
    elClose.addEventListener("click", closePdfOptions);
  }
  btnPdfOptCancel?.addEventListener("click", closePdfOptions);
  btnPdfOptGo?.addEventListener("click", async () => {
    const showLogo = Boolean(pdfOptShowLogo?.checked);
    closePdfOptions();
    await exportPdfA4OnePage({ showLogo });
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

  initCustomUpload();
  initStyleUpload();
  renderAll();
}

document.addEventListener("DOMContentLoaded", init);
