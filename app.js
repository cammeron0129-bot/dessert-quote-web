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
    finalPrice: "",
    note: "不含税",
    orderNotes:
      "1. 甜品台为预约项目，建议提前预定档期。\n2. 动物奶油易融，空调环境建议摆放 2–3 小时，请合理预约时间。\n3. 交通/摆台/撤场等费用请按实际情况另行确认。",
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
  const metaFinalPrice = el("#metaFinalPrice");
  const metaNote = el("#metaNote");
  const metaOrderNotes = el("#metaOrderNotes");

  const qDate = el("#qDate");
  const qLocation = el("#qLocation");
  const qCustomer = el("#qCustomer");
  const qContact = el("#qContact");
  const qNote = el("#qNote");
  const qOrderNotes = el("#qOrderNotes");

  const quoteTbody = el("#quoteTbody");
  const subtotalEl = el("#subtotal");
  const totalAfterDiscountEl = el("#totalAfterDiscount");

  const sidebarResizer = document.querySelector("#sidebarResizer");

  const imageModal = el("#imageModal");
  const modalClose = el("#modalClose");
  const modalImg = el("#modalImg");
  const modalCaption = el("#modalCaption");

  const btnPrint = el("#btnPrint");
  const btnExportCsv = el("#btnExportCsv");
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

  function exportCsv() {
    const header = ["序号", "内容", "数量", "单价", "总价"];
    const rows = state.quoteLines.map((l, idx) => {
      const qty = toNumber(l.qty);
      const unit = toNumber(l.unitPrice);
      const total = qty * unit;
      return [String(idx + 1), String(l.name || ""), String(qty), String(unit), String(total)];
    });
    const csv = [header, ...rows]
      .map((r) => r.map((cell) => `"${String(cell).replace(/\"/g, '""')}"`).join(","))
      .join("\n");
    downloadText(`${buildExportName()}.csv`, "text/csv;charset=utf-8", csv);
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
    metaFinalPrice.value = state.meta.finalPrice || "";
    metaNote.value = state.meta.note || "";
    metaOrderNotes.value = state.meta.orderNotes || "";

    qDate.textContent = state.meta.date || "";
    qLocation.textContent = state.meta.location || "";
    qCustomer.textContent = state.meta.customer || "";
    qContact.textContent = state.meta.contact || "";
    qNote.textContent = state.meta.note || "";
    qOrderNotes.textContent = state.meta.orderNotes || "";
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
        nameInput.addEventListener("change", () => {
          const target = state.quoteLines.find((x) => x.id === line.id);
          if (!target) return;
          target.name = nameInput.value.trim();
          saveState(state);
          renderQuote();
        });

        const grid = document.createElement("div");
        grid.className = "quoteCard__grid";

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
    const discountPercent = clamp(toNumber(state.meta.discountPercent || 100), 0, 100);
    const computedAfterDiscount = subtotal * (discountPercent / 100);
    const manualFinal = toNumber(state.meta.finalPrice);
    const totalAfterDiscount = manualFinal > 0 ? manualFinal : computedAfterDiscount;

    subtotalEl.textContent = formatMoney(subtotal);
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
    state.meta.finalPrice = metaFinalPrice.value.trim();
    state.meta.note = metaNote.value.trim();
    state.meta.orderNotes = metaOrderNotes.value || "";
    saveState(state);
    renderMeta();
    renderQuote();
  }

  for (const t of tabs) {
    t.addEventListener("click", () => setActiveTab(t.dataset.tab));
  }

  menuSearch.addEventListener("input", renderMenu);
  onlySelected.addEventListener("change", renderMenu);

  metaDate.addEventListener("change", onMetaChange);
  metaLocation.addEventListener("change", onMetaChange);
  metaCustomer.addEventListener("change", onMetaChange);
  metaContact.addEventListener("change", onMetaChange);
  metaDiscount.addEventListener("change", onMetaChange);
  metaFinalPrice.addEventListener("change", onMetaChange);
  metaNote.addEventListener("change", onMetaChange);
  metaOrderNotes.addEventListener("change", onMetaChange);

  qOrderNotes.addEventListener("input", () => {
    state.meta.orderNotes = qOrderNotes.textContent || "";
    metaOrderNotes.value = state.meta.orderNotes;
    saveState(state);
  });

  btnExportCsv.addEventListener("click", exportCsv);

  btnPrint.addEventListener("click", () => {
    const old = document.title;
    document.title = buildExportName();
    setTimeout(() => {
      window.print();
      setTimeout(() => {
        document.title = old;
      }, 500);
    }, 50);
  });

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
