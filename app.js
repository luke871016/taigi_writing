/**
 * 台語字練習簿 Worksheet 產生器
 * 所見即所得編輯、可輸出 A4 PDF
 */

(function () {
  var nextItemId = 1;
  function ensureItemId(item) {
    if (!item.id) {
      item.id = "item-" + String(nextItemId++);
    }
    return item;
  }
  function ensureAllItemIds() {
    state.items.forEach(ensureItemId);
  }

  const state = {
    items: [
      { description: "台語", exampleText: "台語 Tâi-gí", lineCount: 1 },
      { description: "寫字", exampleText: "寫字 Siá-jī", lineCount: 1 },
      { description: "練習", exampleText: "練習 Liān-si̍p", lineCount: 1 },
      { description: "筆記", exampleText: "", lineCount: 15 },
    ],
    lineStyle: "single",
    lineSpacingDelta: 0,
    fontSize: 24,
    pageHeader: "",
    focusedItemIndex: null,
  };
  ensureAllItemIds();

  /** 依目前字級計算單行行高（與字級成比例，比例同預設 24px -> 40px） */
  function getLineHeightPx() {
    return Math.round((state.fontSize * 40) / 24);
  }

  /** 取得目前字型 1ex 的 px 值（x-height），用於行距預設 */
  function getExPx() {
    var span = document.createElement("span");
    span.style.cssText =
      "position:absolute;left:-9999px;font-family:'Iansui',sans-serif;font-size:" +
      state.fontSize +
      "px;line-height:1;visibility:hidden;";
    span.textContent = "x";
    document.body.appendChild(span);
    var h = span.offsetHeight;
    document.body.removeChild(span);
    return h;
  }

  /**
   * 期望的視覺間距（px）：「這組底線」到「下一組頂線」= 1ex + delta×lineHeight（delta 為 -0.2～+0.2，不小於 0）
   */
  function getDesiredGapPx() {
    var ex = getExPx();
    var lh = getLineHeightPx();
    return Math.max(0, ex + state.lineSpacingDelta * lh);
  }

  /**
   * 計算 --line-spacing（margin-top px）。
   * 期望視覺間距 desired = 1ex + delta×0.2×lineHeight；換算成 margin 時要加上比例修正，
   * 使 24px、行距 0 時約為 -15px，並依字級縮放。
   */
  function getLineSpacingPx() {
    var desired = getDesiredGapPx();
    var ex = getExPx();
    var correctionPx = 15 * (state.fontSize / 24);
    return Math.round(desired - ex - correctionPx);
  }

  const $itemList = document.getElementById("itemList");
  const $addItem = document.getElementById("addItem");
  const $insertMenuToggle = document.getElementById("insertMenuToggle");
  const $insertMenu = document.getElementById("insertMenu");
  const $insertPageBreak = document.getElementById("insertPageBreak");
  const $insertToneChart = document.getElementById("insertToneChart");
  const $insertSandhiChart = document.getElementById("insertSandhiChart");
  const $exportPdf = document.getElementById("exportPdf");
  const $loadTextbookBtn = document.getElementById("loadTextbookBtn");
  const $loadTextbookMenu = document.getElementById("loadTextbookMenu");
  const $importSettings = document.getElementById("importSettings");
  const $importSettingsFile = document.getElementById("importSettingsFile");
  const $exportSettings = document.getElementById("exportSettings");
  const $worksheetPreview = document.getElementById("worksheetPreview");
  const $main = document.querySelector(".main");
  const $fontSize = document.getElementById("fontSize");
  const $fontSizeValue = document.getElementById("fontSizeValue");
  const $lineSpacing = document.getElementById("lineSpacing");
  const $lineSpacingValue = document.getElementById("lineSpacingValue");
  const $settingsToggle = document.getElementById("settingsToggle");
  const $settingsPanel = document.getElementById("settingsPanel");
  const $pageHeader = document.getElementById("pageHeader");

  /** 將練習項目展開為一維「顯示條目」陣列：一般行、分頁標記、圖片。供分頁與預覽使用 */
  function getFlatEntries() {
    var flat = [];
    state.items.forEach(function (item, itemIndex) {
      if (item.type === "pageBreak") {
        flat.push({ pageBreak: true });
        return;
      }
      if (item.type === "image" && item.imagePath) {
        flat.push({ image: item.imagePath, itemIndex: itemIndex });
        return;
      }
      var n = Math.max(1, parseInt(item.lineCount, 10) || 1);
      for (var i = 0; i < n; i++) {
        flat.push({
          exampleText: i === 0 ? item.exampleText || "" : "",
          descriptionAbove: i === 0 ? item.description || null : null,
          itemIndex: itemIndex,
        });
      }
    });
    return flat;
  }

  /** 取得單頁內容區高度 257mm 在畫面上的 px 值；異常時用合理 fallback 避免只出一頁 */
  function getContentHeightPx() {
    var ruler = document.createElement("div");
    ruler.style.cssText =
      "position:absolute;left:-9999px;width:210mm;height:257mm;visibility:hidden;";
    document.body.appendChild(ruler);
    var h = ruler.offsetHeight;
    document.body.removeChild(ruler);
    if (!h || h < 100) h = 900;
    if (h > 2500) h = 1200;
    return h;
  }

  /** 插入圖片區塊在預覽/分頁時佔用的高度（70mm 換算成 px，與單頁內容區 257mm 同比例） */
  function getImageBlockHeightPx() {
    var contentH = getContentHeightPx();
    return Math.round((100 / 257) * contentH);
  }

  /** 1cm 的 px 值（用於需要時換算）；ruler 須設 10mm 才會量到正確高度，0mm 會得到 0 */
  function getBottomSafetyPx() {
    var ruler = document.createElement("div");
    ruler.style.cssText =
      "position:absolute;left:-9999px;width:1px;height:10mm;visibility:hidden;";
    document.body.appendChild(ruler);
    var h = ruler.offsetHeight;
    document.body.removeChild(ruler);
    return h && h > 0 ? h : 38;
  }

  /** 說明文字區塊高度（與行高成比例，用於分頁時計算該行總高度） */
  function getDescriptionHeightPx() {
    return Math.round(getLineHeightPx() * 0.65);
  }

  /**
   * 換頁用每行佔高：行距 > 0 用 lh+delta×lh；行距 ≤ 0 一律用 lh，不假設 slot 小於 lh，
   * 避免「中間一大段被吃掉」的裁切問題（實際版面每行至少 lh 高）。
   */
  function getSlotHeightForPagination() {
    var lh = getLineHeightPx();
    if (state.lineSpacingDelta > 0) {
      return lh + state.lineSpacingDelta * lh;
    }
    return lh;
  }

  /**
   * 依「文字大小、行距、說明文字高度」即時計算每頁可放幾行。
   * 支援分頁標記（強制換頁）與插入圖片區塊。
   */
  function computePages() {
    var fullH = getContentHeightPx();
    var contentH = Math.max(0, fullH);
    var lh = getLineHeightPx();
    var slotH = getSlotHeightForPagination();
    var descH = getDescriptionHeightPx();
    var imgH = getImageBlockHeightPx();
    var flat = getFlatEntries();
    var pages = [];
    var page = [];
    var used = 0;
    for (var i = 0; i < flat.length; i++) {
      var entry = flat[i];
      if (entry.pageBreak) {
        if (page.length > 0) {
          pages.push(page);
          page = [];
        }
        used = 0;
        continue;
      }
      if (entry.image) {
        if (used + imgH > contentH && page.length > 0) {
          pages.push(page);
          page = [];
          used = 0;
        }
        page.push(entry);
        used += imgH;
        continue;
      }
      var lineH = page.length === 0 ? lh : slotH;
      if (entry.descriptionAbove) {
        lineH += descH;
      }
      if (used + lineH > contentH && page.length > 0) {
        pages.push(page);
        page = [];
        used = 0;
        lineH = lh + (entry.descriptionAbove ? descH : 0);
      }
      page.push(entry);
      used += lineH;
    }
    if (page.length > 0) pages.push(page);
    if (pages.length === 0) pages.push([]);
    return pages;
  }

  /** 更新預覽區的 CSS 變數（字級、行高、行距 = 1ex + Δ×0.2×行高） */
  function applyPreviewVars() {
    var lh = getLineHeightPx();
    $worksheetPreview.style.setProperty(
      "--font-size-px",
      state.fontSize + "px",
    );
    $worksheetPreview.style.setProperty("--line-height-px", lh + "px");
    $worksheetPreview.style.setProperty(
      "--line-spacing",
      getLineSpacingPx() + "px",
    );
  }

  // 文字大小拉桿
  if ($fontSize && $fontSizeValue) {
    function updateFontSize() {
      state.fontSize = Number($fontSize.value);
      $fontSizeValue.textContent = state.fontSize + "px";
      applyPreviewVars();
      renderPreview();
    }
    $fontSize.addEventListener("input", updateFontSize);
    updateFontSize();
  }

  // 行距拉桿：範圍 -0.2～+0.2，step 0.01（內部 -20～20 再 ÷100）
  if ($lineSpacing && $lineSpacingValue) {
    function updateLineSpacing() {
      state.lineSpacingDelta = Number($lineSpacing.value) / 100;
      var v = state.lineSpacingDelta.toFixed(2);
      $lineSpacingValue.textContent =
        (state.lineSpacingDelta > 0 ? "+" : "") + v;
      applyPreviewVars();
      renderPreview();
    }
    $lineSpacing.addEventListener("input", updateLineSpacing);
    if (!$fontSize || !$fontSizeValue) applyPreviewVars();
    updateLineSpacing();
  }

  // 頁首
  if ($pageHeader) {
    function updatePageHeader() {
      state.pageHeader = $pageHeader.value.trim();
      renderPreview();
    }
    $pageHeader.addEventListener("input", updatePageHeader);
    $pageHeader.addEventListener("change", updatePageHeader);
    if (state.pageHeader) $pageHeader.value = state.pageHeader;
  }

  // 底線樣式
  document
    .querySelectorAll('input[name="lineStyle"]')
    .forEach(function (radio) {
      radio.addEventListener("change", function () {
        state.lineStyle = this.value;
        renderPreview();
      });
    });

  function addItem() {
    var newItem = ensureItemId({
      description: "",
      exampleText: "",
      lineCount: 1,
    });
    var insertIndex;
    if (
      state.focusedItemIndex != null &&
      state.focusedItemIndex >= 0 &&
      state.focusedItemIndex < state.items.length
    ) {
      insertIndex = state.focusedItemIndex + 1;
      state.items.splice(insertIndex, 0, newItem);
    } else {
      insertIndex = state.items.length;
      state.items.push(newItem);
    }
    state.focusedItemIndex = insertIndex;
    renderItemList();
    renderPreview();
    scrollPreviewToItem(insertIndex);
    var newBlock = $itemList.querySelector(
      '.item-block[data-item-index="' + insertIndex + '"]',
    );
    if (newBlock) {
      var firstInput = newBlock.querySelector("input");
      if (firstInput) firstInput.focus();
    }
  }

  function removeItem(index) {
    state.items.splice(index, 1);
    renderItemList();
    renderPreview();
  }

  function setItemField(index, field, value) {
    if (!state.items[index]) return;
    if (field === "lineCount") {
      state.items[index].lineCount = Math.max(1, parseInt(value, 10) || 1);
    } else {
      state.items[index][field] = value;
    }
    renderPreview();
  }

  var lastDragOverId = null;
  var draggedItemId = null;
  var flipCleanupTimeout = null;

  /** 依 state.items 順序重排左側列表 DOM，並更新 data-item-index */
  function reorderItemListDOM() {
    if (!$itemList || !state.items.length) return;
    var blocksById = {};
    var list = $itemList;
    for (var k = 0; k < list.children.length; k++) {
      var block = list.children[k];
      var id = block.getAttribute("data-item-id");
      if (id) blocksById[id] = block;
    }
    state.items.forEach(function (item, idx) {
      var block = blocksById[item.id];
      if (block) {
        block.setAttribute("data-item-index", String(idx));
        list.appendChild(block);
      }
    });
  }

  /** 重排並播放 FLIP 動畫（用雙 rAF 記錄「前一格」確保方向反轉時也能正確觸發動畫） */
  function reorderItemListWithFLIP() {
    if (!$itemList || !state.items.length) return;
    var list = $itemList;
    if (flipCleanupTimeout) {
      clearTimeout(flipCleanupTimeout);
      flipCleanupTimeout = null;
    }
    for (var k = 0; k < list.children.length; k++) {
      var block = list.children[k];
      block.style.transition = "";
      block.style.transform = "";
    }
    /* 第一個 rAF：僅讓瀏覽器有機會套用清除後的樣式與 layout */
    requestAnimationFrame(function () {
      requestAnimationFrame(function () {
        var blockRects = {};
        for (var k = 0; k < list.children.length; k++) {
          var block = list.children[k];
          var id = block.getAttribute("data-item-id");
          if (id) blockRects[id] = block.getBoundingClientRect();
        }
        reorderItemListDOM();
        list.offsetHeight; /* 強制 reflow，確保新順序的 layout 已套用 */
        for (var k = 0; k < list.children.length; k++) {
          var block = list.children[k];
          var id = block.getAttribute("data-item-id");
          var newRect = block.getBoundingClientRect();
          var oldRect = id ? blockRects[id] : null;
          if (oldRect) {
            var deltaY = oldRect.top - newRect.top;
            block.style.transform = "translateY(" + deltaY + "px)";
          }
        }
        requestAnimationFrame(function () {
          for (var k = 0; k < list.children.length; k++) {
            var block = list.children[k];
            block.style.transition = "transform 0.25s ease";
            block.style.transform = "translateY(0)";
          }
          flipCleanupTimeout = setTimeout(function () {
            flipCleanupTimeout = null;
            for (var k = 0; k < list.children.length; k++) {
              var block = list.children[k];
              block.style.transition = "";
              block.style.transform = "";
            }
          }, 260);
        });
      });
    });
  }

  function getBlockIndex(wrap) {
    return parseInt(wrap.getAttribute("data-item-index"), 10);
  }

  /** 拖曳重排後更新 focusedItemIndex */
  function reorderItemListUpdateFocus(fromIndex, toIndex) {
    if (state.focusedItemIndex === fromIndex) {
      state.focusedItemIndex = toIndex;
    } else if (
      state.focusedItemIndex != null &&
      fromIndex < state.focusedItemIndex &&
      toIndex >= state.focusedItemIndex
    ) {
      state.focusedItemIndex = state.focusedItemIndex - 1;
    } else if (
      state.focusedItemIndex != null &&
      fromIndex > state.focusedItemIndex &&
      toIndex <= state.focusedItemIndex
    ) {
      state.focusedItemIndex = state.focusedItemIndex + 1;
    }
  }

  function renderItemList() {
    $itemList.innerHTML = "";
    lastDragOverId = null;
    state.items.forEach(function (item, i) {
      var wrap = document.createElement("div");
      wrap.className = "item-block";
      wrap.setAttribute("data-item-index", String(i));
      wrap.setAttribute("data-item-id", item.id);
      if (item.type === "pageBreak") {
        wrap.classList.add("item-type-pagebreak");
        var row = document.createElement("div");
        row.className = "item-field item-field-last";
        var label = document.createElement("span");
        label.className = "item-special-label";
        label.textContent = "分頁標記";
        var btn = document.createElement("button");
        btn.type = "button";
        btn.className = "btn-remove";
        btn.setAttribute("aria-label", "移除此分頁標記");
        btn.textContent = "×";
        btn.addEventListener("click", function () {
          removeItem(getBlockIndex(wrap));
        });
        var dragHandle = document.createElement("span");
        dragHandle.className = "item-drag-handle";
        dragHandle.setAttribute("aria-label", "拖曳以調整順序");
        dragHandle.draggable = true;
        dragHandle.addEventListener("dragstart", function (e) {
          draggedItemId = item.id;
          e.dataTransfer.setData("text/plain", item.id);
          e.dataTransfer.effectAllowed = "move";
          wrap.classList.add("item-dragging");
          lastDragOverId = null;
        });
        dragHandle.addEventListener("dragend", function () {
          draggedItemId = null;
          wrap.classList.remove("item-dragging");
          lastDragOverId = null;
          $itemList.querySelectorAll(".item-block").forEach(function (el) {
            el.classList.remove("item-drag-over");
          });
          updatePreviewEditingOutline();
        });
        wrap.addEventListener("focusin", function () {
          state.focusedItemIndex = getBlockIndex(wrap);
          updatePreviewEditingOutline();
        });
        wrap.addEventListener("focusout", function (e) {
          var target = e.relatedTarget;
          if (wrap.contains(target)) return;
          if (
            target === $addItem ||
            (target && $addItem && $addItem.contains(target))
          )
            return;
          state.focusedItemIndex = null;
        });
        wrap.addEventListener("dragover", function (e) {
          if (e.dataTransfer.types.indexOf("text/plain") === -1) return;
          e.preventDefault();
          e.dataTransfer.dropEffect = "move";
          var blockId = wrap.getAttribute("data-item-id");
          if (!blockId) return;
          wrap.classList.add("item-drag-over");
          if (blockId === lastDragOverId) return;
          if (!draggedItemId || blockId === draggedItemId) return;
          lastDragOverId = blockId;
          var fromIndex = state.items.findIndex(function (it) {
            return it.id === draggedItemId;
          });
          var toIndex = getBlockIndex(wrap);
          if (fromIndex === -1 || fromIndex === toIndex) return;
          var moved = state.items[fromIndex];
          state.items.splice(fromIndex, 1);
          state.items.splice(toIndex, 0, moved);
          reorderItemListUpdateFocus(fromIndex, toIndex);
          reorderItemListWithFLIP();
        });
        wrap.addEventListener("dragleave", function () {
          wrap.classList.remove("item-drag-over");
        });
        wrap.addEventListener("drop", function (e) {
          e.preventDefault();
          wrap.classList.remove("item-drag-over");
          lastDragOverId = null;
          var draggedId = e.dataTransfer.getData("text/plain");
          var fromIndex = state.items.findIndex(function (it) {
            return it.id === draggedId;
          });
          var toIndex = getBlockIndex(wrap);
          if (fromIndex === -1) return;
          if (fromIndex !== toIndex) {
            var moved = state.items[fromIndex];
            state.items.splice(fromIndex, 1);
            state.items.splice(toIndex, 0, moved);
            reorderItemListUpdateFocus(fromIndex, toIndex);
            reorderItemListDOM();
          }
          renderPreview();
        });
        row.appendChild(label);
        row.appendChild(btn);
        row.appendChild(dragHandle);
        wrap.appendChild(row);
        $itemList.appendChild(wrap);
        return;
      }
      if (item.type === "image") {
        wrap.classList.add("item-type-image");
        var row = document.createElement("div");
        row.className = "item-field item-field-last";
        var label = document.createElement("span");
        label.className = "item-special-label";
        label.textContent =
          item.imagePath && item.imagePath.indexOf("變調") !== -1
            ? "變調圖"
            : "聲調音值圖";
        var btn = document.createElement("button");
        btn.type = "button";
        btn.className = "btn-remove";
        btn.setAttribute("aria-label", "移除此圖片");
        btn.textContent = "×";
        btn.addEventListener("click", function () {
          removeItem(getBlockIndex(wrap));
        });
        var dragHandle = document.createElement("span");
        dragHandle.className = "item-drag-handle";
        dragHandle.setAttribute("aria-label", "拖曳以調整順序");
        dragHandle.draggable = true;
        dragHandle.addEventListener("dragstart", function (e) {
          draggedItemId = item.id;
          e.dataTransfer.setData("text/plain", item.id);
          e.dataTransfer.effectAllowed = "move";
          wrap.classList.add("item-dragging");
          lastDragOverId = null;
        });
        dragHandle.addEventListener("dragend", function () {
          draggedItemId = null;
          wrap.classList.remove("item-dragging");
          lastDragOverId = null;
          $itemList.querySelectorAll(".item-block").forEach(function (el) {
            el.classList.remove("item-drag-over");
          });
          updatePreviewEditingOutline();
        });
        wrap.addEventListener("focusin", function () {
          state.focusedItemIndex = getBlockIndex(wrap);
          updatePreviewEditingOutline();
          scrollPreviewToItem(getBlockIndex(wrap));
        });
        wrap.addEventListener("focusout", function (e) {
          var target = e.relatedTarget;
          if (wrap.contains(target)) return;
          if (
            target === $addItem ||
            (target && $addItem && $addItem.contains(target))
          )
            return;
          state.focusedItemIndex = null;
        });
        wrap.addEventListener("dragover", function (e) {
          if (e.dataTransfer.types.indexOf("text/plain") === -1) return;
          e.preventDefault();
          e.dataTransfer.dropEffect = "move";
          var blockId = wrap.getAttribute("data-item-id");
          if (!blockId) return;
          wrap.classList.add("item-drag-over");
          if (blockId === lastDragOverId) return;
          if (!draggedItemId || blockId === draggedItemId) return;
          lastDragOverId = blockId;
          var fromIndex = state.items.findIndex(function (it) {
            return it.id === draggedItemId;
          });
          var toIndex = getBlockIndex(wrap);
          if (fromIndex === -1 || fromIndex === toIndex) return;
          var moved = state.items[fromIndex];
          state.items.splice(fromIndex, 1);
          state.items.splice(toIndex, 0, moved);
          reorderItemListUpdateFocus(fromIndex, toIndex);
          reorderItemListWithFLIP();
        });
        wrap.addEventListener("dragleave", function () {
          wrap.classList.remove("item-drag-over");
        });
        wrap.addEventListener("drop", function (e) {
          e.preventDefault();
          wrap.classList.remove("item-drag-over");
          lastDragOverId = null;
          var draggedId = e.dataTransfer.getData("text/plain");
          var fromIndex = state.items.findIndex(function (it) {
            return it.id === draggedId;
          });
          var toIndex = getBlockIndex(wrap);
          if (fromIndex === -1) return;
          if (fromIndex !== toIndex) {
            var moved = state.items[fromIndex];
            state.items.splice(fromIndex, 1);
            state.items.splice(toIndex, 0, moved);
            reorderItemListUpdateFocus(fromIndex, toIndex);
            reorderItemListDOM();
          }
          renderPreview();
        });
        row.appendChild(label);
        row.appendChild(btn);
        row.appendChild(dragHandle);
        wrap.appendChild(row);
        $itemList.appendChild(wrap);
        return;
      }
      function setFocusedItemIndex(value) {
        state.focusedItemIndex = value;
        updatePreviewEditingOutline();
      }
      wrap.addEventListener("focusin", function () {
        var idx = getBlockIndex(wrap);
        setFocusedItemIndex(idx);
        scrollPreviewToItem(idx);
      });
      wrap.addEventListener("focusout", function (e) {
        var target = e.relatedTarget;
        if (wrap.contains(target)) return;
        if (
          target === $addItem ||
          (target && $addItem && $addItem.contains(target))
        )
          return;
        if (
          (target && $insertMenu && $insertMenu.contains(target)) ||
          (target && $insertMenuToggle && $insertMenuToggle.contains(target))
        )
          return;
        setFocusedItemIndex(null);
      });
      var descRow = document.createElement("div");
      descRow.className = "item-field";
      var descLabel = document.createElement("label");
      descLabel.textContent = "項目標題";
      var descInput = document.createElement("input");
      descInput.type = "text";
      descInput.placeholder = "寫佇項目頂面的標題";
      descInput.value = item.description || "";
      descInput.addEventListener("input", function () {
        setItemField(getBlockIndex(wrap), "description", descInput.value);
      });
      descRow.appendChild(descLabel);
      descRow.appendChild(descInput);
      var exRow = document.createElement("div");
      exRow.className = "item-field";
      var exLabel = document.createElement("label");
      exLabel.textContent = "字詞見本";
      var exInput = document.createElement("input");
      exInput.type = "text";
      exInput.placeholder = "Kiàn-pún";
      exInput.value = item.exampleText || "";
      exInput.addEventListener("input", function () {
        setItemField(getBlockIndex(wrap), "exampleText", exInput.value);
      });
      exRow.appendChild(exLabel);
      exRow.appendChild(exInput);
      var lcRow = document.createElement("div");
      lcRow.className = "item-field item-field-last";
      var lcLabel = document.createElement("label");
      lcLabel.textContent = "練幾逝";
      var lcInput = document.createElement("input");
      lcInput.type = "number";
      lcInput.min = 1;
      lcInput.value = item.lineCount;
      lcInput.addEventListener("input", function () {
        var idx = getBlockIndex(wrap);
        setItemField(idx, "lineCount", lcInput.value);
        lcInput.value = state.items[idx].lineCount;
      });
      var btn = document.createElement("button");
      btn.type = "button";
      btn.className = "btn-remove";
      btn.setAttribute("aria-label", "移除此練習項目");
      btn.textContent = "×";
      btn.addEventListener("click", function () {
        removeItem(getBlockIndex(wrap));
      });
      var dragHandle = document.createElement("span");
      dragHandle.className = "item-drag-handle";
      dragHandle.setAttribute("aria-label", "拖曳以調整順序");
      dragHandle.draggable = true;
      dragHandle.addEventListener("dragstart", function (e) {
        draggedItemId = item.id;
        e.dataTransfer.setData("text/plain", item.id);
        e.dataTransfer.effectAllowed = "move";
        wrap.classList.add("item-dragging");
        lastDragOverId = null;
      });
      dragHandle.addEventListener("dragend", function () {
        draggedItemId = null;
        wrap.classList.remove("item-dragging");
        lastDragOverId = null;
        $itemList.querySelectorAll(".item-block").forEach(function (el) {
          el.classList.remove("item-drag-over");
        });
        updatePreviewEditingOutline();
      });
      lcRow.appendChild(lcLabel);
      lcRow.appendChild(lcInput);
      lcRow.appendChild(btn);
      lcRow.appendChild(dragHandle);
      wrap.addEventListener("dragover", function (e) {
        if (e.dataTransfer.types.indexOf("text/plain") === -1) return;
        e.preventDefault();
        e.dataTransfer.dropEffect = "move";
        var blockId = wrap.getAttribute("data-item-id");
        if (!blockId) return;
        wrap.classList.add("item-drag-over");
        if (blockId === lastDragOverId) return;
        if (!draggedItemId || blockId === draggedItemId) return;
        lastDragOverId = blockId;
        var fromIndex = state.items.findIndex(function (it) {
          return it.id === draggedItemId;
        });
        var toIndex = getBlockIndex(wrap);
        if (fromIndex === -1 || fromIndex === toIndex) return;
        var moved = state.items[fromIndex];
        state.items.splice(fromIndex, 1);
        state.items.splice(toIndex, 0, moved);
        var newIndex = toIndex;
        if (state.focusedItemIndex === fromIndex) {
          state.focusedItemIndex = newIndex;
        } else if (
          state.focusedItemIndex != null &&
          fromIndex < state.focusedItemIndex &&
          newIndex >= state.focusedItemIndex
        ) {
          state.focusedItemIndex = state.focusedItemIndex - 1;
        } else if (
          state.focusedItemIndex != null &&
          fromIndex > state.focusedItemIndex &&
          newIndex <= state.focusedItemIndex
        ) {
          state.focusedItemIndex = state.focusedItemIndex + 1;
        }
        reorderItemListWithFLIP();
      });
      wrap.addEventListener("dragleave", function () {
        wrap.classList.remove("item-drag-over");
      });
      wrap.addEventListener("drop", function (e) {
        e.preventDefault();
        wrap.classList.remove("item-drag-over");
        lastDragOverId = null;
        var draggedId = e.dataTransfer.getData("text/plain");
        var fromIndex = state.items.findIndex(function (it) {
          return it.id === draggedId;
        });
        var toIndex = getBlockIndex(wrap);
        if (fromIndex === -1) return;
        if (fromIndex !== toIndex) {
          var moved = state.items[fromIndex];
          state.items.splice(fromIndex, 1);
          var newIndex = toIndex;
          state.items.splice(toIndex, 0, moved);
          if (state.focusedItemIndex === fromIndex) {
            state.focusedItemIndex = newIndex;
          } else if (
            state.focusedItemIndex != null &&
            fromIndex < state.focusedItemIndex &&
            newIndex >= state.focusedItemIndex
          ) {
            state.focusedItemIndex = state.focusedItemIndex - 1;
          } else if (
            state.focusedItemIndex != null &&
            fromIndex > state.focusedItemIndex &&
            newIndex <= state.focusedItemIndex
          ) {
            state.focusedItemIndex = state.focusedItemIndex + 1;
          }
          reorderItemListDOM();
        }
        renderPreview();
      });
      wrap.appendChild(descRow);
      wrap.appendChild(exRow);
      wrap.appendChild(lcRow);
      $itemList.appendChild(wrap);
    });
  }

  /** 建立一筆顯示行（可選說明在上方 + 範例字 + 練習區 + 底線） */
  function createLineEl(flatLine, styleClass) {
    var wrap = document.createElement("div");
    wrap.className = "preview-line-wrap";
    if (flatLine.descriptionAbove) {
      var desc = document.createElement("div");
      desc.className = "line-description";
      desc.textContent = flatLine.descriptionAbove;
      wrap.appendChild(desc);
    }
    var row = document.createElement("div");
    row.className = "preview-line " + styleClass;
    var content = document.createElement("div");
    content.className = "row-content";
    var example = document.createElement("span");
    example.className = "example-text";
    example.textContent = flatLine.exampleText || "\u00A0";
    if (!flatLine.exampleText) {
      example.setAttribute("aria-hidden", "true");
      example.classList.add("example-text-placeholder");
    }
    content.appendChild(example);
    var zone = document.createElement("div");
    zone.className = "practice-zone";
    var baselineLine = document.createElement("span");
    baselineLine.className = "baseline-line";
    content.appendChild(zone);
    row.appendChild(content);
    var guideWrap = document.createElement("div");
    guideWrap.className = "baseline-guides";
    var baselineTop = document.createElement("span");
    baselineTop.className = "baseline-top";
    var baselineBottom = document.createElement("span");
    baselineBottom.className = "baseline-bottom";
    guideWrap.appendChild(baselineTop);
    guideWrap.appendChild(baselineLine);
    guideWrap.appendChild(baselineBottom);
    row.appendChild(guideWrap);
    wrap.appendChild(row);
    return wrap;
  }

  /** 依目前 focusedItemIndex 在預覽區為對應項目加上/移除編輯中外框 */
  function updatePreviewEditingOutline() {
    var index = state.focusedItemIndex;
    $worksheetPreview
      .querySelectorAll(".preview-item-group")
      .forEach(function (el) {
        var isActive =
          index != null &&
          String(el.getAttribute("data-item-index")) === String(index);
        el.classList.toggle("preview-item-editing", !!isActive);
      });
  }

  /** 將預覽區捲動到指定項目的第一個區塊 */
  function scrollPreviewToItem(itemIndex) {
    if (!$worksheetPreview || itemIndex == null) return;
    var first = $worksheetPreview.querySelector(
      '.preview-item-group[data-item-index="' + itemIndex + '"]',
    );
    if (first) {
      first.scrollIntoView({ behavior: "smooth", block: "start" });
    }
  }

  function renderPreview() {
    var savedScrollTop = $main ? $main.scrollTop : 0;
    $worksheetPreview.innerHTML = "";
    applyPreviewVars();
    var styleClass = state.lineStyle === "triple" ? "triple" : "single";
    var pages = computePages();
    pages.forEach(function (pageEntries) {
      var pageEl = document.createElement("div");
      pageEl.className = "worksheet-page";
      if (state.pageHeader) {
        var headerEl = document.createElement("div");
        headerEl.className = "worksheet-page-header";
        headerEl.textContent = state.pageHeader;
        pageEl.appendChild(headerEl);
      }
      var contentEl = document.createElement("div");
      contentEl.className = "worksheet-page-content";
      var linesEl = document.createElement("div");
      linesEl.className = "preview-lines";
      var currentGroup = null;
      var currentItemIndex = -1;
      for (var j = 0; j < pageEntries.length; j++) {
        var entry = pageEntries[j];
        if (entry.image) {
          if (currentGroup) linesEl.appendChild(currentGroup);
          currentGroup = document.createElement("div");
          currentGroup.className = "preview-item-group";
          currentGroup.setAttribute("data-item-index", String(entry.itemIndex));
          var imgWrap = document.createElement("div");
          imgWrap.className = "preview-image-block";
          var img = document.createElement("img");
          img.src = entry.image;
          img.alt =
            entry.image.indexOf("聲調音值") !== -1 ? "聲調音值圖" : "變調圖";
          imgWrap.appendChild(img);
          currentGroup.appendChild(imgWrap);
          linesEl.appendChild(currentGroup);
          currentGroup = null;
          currentItemIndex = -1;
          continue;
        }
        if (entry.itemIndex !== currentItemIndex) {
          if (currentGroup) linesEl.appendChild(currentGroup);
          currentGroup = document.createElement("div");
          currentGroup.className = "preview-item-group";
          currentGroup.setAttribute("data-item-index", String(entry.itemIndex));
          currentItemIndex = entry.itemIndex;
        }
        currentGroup.appendChild(createLineEl(entry, styleClass));
      }
      if (currentGroup) linesEl.appendChild(currentGroup);
      contentEl.appendChild(linesEl);
      pageEl.appendChild(contentEl);
      $worksheetPreview.appendChild(pageEl);
    });
    updatePreviewEditingOutline();
    if ($main) $main.scrollTop = savedScrollTop;
  }

  $addItem.addEventListener("click", addItem);

  if ($settingsToggle && $settingsPanel) {
    $settingsToggle.addEventListener("click", function () {
      var isCollapsed = $settingsPanel.classList.toggle("is-collapsed");
      $settingsToggle.setAttribute(
        "aria-expanded",
        isCollapsed ? "false" : "true",
      );
    });
  }

  /** 在目前項目下方插入一筆項目（分頁標記或圖片） */
  function insertBelowCurrent(item) {
    var insertIndex;
    if (
      state.focusedItemIndex != null &&
      state.focusedItemIndex >= 0 &&
      state.focusedItemIndex < state.items.length
    ) {
      insertIndex = state.focusedItemIndex + 1;
    } else {
      insertIndex = state.items.length;
    }
    ensureItemId(item);
    state.items.splice(insertIndex, 0, item);
    state.focusedItemIndex = insertIndex;
    renderItemList();
    renderPreview();
    scrollPreviewToItem(insertIndex);
    closeInsertMenu();
  }

  function openInsertMenu() {
    if (!$insertMenu || !$insertMenuToggle) return;
    $insertMenu.hidden = false;
    $insertMenuToggle.setAttribute("aria-expanded", "true");
    var rect = $insertMenuToggle.getBoundingClientRect();
    var menuTop = rect.bottom + 4;
    var menuLeft = rect.left;
    $insertMenu.style.top = menuTop + "px";
    $insertMenu.style.left = menuLeft + "px";
    requestAnimationFrame(function () {
      var menuRect = $insertMenu.getBoundingClientRect();
      if (menuRect.right > window.innerWidth) {
        $insertMenu.style.left = window.innerWidth - menuRect.width - 8 + "px";
      }
      if (menuRect.bottom > window.innerHeight) {
        $insertMenu.style.top = rect.top - menuRect.height - 4 + "px";
      }
    });
  }

  function closeInsertMenu() {
    if ($insertMenu) {
      $insertMenu.hidden = true;
      if ($insertMenuToggle)
        $insertMenuToggle.setAttribute("aria-expanded", "false");
    }
  }

  if ($insertMenuToggle && $insertMenu) {
    $insertMenuToggle.addEventListener("click", function (e) {
      e.stopPropagation();
      if ($insertMenu.hidden) {
        openInsertMenu();
      } else {
        closeInsertMenu();
      }
    });
    document.addEventListener("click", function () {
      closeInsertMenu();
    });
    $insertMenu.addEventListener("click", function (e) {
      e.stopPropagation();
    });
  }
  if ($insertPageBreak) {
    $insertPageBreak.addEventListener("click", function () {
      insertBelowCurrent({ type: "pageBreak" });
    });
  }
  if ($insertToneChart) {
    $insertToneChart.addEventListener("click", function () {
      insertBelowCurrent({ type: "image", imagePath: "images/聲調音值圖.png" });
    });
  }
  if ($insertSandhiChart) {
    $insertSandhiChart.addEventListener("click", function () {
      insertBelowCurrent({ type: "image", imagePath: "images/變調圖.png" });
    });
  }

  /** 載入教材：教材清單（textbooks/index.json 取得失敗時使用） */
  var defaultTextbookList = [
    { name: "台羅書寫練習", file: "台羅書寫練習.json" },
  ];

  function openLoadTextbookMenu(list) {
    if (!$loadTextbookMenu || !$loadTextbookBtn || !list || !list.length)
      return;
    $loadTextbookMenu.innerHTML = "";
    list.forEach(function (entry) {
      var name = entry.name || entry.file || "未命名教材";
      var file = entry.file;
      if (!file) return;
      var li = document.createElement("li");
      li.setAttribute("role", "none");
      var btn = document.createElement("button");
      btn.type = "button";
      btn.setAttribute("role", "menuitem");
      btn.setAttribute("data-file", file);
      btn.textContent = name;
      btn.addEventListener("click", function () {
        var path = "textbooks/" + file;
        fetch(path)
          .then(function (r) {
            if (!r.ok) throw new Error(r.statusText);
            return r.json();
          })
          .then(function (data) {
            applyImportedSettings(data);
            closeLoadTextbookMenu();
          })
          .catch(function (err) {
            alert(
              "無法載入教材：「" +
                name +
                "」\n" +
                (err && err.message ? err.message : "請確認檔案存在且可讀取。"),
            );
          });
      });
      li.appendChild(btn);
      $loadTextbookMenu.appendChild(li);
    });
    $loadTextbookMenu.hidden = false;
    $loadTextbookBtn.setAttribute("aria-expanded", "true");
    var rect = $loadTextbookBtn.getBoundingClientRect();
    $loadTextbookMenu.style.top = rect.bottom + 4 + "px";
    $loadTextbookMenu.style.left = rect.left + "px";
    requestAnimationFrame(function () {
      var menuRect = $loadTextbookMenu.getBoundingClientRect();
      if (menuRect.right > window.innerWidth) {
        $loadTextbookMenu.style.left =
          window.innerWidth - menuRect.width - 8 + "px";
      }
      if (menuRect.bottom > window.innerHeight) {
        $loadTextbookMenu.style.top = rect.top - menuRect.height - 4 + "px";
      }
    });
  }

  function closeLoadTextbookMenu() {
    if ($loadTextbookMenu) {
      $loadTextbookMenu.hidden = true;
      if ($loadTextbookBtn)
        $loadTextbookBtn.setAttribute("aria-expanded", "false");
    }
  }

  if ($loadTextbookBtn && $loadTextbookMenu) {
    $loadTextbookBtn.addEventListener("click", function (e) {
      e.stopPropagation();
      if ($loadTextbookMenu.hidden) {
        fetch("textbooks/index.json")
          .then(function (r) {
            if (!r.ok) throw new Error("無法取得教材清單");
            return r.json();
          })
          .then(function (list) {
            if (Array.isArray(list) && list.length > 0) {
              openLoadTextbookMenu(list);
            } else {
              openLoadTextbookMenu(defaultTextbookList);
            }
          })
          .catch(function () {
            openLoadTextbookMenu(defaultTextbookList);
          });
      } else {
        closeLoadTextbookMenu();
      }
    });
    document.addEventListener("click", function () {
      closeLoadTextbookMenu();
    });
    $loadTextbookMenu.addEventListener("click", function (e) {
      e.stopPropagation();
    });
  }

  /** 匯出設定為 JSON 檔 */
  function exportSettingsToJson() {
    var data = {
      version: 1,
      items: state.items.map(function (item) {
        if (item.type === "pageBreak") {
          return { type: "pageBreak", id: item.id };
        }
        if (item.type === "image" && item.imagePath) {
          return { type: "image", imagePath: item.imagePath, id: item.id };
        }
        return {
          description: item.description || "",
          exampleText: item.exampleText || "",
          lineCount: Math.max(1, parseInt(item.lineCount, 10) || 1),
          id: item.id,
        };
      }),
      lineStyle: state.lineStyle,
      lineSpacingDelta: state.lineSpacingDelta,
      fontSize: state.fontSize,
      pageHeader: state.pageHeader || "",
    };
    var json = JSON.stringify(data, null, 2);
    var blob = new Blob([json], { type: "application/json" });
    var url = URL.createObjectURL(blob);
    var a = document.createElement("a");
    a.href = url;
    a.download = "台語字練習簿設定.json";
    a.click();
    URL.revokeObjectURL(url);
  }

  /** 從 JSON 套用設定並更新介面 */
  function applyImportedSettings(data) {
    if (!data || typeof data !== "object") return;
    if (Array.isArray(data.items) && data.items.length > 0) {
      state.items = data.items.map(function (item) {
        if (item.type === "pageBreak") {
          return ensureItemId({ type: "pageBreak", id: item.id });
        }
        if (item.type === "image" && item.imagePath) {
          return ensureItemId({
            type: "image",
            imagePath: String(item.imagePath),
            id: item.id,
          });
        }
        return ensureItemId({
          description: String(item.description != null ? item.description : ""),
          exampleText: String(item.exampleText != null ? item.exampleText : ""),
          lineCount: Math.max(1, parseInt(item.lineCount, 10) || 1),
          id: item.id,
        });
      });
    }
    if (data.lineStyle === "single" || data.lineStyle === "triple") {
      state.lineStyle = data.lineStyle;
      document
        .querySelectorAll('input[name="lineStyle"]')
        .forEach(function (radio) {
          radio.checked = radio.value === state.lineStyle;
        });
    }
    if (
      typeof data.lineSpacingDelta === "number" &&
      data.lineSpacingDelta >= -0.2 &&
      data.lineSpacingDelta <= 0.2
    ) {
      state.lineSpacingDelta = data.lineSpacingDelta;
      if ($lineSpacing)
        $lineSpacing.value = Math.round(state.lineSpacingDelta * 100);
      if ($lineSpacingValue) {
        var v = state.lineSpacingDelta.toFixed(2);
        $lineSpacingValue.textContent =
          (state.lineSpacingDelta > 0 ? "+" : "") + v;
      }
    }
    if (
      typeof data.fontSize === "number" &&
      data.fontSize >= 16 &&
      data.fontSize <= 40
    ) {
      state.fontSize = data.fontSize;
      if ($fontSize) $fontSize.value = state.fontSize;
      if ($fontSizeValue) $fontSizeValue.textContent = state.fontSize + "px";
    }
    if (typeof data.pageHeader === "string") {
      state.pageHeader = data.pageHeader;
      if ($pageHeader) $pageHeader.value = state.pageHeader;
    }
    applyPreviewVars();
    renderItemList();
    renderPreview();
  }

  if ($importSettings && $importSettingsFile) {
    $importSettings.addEventListener("click", function () {
      $importSettingsFile.value = "";
      $importSettingsFile.click();
    });
    $importSettingsFile.addEventListener("change", function () {
      var file = this.files && this.files[0];
      if (!file) return;
      var reader = new FileReader();
      reader.onload = function () {
        try {
          var data = JSON.parse(reader.result);
          applyImportedSettings(data);
        } catch (e) {
          alert("無法解析設定檔，請確認是有效的 JSON 格式。");
        }
      };
      reader.readAsText(file, "UTF-8");
    });
  }
  if ($exportSettings) {
    $exportSettings.addEventListener("click", exportSettingsToJson);
  }

  $exportPdf.addEventListener("click", function () {
    var pageEls = $worksheetPreview.querySelectorAll(".worksheet-page");
    var pageArray = Array.prototype.slice.call(pageEls);
    if (pageArray.length === 0) return;
    var opt = {
      margin: 0,
      filename: "台語字練習簿.pdf",
      image: { type: "jpeg", quality: 0.98 },
      html2canvas: {
        scale: 2,
        useCORS: true,
        letterRendering: true,
        logging: false,
      },
      jsPDF: { unit: "mm", format: "a4", orientation: "portrait" },
    };

    /** 與預覽區使用相同計算方式，並把實際 px 值內聯到 clone 內所有相關節點，避免 html2canvas 對 CSS 變數解析不一致 */
    function applyPreviewVarsToClone(clone) {
      var lh = getLineHeightPx();
      var spacing = getLineSpacingPx();
      var fs = state.fontSize + "px";
      var lhPx = lh + "px";
      var spacingPx = spacing + "px";
      clone.style.setProperty("--font-size-px", fs);
      clone.style.setProperty("--line-height-px", lhPx);
      clone.style.setProperty("--line-spacing", spacingPx);
      // 內聯寫入 px，確保 html2canvas 擷取時與預覽一致
      var sheet = clone.querySelector(".worksheet-page-content");
      if (sheet) sheet.style.fontSize = fs;
      clone.querySelectorAll(".preview-line-wrap").forEach(function (el) {
        el.style.marginBottom = spacingPx;
      });
      clone
        .querySelectorAll(
          ".preview-lines > .preview-line-wrap:last-child, .preview-item-group .preview-line-wrap:last-child",
        )
        .forEach(function (el) {
          el.style.marginBottom = "0";
        });
      clone.querySelectorAll(".preview-line").forEach(function (el) {
        el.style.minHeight = lhPx;
        el.style.fontSize = fs;
      });
      clone
        .querySelectorAll(".preview-line .example-text")
        .forEach(function (el) {
          el.style.fontSize = fs;
          el.style.lineHeight = lhPx;
        });
      clone
        .querySelectorAll(".preview-line .practice-zone")
        .forEach(function (el) {
          el.style.minHeight = lhPx;
        });
      clone
        .querySelectorAll(".preview-line .baseline-guides")
        .forEach(function (el) {
          el.style.fontSize = fs;
        });
      // 讓 html2canvas 依 DOM 順序繪製時範例文字在底線上方：把 .row-content 移到 .baseline-guides 後面
      clone.querySelectorAll(".preview-line").forEach(function (line) {
        var content = line.querySelector(".row-content");
        var guides = line.querySelector(".baseline-guides");
        if (content && guides && content.nextSibling === guides) {
          content.remove();
          line.appendChild(content);
        }
      });
    }

    /** 將一頁複製到固定容器並轉成 canvas，避免在可捲動區外擷取導致第二頁以後空白 */
    function capturePageToCanvas(pageEl) {
      var wrap = document.createElement("div");
      wrap.style.cssText =
        "position:fixed;left:0;top:0;width:210mm;height:297mm;z-index:9999;pointer-events:none;";
      document.body.appendChild(wrap);
      var clone = pageEl.cloneNode(true);
      clone.style.boxShadow = "none";
      applyPreviewVarsToClone(clone);
      wrap.appendChild(clone);
      return new Promise(function (resolve, reject) {
        requestAnimationFrame(function () {
          requestAnimationFrame(function () {
            html2pdf()
              .set(opt)
              .from(clone)
              .toCanvas()
              .get("canvas")
              .then(function (canvas) {
                document.body.removeChild(wrap);
                resolve(canvas);
              })
              .catch(function (err) {
                if (document.body.contains(wrap))
                  document.body.removeChild(wrap);
                reject(err);
              });
          });
        });
      });
    }

    /** 第一頁也用 clone 在固定容器內產生 PDF，避免從多頁預覽取第一枚時 html2pdf 多插一張空白頁 */
    function createPdfFromFirstPage() {
      var wrap = document.createElement("div");
      wrap.style.cssText =
        "position:fixed;left:0;top:0;width:210mm;height:297mm;z-index:9999;pointer-events:none;";
      document.body.appendChild(wrap);
      var firstClone = pageArray[0].cloneNode(true);
      firstClone.style.boxShadow = "none";
      applyPreviewVarsToClone(firstClone);
      wrap.appendChild(firstClone);
      return new Promise(function (resolve, reject) {
        requestAnimationFrame(function () {
          requestAnimationFrame(function () {
            html2pdf()
              .set(opt)
              .from(firstClone)
              .toPdf()
              .get("pdf")
              .then(function (pdf) {
                if (document.body.contains(wrap))
                  document.body.removeChild(wrap);
                resolve(pdf);
              })
              .catch(function (err) {
                if (document.body.contains(wrap))
                  document.body.removeChild(wrap);
                reject(err);
              });
          });
        });
      });
    }

    createPdfFromFirstPage()
      .then(function (pdf) {
        // html2pdf 有時會多產生一頁空白，只保留第一頁
        while (pdf.getNumberOfPages && pdf.getNumberOfPages() > 1) {
          pdf.deletePage(pdf.getNumberOfPages());
        }
        function addNextPage(index) {
          if (index >= pageArray.length) {
            pdf.save(opt.filename);
            return;
          }
          return capturePageToCanvas(pageArray[index]).then(function (canvas) {
            var imgData = canvas.toDataURL("image/jpeg", 0.98);
            pdf.addPage();
            pdf.addImage(imgData, "JPEG", 0, 0, 210, 297);
            return addNextPage(index + 1);
          });
        }
        return addNextPage(1);
      })
      .catch(function (err) {
        console.error("PDF 匯出失敗", err);
        alert(
          "無法匯出 PDF：" + (err && err.message ? err.message : "請稍後再試"),
        );
      });
  });

  // 初始渲染
  renderItemList();
  renderPreview();
})();
