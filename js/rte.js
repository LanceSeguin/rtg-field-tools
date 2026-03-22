// ─────────────────────────────────────────────────────────────────────────────
// rte.js — Rich Text Editor with proper resizable/draggable images
//
// Images are managed as absolutely-positioned overlays on top of the editor.
// This gives full drag-to-move and corner-handle resize without fighting
// the browser's contenteditable limitations.
// ─────────────────────────────────────────────────────────────────────────────

const RTE = (() => {

  const TOOLBAR_GROUPS = [
    [
      { cmd: 'bold',          icon: '<b>B</b>',   tip: 'Bold (Ctrl+B)' },
      { cmd: 'italic',        icon: '<i>I</i>',   tip: 'Italic (Ctrl+I)' },
      { cmd: 'underline',     icon: '<u>U</u>',   tip: 'Underline (Ctrl+U)' },
      { cmd: 'strikeThrough', icon: '<s>S</s>',   tip: 'Strikethrough' },
    ],
    'sep',
    [
      { cmd: 'insertUnorderedList', icon: '• List',  tip: 'Bullet list' },
      { cmd: 'insertOrderedList',   icon: '1. List', tip: 'Numbered list' },
      { cmd: 'outdent',  icon: '⇤', tip: 'Decrease indent' },
      { cmd: 'indent',   icon: '⇥', tip: 'Increase indent' },
    ],
    'sep',
    [
      { cmd: 'justifyLeft',   icon: '≡L', tip: 'Align left' },
      { cmd: 'justifyCenter', icon: '≡C', tip: 'Center' },
      { cmd: 'justifyRight',  icon: '≡R', tip: 'Align right' },
    ],
    'sep',
    [
      { special: 'fontSize' },
      { special: 'foreColor' },
    ],
    'sep',
    [
      { special: 'image',        icon: '🖼',     tip: 'Insert image from file' },
      { cmd:     'removeFormat', icon: '✕ fmt', tip: 'Clear formatting' },
    ],
  ];

  // ── Each editor instance tracks its own images ────────────────────────────
  const _instances = {};  // editorId → { wrapper, editor, images[] }

  // ── Init ──────────────────────────────────────────────────────────────────
  function init(editorId, toolbarId) {
    if (_instances[editorId]) return;  // already initialized

    const ed = document.getElementById(editorId);
    const tb = document.getElementById(toolbarId);
    if (!ed || !tb) return;

    // Wrap editor in a relative-positioned container for image overlays
    const wrapper = document.createElement('div');
    wrapper.style.cssText = 'position:relative;';
    ed.parentNode.insertBefore(wrapper, ed);
    wrapper.appendChild(ed);

    const inst = { wrapper, editor: ed, images: [], selectedImg: null };
    _instances[editorId] = inst;

    _buildToolbar(tb, editorId);
    _bindPaste(ed, editorId);

    // Click on editor text deselects images
    ed.addEventListener('mousedown', () => _deselectAll(inst));
  }

  // ── Toolbar ───────────────────────────────────────────────────────────────
  function _buildToolbar(tb, editorId) {
    TOOLBAR_GROUPS.forEach(group => {
      if (group === 'sep') {
        const s = document.createElement('div');
        s.className = 'rte-sep';
        tb.appendChild(s);
        return;
      }
      group.forEach(item => {
        if      (item.special === 'fontSize')  _addFontSize(tb, editorId);
        else if (item.special === 'foreColor') _addColorPicker(tb, editorId);
        else if (item.special === 'image')     _addImageBtn(tb, editorId, item);
        else                                   _addCmdBtn(tb, editorId, item);
      });
    });
  }

  function _addCmdBtn(tb, editorId, item) {
    const btn = document.createElement('button');
    btn.type = 'button'; btn.className = 'rte-btn';
    btn.title = item.tip || ''; btn.innerHTML = item.icon;
    btn.addEventListener('mousedown', e => {
      e.preventDefault();
      document.getElementById(editorId)?.focus();
      document.execCommand(item.cmd, false, null);
    });
    tb.appendChild(btn);
  }

  function _addFontSize(tb, editorId) {
    const sel = document.createElement('select');
    sel.className = 'rte-select'; sel.title = 'Font size';
    [8,9,10,11,12,14,16,18,20,24,28,32,36].forEach(sz => {
      const o = document.createElement('option');
      o.value = sz; o.textContent = sz + 'pt';
      if (sz === 11) o.selected = true;
      sel.appendChild(o);
    });
    sel.addEventListener('change', () => {
      const ed = document.getElementById(editorId);
      ed?.focus();
      document.execCommand('fontSize', false, '7');
      ed?.querySelectorAll('font[size="7"]').forEach(f => {
        f.removeAttribute('size');
        f.style.fontSize = sel.value + 'pt';
      });
    });
    tb.appendChild(sel);
  }

  function _addColorPicker(tb, editorId) {
    const lbl = document.createElement('label');
    lbl.className = 'rte-btn'; lbl.title = 'Text color';
    lbl.innerHTML = '🎨'; lbl.style.cursor = 'pointer';
    const inp = document.createElement('input');
    inp.type = 'color';
    inp.style.cssText = 'width:0;height:0;opacity:0;position:absolute;pointer-events:none;';
    inp.addEventListener('input', () => {
      document.getElementById(editorId)?.focus();
      document.execCommand('foreColor', false, inp.value);
    });
    lbl.appendChild(inp);
    lbl.addEventListener('mousedown', e => { e.preventDefault(); inp.click(); });
    tb.appendChild(lbl);
  }

  function _addImageBtn(tb, editorId, item) {
    const btn = document.createElement('button');
    btn.type = 'button'; btn.className = 'rte-btn';
    btn.title = item.tip; btn.innerHTML = item.icon;
    btn.addEventListener('click', () => {
      const fi = document.createElement('input');
      fi.type = 'file'; fi.accept = 'image/*'; fi.multiple = true;
      fi.addEventListener('change', () => {
        Array.from(fi.files).forEach(f => _insertImageFile(editorId, f));
      });
      fi.click();
    });
    tb.appendChild(btn);
  }

  // ── Paste images ──────────────────────────────────────────────────────────
  function _bindPaste(ed, editorId) {
    ed.addEventListener('paste', e => {
      const items = e.clipboardData?.items;
      if (!items) return;
      for (const item of items) {
        if (item.type.startsWith('image/')) {
          e.preventDefault();
          _insertImageFile(editorId, item.getAsFile());
          return;
        }
      }
    });
  }

  // ── Image insertion ───────────────────────────────────────────────────────
  function _insertImageFile(editorId, file) {
    const reader = new FileReader();
    reader.onload = ev => _insertImageSrc(editorId, ev.target.result);
    reader.readAsDataURL(file);
  }

  function _insertImageSrc(editorId, src) {
    const inst = _instances[editorId];
    if (!inst) return;

    const offset = inst.images.length * 20;
    const imgData = { src, x: 10 + offset, y: 10 + offset, w: 300, h: 200 };

    // Overlay container — drag target
    const overlay = document.createElement('div');
    overlay.style.cssText =
      `position:absolute;left:${imgData.x}px;top:${imgData.y}px;` +
      `width:${imgData.w}px;height:${imgData.h}px;` +
      'cursor:move;user-select:none;z-index:10;';

    // Image
    const img = document.createElement('img');
    img.src = src;
    img.draggable = false;
    img.style.cssText = 'width:100%;height:100%;display:block;object-fit:contain;' +
      'border:2px solid transparent;box-sizing:border-box;pointer-events:none;';

    img.onload = () => {
      const ar = img.naturalHeight / Math.max(img.naturalWidth, 1);
      imgData.h = Math.round(imgData.w * ar);
      overlay.style.height = imgData.h + 'px';
      _updateHandlePos(entry);
    };

    // Delete button — inside overlay top-right
    const delBtn = document.createElement('div');
    delBtn.innerHTML = '✕';
    delBtn.style.cssText =
      'position:absolute;top:-10px;right:-10px;width:20px;height:20px;' +
      'background:#e53e3e;color:#fff;border-radius:50%;font-size:11px;' +
      'display:none;align-items:center;justify-content:center;' +
      'cursor:pointer;z-index:30;line-height:20px;text-align:center;' +
      'pointer-events:all;font-weight:bold;';

    // Resize handle — SEPARATE element appended to wrapper, not inside overlay
    // This avoids all event bubbling issues between handle and overlay
    const handle = document.createElement('div');
    handle.style.cssText =
      'position:absolute;width:18px;height:18px;' +
      'background:#00b8d9;border:2px solid #fff;border-radius:50%;' +
      'cursor:se-resize;display:none;z-index:30;' +
      'box-shadow:0 2px 6px rgba(0,0,0,0.4);pointer-events:all;';

    overlay.appendChild(img);
    overlay.appendChild(delBtn);
    inst.wrapper.appendChild(overlay);
    inst.wrapper.appendChild(handle);  // handle is sibling of overlay, not child

    const entry = { overlay, img, handle, delBtn, imgData };
    inst.images.push(entry);

    function _updateHandlePos(en) {
      // Position handle at bottom-right corner of overlay
      const x = en.imgData.x + en.imgData.w - 6;
      const y = en.imgData.y + en.imgData.h - 6;
      en.handle.style.left = x + 'px';
      en.handle.style.top  = y + 'px';
    }
    entry._updateHandlePos = _updateHandlePos;

    // ── Handle mousedown — no bubbling issues since it's not inside overlay ─
    handle.addEventListener('mousedown', e => {
      e.preventDefault();
      e.stopPropagation();
      _selectImg(inst, entry);
      _startResize(e, inst, entry, _updateHandlePos);
    });
    handle.addEventListener('touchstart', e => {
      e.preventDefault();
      e.stopPropagation();
      _selectImg(inst, entry);
      _startResize(e.touches[0], inst, entry, _updateHandlePos);
    }, { passive: false });

    // ── Delete
    delBtn.addEventListener('mousedown', e => {
      e.preventDefault();
      e.stopPropagation();
      handle.remove();
      _removeImg(inst, entry);
    });

    // ── Overlay drag
    overlay.addEventListener('mousedown', e => {
      e.stopPropagation();
      _selectImg(inst, entry);
      _startDrag(e, inst, entry, _updateHandlePos);
    });
    overlay.addEventListener('touchstart', e => {
      e.stopPropagation();
      _selectImg(inst, entry);
      _startDrag(e.touches[0], inst, entry, _updateHandlePos);
    }, { passive: false });

    setTimeout(() => _selectImg(inst, entry), 50);
  }


  // ── Select / deselect ─────────────────────────────────────────────────────
  function _selectImg(inst, entry) {
    _deselectAll(inst);
    inst.selectedImg = entry;
    entry.img.style.border     = '2px solid #00b8d9';
    entry.handle.style.display = 'block';
    entry.delBtn.style.display = 'flex';
    if (entry._updateHandlePos) entry._updateHandlePos(entry);
  }

  function _deselectAll(inst) {
    inst.images.forEach(e => {
      e.img.style.border     = '2px solid transparent';
      e.handle.style.display = 'none';
      e.delBtn.style.display = 'none';
    });
    inst.selectedImg = null;
  }

  function _removeImg(inst, entry) {
    entry.overlay.remove();
    entry.handle.remove();
    inst.images = inst.images.filter(e => e !== entry);
    inst.selectedImg = null;
  }

  // ── Drag to move ──────────────────────────────────────────────────────────
  function _startDrag(e, inst, entry, updateHandle) {
    const startX  = e.clientX;
    const startY  = e.clientY;
    const startOX = entry.imgData.x;
    const startOY = entry.imgData.y;

    function _move(clientX, clientY) {
      const dx = clientX - startX;
      const dy = clientY - startY;
      entry.imgData.x = Math.max(0, startOX + dx);
      entry.imgData.y = Math.max(0, startOY + dy);
      entry.overlay.style.left = entry.imgData.x + 'px';
      entry.overlay.style.top  = entry.imgData.y + 'px';
      if (updateHandle) updateHandle(entry);
    }

    const onMove      = ev => _move(ev.clientX, ev.clientY);
    const onTouchMove = ev => { ev.preventDefault(); _move(ev.touches[0].clientX, ev.touches[0].clientY); };
    const onUp        = () => {
      document.removeEventListener('mousemove',  onMove);
      document.removeEventListener('mouseup',    onUp);
      document.removeEventListener('touchmove',  onTouchMove);
      document.removeEventListener('touchend',   onUp);
    };

    document.addEventListener('mousemove', onMove);
    document.addEventListener('mouseup',   onUp);
    document.addEventListener('touchmove', onTouchMove, { passive: false });
    document.addEventListener('touchend',  onUp);
  }

  // ── Resize from corner handle ─────────────────────────────────────────────
  function _startResize(e, inst, entry, updateHandle) {
    e.preventDefault();
    const startX  = e.clientX;
    const startY  = e.clientY;
    const startW  = entry.imgData.w;
    const startH  = entry.imgData.h || entry.overlay.offsetHeight;
    const ar      = startH / Math.max(startW, 1);

    const onMove = ev => {
      const dx   = ev.clientX - startX;
      const dy   = ev.clientY - startY;
      // Use whichever axis has more movement
      const delta = Math.abs(dx) >= Math.abs(dy) ? dx : dy / ar;
      const newW = Math.max(60, startW + delta);
      const newH = Math.round(newW * ar);
      entry.imgData.w = newW;
      entry.imgData.h = newH;
      // Update both the overlay container AND the image element
      entry.overlay.style.width  = newW + 'px';
      entry.overlay.style.height = newH + 'px';
      entry.img.style.width  = '100%';
      entry.img.style.height = '100%';
      if (updateHandle) updateHandle(entry);
    };

    const onTouchMove = ev => {
      ev.preventDefault();
      onMove({ clientX: ev.touches[0].clientX, clientY: ev.touches[0].clientY });
    };
    const onUp = () => {
      entry._resizing = false;
      document.removeEventListener('mousemove',  onMove);
      document.removeEventListener('mouseup',    onUp);
      document.removeEventListener('touchmove',  onTouchMove);
      document.removeEventListener('touchend',   onUp);
    };

    document.addEventListener('mousemove', onMove);
    document.addEventListener('mouseup',   onUp);
    document.addEventListener('touchmove', onTouchMove, { passive: false });
    document.addEventListener('touchend',  onUp);
  }

  // ── Get image data for docx export ───────────────────────────────────────
  function getImages(editorId) {
    const inst = _instances[editorId];
    if (!inst) return [];
    return inst.images.map(e => ({
      src: e.imgData.src,
      w:   e.imgData.w,
      h:   e.imgData.h || e.overlay.offsetHeight,
      x:   e.imgData.x,
      y:   e.imgData.y,
    }));
  }

  return { init, getImages };
})();
