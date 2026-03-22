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

    // Stagger position so multiple images don't stack exactly
    const offset = inst.images.length * 20;
    const imgData = {
      src,
      x: 10 + offset,
      y: 10 + offset,
      w: 300,
      h: null,   // computed after image loads
    };

    // Create overlay container
    const overlay = document.createElement('div');
    overlay.style.cssText = `
      position: absolute;
      left: ${imgData.x}px;
      top:  ${imgData.y}px;
      width: ${imgData.w}px;
      cursor: move;
      user-select: none;
      z-index: 10;
    `;

    // The image itself
    const img = document.createElement('img');
    img.src = src;
    img.style.cssText = 'width:100%;display:block;border:2px solid transparent;box-sizing:border-box;';
    img.draggable = false;  // prevent browser native drag

    // Set height once loaded
    img.onload = () => {
      const ar = img.naturalHeight / img.naturalWidth;
      imgData.h = Math.round(imgData.w * ar);
      overlay.style.height = imgData.h + 'px';
    };

    // Resize handle (bottom-right corner)
    const handle = document.createElement('div');
    handle.style.cssText = `
      position: absolute;
      right: -6px; bottom: -6px;
      width: 14px; height: 14px;
      background: #00b8d9;
      border: 2px solid #fff;
      border-radius: 50%;
      cursor: se-resize;
      display: none;
      z-index: 11;
    `;

    // Delete button (top-right)
    const delBtn = document.createElement('div');
    delBtn.innerHTML = '✕';
    delBtn.style.cssText = `
      position: absolute;
      right: -8px; top: -8px;
      width: 18px; height: 18px;
      background: #e53e3e;
      color: #fff;
      border-radius: 50%;
      font-size: 10px;
      display: none;
      align-items: center;
      justify-content: center;
      cursor: pointer;
      z-index: 12;
      line-height: 18px;
      text-align: center;
    `;

    overlay.appendChild(img);
    overlay.appendChild(handle);
    overlay.appendChild(delBtn);
    inst.wrapper.appendChild(overlay);

    const entry = { overlay, img, handle, delBtn, imgData };
    inst.images.push(entry);

    // ── Click to select ───────────────────────────────────────────────────
    overlay.addEventListener('mousedown', e => {
      if (e.target === handle) return;  // let resize handle its own event
      if (e.target === delBtn) return;
      e.stopPropagation();
      _selectImg(inst, entry);
      _startDrag(e, inst, entry);
    });

    // ── Delete ────────────────────────────────────────────────────────────
    delBtn.addEventListener('mousedown', e => {
      e.stopPropagation();
      _removeImg(inst, entry);
    });

    // ── Resize ───────────────────────────────────────────────────────────
    handle.addEventListener('mousedown', e => {
      e.stopPropagation();
      e.preventDefault();
      _startResize(e, inst, entry);
    });

    // Auto-select when first inserted
    setTimeout(() => _selectImg(inst, entry), 50);
  }

  // ── Select / deselect ─────────────────────────────────────────────────────
  function _selectImg(inst, entry) {
    _deselectAll(inst);
    inst.selectedImg = entry;
    entry.img.style.border       = '2px solid #00b8d9';
    entry.handle.style.display   = 'block';
    entry.delBtn.style.display   = 'flex';
  }

  function _deselectAll(inst) {
    inst.images.forEach(e => {
      e.img.style.border       = '2px solid transparent';
      e.handle.style.display   = 'none';
      e.delBtn.style.display   = 'none';
    });
    inst.selectedImg = null;
  }

  function _removeImg(inst, entry) {
    entry.overlay.remove();
    inst.images = inst.images.filter(e => e !== entry);
    inst.selectedImg = null;
  }

  // ── Drag to move ──────────────────────────────────────────────────────────
  function _startDrag(e, inst, entry) {
    const startX  = e.clientX;
    const startY  = e.clientY;
    const startOX = entry.imgData.x;
    const startOY = entry.imgData.y;

    const onMove = ev => {
      const dx = ev.clientX - startX;
      const dy = ev.clientY - startY;
      entry.imgData.x = Math.max(0, startOX + dx);
      entry.imgData.y = Math.max(0, startOY + dy);
      entry.overlay.style.left = entry.imgData.x + 'px';
      entry.overlay.style.top  = entry.imgData.y + 'px';
    };

    const onUp = () => {
      document.removeEventListener('mousemove', onMove);
      document.removeEventListener('mouseup',   onUp);
    };

    document.addEventListener('mousemove', onMove);
    document.addEventListener('mouseup',   onUp);
  }

  // ── Resize from corner handle ─────────────────────────────────────────────
  function _startResize(e, inst, entry) {
    e.preventDefault();
    const startX  = e.clientX;
    const startW  = entry.imgData.w;
    const startH  = entry.imgData.h || entry.overlay.offsetHeight;
    const ar      = startH / Math.max(startW, 1);

    const onMove = ev => {
      const dx   = ev.clientX - startX;
      const newW = Math.max(60, startW + dx);
      const newH = Math.round(newW * ar);
      entry.imgData.w = newW;
      entry.imgData.h = newH;
      entry.overlay.style.width  = newW + 'px';
      entry.overlay.style.height = newH + 'px';
    };

    const onUp = () => {
      document.removeEventListener('mousemove', onMove);
      document.removeEventListener('mouseup',   onUp);
    };

    document.addEventListener('mousemove', onMove);
    document.addEventListener('mouseup',   onUp);
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
