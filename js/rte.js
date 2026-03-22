// ─────────────────────────────────────────────────────────────────────────────
// rte.js — Rich Text Editor with draggable/resizable images
// ─────────────────────────────────────────────────────────────────────────────

const RTE = (() => {

  const TOOLBAR_GROUPS = [
    [
      { cmd: 'bold',          icon: '<b>B</b>',   tip: 'Bold' },
      { cmd: 'italic',        icon: '<i>I</i>',   tip: 'Italic' },
      { cmd: 'underline',     icon: '<u>U</u>',   tip: 'Underline' },
      { cmd: 'strikeThrough', icon: '<s>S</s>',   tip: 'Strikethrough' },
    ], 'sep', [
      { cmd: 'insertUnorderedList', icon: '• List',  tip: 'Bullet list' },
      { cmd: 'insertOrderedList',   icon: '1. List', tip: 'Numbered list' },
      { cmd: 'outdent', icon: '⇤', tip: 'Outdent' },
      { cmd: 'indent',  icon: '⇥', tip: 'Indent' },
    ], 'sep', [
      { cmd: 'justifyLeft',   icon: '≡L', tip: 'Left' },
      { cmd: 'justifyCenter', icon: '≡C', tip: 'Center' },
      { cmd: 'justifyRight',  icon: '≡R', tip: 'Right' },
    ], 'sep', [
      { special: 'fontSize' },
      { special: 'foreColor' },
    ], 'sep', [
      { special: 'image',        icon: '🖼',    tip: 'Insert image' },
      { cmd: 'removeFormat', icon: '✕ fmt', tip: 'Clear formatting' },
    ],
  ];

  const _instances = {};

  // ── Init ──────────────────────────────────────────────────────────────────
  function init(editorId, toolbarId) {
    if (_instances[editorId]) return;
    const ed = document.getElementById(editorId);
    const tb = document.getElementById(toolbarId);
    if (!ed || !tb) return;

    // Wrap in relative container for image overlays
    const wrapper = document.createElement('div');
    wrapper.style.cssText = 'position:relative;';
    ed.parentNode.insertBefore(wrapper, ed);
    wrapper.appendChild(ed);

    const inst = { wrapper, editor: ed, images: [] };
    _instances[editorId] = inst;

    _buildToolbar(tb, editorId);

    // Paste images
    ed.addEventListener('paste', e => {
      const items = e.clipboardData?.items || [];
      for (const item of items) {
        if (item.type.startsWith('image/')) {
          e.preventDefault();
          const reader = new FileReader();
          reader.onload = ev => _addImage(editorId, ev.target.result);
          reader.readAsDataURL(item.getAsFile());
          return;
        }
      }
    });

    // Click on editor text deselects images
    ed.addEventListener('mousedown', () => _deselectAll(inst));
  }

  // ── Toolbar ───────────────────────────────────────────────────────────────
  function _buildToolbar(tb, editorId) {
    TOOLBAR_GROUPS.forEach(group => {
      if (group === 'sep') {
        const s = document.createElement('div');
        s.className = 'rte-sep'; tb.appendChild(s); return;
      }
      group.forEach(item => {
        if (item.special === 'fontSize') {
          const sel = document.createElement('select');
          sel.className = 'rte-select'; sel.title = 'Font size';
          [8,9,10,11,12,14,16,18,20,24,28,32,36].forEach(sz => {
            const o = document.createElement('option');
            o.value = sz; o.textContent = sz + 'pt';
            if (sz === 11) o.selected = true;
            sel.appendChild(o);
          });
          sel.addEventListener('change', () => {
            const ed = document.getElementById(editorId); ed?.focus();
            document.execCommand('fontSize', false, '7');
            ed?.querySelectorAll('font[size="7"]').forEach(f => {
              f.removeAttribute('size'); f.style.fontSize = sel.value + 'pt';
            });
          });
          tb.appendChild(sel);
        } else if (item.special === 'foreColor') {
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
        } else if (item.special === 'image') {
          const btn = document.createElement('button');
          btn.type = 'button'; btn.className = 'rte-btn';
          btn.title = item.tip; btn.innerHTML = item.icon;
          btn.addEventListener('click', () => {
            const fi = document.createElement('input');
            fi.type = 'file'; fi.accept = 'image/*'; fi.multiple = true;
            fi.addEventListener('change', () => {
              Array.from(fi.files).forEach(f => {
                const r = new FileReader();
                r.onload = ev => _addImage(editorId, ev.target.result);
                r.readAsDataURL(f);
              });
            });
            fi.click();
          });
          tb.appendChild(btn);
        } else {
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
      });
    });
  }

  // ── Add image ─────────────────────────────────────────────────────────────
  function _addImage(editorId, src) {
    const inst = _instances[editorId];
    if (!inst) return;

    const n   = inst.images.length;
    const x   = 10 + n * 22;
    const y   = 10 + n * 22;
    const w   = 320;

    // We need the natural dimensions — load first
    const probe = new Image();
    probe.onload = () => {
      const ar = probe.naturalHeight / Math.max(probe.naturalWidth, 1);
      const h  = Math.round(w * ar);
      _createImageEntry(inst, src, x, y, w, h);
    };
    probe.onerror = () => _createImageEntry(inst, src, x, y, w, 240);
    probe.src = src;
  }

  function _createImageEntry(inst, src, x, y, w, h) {
    // ── Overlay: the draggable image container ────────────────────────────
    const overlay = document.createElement('div');
    overlay.style.cssText =
      `position:absolute;left:${x}px;top:${y}px;width:${w}px;height:${h}px;` +
      'user-select:none;cursor:move;z-index:10;box-sizing:border-box;' +
      'border:2px solid transparent;';

    const img = document.createElement('img');
    img.src = src;
    img.draggable = false;
    img.style.cssText = 'width:100%;height:100%;display:block;object-fit:fill;pointer-events:none;';
    overlay.appendChild(img);

    // ── Delete button (top-right, inside overlay) ──────────────────────
    const del = document.createElement('div');
    del.innerHTML = '✕';
    del.style.cssText =
      'position:absolute;top:-10px;right:-10px;width:20px;height:20px;' +
      'background:#e53e3e;color:#fff;border-radius:50%;font-size:11px;font-weight:bold;' +
      'display:none;align-items:center;justify-content:center;text-align:center;' +
      'line-height:20px;cursor:pointer;z-index:20;pointer-events:all;';
    overlay.appendChild(del);

    // ── Resize handle (bottom-right, inside overlay) ───────────────────
    // Key: use pointer-events:all on handle, pointer-events:none on img
    // so clicks in the corner go to the handle div, not to overlay drag
    const res = document.createElement('div');
    res.style.cssText =
      'position:absolute;bottom:-10px;right:-10px;width:20px;height:20px;' +
      'background:#00b8d9;border:2px solid #fff;border-radius:50%;' +
      'cursor:se-resize;display:none;z-index:20;pointer-events:all;';
    overlay.appendChild(res);

    inst.wrapper.appendChild(overlay);

    const entry = { overlay, img, del, res, src, x, y, w, h };
    inst.images.push(entry);

    // ── Interactions ───────────────────────────────────────────────────

    // Select on click
    overlay.addEventListener('mousedown', e => {
      if (e.target === res) return; // resize handles itself
      if (e.target === del) return; // delete handles itself
      e.stopPropagation();
      e.preventDefault();
      _select(inst, entry);
      _drag(e.clientX, e.clientY, entry);
    });

    // Delete
    del.addEventListener('mousedown', e => {
      e.stopPropagation(); e.preventDefault();
      overlay.remove();
      inst.images = inst.images.filter(en => en !== entry);
      inst.selectedImg = null;
    });

    // Resize — ONLY on the res handle div
    res.addEventListener('mousedown', e => {
      e.stopPropagation(); e.preventDefault();
      _resize(e.clientX, e.clientY, entry);
    });
    res.addEventListener('touchstart', e => {
      e.stopPropagation(); e.preventDefault();
      _resize(e.touches[0].clientX, e.touches[0].clientY, entry);
    }, { passive: false });

    // Touch drag
    overlay.addEventListener('touchstart', e => {
      if (e.target === res || e.target === del) return;
      e.stopPropagation(); e.preventDefault();
      _select(inst, entry);
      _drag(e.touches[0].clientX, e.touches[0].clientY, entry);
    }, { passive: false });

    // Auto-select
    _select(inst, entry);
  }

  // ── Select ────────────────────────────────────────────────────────────────
  function _select(inst, entry) {
    _deselectAll(inst);
    inst.selectedImg = entry;
    entry.overlay.style.border = '2px solid #00b8d9';
    entry.del.style.display    = 'flex';
    entry.res.style.display    = 'block';
  }

  function _deselectAll(inst) {
    inst.images.forEach(e => {
      e.overlay.style.border = '2px solid transparent';
      e.del.style.display    = 'none';
      e.res.style.display    = 'none';
    });
    inst.selectedImg = null;
  }

  // ── Drag ──────────────────────────────────────────────────────────────────
  function _drag(startClientX, startClientY, entry) {
    const ox = entry.x, oy = entry.y;

    const move = (cx, cy) => {
      entry.x = Math.max(0, ox + (cx - startClientX));
      entry.y = Math.max(0, oy + (cy - startClientY));
      entry.overlay.style.left = entry.x + 'px';
      entry.overlay.style.top  = entry.y + 'px';
    };

    const onMouseMove = e => move(e.clientX, e.clientY);
    const onTouchMove = e => { e.preventDefault(); move(e.touches[0].clientX, e.touches[0].clientY); };
    const onUp = () => {
      document.removeEventListener('mousemove', onMouseMove);
      document.removeEventListener('mouseup',   onUp);
      document.removeEventListener('touchmove', onTouchMove);
      document.removeEventListener('touchend',  onUp);
    };
    document.addEventListener('mousemove', onMouseMove);
    document.addEventListener('mouseup',   onUp);
    document.addEventListener('touchmove', onTouchMove, { passive: false });
    document.addEventListener('touchend',  onUp);
  }

  // ── Resize ────────────────────────────────────────────────────────────────
  function _resize(startClientX, startClientY, entry) {
    const startW = entry.w;
    const startH = entry.h;
    const ar     = startH / Math.max(startW, 1);

    const move = (cx, cy) => {
      const dx   = cx - startClientX;
      const dy   = cy - startClientY;
      const delta = Math.abs(dx) >= Math.abs(dy) ? dx : dy / ar;
      const newW  = Math.max(60, startW + delta);
      const newH  = Math.round(newW * ar);
      entry.w = newW; entry.h = newH;
      entry.overlay.style.width  = newW + 'px';
      entry.overlay.style.height = newH + 'px';
    };

    const onMouseMove = e => move(e.clientX, e.clientY);
    const onTouchMove = e => { e.preventDefault(); move(e.touches[0].clientX, e.touches[0].clientY); };
    const onUp = () => {
      document.removeEventListener('mousemove', onMouseMove);
      document.removeEventListener('mouseup',   onUp);
      document.removeEventListener('touchmove', onTouchMove);
      document.removeEventListener('touchend',  onUp);
    };
    document.addEventListener('mousemove', onMouseMove);
    document.addEventListener('mouseup',   onUp);
    document.addEventListener('touchmove', onTouchMove, { passive: false });
    document.addEventListener('touchend',  onUp);
  }

  // ── Get images for docx export ────────────────────────────────────────────
  function getImages(editorId) {
    const inst = _instances[editorId];
    if (!inst) return [];
    return inst.images.map(e => ({ src: e.src, w: e.w, h: e.h, x: e.x, y: e.y }));
  }

  return { init, getImages };
})();
