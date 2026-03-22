// ─────────────────────────────────────────────────────────────────────────────
// rte.js — Rich Text Editor with draggable/resizable images
// Resize by hovering near any edge/corner and dragging.
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
      { special: 'image',    icon: '🖼',    tip: 'Insert image' },
      { cmd: 'removeFormat', icon: '✕ fmt', tip: 'Clear formatting' },
    ],
  ];

  const _instances = {};
  const EDGE = 14; // pixels from edge that trigger resize cursor

  // ── Init ──────────────────────────────────────────────────────────────────
  function init(editorId, toolbarId) {
    if (_instances[editorId]) return;
    const ed = document.getElementById(editorId);
    const tb = document.getElementById(toolbarId);
    if (!ed || !tb) return;

    const wrapper = document.createElement('div');
    wrapper.style.cssText = 'position:relative;';
    ed.parentNode.insertBefore(wrapper, ed);
    wrapper.appendChild(ed);

    const inst = { wrapper, editor: ed, images: [] };
    _instances[editorId] = inst;

    _buildToolbar(tb, editorId);

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
    const n = inst.images.length;
    const probe = new Image();
    probe.onload = () => {
      const ar = probe.naturalHeight / Math.max(probe.naturalWidth, 1);
      const w = 300, h = Math.round(w * ar);
      _createEntry(inst, src, 10 + n * 22, 10 + n * 22, w, h);
    };
    probe.onerror = () => _createEntry(inst, src, 10 + n * 22, 10 + n * 22, 300, 220);
    probe.src = src;
  }

  // ── Determine resize direction from mouse position within overlay ─────────
  // Returns a direction string like 'se', 'e', 's', '' (center = drag)
  function _getEdge(e, overlay) {
    const r  = overlay.getBoundingClientRect();
    const mx = e.clientX - r.left;   // mouse x relative to overlay
    const my = e.clientY - r.top;    // mouse y relative to overlay
    const w  = r.width;
    const h  = r.height;

    const nearRight  = mx >= w - EDGE;
    const nearBottom = my >= h - EDGE;
    const nearLeft   = mx <= EDGE;
    const nearTop    = my <= EDGE;

    if (nearBottom && nearRight) return 'se';
    if (nearBottom && nearLeft)  return 'sw';
    if (nearTop    && nearRight) return 'ne';
    if (nearTop    && nearLeft)  return 'nw';
    if (nearRight)  return 'e';
    if (nearBottom) return 's';
    if (nearLeft)   return 'w';
    if (nearTop)    return 'n';
    return ''; // center — drag
  }

  // ── Cursor for each edge ──────────────────────────────────────────────────
  const CURSORS = {
    se: 'se-resize', sw: 'sw-resize', ne: 'ne-resize', nw: 'nw-resize',
    e:  'e-resize',  w:  'w-resize',  s:  's-resize',  n:  'n-resize',
    '': 'move',
  };

  // ── Create image entry ────────────────────────────────────────────────────
  function _createEntry(inst, src, x, y, w, h) {
    const overlay = document.createElement('div');
    overlay.style.cssText =
      `position:absolute;left:${x}px;top:${y}px;width:${w}px;height:${h}px;` +
      'user-select:none;z-index:10;box-sizing:border-box;border:2px solid transparent;cursor:move;';

    const img = document.createElement('img');
    img.src = src; img.draggable = false;
    img.style.cssText = 'width:100%;height:100%;display:block;object-fit:fill;pointer-events:none;';
    overlay.appendChild(img);

    // Delete button
    const del = document.createElement('div');
    del.innerHTML = '✕';
    del.style.cssText =
      'position:absolute;top:-10px;right:-10px;width:20px;height:20px;' +
      'background:#e53e3e;color:#fff;border-radius:50%;font-size:11px;font-weight:bold;' +
      'display:none;align-items:center;justify-content:center;text-align:center;' +
      'line-height:20px;cursor:pointer;z-index:30;pointer-events:all;';
    del.addEventListener('mousedown', e => {
      e.stopPropagation(); e.preventDefault();
      overlay.remove();
      inst.images = inst.images.filter(en => en !== entry);
      inst.selectedImg = null;
    });
    overlay.appendChild(del);

    inst.wrapper.appendChild(overlay);

    const entry = { overlay, img, del, src, x, y, w, h };
    inst.images.push(entry);

    // ── Mousemove: update cursor based on edge proximity ─────────────────
    overlay.addEventListener('mousemove', e => {
      const edge = _getEdge(e, overlay);
      overlay.style.cursor = CURSORS[edge];
    });

    overlay.addEventListener('mouseleave', () => {
      overlay.style.cursor = 'move';
    });

    // ── Mousedown: select then drag or resize based on edge ───────────────
    overlay.addEventListener('mousedown', e => {
      // Skip if it's the delete button
      if (e.target === del || del.contains(e.target)) return;

      e.stopPropagation();
      e.preventDefault();
      _select(inst, entry);

      const edge = _getEdge(e, overlay);

      if (edge === '') {
        // Center — drag
        _drag(e.clientX, e.clientY, entry);
      } else {
        // Edge — resize
        _resize(e.clientX, e.clientY, edge, entry);
      }
    });

    // Touch
    overlay.addEventListener('touchstart', e => {
      if (e.target === del || del.contains(e.target)) return;
      e.stopPropagation(); e.preventDefault();
      _select(inst, entry);
      const t = e.touches[0];
      const edge = _getEdge(t, overlay);
      if (edge === '') {
        _drag(t.clientX, t.clientY, entry);
      } else {
        _resize(t.clientX, t.clientY, edge, entry);
      }
    }, { passive: false });

    _select(inst, entry);
  }

  // ── Select / deselect ─────────────────────────────────────────────────────
  function _select(inst, entry) {
    _deselectAll(inst);
    inst.selectedImg = entry;
    entry.overlay.style.border = '2px solid #00b8d9';
    entry.del.style.display    = 'flex';
  }

  function _deselectAll(inst) {
    inst.images.forEach(e => {
      e.overlay.style.border = '2px solid transparent';
      e.del.style.display    = 'none';
    });
    inst.selectedImg = null;
  }

  // ── Drag ──────────────────────────────────────────────────────────────────
  function _drag(startX, startY, entry) {
    const ox = entry.x, oy = entry.y;
    const move = (cx, cy) => {
      entry.x = Math.max(0, ox + cx - startX);
      entry.y = Math.max(0, oy + cy - startY);
      entry.overlay.style.left = entry.x + 'px';
      entry.overlay.style.top  = entry.y + 'px';
    };
    const mm = e => move(e.clientX, e.clientY);
    const tm = e => { e.preventDefault(); move(e.touches[0].clientX, e.touches[0].clientY); };
    const up = () => {
      document.removeEventListener('mousemove', mm);
      document.removeEventListener('mouseup',   up);
      document.removeEventListener('touchmove', tm);
      document.removeEventListener('touchend',  up);
    };
    document.addEventListener('mousemove', mm);
    document.addEventListener('mouseup',   up);
    document.addEventListener('touchmove', tm, { passive: false });
    document.addEventListener('touchend',  up);
  }

  // ── Resize ────────────────────────────────────────────────────────────────
  function _resize(startX, startY, edge, entry) {
    const sw = entry.w, sh = entry.h;
    const sx = entry.x, sy = entry.y;
    const ar = sh / Math.max(sw, 1);

    const move = (cx, cy) => {
      const dx = cx - startX;
      const dy = cy - startY;
      let nw = sw, nh = sh, nx = sx, ny = sy;

      // Calculate new dimensions based on which edge is being dragged
      if (edge.includes('e'))  nw = Math.max(60, sw + dx);
      if (edge.includes('s'))  nh = Math.max(40, sh + dy);
      if (edge.includes('w')) { nw = Math.max(60, sw - dx); nx = sx + (sw - nw); }
      if (edge.includes('n')) { nh = Math.max(40, sh - dy); ny = sy + (sh - nh); }

      // For corner drags, maintain aspect ratio using the larger delta
      if (edge.length === 2) {
        const delta = Math.abs(dx) >= Math.abs(dy) ? nw / sw : nh / sh;
        nw = Math.round(sw * delta);
        nh = Math.round(sh * delta);
        nw = Math.max(60, nw);
        nh = Math.max(40, nh);
        if (edge.includes('w')) nx = sx + (sw - nw);
        if (edge.includes('n')) ny = sy + (sh - nh);
      }

      entry.w = nw; entry.h = nh;
      entry.x = nx; entry.y = ny;
      entry.overlay.style.width  = nw + 'px';
      entry.overlay.style.height = nh + 'px';
      entry.overlay.style.left   = nx + 'px';
      entry.overlay.style.top    = ny + 'px';
    };

    const mm = e => move(e.clientX, e.clientY);
    const tm = e => { e.preventDefault(); move(e.touches[0].clientX, e.touches[0].clientY); };
    const up = () => {
      document.removeEventListener('mousemove', mm);
      document.removeEventListener('mouseup',   up);
      document.removeEventListener('touchmove', tm);
      document.removeEventListener('touchend',  up);
    };
    document.addEventListener('mousemove', mm);
    document.addEventListener('mouseup',   up);
    document.addEventListener('touchmove', tm, { passive: false });
    document.addEventListener('touchend',  up);
  }

  // ── Get images for docx export ────────────────────────────────────────────
  function getImages(editorId) {
    const inst = _instances[editorId];
    if (!inst) return [];
    return inst.images.map(e => ({ src: e.src, w: e.w, h: e.h, x: e.x, y: e.y }));
  }

  return { init, getImages };
})();
