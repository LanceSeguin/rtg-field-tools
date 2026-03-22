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
      { special: 'image',    icon: '🖼',    tip: 'Insert image' },
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
      _createEntry(inst, src, 10 + n*22, 10 + n*22, w, h);
    };
    probe.onerror = () => _createEntry(inst, src, 10 + n*22, 10 + n*22, 300, 220);
    probe.src = src;
  }

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

    // Delete button — top right
    const del = document.createElement('div');
    del.innerHTML = '✕';
    del.style.cssText =
      'position:absolute;top:-10px;right:-10px;width:20px;height:20px;' +
      'background:#e53e3e;color:#fff;border-radius:50%;font-size:11px;font-weight:bold;' +
      'display:none;align-items:center;justify-content:center;text-align:center;' +
      'line-height:20px;cursor:pointer;z-index:30;';
    overlay.appendChild(del);

    // Resize handle — bottom right, LARGER hit area, different cursor
    const res = document.createElement('div');
    res.style.cssText =
      'position:absolute;bottom:-10px;right:-10px;width:24px;height:24px;' +
      'background:#00b8d9;border:3px solid #fff;border-radius:50%;' +
      'cursor:se-resize;display:none;z-index:30;';
    overlay.appendChild(res);

    inst.wrapper.appendChild(overlay);

    const entry = { overlay, img, del, res, src, x, y, w, h };
    inst.images.push(entry);

    // ── Single mousedown handler on overlay ───────────────────────────────
    // Determine action by checking what element was clicked using getBoundingClientRect
    overlay.addEventListener('mousedown', e => {
      e.stopPropagation();
      e.preventDefault();
      _select(inst, entry);

      // Check if click landed within the res handle's bounding box
      const resRect = res.getBoundingClientRect();
      const inRes = (
        e.clientX >= resRect.left && e.clientX <= resRect.right &&
        e.clientY >= resRect.top  && e.clientY <= resRect.bottom
      );

      // Check if click landed within the del button's bounding box
      const delRect = del.getBoundingClientRect();
      const inDel = (
        e.clientX >= delRect.left && e.clientX <= delRect.right &&
        e.clientY >= delRect.top  && e.clientY <= delRect.bottom
      );

      if (inDel) {
        overlay.remove();
        inst.images = inst.images.filter(en => en !== entry);
        inst.selectedImg = null;
        return;
      }

      if (inRes) {
        _resize(e.clientX, e.clientY, entry);
        return;
      }

      _drag(e.clientX, e.clientY, entry);
    });

    // Touch version — same coordinate-based hit test
    overlay.addEventListener('touchstart', e => {
      e.stopPropagation(); e.preventDefault();
      _select(inst, entry);
      const t = e.touches[0];

      const resRect = res.getBoundingClientRect();
      const inRes = (
        t.clientX >= resRect.left && t.clientX <= resRect.right &&
        t.clientY >= resRect.top  && t.clientY <= resRect.bottom
      );

      if (inRes) { _resize(t.clientX, t.clientY, entry); return; }
      _drag(t.clientX, t.clientY, entry);
    }, { passive: false });

    _select(inst, entry);
  }

  // ── Select / deselect ─────────────────────────────────────────────────────
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
      document.removeEventListener('mouseup', up);
      document.removeEventListener('touchmove', tm);
      document.removeEventListener('touchend', up);
    };
    document.addEventListener('mousemove', mm);
    document.addEventListener('mouseup', up);
    document.addEventListener('touchmove', tm, { passive: false });
    document.addEventListener('touchend', up);
  }

  // ── Resize ────────────────────────────────────────────────────────────────
  function _resize(startX, startY, entry) {
    const sw = entry.w, sh = entry.h;
    const ar = sh / Math.max(sw, 1);
    const move = (cx, cy) => {
      const dx = cx - startX;
      const dy = cy - startY;
      const delta = Math.abs(dx) >= Math.abs(dy) ? dx : dy / ar;
      const nw = Math.max(60, sw + delta);
      const nh = Math.round(nw * ar);
      entry.w = nw; entry.h = nh;
      entry.overlay.style.width  = nw + 'px';
      entry.overlay.style.height = nh + 'px';
    };
    const mm = e => move(e.clientX, e.clientY);
    const tm = e => { e.preventDefault(); move(e.touches[0].clientX, e.touches[0].clientY); };
    const up = () => {
      document.removeEventListener('mousemove', mm);
      document.removeEventListener('mouseup', up);
      document.removeEventListener('touchmove', tm);
      document.removeEventListener('touchend', up);
    };
    document.addEventListener('mousemove', mm);
    document.addEventListener('mouseup', up);
    document.addEventListener('touchmove', tm, { passive: false });
    document.addEventListener('touchend', up);
  }

  // ── Get images for docx export ────────────────────────────────────────────
  function getImages(editorId) {
    const inst = _instances[editorId];
    if (!inst) return [];
    return inst.images.map(e => ({ src: e.src, w: e.w, h: e.h, x: e.x, y: e.y }));
  }

  return { init, getImages };
})();
