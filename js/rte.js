// ─────────────────────────────────────────────────────────────────────────────
// rte.js — Rich Text Editor (built from scratch, no libraries)
// Edit this file to add toolbar buttons or change editor behavior.
// ─────────────────────────────────────────────────────────────────────────────

const RTE = (() => {

  // ── Toolbar button definitions ────────────────────────────────────────────
  // To add a button: add an entry here with cmd (execCommand name) and icon/label
  const TOOLBAR_GROUPS = [
    [
      { cmd: 'bold',          icon: '<b>B</b>',  tip: 'Bold (Ctrl+B)' },
      { cmd: 'italic',        icon: '<i>I</i>',  tip: 'Italic (Ctrl+I)' },
      { cmd: 'underline',     icon: '<u>U</u>',  tip: 'Underline (Ctrl+U)' },
      { cmd: 'strikeThrough', icon: '<s>S</s>',  tip: 'Strikethrough' },
    ],
    'sep',
    [
      { cmd: 'insertUnorderedList', icon: '• List', tip: 'Bullet list' },
      { cmd: 'insertOrderedList',   icon: '1. List', tip: 'Numbered list' },
      { cmd: 'outdent', icon: '⇤', tip: 'Decrease indent' },
      { cmd: 'indent',  icon: '⇥', tip: 'Increase indent' },
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
      { special: 'image',  icon: '🖼', tip: 'Insert image from file' },
      { cmd: 'removeFormat', icon: '✕ fmt', tip: 'Clear all formatting' },
    ],
  ];

  // ── Initialize an editor + toolbar pair ──────────────────────────────────
  function init(editorId, toolbarId) {
    const ed = document.getElementById(editorId);
    const tb = document.getElementById(toolbarId);

    _buildToolbar(tb, ed);
    _bindPaste(ed);
    _bindPlaceholder(ed);
  }

  // ── Toolbar builder ───────────────────────────────────────────────────────
  function _buildToolbar(tb, ed) {
    TOOLBAR_GROUPS.forEach(group => {
      if (group === 'sep') {
        const sep = document.createElement('div');
        sep.className = 'rte-sep';
        tb.appendChild(sep);
        return;
      }

      group.forEach(item => {
        if      (item.special === 'fontSize')  _addFontSize(tb, ed);
        else if (item.special === 'foreColor') _addColorPicker(tb, ed);
        else if (item.special === 'image')     _addImageBtn(tb, ed, item);
        else                                   _addCmdBtn(tb, ed, item);
      });
    });
  }

  function _addCmdBtn(tb, ed, item) {
    const btn = document.createElement('button');
    btn.type      = 'button';
    btn.className = 'rte-btn';
    btn.title     = item.tip || '';
    btn.innerHTML = item.icon;
    btn.addEventListener('mousedown', e => {
      e.preventDefault(); // keep focus in editor
      ed.focus();
      document.execCommand(item.cmd, false, null);
    });
    tb.appendChild(btn);
  }

  function _addFontSize(tb, ed) {
    const sel = document.createElement('select');
    sel.className = 'rte-select';
    sel.title = 'Font size';
    [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36].forEach(sz => {
      const o = document.createElement('option');
      o.value = sz; o.textContent = sz + 'pt';
      if (sz === 11) o.selected = true;
      sel.appendChild(o);
    });
    sel.addEventListener('change', () => {
      ed.focus();
      // execCommand fontSize only accepts 1-7; we override with CSS
      document.execCommand('fontSize', false, '7');
      ed.querySelectorAll('font[size="7"]').forEach(f => {
        f.removeAttribute('size');
        f.style.fontSize = sel.value + 'pt';
      });
    });
    tb.appendChild(sel);
  }

  function _addColorPicker(tb, ed) {
    const lbl = document.createElement('label');
    lbl.className = 'rte-btn';
    lbl.title     = 'Text color';
    lbl.innerHTML = '🎨';
    lbl.style.cursor = 'pointer';

    const inp = document.createElement('input');
    inp.type  = 'color';
    inp.style.cssText = 'width:0;height:0;opacity:0;position:absolute;pointer-events:none;';
    inp.addEventListener('input', () => {
      ed.focus();
      document.execCommand('foreColor', false, inp.value);
    });

    lbl.appendChild(inp);
    lbl.addEventListener('mousedown', e => {
      e.preventDefault();
      inp.click();
    });
    tb.appendChild(lbl);
  }

  function _addImageBtn(tb, ed, item) {
    const btn = document.createElement('button');
    btn.type      = 'button';
    btn.className = 'rte-btn';
    btn.title     = item.tip;
    btn.innerHTML = item.icon;
    btn.addEventListener('click', () => {
      const fi = document.createElement('input');
      fi.type     = 'file';
      fi.accept   = 'image/*';
      fi.multiple = true;
      fi.addEventListener('change', () => {
        Array.from(fi.files).forEach(f => _insertImageFile(ed, f));
      });
      fi.click();
    });
    tb.appendChild(btn);
  }

  // ── Image insertion ───────────────────────────────────────────────────────
  function _insertImageFile(ed, file) {
    const reader = new FileReader();
    reader.onload = ev => _insertImageSrc(ed, ev.target.result);
    reader.readAsDataURL(file);
  }

  function _insertImageSrc(ed, src) {
    ed.focus();
    const img = document.createElement('img');
    img.src = src;
    img.style.cssText = 'max-width:100%;cursor:move;display:block;margin:4px 0;';
    img.draggable = true;

    // Click to select
    img.addEventListener('click', e => {
      e.stopPropagation();
      ed.querySelectorAll('img').forEach(i => i.classList.remove('selected'));
      img.classList.add('selected');
    });

    // Insert at cursor, or append if no selection
    const sel = window.getSelection();
    if (sel && sel.rangeCount) {
      const range = sel.getRangeAt(0);
      range.insertNode(img);
      range.collapse(false);
    } else {
      ed.appendChild(img);
    }
  }

  // ── Paste handler (Ctrl+V images) ─────────────────────────────────────────
  function _bindPaste(ed) {
    ed.addEventListener('paste', e => {
      const items = e.clipboardData?.items;
      if (!items) return;
      for (const item of items) {
        if (item.type.startsWith('image/')) {
          e.preventDefault();
          _insertImageFile(ed, item.getAsFile());
          return;
        }
      }
    });

    // Deselect images when clicking on text
    ed.addEventListener('click', e => {
      if (e.target.tagName !== 'IMG') {
        ed.querySelectorAll('img.selected').forEach(i => i.classList.remove('selected'));
      }
    });
  }

  // ── Placeholder ───────────────────────────────────────────────────────────
  function _bindPlaceholder(ed) {
    const updatePlaceholder = () => {
      const empty = !ed.textContent.trim() && !ed.querySelector('img');
      ed.dataset.empty = empty ? 'true' : 'false';
    };
    ed.addEventListener('input', updatePlaceholder);
    ed.addEventListener('focus', updatePlaceholder);
    ed.addEventListener('blur',  updatePlaceholder);
    updatePlaceholder();
  }

  return { init };
})();
