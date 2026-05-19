'use strict';

// ── Constantes ──────────────────────────────────────────────
const COLORS = ['#4285F4','#EA4335','#FBBC05','#34A853','#FF6D00','#9C27B0','#00BCD4','#795548'];
const SETTINGS_KEY = 'sheetGroupsData_v1';

// ── Estado ───────────────────────────────────────────────────
let groups    = [];   // [{ id, name, color, collapsed, sheets: [name,...] }]
let allSheets = [];   // [{ name, position }]
let selColor  = COLORS[0];

// ── Inicialización Office ─────────────────────────────────────
Office.onReady(async info => {
  if (info.host !== Office.HostType.Excel) return;

  buildColorPicker();
  setupListeners();

  await loadGroups();
  await refreshSheets();
  renderAll();
  highlightActive();

  // Eventos del libro
  await Excel.run(async ctx => {
    ctx.workbook.worksheets.onAdded.add(async () => { await refreshSheets(); renderAll(); });
    ctx.workbook.worksheets.onDeleted.add(async () => { await refreshSheets(); saveGroups(); renderAll(); });
    ctx.workbook.worksheets.onActivated.add(highlightActive);
    try {
      ctx.workbook.worksheets.onNameChanged.add(async ev => {
        if (ev.nameBefore && ev.nameAfter) {
          groups.forEach(g => {
            const i = g.sheets.indexOf(ev.nameBefore);
            if (i !== -1) g.sheets[i] = ev.nameAfter;
          });
          await saveGroups();
        }
        await refreshSheets();
        renderAll();
      });
    } catch(_) {}
    await ctx.sync();
  });
});

// ── Persistencia ──────────────────────────────────────────────
async function loadGroups() {
  try {
    await Excel.run(async ctx => {
      const s = ctx.workbook.settings.getItemOrNullObject(SETTINGS_KEY);
      s.load('value');
      await ctx.sync();
      groups = (!s.isNullObject && s.value) ? JSON.parse(s.value) : [];
    });
  } catch(_) { groups = []; }
}

async function saveGroups() {
  try {
    await Excel.run(async ctx => {
      ctx.workbook.settings.add(SETTINGS_KEY, JSON.stringify(groups));
      await ctx.sync();
    });
  } catch(_) {}
}

// ── Hojas ─────────────────────────────────────────────────────
async function refreshSheets() {
  try {
    await Excel.run(async ctx => {
      const ws = ctx.workbook.worksheets;
      ws.load('items/name,items/position');
      await ctx.sync();
      allSheets = ws.items
        .map(s => ({ name: s.name, position: s.position }))
        .sort((a,b) => a.position - b.position);
      // Limpia del estado hojas que ya no existen
      const names = new Set(allSheets.map(s => s.name));
      let changed = false;
      groups.forEach(g => {
        const before = g.sheets.length;
        g.sheets = g.sheets.filter(n => names.has(n));
        if (g.sheets.length !== before) changed = true;
      });
      if (changed) await saveGroups();
    });
  } catch(_) {}
}

async function goToSheet(name) {
  try {
    await Excel.run(async ctx => {
      ctx.workbook.worksheets.getItem(name).activate();
      await ctx.sync();
    });
  } catch(_) {}
}

async function highlightActive() {
  try {
    await Excel.run(async ctx => {
      const ws = ctx.workbook.worksheets.getActiveWorksheet();
      ws.load('name');
      await ctx.sync();
      document.querySelectorAll('.sheet-item').forEach(el =>
        el.classList.toggle('active', el.dataset.name === ws.name)
      );
    });
  } catch(_) {}
}

// ── Operaciones de grupos ─────────────────────────────────────
function createGroup(name, color) {
  groups.push({ id: 'g' + Date.now(), name, color, collapsed: false, sheets: [] });
  saveGroups(); renderAll();
}

function deleteGroup(id) {
  groups = groups.filter(g => g.id !== id);
  saveGroups(); renderAll();
}

function renameGroup(id, name) {
  const g = groups.find(g => g.id === id);
  if (g && name) { g.name = name; saveGroups(); renderAll(); }
}

function toggleCollapse(id) {
  const g = groups.find(g => g.id === id);
  if (g) { g.collapsed = !g.collapsed; saveGroups(); renderAll(); }
}

function addToGroup(sheetName, groupId) {
  groups.forEach(g => { g.sheets = g.sheets.filter(s => s !== sheetName); });
  const g = groups.find(g => g.id === groupId);
  if (g && !g.sheets.includes(sheetName)) g.sheets.push(sheetName);
  saveGroups(); renderAll();
}

function removeFromGroup(sheetName) {
  groups.forEach(g => { g.sheets = g.sheets.filter(s => s !== sheetName); });
  saveGroups(); renderAll();
}

// ── Render principal ──────────────────────────────────────────
function renderAll() {
  renderGroups();
  renderUngrouped();
}

function renderGroups() {
  const c = document.getElementById('groups-container');
  c.innerHTML = '';
  groups.forEach(g => c.appendChild(buildGroupEl(g)));
}

function buildGroupEl(g) {
  const wrap = el('div', 'group');
  wrap.dataset.gid = g.id;

  // Cabecera
  const hdr = el('div', 'group-header');
  hdr.style.borderLeftColor = g.color;

  const chev = el('span','chevron'); chev.textContent = g.collapsed ? '▶' : '▼';
  chev.onclick = () => toggleCollapse(g.id);

  const dot = el('span','color-dot'); dot.style.backgroundColor = g.color;

  const nameEl = el('span','group-name'); nameEl.textContent = g.name;
  nameEl.ondblclick = () => startRename(g.id, nameEl);

  const cnt = el('span','g-count'); cnt.textContent = '(' + g.sheets.length + ')';

  const delBtn = el('button','btn-del-group'); delBtn.textContent = '×'; delBtn.title = 'Eliminar grupo';
  delBtn.onclick = e => {
    e.stopPropagation();
    if (confirm('¿Eliminar el grupo "' + g.name + '"?\n(Las hojas no se borrarán)')) deleteGroup(g.id);
  };

  append(hdr, chev, dot, nameEl, cnt, delBtn);

  // Lista de hojas
  const list = el('div', 'sheets-list' + (g.collapsed ? ' collapsed' : ''));
  const ordered = allSheets.filter(s => g.sheets.includes(s.name));
  ordered.forEach(s => list.appendChild(buildSheetEl(s.name, g.id)));

  // Drop zone (arrastrar hojas al grupo)
  list.addEventListener('dragover', e => { e.preventDefault(); list.classList.add('drag-over'); });
  list.addEventListener('dragleave', () => list.classList.remove('drag-over'));
  list.addEventListener('drop', e => {
    e.preventDefault(); list.classList.remove('drag-over');
    const sn = e.dataTransfer.getData('sheetName');
    if (sn) addToGroup(sn, g.id);
  });

  append(wrap, hdr, list);
  return wrap;
}

function buildSheetEl(name, groupId) {
  const d = el('div','sheet-item');
  d.textContent = name;
  d.dataset.name = name;
  d.draggable = true;
  d.addEventListener('click', () => goToSheet(name));
  d.addEventListener('contextmenu', e => { e.preventDefault(); showCtx(e, name, groupId); });
  d.addEventListener('dragstart', e => e.dataTransfer.setData('sheetName', name));
  return d;
}

function renderUngrouped() {
  const grouped = new Set(groups.flatMap(g => g.sheets));
  const free    = allSheets.filter(s => !grouped.has(s.name));
  const sec     = document.getElementById('ungrouped-section');
  const list    = document.getElementById('ungrouped-list');
  sec.style.display = free.length ? 'block' : 'none';
  list.innerHTML = '';
  free.forEach(s => list.appendChild(buildSheetEl(s.name, null)));
}

// ── Menú contextual ───────────────────────────────────────────
function showCtx(e, sheetName, currentGid) {
  const menu = document.getElementById('ctx-menu');
  const ul   = document.getElementById('ctx-list');
  ul.innerHTML = '';

  const others = groups.filter(g => g.id !== currentGid);
  if (others.length) {
    addCtxLabel(ul, 'Mover a grupo:');
    others.forEach(g => addCtxItem(ul, g.name, g.color, () => { addToGroup(sheetName, g.id); hideCtx(); }));
  }
  if (currentGid) {
    if (others.length) ul.appendChild(el('li','ctx-divider'));
    addCtxItem(ul, 'Quitar del grupo', null, () => { removeFromGroup(sheetName); hideCtx(); }, 'ctx-danger');
  }
  if (!others.length && !currentGid) addCtxLabel(ul, 'Crea un grupo primero');

  menu.style.left = Math.min(e.clientX, window.innerWidth - 185) + 'px';
  menu.style.top  = Math.min(e.clientY, window.innerHeight - 120) + 'px';
  menu.classList.remove('hidden');
}

function addCtxLabel(ul, txt) {
  const li = el('li','ctx-label'); li.textContent = txt; ul.appendChild(li);
}
function addCtxItem(ul, text, color, cb, extraClass) {
  const li = el('li', 'ctx-item' + (extraClass ? ' ' + extraClass : ''));
  if (color) { const d = el('span','color-dot'); d.style.backgroundColor = color; li.appendChild(d); }
  li.appendChild(document.createTextNode(text));
  li.addEventListener('click', cb);
  ul.appendChild(li);
}
function hideCtx() { document.getElementById('ctx-menu').classList.add('hidden'); }

// ── Renombrar grupo inline ────────────────────────────────────
function startRename(id, nameEl) {
  const orig = nameEl.textContent;
  const inp  = el('input','rename-input');
  inp.value  = orig;
  nameEl.replaceWith(inp);
  inp.focus(); inp.select();
  const done = () => renameGroup(id, inp.value.trim() || orig);
  inp.addEventListener('blur', done);
  inp.addEventListener('keydown', e => {
    if (e.key === 'Enter') done();
    if (e.key === 'Escape') { inp.value = orig; done(); }
  });
}

// ── Modal nuevo grupo ─────────────────────────────────────────
function showModal() {
  selColor = COLORS[0];
  document.getElementById('inp-name').value = '';
  document.querySelectorAll('.swatch').forEach((s,i) => s.classList.toggle('selected', i === 0));
  document.getElementById('modal-overlay').classList.remove('hidden');
  setTimeout(() => document.getElementById('inp-name').focus(), 50);
}
function hideModal() { document.getElementById('modal-overlay').classList.add('hidden'); }

function buildColorPicker() {
  const cp = document.getElementById('color-picker');
  COLORS.forEach((c,i) => {
    const s = el('div', 'swatch' + (i === 0 ? ' selected' : ''));
    s.style.backgroundColor = c;
    s.addEventListener('click', () => {
      selColor = c;
      document.querySelectorAll('.swatch').forEach(x => x.classList.remove('selected'));
      s.classList.add('selected');
    });
    cp.appendChild(s);
  });
}

function setupListeners() {
  document.getElementById('btn-new-group').onclick = showModal;
  document.getElementById('btn-cancel').onclick    = hideModal;
  document.getElementById('btn-refresh').onclick   = async () => {
    await refreshSheets(); renderAll(); highlightActive();
  };
  document.getElementById('btn-save').onclick = () => {
    const name = document.getElementById('inp-name').value.trim();
    if (!name) { document.getElementById('inp-name').focus(); return; }
    createGroup(name, selColor);
    hideModal();
  };
  document.getElementById('inp-name').addEventListener('keydown', e => {
    if (e.key === 'Enter')  document.getElementById('btn-save').click();
    if (e.key === 'Escape') hideModal();
  });
  document.getElementById('modal-overlay').addEventListener('click', e => {
    if (e.target === document.getElementById('modal-overlay')) hideModal();
  });
  document.addEventListener('click', e => {
    if (!e.target.closest('#ctx-menu')) hideCtx();
  });
}

// ── Utilidades DOM ────────────────────────────────────────────
function el(tag, cls) {
  const e = document.createElement(tag);
  if (cls) e.className = cls;
  return e;
}
function append(parent, ...children) { children.forEach(c => parent.appendChild(c)); }
