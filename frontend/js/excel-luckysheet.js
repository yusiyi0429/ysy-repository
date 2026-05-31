/**
 * Luckysheet 在线 Excel 编辑器封装
 * 支持：增删行列、合并/取消合并单元格、多工作表、中文工具栏
 */
(function (global) {
  'use strict';

  /** 必须为 luckysheet（官方硬编码），不可用其他 id */
  const CONTAINER_ID = 'luckysheet';
  const LUCKYSHEET_CONTAINER_STYLE =
    'margin:0;padding:0;position:absolute;width:100%;height:100%;left:0;top:0;';
  let _mounted = false;

  /** 与 Luckysheet 官方 demo 一致，不要单独加载 jquery.min.js（会冲突） */
  const LUCKYSHEET_SCRIPTS = [
    'vendor/luckysheet/plugins/js/plugin.js',
    'vendor/luckysheet/luckysheet.umd.js',
  ];

  function assetUrl(relativePath) {
    const base = (global.location && global.location.origin) ? global.location.origin : '';
    return base + '/' + relativePath.replace(/^\//, '');
  }

  function isLuckysheetReady() {
    return typeof global.luckysheet !== 'undefined' && typeof global.luckysheet.create === 'function';
  }

  function scriptTagExists(src) {
    return Array.from(document.querySelectorAll('script[src]')).some(el => {
      const s = el.getAttribute('src') || '';
      return s === src || s.endsWith(src.split('/').slice(-2).join('/'));
    });
  }

  function loadScriptTag(src) {
    if (scriptTagExists(src)) {
      return Promise.resolve();
    }
    return new Promise((resolve, reject) => {
      const el = document.createElement('script');
      el.src = src;
      el.async = false;
      el.onload = () => resolve();
      el.onerror = () => reject(new Error('无法加载: ' + src));
      document.body.appendChild(el);
    });
  }

  let _loadPromise = null;

  async function probeAsset(relativePath) {
    try {
      const r = await fetch(assetUrl(relativePath), { method: 'HEAD', cache: 'no-store' });
      return r.ok;
    } catch (_) {
      return false;
    }
  }

  function setLoadingMessage(containerEl, msg) {
    if (!containerEl) return;
    containerEl.innerHTML = '<div class="excel-editor-loading"><div class="excel-loading-spinner"></div><div>' + msg + '</div></div>';
  }

  async function ensureLuckysheetLoaded(statusEl) {
    if (isLuckysheetReady()) {
      return { ok: true };
    }
    if (_loadPromise) {
      return _loadPromise;
    }

    _loadPromise = (async () => {
      const missing = [];
      for (const p of LUCKYSHEET_SCRIPTS) {
        if (statusEl) setLoadingMessage(statusEl, '检查组件资源…');
        if (!(await probeAsset(p))) {
          missing.push(p);
        }
      }
      if (missing.length) {
        _loadPromise = null;
        return {
          ok: false,
          missing,
          hint: '静态资源 404：请重启后端（需含 /vendor/ 路由），并确认 frontend/vendor 目录存在。',
        };
      }

      try {
        if (statusEl) setLoadingMessage(statusEl, '加载表格组件 (1/2)…');
        await loadScriptTag(assetUrl(LUCKYSHEET_SCRIPTS[0]));
        if (statusEl) setLoadingMessage(statusEl, '加载表格组件 (2/2)…');
        await loadScriptTag(assetUrl(LUCKYSHEET_SCRIPTS[1]));
      } catch (e) {
        _loadPromise = null;
        return { ok: false, error: e.message, hint: '脚本执行失败，请查看 F12 Console。' };
      }

      if (!isLuckysheetReady()) {
        _loadPromise = null;
        return {
          ok: false,
          hint: '组件未就绪，请勿单独加载 jquery.min.js，仅需 plugin.js + luckysheet.umd.js。',
        };
      }
      return { ok: true };
    })();

    const result = await _loadPromise;
    if (!result.ok) {
      _loadPromise = null;
    }
    return result;
  }

  function preloadLuckysheet() {
    if (isLuckysheetReady() || _loadPromise) return;
    _loadPromise = ensureLuckysheetLoaded(null).then(function (r) {
      if (!r.ok) _loadPromise = null;
      return r;
    }).catch(function () {
      _loadPromise = null;
    });
  }

  function renderLoadError(containerEl, status) {
    let html = '<div class="excel-editor-loading"><strong>表格编辑器未加载</strong><ul style="text-align:left;margin:12px auto;max-width:520px;">';
    if (status.missing && status.missing.length) {
      html += '<li>以下资源不可访问（应为 HTTP 200）：<br><code>' + status.missing.join('</code><br><code>') + '</code></li>';
    }
    if (status.error) {
      html += '<li>' + status.error + '</li>';
    }
    if (status.hint) {
      html += '<li>' + status.hint + '</li>';
    }
    html += '<li>处理：① 停止旧进程后重新运行 <code>python backend/app_server.py</code>；② Ctrl+F5 强刷；③ 仍失败则执行 <code>node scripts/copy-frontend-vendor.js</code></li>';
    html += '</ul></div>';
    containerEl.innerHTML = html;
  }

  function normalizeSheetPayload(val) {
    if (Array.isArray(val)) {
      return { rows: val.map(r => (Array.isArray(r) ? r.map(c => (c == null ? '' : c)) : [])), merges: [] };
    }
    if (val && typeof val === 'object') {
      return {
        rows: (val.rows || []).map(r => (Array.isArray(r) ? r.map(c => (c == null ? '' : c)) : [])),
        merges: val.merges || [],
      };
    }
    return { rows: [], merges: [] };
  }

  function normalizeSheetsFromApi(sheets) {
    const out = {};
    if (!sheets || typeof sheets !== 'object') return out;
    for (const name of Object.keys(sheets)) {
      out[name] = normalizeSheetPayload(sheets[name]);
    }
    return out;
  }

  function trimRows(rows) {
    if (!rows.length) return [[]];
    let lastR = rows.length - 1;
    while (lastR > 0) {
      const row = rows[lastR];
      if (row && row.some(c => String(c ?? '').trim() !== '')) break;
      lastR--;
    }
    const trimmed = rows.slice(0, lastR + 1).map(row => {
      if (!row) return [];
      let lastC = row.length - 1;
      while (lastC >= 0 && String(row[lastC] ?? '').trim() === '') lastC--;
      return row.slice(0, lastC + 1);
    });
    const maxCols = Math.max(1, ...trimmed.map(r => r.length));
    return trimmed.map(r => {
      const copy = r.slice();
      while (copy.length < maxCols) copy.push('');
      return copy;
    });
  }

  function makeCell(raw) {
    const s = raw === null || raw === undefined ? '' : String(raw);
    return { v: s, m: s, ct: { fa: 'General', t: 'g' } };
  }

  function upsertCelldata(celldata, r, c, vObj) {
    const found = celldata.find(cell => cell.r === r && cell.c === c);
    if (found) {
      found.v = Object.assign({}, found.v, vObj);
      return found;
    }
    const item = { r, c, v: vObj };
    celldata.push(item);
    return item;
  }

  function rowsToLuckysheetSheet(name, payload, index) {
    const { rows, merges } = normalizeSheetPayload(payload);
    const trimmed = trimRows(rows);
    const celldata = [];
    let maxR = Math.max(0, trimmed.length - 1);
    let maxC = 0;

    trimmed.forEach((row, r) => {
      row.forEach((val, c) => {
        maxR = Math.max(maxR, r);
        maxC = Math.max(maxC, c);
        const s = val === null || val === undefined ? '' : String(val);
        if (s.trim() === '') return;
        celldata.push({ r, c, v: makeCell(val) });
      });
    });

    const mergeConfig = {};
    (merges || []).forEach(m => {
      if (!m || m.rs < 1 || m.cs < 1) return;
      const r = Number(m.r) || 0;
      const c = Number(m.c) || 0;
      const rs = Number(m.rs) || 1;
      const cs = Number(m.cs) || 1;
      mergeConfig[r + '_' + c] = { r, c, rs, cs };
      maxR = Math.max(maxR, r + rs - 1);
      maxC = Math.max(maxC, c + cs - 1);

      const anchorText = trimmed[r] && trimmed[r][c] !== undefined ? String(trimmed[r][c] ?? '') : '';
      const anchor = upsertCelldata(celldata, r, c, makeCell(anchorText));
      anchor.v.mc = { r, c, rs, cs };

      for (let dr = r; dr < r + rs; dr++) {
        for (let dc = c; dc < c + cs; dc++) {
          if (dr === r && dc === c) continue;
          upsertCelldata(celldata, dr, dc, { mc: { r, c } });
        }
      }
    });

    const rowCount = Math.min(Math.max(maxR + 6, 12), 150);
    const colCount = Math.min(Math.max(maxC + 3, 10), 36);
    const idx = String(index);

    return {
      name: name || 'Sheet1',
      color: '',
      index: idx,
      status: index === 0 ? 1 : 0,
      order: idx,
      celldata: celldata,
      row: rowCount,
      column: colCount,
      defaultRowHeight: 22,
      defaultColWidth: 100,
      scrollLeft: 0,
      scrollTop: 0,
      luckysheet_select_save: [],
      calcChain: [],
      isPivotTable: false,
      pivotTable: {},
      filter_select: null,
      luckysheet_conditionformat_save: [],
      luckysheet_alternateformat_save: [],
      frozen: {},
      chart: [],
      zoomRatio: 1,
      image: [],
      showGridLines: 1,
      dataVerification: {},
      hyperlink: {},
      config: {
        merge: mergeConfig,
        rowlen: {},
        columnlen: {},
        rowhidden: {},
        colhidden: {},
        borderInfo: [],
      },
    };
  }

  function cellValue(cell) {
    if (cell == null) return '';
    if (typeof cell === 'string' || typeof cell === 'number') return String(cell);
    if (cell.v != null && cell.v !== '') return String(cell.v);
    if (cell.m != null) return String(cell.m);
    return '';
  }

  function luckysheetSheetToPayload(sheet) {
    const merges = [];
    const mergeObj = (sheet.config && sheet.config.merge) || {};
    Object.keys(mergeObj).forEach(k => {
      const m = mergeObj[k];
      if (m && m.rs >= 1 && m.cs >= 1) {
        merges.push({ r: m.r, c: m.c, rs: m.rs, cs: m.cs });
      }
    });

    const rows = [];
    if (sheet.data && sheet.data.length) {
      for (let r = 0; r < sheet.data.length; r++) {
        const srcRow = sheet.data[r] || [];
        const out = [];
        for (let c = 0; c < srcRow.length; c++) {
          out.push(cellValue(srcRow[c]));
        }
        rows.push(out);
      }
    } else if (sheet.celldata && sheet.celldata.length) {
      sheet.celldata.forEach(item => {
        const r = item.r;
        const c = item.c;
        while (rows.length <= r) rows.push([]);
        while (rows[r].length <= c) rows[r].push('');
        rows[r][c] = cellValue(item.v);
      });
    }

    return { rows: trimRows(rows), merges };
  }

  function destroy() {
    if (typeof global.luckysheet !== 'undefined' && global.luckysheet.destroy) {
      try {
        global.luckysheet.destroy();
      } catch (_) { /* ignore */ }
    }
    _mounted = false;
    const root = document.getElementById(CONTAINER_ID);
    if (root) {
      root.innerHTML = '';
      root.removeAttribute('style');
    }
    document.querySelectorAll('.luckysheet-wa-editor').forEach(function (node) {
      if (node.parentNode) node.parentNode.removeChild(node);
    });
  }

  function waitForContainerReady(el, maxFrames) {
    maxFrames = maxFrames || 80;
    return new Promise(function (resolve) {
      var frames = 0;
      function tick() {
        frames++;
        var rect = el.getBoundingClientRect();
        var modal = document.getElementById('excel-editor-modal');
        var visible = modal && modal.classList.contains('active');
        if (visible && rect.width >= 80 && rect.height >= 120) {
          resolve();
        } else if (frames >= maxFrames) {
          resolve();
        } else {
          requestAnimationFrame(tick);
        }
      }
      tick();
    });
  }

  function createLuckysheetInstance(sheetHost, containerEl, luckysheetData, names, host) {
    return waitForContainerReady(sheetHost).then(function () {
      global.luckysheet.create({
        container: CONTAINER_ID,
        lang: 'zh',
        data: luckysheetData,
        showinfobar: false,
        showsheetbar: names.length > 1,
        showstatisticBar: true,
        showtoolbar: true,
        allowUpdate: false,
        updateUrl: '',
        loadUrl: '',
        // 使用默认工具栏，避免传入不兼容键导致 resize.js 工具栏定位报错
        sheetFormulaBar: true,
        enableAddRow: true,
        enableAddBackTop: true,
        allowEdit: true,
        forceCalculation: false,
        hook: {
          workbookCreateAfter: function () {
            setTimeout(function () {
              try {
                if (global.luckysheet && global.luckysheet.resize) {
                  global.luckysheet.resize();
                }
                if (global.luckysheet && global.luckysheet.refresh) {
                  global.luckysheet.refresh();
                }
              } catch (_) { /* ignore */ }
            }, 200);
          },
          updated: function () {
            notifyModified();
          },
          sheetActivate: function () {
            notifyModified();
          },
        },
      });
      _mounted = true;
      var panel = host.querySelector('#excel-guide-panel');
      if (panel) panel.hidden = false;
      var toggle = host.querySelector('#excel-guide-toggle');
      if (toggle) {
        toggle.setAttribute('aria-expanded', 'true');
        toggle.textContent = '操作引导 ▴';
      }
    });
  }

  function notifyModified() {
    if (typeof global.markExcelModified === 'function') {
      global.markExcelModified();
    }
  }

  function getSelectionRange() {
    if (typeof global.luckysheet === 'undefined' || !global.luckysheet.getRange) {
      return { r0: 0, r1: 0, c0: 0, c1: 0 };
    }
    const ranges = global.luckysheet.getRange();
    const block = ranges && ranges[0];
    if (!block || !block.row) {
      return { r0: 0, r1: 0, c0: 0, c1: 0 };
    }
    return {
      r0: block.row[0],
      r1: block.row[1],
      c0: block.column[0],
      c1: block.column[1],
    };
  }

  function runAction(action) {
    if (typeof global.luckysheet === 'undefined') {
      return { ok: false, msg: '表格组件未加载。请确认 frontend/vendor 已部署，或执行 node scripts/copy-frontend-vendor.js 后重新构建镜像。' };
    }
    const { r0, r1, c0, c1 } = getSelectionRange();
    const rowSpan = Math.max(1, r1 - r0 + 1);
    const colSpan = Math.max(1, c1 - c0 + 1);
    try {
      switch (action) {
        case 'merge':
          if (rowSpan === 1 && colSpan === 1) {
            return { ok: false, msg: '请先拖动选中多个单元格，再点击「合并单元格」。' };
          }
          global.luckysheet.setRangeMerge('all');
          break;
        case 'unmerge':
          global.luckysheet.setRangeMerge('merge-cancel');
          break;
        case 'insertRow':
          if (global.luckysheet.insertRow) global.luckysheet.insertRow(r0, 1);
          else return { ok: false, msg: '当前版本不支持插入行' };
          break;
        case 'deleteRow':
          if (global.luckysheet.deleteRow) global.luckysheet.deleteRow(r0, rowSpan);
          else return { ok: false, msg: '当前版本不支持删除行' };
          break;
        case 'insertCol':
          if (global.luckysheet.insertColumn) global.luckysheet.insertColumn(c0, 1);
          else return { ok: false, msg: '当前版本不支持插入列' };
          break;
        case 'deleteCol':
          if (global.luckysheet.deleteColumn) global.luckysheet.deleteColumn(c0, colSpan);
          else return { ok: false, msg: '当前版本不支持删除列' };
          break;
        default:
          return { ok: false, msg: '未知操作' };
      }
      notifyModified();
      return { ok: true };
    } catch (e) {
      console.error('ExcelEditor.runAction', action, e);
      return { ok: false, msg: e.message || String(e) };
    }
  }

  function bindGuideBar(wrapEl) {
    const tpl = document.getElementById('excel-editor-guide-template');
    if (!tpl || !wrapEl) return;
    const guide = tpl.content.cloneNode(true);
    wrapEl.insertBefore(guide, wrapEl.firstChild);

    wrapEl.querySelectorAll('[data-excel-action]').forEach(btn => {
      btn.addEventListener('click', () => {
        const res = runAction(btn.getAttribute('data-excel-action'));
        if (res && !res.ok && res.msg) alert(res.msg);
      });
    });

    const toggle = wrapEl.querySelector('#excel-guide-toggle');
    const panel = wrapEl.querySelector('#excel-guide-panel');
    if (toggle && panel) {
      toggle.addEventListener('click', () => {
        const open = panel.hidden;
        panel.hidden = !open;
        toggle.setAttribute('aria-expanded', open ? 'true' : 'false');
        toggle.textContent = open ? '操作引导 ▴' : '操作引导 ▾';
      });
    }
  }

  global.toggleExcelGuide = function () {
    const panel = document.getElementById('excel-guide-panel');
    const toggle = document.getElementById('excel-guide-toggle');
    if (!panel || !toggle) return;
    const open = panel.hidden;
    panel.hidden = !open;
    toggle.setAttribute('aria-expanded', open ? 'true' : 'false');
    toggle.textContent = open ? '操作引导 ▴' : '操作引导 ▾';
  };

  async function mount(containerEl, sheets, activeSheetName) {
    setLoadingMessage(containerEl, '正在加载表格组件…');
    const loadStatus = await ensureLuckysheetLoaded(containerEl);
    if (!loadStatus.ok) {
      renderLoadError(containerEl, loadStatus);
      return false;
    }

    setLoadingMessage(containerEl, '正在渲染表格…');
    destroy();

    const normalized = normalizeSheetsFromApi(sheets);
    const names = Object.keys(normalized);
    if (!names.length) {
      containerEl.innerHTML = '<div class="excel-editor-loading">无工作表数据</div>';
      return false;
    }

    const luckysheetData = names.map((name, i) =>
      rowsToLuckysheetSheet(name, normalized[name], i)
    );

    let activeIndex = names.indexOf(activeSheetName);
    if (activeIndex < 0) activeIndex = 0;
    luckysheetData.forEach((s, i) => {
      s.status = i === activeIndex ? 1 : 0;
    });

    containerEl.innerHTML =
      '<div class="excel-luckysheet-inner">' +
      '<div id="excel-luckysheet-mount-host" class="excel-luckysheet-mount-host"></div>' +
      '</div>';
    const host = containerEl.querySelector('#excel-luckysheet-mount-host');
    bindGuideBar(host);
    const sheetHost = document.createElement('div');
    sheetHost.id = CONTAINER_ID;
    sheetHost.className = 'luckysheet-host';
    sheetHost.setAttribute('style', LUCKYSHEET_CONTAINER_STYLE);
    host.appendChild(sheetHost);

    try {
      await createLuckysheetInstance(sheetHost, containerEl, luckysheetData, names, host);
    } catch (e) {
      console.error('Luckysheet init failed:', e);
      containerEl.innerHTML = '<div class="excel-editor-loading">表格编辑器初始化失败：' + (e.message || e) + '</div>';
      return false;
    }

    return true;
  }

  function exportAllSheets() {
    if (typeof global.luckysheet === 'undefined' || !global.luckysheet.getAllSheets) {
      return null;
    }
    try {
      const all = global.luckysheet.getAllSheets();
      const out = {};
      all.forEach(sheet => {
        const name = sheet.name || 'Sheet1';
        out[name] = luckysheetSheetToPayload(sheet);
      });
      return out;
    } catch (e) {
      console.error('exportAllSheets failed:', e);
      return null;
    }
  }

  function syncBeforeSave(fallbackSheets) {
    const exported = exportAllSheets();
    return exported && Object.keys(exported).length ? exported : normalizeSheetsFromApi(fallbackSheets);
  }

  global.ExcelEditor = {
    normalizeSheetsFromApi,
    normalizeSheetPayload,
    mount,
    ensureLuckysheetLoaded,
    preloadLuckysheet,
    isLuckysheetReady,
    destroy,
    exportAllSheets,
    syncBeforeSave,
    runAction,
  };

  if (global.document) {
    global.document.addEventListener('DOMContentLoaded', function () {
      setTimeout(preloadLuckysheet, 2000);
    });
  }
})(typeof window !== 'undefined' ? window : this);
