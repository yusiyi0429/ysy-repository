/**
 * 公共工具函数
 * escapeHtml / showToast / apiCall / apiCallJSON / copyTextToClipboard / 按钮锁
 */
(function (global) {
  'use strict';

  var API_BASE = global.location.origin;

  // ── HTML 转义 ──
  function escapeHtml(str) {
    if (str == null) return '';
    return String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }

  // ── Toast 提示 ──
  function showToast(msg, type) {
    type = type || 'success';
    var existing = document.querySelector('.s-toast');
    if (existing) existing.remove();
    var toast = document.createElement('div');
    toast.className = 's-toast' + (type === 'error' ? ' s-toast-error' : '');
    toast.textContent = msg;
    document.body.appendChild(toast);
    setTimeout(function () { toast.classList.add('s-toast-show'); }, 10);
    setTimeout(function () {
      toast.classList.remove('s-toast-show');
      setTimeout(function () { toast.remove(); }, 300);
    }, 2000);
  }

  // ── 复制到剪贴板 ──
  async function copyTextToClipboard(text) {
    var value = text == null ? '' : String(text);
    if (!value) return false;
    try {
      if (navigator.clipboard && window.isSecureContext) {
        await navigator.clipboard.writeText(value);
        return true;
      }
    } catch (_) { /* fallback */ }
    var ta = document.createElement('textarea');
    ta.value = value;
    ta.setAttribute('readonly', '');
    ta.style.position = 'fixed';
    ta.style.left = '-9999px';
    document.body.appendChild(ta);
    ta.select();
    var ok = false;
    try { ok = document.execCommand('copy'); } catch (_) { ok = false; }
    document.body.removeChild(ta);
    return ok;
  }

  // ── 质量条 ──
  function qualityBar(pct) {
    var cls = pct >= 80 ? 'green' : pct >= 60 ? 'orange' : 'red';
    return '<div class="quality-bar"><div class="quality-fill ' + cls + '" style="width:' + pct + '%"></div></div>';
  }

  // ── 统计行 ──
  function statRow(label, value, cls) {
    return '<div class="stat-row"><span class="stat-label">' + label + '</span><span class="stat-value ' + (cls || '') + '">' + value + '</span></div>';
  }

  // ── API 调用（FormData，带超时） ──
  async function apiCall(endpoint, formData, timeoutMs) {
    timeoutMs = timeoutMs || 120000;
    var controller = new AbortController();
    var timer = setTimeout(function () { controller.abort(); }, timeoutMs);
    try {
      var resp = await fetch(API_BASE + endpoint, { method: 'POST', body: formData, signal: controller.signal });
      var text = await resp.text();
      try { return JSON.parse(text); } catch (_) { return { raw: text }; }
    } catch (e) {
      if (e.name === 'AbortError') {
        return { status: 'error', error: '请求超时，请检查网络后重试' };
      }
      throw e;
    } finally {
      clearTimeout(timer);
    }
  }

  // ── API 调用（JSON，带可选超时） ──
  async function apiCallJSON(endpoint, body, method, timeoutMs) {
    method = method || 'POST';
    timeoutMs = timeoutMs || 0;
    var controller = new AbortController();
    var timer = null;
    if (timeoutMs > 0) {
      timer = setTimeout(function () { controller.abort(); }, timeoutMs);
    }
    try {
      var resp = await fetch(API_BASE + endpoint, {
        method: method,
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(body),
        signal: controller.signal,
      });
      var text = await resp.text();
      try { return JSON.parse(text); } catch (_) { return { raw: text }; }
    } catch (e) {
      if (e.name === 'AbortError') {
        return { status: 'error', error: '请求超时，请检查网络后重试' };
      }
      throw e;
    } finally {
      if (timer) clearTimeout(timer);
    }
  }

  // ── 按钮防重复点击锁 ──
  function withButtonLock(btnEl, asyncFn) {
    if (!btnEl) return asyncFn();
    if (btnEl._locked) return Promise.resolve();
    btnEl._locked = true;
    btnEl.disabled = true;
    try {
      return asyncFn().finally(function () {
        btnEl.disabled = false;
        btnEl._locked = false;
      });
    } catch (e) {
      btnEl.disabled = false;
      btnEl._locked = false;
      throw e;
    }
  }

  // ── 渲染输出区域 ──
  function renderOutput(containerId, html) {
    var el = document.getElementById(containerId);
    if (el) el.innerHTML = '<div class="output-result">' + html + '</div>';
  }

  function renderLoading(containerId) {
    var el = document.getElementById(containerId);
    if (el) el.innerHTML = '<div class="loading"><div class="spinner"></div>处理中...</div>';
  }

  // ── 暴露到全局 ──
  global.App = global.App || {};
  global.App.escapeHtml = escapeHtml;
  global.App.showToast = showToast;
  global.App.copyTextToClipboard = copyTextToClipboard;
  global.App.qualityBar = qualityBar;
  global.App.statRow = statRow;
  global.App.apiCall = apiCall;
  global.App.apiCallJSON = apiCallJSON;
  global.App.withButtonLock = withButtonLock;
  global.App.renderOutput = renderOutput;
  global.App.renderLoading = renderLoading;

})(window);
