/* ===== State ===== */
let currentStep = 0;
const API_BASE = window.location.origin;
let allModels = [];
let currentPipeline = null; // { id, name, scenario, domain, current_step, step_status, step_data }
const MAX_STEP = 4;
const MAX_FORM_STEP = 3;
const STEP_NAMES = { 1: "场景锚定", 2: "知识萃取", 3: "知识对齐", 4: "智能转化" };
let _formSaveTimer = null;
let _lastStep2ExtractedText = '';
let _step2InputMode = 'doc'; // 'doc' | 'case'
let _step2ActiveSkill = 'knowledge-extraction'; // 当前选中的 Skill
let _alignTacitAnnotations = {}; // { noteId: { question, answer } } — Step3 隐性注释缓存

/* ===== Step2 Skill 卡片选择 ===== */
function selectStep2Skill(skillId) {
  _step2ActiveSkill = skillId;
  // 更新卡片选中态
  document.querySelectorAll('.s2-skill-card').forEach(function (c) {
    c.classList.toggle('active', c.dataset.skill === skillId);
  });
  // 同步隐藏下拉框
  var sel = document.getElementById('s2-skill-select');
  if (sel) sel.value = skillId;

  // 显示/隐藏条件元素
  var isExtraction = (skillId === 'knowledge-extraction');
  var isPattern = (skillId === 'knowledge-pattern-mining');
  var isGap = (skillId === 'knowledge-gap-analysis');

  // 萃取风格（仅知识萃取）
  var styleGroup = document.getElementById('s2-group-style');
  if (styleGroup) styleGroup.classList.toggle('hidden', !isExtraction);

  // 模式发现多文件
  var patternFiles = document.getElementById('s2-group-pattern-files');
  var patternText = document.getElementById('s2-group-pattern-text');
  if (patternFiles) patternFiles.classList.toggle('hidden', !isPattern);
  if (patternText) patternText.classList.toggle('hidden', !isPattern);

  // 盲区检测说明
  var gapInfo = document.getElementById('s2-group-gap-info');
  if (gapInfo) gapInfo.classList.toggle('hidden', !isGap);

  // 模式切换 Tab（仅知识萃取时显示）
  var modeTabs = document.getElementById('s2-mode-tabs');
  if (modeTabs) modeTabs.classList.toggle('hidden', !isExtraction);

  // 文档输入面板（非知识萃取时隐藏文档/案例面板，使用skill自有输入）
  var docPanel = document.getElementById('s2-panel-doc');
  var casePanel = document.getElementById('s2-panel-case');
  if (docPanel) docPanel.classList.toggle('hidden', !isExtraction);
  if (casePanel) casePanel.classList.toggle('hidden', !isExtraction || _step2InputMode !== 'case');

  // 更新按钮文字
  var btnText = document.getElementById('s2-btn-text');
  var labels = {
    'knowledge-extraction': '执行知识萃取',
    'knowledge-pattern-mining': '执行模式发现',
    'knowledge-gap-analysis': '执行盲区检测'
  };
  if (btnText) btnText.textContent = labels[skillId] || '执行';

  updateStep2Readiness();
}

/* ===== 模式发现：文件列表展示 ===== */
function refreshPatternFileList() {
  var input = document.getElementById('s2-pattern-files');
  var list = document.getElementById('s2-pattern-filelist');
  var items = document.getElementById('s2-pattern-fileitems');
  var count = document.getElementById('s2-pattern-filecount');
  var status = document.getElementById('s2-pattern-filestatus');
  if (!input || !list || !items) return;

  var files = input.files || [];
  if (files.length === 0) { list.classList.add('hidden'); return; }
  list.classList.remove('hidden');

  count.textContent = files.length + ' 个文件';
  var html = '';
  for (var i = 0; i < files.length; i++) {
    var size = files[i].size > 1024 ? (files[i].size / 1024).toFixed(1) + ' KB' : files[i].size + ' B';
    html += '<div class="s2-pattern-fileitem"><span class="s2-pattern-fileitem-icon">📄</span><span class="s2-pattern-fileitem-name">' + escapeHtml(files[i].name) + '</span><span class="s2-pattern-fileitem-size">' + size + '</span></div>';
  }
  items.innerHTML = html;

  if (status) {
    if (files.length >= 2) { status.className = 's2-pattern-filelist-status ok'; status.textContent = '✅ 已满足最低要求（≥2个案例）'; }
    else { status.className = 's2-pattern-filelist-status warn'; status.textContent = '⚠️ 至少需要 2 个案例文件，请继续添加'; }
  }
  updateStep2Readiness();
}

function clearPatternFiles() {
  var input = document.getElementById('s2-pattern-files');
  if (input) input.value = '';
  refreshPatternFileList();
}

/** 各步骤产出物字段：保存表单时不得覆盖丢失 */
const PIPELINE_OUTPUT_KEYS = [
  'step1_output_file', 'step1_download_url', 'step1_md_file', 'step1_md_download_url', 'step1_output_format',
  'step2_output_file', 'step2_download_url', 'step2_md_file', 'step2_md_download_url', 'step2_extracted_count',
  'skill_extract_result', 'skill_extract_style',
  'step3_revision_file', 'step3_download_url', 'step3_md_file', 'step3_md_download_url', 'step3_revision_notes', 'step3_revision_style', 'step3_revision_count', 'step3_excel_path',
  'step4_final_file', 'step4_download_url', 'step4_md_file', 'step4_md_download_url', 'step4_final_notes', 'step4_final_style', 'step4_final_count',
  'step5_skill_file', 'step5_download_url',
  'step5_cot_file', 'step5_cot_download_url',
  'step5_qa_file', 'step5_qa_download_url', 'step5_qa_md_file', 'step5_qa_md_download_url',
  'step5_openclaw_manifest_file', 'step5_openclaw_manifest_url',
];

const DOWNSTREAM_OUTPUT_KEYS = [
  'step2_output_file', 'step2_download_url', 'step2_md_file', 'step2_md_download_url', 'step2_extracted_count', 'skill_extract_result', 'skill_extract_style',
  'step3_revision_file', 'step3_download_url', 'step3_md_file', 'step3_md_download_url', 'step3_revision_notes', 'step3_revision_style', 'step3_revision_count', 'step3_excel_path',
  'step4_final_file', 'step4_download_url', 'step4_md_file', 'step4_md_download_url', 'step4_final_notes', 'step4_final_style', 'step4_final_count',
  'step5_skill_file', 'step5_download_url',
  'step5_cot_file', 'step5_cot_download_url',
  'step5_qa_file', 'step5_qa_download_url', 'step5_qa_md_file', 'step5_qa_md_download_url',
  'step5_openclaw_manifest_file', 'step5_openclaw_manifest_url',
];

function mergeStepDataPreserveOutputs(serverData, localData, options) {
  const preferServer = options?.preferServer === true;
  const server = serverData || {};
  const local = localData || {};
  const merged = preferServer ? { ...local, ...server } : { ...server, ...local };
  for (const key of PIPELINE_OUTPUT_KEYS) {
    const v = preferServer ? (server[key] || local[key]) : (local[key] || server[key]);
    if (v) merged[key] = v;
  }
  return merged;
}

function isStep3RevisionFile(fileName) {
  const n = String(fileName || '').toLowerCase();
  return n.endsWith('.xlsx') && (n.startsWith('revision_') || n.startsWith('edited_step3_'));
}

function isStep4FinalFile(fileName) {
  const n = String(fileName || '').toLowerCase();
  return n.endsWith('.xlsx') && (n.startsWith('final_') || n.startsWith('edited_step4_'));
}

function prefersMarkdownFlow() {
  const sd = currentPipeline?.step_data || {};
  const direct = String(sd.step1_output_format || '').toLowerCase();
  if (direct) return direct === 'markdown';
  const formFmt = String(sd.step1_form_data?.output_format || '').toLowerCase();
  return formFmt === 'markdown';
}

function resolveArtifact(result) {
  const name = result?.download_name || result?.output_file || result?.file_name || '';
  let url = result?.download_url || '';
  if (name && !url) url = '/downloads/' + name;
  if (url && !/^https?:\/\//i.test(url)) url = url.startsWith('/') ? url : ('/downloads/' + url);
  return { name, url };
}

function clearDownstreamOutputs(fromStep) {
  if (!currentPipeline?.step_data) return;
  const start = fromStep <= 1 ? 2 : (fromStep + 1);
  const keysByStep = {
    2: DOWNSTREAM_OUTPUT_KEYS.filter(k => k.startsWith('step2_') || k.startsWith('skill_')),
    3: DOWNSTREAM_OUTPUT_KEYS.filter(k => k.startsWith('step3_')),
    4: DOWNSTREAM_OUTPUT_KEYS.filter(k => k.startsWith('step4_')),
    5: DOWNSTREAM_OUTPUT_KEYS.filter(k => k.startsWith('step5_')),
  };
  for (let s = start; s <= 5; s++) {
    (keysByStep[s] || []).forEach(k => delete currentPipeline.step_data[k]);
  }
  if (fromStep <= 1) _lastStep2ExtractedText = '';
}

function rememberStep1Output(pipelineId, fileName, downloadUrl) {
  if (!pipelineId || !fileName) return;
  try {
    sessionStorage.setItem('step1_output:' + pipelineId, JSON.stringify({
      file_name: fileName,
      download_url: downloadUrl || ('/downloads/' + fileName),
    }));
  } catch (_) { /* ignore */ }
}

function recallStep1Output(pipelineId) {
  if (!pipelineId) return null;
  try {
    const raw = sessionStorage.getItem('step1_output:' + pipelineId);
    return raw ? JSON.parse(raw) : null;
  } catch (_) {
    return null;
  }
}

let _s1DefaultKnowledgeColumns = [];
let _s1RichMarkdownColumns = [];
let _s1KnowledgeColumnSeq = 0;

function step1IsAbstractColumn(name) {
  const s = String(name || '').trim();
  if (!s) return true;
  return /^(列[a-zA-Z0-9]{1,3}|(column|field|字段)\s*\d+|[a-zA-Z]\d?)$/i.test(s);
}

function step1MergeRichColumnsForMarkdown() {
  const rich = (_s1RichMarkdownColumns.length ? _s1RichMarkdownColumns : _s1DefaultKnowledgeColumns).slice();
  const current = step1GetKnowledgeColumns();
  const substantive = current.filter(c => !step1IsAbstractColumn(c));
  const seen = new Set();
  const merged = [];
  substantive.forEach(c => {
    if (!seen.has(c)) { seen.add(c); merged.push(c); }
  });
  rich.forEach(c => {
    if (!seen.has(c)) { seen.add(c); merged.push(c); }
  });
  if (merged.length) step1RenderKnowledgeColumns(merged);
  return merged;
}

function step1OnOutputFormatChange() {
  const fmt = document.getElementById('s1-output-format')?.value || 'excel';
  if (fmt === 'markdown') step1MergeRichColumnsForMarkdown();
  scheduleFormSave(1);
}

function step1AddKnowledgeColumn(name, idx) {
  const container = document.getElementById('s1-knowledge-columns');
  if (!container) return;
  const i = idx != null ? idx : (++_s1KnowledgeColumnSeq);
  const div = document.createElement('div');
  div.className = 's1-k-col-row';
  div.id = 's1-k-col-' + i;
  div.innerHTML =
    '<input type="text" class="s1-k-col-input" placeholder="如：具体方法、判断逻辑" value="' + escapeHtml(name || '') + '">' +
    '<button type="button" class="s1-k-col-remove" onclick="step1RemoveKnowledgeColumn(' + i + ')" title="删除">✕</button>';
  container.appendChild(div);
}

function step1RemoveKnowledgeColumn(idx) {
  const el = document.getElementById('s1-k-col-' + idx);
  if (el) el.remove();
}

function step1GetKnowledgeColumns() {
  const cols = [];
  document.querySelectorAll('.s1-k-col-input').forEach(el => {
    const v = el.value.trim();
    if (v) cols.push(v);
  });
  return cols;
}

function step1RenderKnowledgeColumns(columns) {
  const container = document.getElementById('s1-knowledge-columns');
  if (!container) return;
  container.innerHTML = '';
  _s1KnowledgeColumnSeq = 0;
  const list = (columns && columns.length) ? columns : _s1DefaultKnowledgeColumns;
  if (!list.length) {
    step1AddKnowledgeColumn('具体方法', 1);
    return;
  }
  list.forEach((name, i) => step1AddKnowledgeColumn(name, i + 1));
  _s1KnowledgeColumnSeq = list.length;
}

function step1ResetKnowledgeColumns() {
  const fmt = document.getElementById('s1-output-format')?.value || 'excel';
  if (fmt === 'markdown' && _s1RichMarkdownColumns.length) {
    step1RenderKnowledgeColumns(_s1RichMarkdownColumns);
  } else {
    step1RenderKnowledgeColumns(_s1DefaultKnowledgeColumns);
  }
  scheduleFormSave(1);
}

async function loadStep1SchemaAndTemplates(preferredLegacyTemplate, preferredKnowledgeColumns) {
  const legacySelect = document.getElementById('s1-legacy-template');
  try {
    const resp = await fetch(API_BASE + '/api/step1/templates');
    const data = await resp.json();
    if (data.status === 'ok' && Array.isArray(data.schema?.knowledge_columns)) {
      _s1DefaultKnowledgeColumns = data.schema.knowledge_columns.slice();
    }
    if (data.status === 'ok' && Array.isArray(data.schema?.rich_markdown_columns)) {
      _s1RichMarkdownColumns = data.schema.rich_markdown_columns.slice();
    } else if (_s1DefaultKnowledgeColumns.length) {
      _s1RichMarkdownColumns = _s1DefaultKnowledgeColumns.slice();
    }
    if (legacySelect) {
      legacySelect.innerHTML = '<option value="">不使用（按上方自定义列生成）</option>';
      (data.templates || []).forEach(t => {
        const opt = document.createElement('option');
        opt.value = t.name || '';
        opt.textContent = t.label || t.name || '未命名模板';
        legacySelect.appendChild(opt);
      });
      if (preferredLegacyTemplate && Array.from(legacySelect.options).some(o => o.value === preferredLegacyTemplate)) {
        legacySelect.value = preferredLegacyTemplate;
      }
    }
  } catch (e) {
    console.warn('load step1 schema/templates failed:', e);
  }
  if (preferredKnowledgeColumns && preferredKnowledgeColumns.length) {
    step1RenderKnowledgeColumns(preferredKnowledgeColumns);
  } else if (!document.querySelector('.s1-k-col-input')) {
    step1RenderKnowledgeColumns(_s1DefaultKnowledgeColumns);
  }
  if (document.getElementById('s1-output-format')?.value === 'markdown') {
    step1MergeRichColumnsForMarkdown();
  }
}

/** @deprecated 兼容旧调用 */
async function loadStep1DefaultTemplates(preferredTemplate) {
  return loadStep1SchemaAndTemplates(preferredTemplate, null);
}

async function cacheUploadedFile(step, inputEl, fileNameEl) {
  if (!currentPipeline || !inputEl || !inputEl.files || !inputEl.files[0]) return;
  const file = inputEl.files[0];
  const fd = new FormData();
  fd.append('file', file);
  fd.append('pipeline_id', currentPipeline.id);
  fd.append('step', String(step));
  try {
    const resp = await fetch(API_BASE + '/api/files/cache_upload', { method: 'POST', body: fd });
    const data = await resp.json();
    if (data.status === 'ok') {
      currentPipeline.step_data = currentPipeline.step_data || {};
      currentPipeline.step_data['step' + step + '_cached_file'] = data.file_name;
      currentPipeline.step_data['step' + step + '_cached_name'] = file.name;
      await persistPipeline({
        ['step' + step + '_cached_file']: data.file_name,
        ['step' + step + '_cached_name']: file.name,
      });
      if (fileNameEl) fileNameEl.textContent = file.name + '（已缓存）';
    }
  } catch (e) {
    console.warn('cache upload failed:', e);
  }
}

function refreshCachedUploadLabels(step) {
  if (!currentPipeline?.step_data) return;
  const sd = currentPipeline.step_data;
  if (!step || step === 2) {
    const el = document.getElementById('s2-file-name');
    if (el && !el.textContent && sd.step2_cached_name) el.textContent = sd.step2_cached_name + '（已缓存）';
  }
  if (!step || step === 3) {
    const el = document.getElementById('s3-file-name');
    if (el && !el.textContent && sd.step3_cached_name) el.textContent = sd.step3_cached_name + '（已缓存）';
  }
}

async function refreshCurrentPipeline() {
  if (!currentPipeline?.id) return null;
  try {
    const resp = await fetch(API_BASE + '/api/pipelines/' + currentPipeline.id);
    const data = await resp.json();
    if (data.status === 'ok' && data.pipeline) {
      currentPipeline.current_step = Math.min(MAX_STEP, Math.max(1, data.pipeline.current_step || 1));
      currentPipeline.step_status = data.pipeline.step_status;
      currentPipeline.step_data = mergeStepDataPreserveOutputs(
        data.pipeline.step_data,
        currentPipeline.step_data,
        { preferServer: true }
      );
      if (data.pipeline.scenario) currentPipeline.scenario = data.pipeline.scenario;
      if (data.pipeline.domain) currentPipeline.domain = data.pipeline.domain;
    }
    return currentPipeline;
  } catch (e) {
    console.warn('refreshCurrentPipeline failed:', e);
    return currentPipeline;
  }
}

/* ===== Form Auto-Save & Restore ===== */

function collectAllStepsFormData() {
  const merged = {};
  for (let step = 1; step <= MAX_FORM_STEP; step++) {
    merged['step' + step + '_form_data'] = collectStepFormData(step);
  }
  return merged;
}

function restoreAllStepsFormData() {
  if (!currentPipeline?.step_data) return;
  for (let step = 1; step <= MAX_FORM_STEP; step++) {
    const key = 'step' + step + '_form_data';
    if (currentPipeline.step_data[key]) {
      restoreStepFormData(step, currentPipeline.step_data[key]);
    }
  }
}

async function persistPipeline(extraStepData, extraFields) {
  if (!currentPipeline) return;
  const payload = {
    current_step: currentPipeline.current_step,
    step_status: currentPipeline.step_status,
    step_data: mergeStepDataPreserveOutputs(currentPipeline.step_data, extraStepData || {}),
    ...(extraFields || {}),
  };
  currentPipeline.step_data = payload.step_data;
  const result = await apiCallJSON('/api/pipelines/' + currentPipeline.id, payload, 'PUT');
  if (result.pipeline) {
    currentPipeline.step_data = mergeStepDataPreserveOutputs(
      result.pipeline.step_data,
      currentPipeline.step_data
    );
    if (result.pipeline.current_step != null) {
      currentPipeline.current_step = Math.min(MAX_STEP, Math.max(1, result.pipeline.current_step || 1));
    }
    if (result.pipeline.step_status) currentPipeline.step_status = result.pipeline.step_status;
  }
  return result;
}

function applyStep1GenerateResult(result) {
  if (!currentPipeline) return;
  const excelFile = result.excel_file || result.file_name;
  if (!excelFile) return;
  currentPipeline.step_data = currentPipeline.step_data || {};
  clearDownstreamOutputs(1);
  currentPipeline.step_data.step1_output_file = excelFile;
  currentPipeline.step_data.step1_download_url = result.excel_download_url || ('/downloads/' + excelFile);
  if (result.markdown_file) {
    currentPipeline.step_data.step1_md_file = result.markdown_file;
    currentPipeline.step_data.step1_md_download_url = result.markdown_download_url || ('/downloads/' + result.markdown_file);
  }
  if (result.knowledge_columns) {
    currentPipeline.step_data.step1_knowledge_columns = result.knowledge_columns;
  }
  if (result.output_format) currentPipeline.step_data.step1_output_format = result.output_format;
  if (result.scenario) currentPipeline.scenario = result.scenario;
  rememberStep1Output(
    currentPipeline.id,
    excelFile,
    currentPipeline.step_data.step1_download_url
  );
}

function syncPipelineFromStep1Response(result) {
  if (!currentPipeline || !result?.pipeline) return;
  currentPipeline.step_data = mergeStepDataPreserveOutputs(
    result.pipeline.step_data,
    currentPipeline.step_data
  );
  if (result.pipeline.scenario) currentPipeline.scenario = result.pipeline.scenario;
  if (result.pipeline.domain) currentPipeline.domain = result.pipeline.domain;
}

function collectStepFormData(step) {
  const data = {};
  if (step === 1) {
    data.scenario_name = document.getElementById('s1-scenario-name')?.value || '';
    data.scenario_content = document.getElementById('s1-scenario-content')?.value || '';
    data.sub_scenarios = step1GetSubScenarios();
    data.output_format = document.getElementById('s1-output-format')?.value || 'excel';
    data.knowledge_columns = step1GetKnowledgeColumns();
    data.legacy_template = document.getElementById('s1-legacy-template')?.value || '';
  } else if (step === 2) {
    data.doc_text = document.getElementById('s2-doc-text')?.value || '';
    data.extract_style = document.getElementById('s2-extract-style')?.value || '';
  } else if (step === 3) {
    data.expert_text = document.getElementById('s3-expert-text')?.value || '';
    data.revision_style = document.getElementById('s3-revision-style')?.value || '';
  } else if (step === 4) {
    // Step 4 (智能转化) has no form data to collect
  }
  return data;
}

function restoreStepFormData(step, data) {
  if (!data) return;
  if (step === 1) {
    const nameEl = document.getElementById('s1-scenario-name');
    const contentEl = document.getElementById('s1-scenario-content');
    const outputFmtEl = document.getElementById('s1-output-format');
    const legacyTplEl = document.getElementById('s1-legacy-template');
    if (nameEl && data.scenario_name) nameEl.value = data.scenario_name;
    if (contentEl && data.scenario_content) contentEl.value = data.scenario_content;
    if (outputFmtEl && data.output_format) outputFmtEl.value = data.output_format;
    const leg = data.legacy_template || (data.default_template && data.default_template !== '__schema__' ? data.default_template : '');
    loadStep1SchemaAndTemplates(leg, data.knowledge_columns);
    if (legacyTplEl && leg) {
      const hasOption = Array.from(legacyTplEl.options || []).some(o => o.value === leg);
      if (hasOption) legacyTplEl.value = leg;
    }
    // Clear existing sub-scenarios before restoring
    const subContainer = document.getElementById('s1-sub-scenarios');
    if (subContainer) subContainer.innerHTML = '';
    _s1SubScenarioCount = 0;
    if (data.sub_scenarios && data.sub_scenarios.length) {
      data.sub_scenarios.forEach(sub => {
        step1AddSubScenario();
        const items = document.querySelectorAll('.s1-sub-item');
        const last = items[items.length - 1];
        if (last) {
          last.querySelector('.s1-sub-name').value = sub.name || '';
          last.querySelector('.s1-sub-content').value = sub.content || '';
        }
      });
    }
  } else if (step === 2) {
    const docEl = document.getElementById('s2-doc-text');
    const styleEl = document.getElementById('s2-extract-style');
    if (docEl && data.doc_text) docEl.value = data.doc_text;
    if (styleEl && data.extract_style) styleEl.value = data.extract_style;
  } else if (step === 3) {
    const expertEl = document.getElementById('s3-expert-text');
    const styleEl = document.getElementById('s3-revision-style');
    if (expertEl && data.expert_text) expertEl.value = data.expert_text;
    if (styleEl && data.revision_style) styleEl.value = data.revision_style;
  } else if (step === 4) {
    // Step 4 (智能转化) has no form data to restore
  }
}

function scheduleFormSave(step) {
  if (!currentPipeline) return;
  clearTimeout(_formSaveTimer);
  _autoSaveUI('saving');
  _formSaveTimer = setTimeout(async () => {
    try {
      const allForm = collectAllStepsFormData();
      currentPipeline.step_data = currentPipeline.step_data || {};
      let changed = false;
      for (const [key, formData] of Object.entries(allForm)) {
        if (JSON.stringify(currentPipeline.step_data[key]) !== JSON.stringify(formData)) {
          currentPipeline.step_data[key] = formData;
          changed = true;
        }
      }
      if (!changed) { _autoSaveUI('idle'); return; }
      await persistPipeline(allForm);
      _autoSaveUI('saved');
    } catch (e) {
      console.error('Auto-save failed:', e);
      _autoSaveUI('failed');
    }
  }, 1000);
}

let _autoSaveTimer = null;
function _autoSaveUI(state) {
  var el = document.getElementById('auto-save-indicator');
  var txt = document.getElementById('auto-save-text');
  if (!el || !txt) return;
  clearTimeout(_autoSaveTimer);
  el.className = 'auto-save-indicator ' + state + ' visible';
  if (state === 'saving') { txt.textContent = '保存中...'; }
  else if (state === 'saved') {
    var now = new Date();
    txt.textContent = '已保存 ' + now.getHours().toString().padStart(2,'0') + ':' + now.getMinutes().toString().padStart(2,'0');
    _autoSaveTimer = setTimeout(function () { el.classList.remove('visible'); }, 3000);
  }
  else if (state === 'failed') {
    txt.textContent = '保存失败，点击重试';
    el.onclick = function () { scheduleFormSave(currentStep); };
  }
  else {
    el.classList.remove('visible');
    el.onclick = null;
  }
}

function renderStepReadiness(elId, text, level) {
  const el = document.getElementById(elId);
  if (!el) return;
  const lv = level || 'info';
  el.className = `step-readiness ${lv}`;
  el.textContent = text || '';
}

function step3LooksLikeNoOpinion(text) {
  const t = String(text || '').trim();
  if (!t) return true;
  return /(暂无意见|无意见|无需修订|无需修改|保持不变|确认通过|没有意见|无异议)/.test(t);
}

/* ===== Step2 输入模式切换（文档萃取 / 案例复盘） ===== */
function switchStep2Mode(mode) {
  _step2InputMode = mode;
  var tabs = document.querySelectorAll('.s2-mode-tab');
  tabs.forEach(function (t) { t.classList.toggle('active', t.dataset.mode === mode); });
  var panelDoc = document.getElementById('s2-panel-doc');
  var panelCase = document.getElementById('s2-panel-case');
  if (panelDoc) panelDoc.classList.toggle('hidden', mode !== 'doc');
  if (panelCase) panelCase.classList.toggle('hidden', mode !== 'case');
  updateStep2Readiness();
}

function _buildCaseReviewText() {
  var title = (document.getElementById('s2-case-title') || {}).value || '';
  var ctx = (document.getElementById('s2-case-context') || {}).value || '';
  var dec = (document.getElementById('s2-case-decision') || {}).value || '';
  var out = (document.getElementById('s2-case-outcome') || {}).value || '';
  var redo = (document.getElementById('s2-case-redo') || {}).value || '';
  var habit = (document.getElementById('s2-case-habit') || {}).value || '';
  var parts = [];
  if (title) parts.push('## 案例标题\n' + title);
  if (ctx) parts.push('## 背景 / 情境\n' + ctx);
  if (dec) parts.push('## 当时的判断与行动\n' + dec);
  if (out) parts.push('## 结果\n' + out);
  if (redo) parts.push('## 如果重来\n' + redo);
  if (habit) parts.push('## 养成的习惯 / 条件反射\n' + habit);
  return parts.join('\n\n');
}

function updateStep2Readiness() {
  const btn = document.getElementById('s2-skill-extract');
  if (!btn) return;
  const skillId = _step2ActiveSkill || 'knowledge-extraction';
  const model = resolveModelName('s2-model');
  var hasInput = false, hintMissing = '';

  if (skillId === 'knowledge-gap-analysis') {
    hasInput = !!currentPipeline;
    hintMissing = '请先从总览进入一条流水线';
  } else if (skillId === 'knowledge-pattern-mining') {
    var pf = document.getElementById('s2-pattern-files');
    var pt = document.getElementById('s2-pattern-text');
    hasInput = (pf && pf.files && pf.files.length >= 2) || ((pt && pt.value || '').trim().length >= 50);
    hintMissing = '请上传至少2个案例文件，或粘贴多个案例文本（≥50字）';
  } else {
    if (_step2InputMode === 'case') {
      var ct = (document.getElementById('s2-case-title') || {}).value || '';
      var cc = (document.getElementById('s2-case-context') || {}).value || '';
      hasInput = !!(ct.trim() && cc.trim());
      hintMissing = '请至少填写案例标题和背景情境';
    } else {
      var dt = document.getElementById('s2-doc-text')?.value || '';
      var sf = document.getElementById('s2-source-file')?.files?.[0];
      var cf = currentPipeline?.step_data?.step2_cached_file || '';
      hasInput = !!dt.trim() || !!sf || !!cf;
      hintMissing = '请上传知识来源文件或粘贴文档内容';
    }
  }

  var ready = !!currentPipeline && !!model && hasInput;
  btn.disabled = !ready;
  if (!currentPipeline)  { renderStepReadiness('s2-readiness', '请先从总览进入一条流水线后再执行', 'warn'); return; }
  if (!model)            { renderStepReadiness('s2-readiness', '请先配置并选择模型', 'warn'); return; }
  if (!hasInput)         { renderStepReadiness('s2-readiness', hintMissing, 'warn'); return; }
  var labels = {
    'knowledge-extraction': (_step2InputMode === 'case' ? '已就绪：可执行案例复盘萃取' : '已就绪：可执行知识萃取'),
    'knowledge-pattern-mining': '已就绪：可执行跨案例模式发现',
    'knowledge-gap-analysis': '已就绪：可执行知识盲区检测'
  };
  renderStepReadiness('s2-readiness', labels[skillId] || '已就绪', 'ok');
}

function updateStep3AlignModeHint() {
  const text = document.getElementById('s3-expert-text')?.value || '';
  const hasUpload = !!(document.getElementById('s3-expert-file')?.files?.length);
  const hasCached = !!(currentPipeline?.step_data?.step3_cached_file);
  const hasMaterial = hasUpload || hasCached;
  const btn = document.getElementById('s3-revise-btn');
  if (!btn) return;
  if (!text.trim() && !hasMaterial) {
    btn.innerHTML = '<span class="action-icon">&#10003;</span> 无意见直通生成对齐稿';
    renderStepReadiness('s3-align-hint', '未填写意见：将直接按预萃稿生成对齐稿（无修订）', 'info');
    return;
  }
  if (step3LooksLikeNoOpinion(text) && !hasMaterial) {
    btn.innerHTML = '<span class="action-icon">&#10003;</span> 按当前稿生成对齐稿';
    renderStepReadiness('s3-align-hint', '检测到“无修订”表达：将自动确认当前稿为对齐稿', 'info');
    return;
  }
  btn.innerHTML = '<span class="action-icon">&#9881;</span> 发送并智能修订';
  renderStepReadiness('s3-align-hint', '已检测到专家意见/材料：将按意见生成修订建议', 'ok');
}

function setupFormAutoSave() {
  const bind = (id, step) => {
    const el = document.getElementById(id);
    if (!el) return;
    const evt = el.tagName === 'SELECT' ? 'change' : 'input';
    el.addEventListener(evt, () => scheduleFormSave(step));
  };
  ['s1-scenario-name', 's1-scenario-content', 's1-legacy-template'].forEach(id => bind(id, 1));
  const s1Fmt = document.getElementById('s1-output-format');
  if (s1Fmt) {
    s1Fmt.addEventListener('change', step1OnOutputFormatChange);
    bind('s1-output-format', 1);
  }
  const subBox = document.getElementById('s1-sub-scenarios');
  if (subBox) subBox.addEventListener('input', () => scheduleFormSave(1));
  const kBox = document.getElementById('s1-knowledge-columns');
  if (kBox) kBox.addEventListener('input', () => scheduleFormSave(1));
  ['s2-doc-text', 's2-extract-style'].forEach(id => bind(id, 2));
  ['s3-expert-text', 's3-revision-style'].forEach(id => bind(id, 3));
  const s2Skill = document.getElementById('s2-skill-select');
  const s2Model = document.getElementById('s2-model');
  const s2Source = document.getElementById('s2-source-file');
  const s2Doc = document.getElementById('s2-doc-text');
  const s3Expert = document.getElementById('s3-expert-text');
  const s3File = document.getElementById('s3-expert-file');
  [s2Skill, s2Model].forEach(el => el && el.addEventListener('change', updateStep2Readiness));
  if (s2Source) s2Source.addEventListener('change', updateStep2Readiness);
  if (s2Doc) s2Doc.addEventListener('input', updateStep2Readiness);
  if (s3Expert) s3Expert.addEventListener('input', updateStep3AlignModeHint);
  if (s3File) s3File.addEventListener('change', updateStep3AlignModeHint);
  // 模式发现文件列表
  var pfInput = document.getElementById('s2-pattern-files');
  if (pfInput) pfInput.addEventListener('change', refreshPatternFileList);
}

/* ===== File upload name display ===== */
document.addEventListener('DOMContentLoaded', () => {
  [['s1-template-file', 's1-template-file-name'], ['s2-source-file', 's2-file-name'], ['s3-expert-file', 's3-file-name']].forEach(([inputId, nameId]) => {
    const input = document.getElementById(inputId);
    const nameEl = document.getElementById(nameId);
    if (input && nameEl) {
      input.addEventListener('change', async () => {
        nameEl.textContent = input.files.length ? input.files[0].name : '';
        if (!input.files.length) return;
        if (inputId === 's2-source-file') await cacheUploadedFile(2, input, nameEl);
        if (inputId === 's3-expert-file') await cacheUploadedFile(3, input, nameEl);
        if (inputId === 's2-source-file') updateStep2Readiness();
        if (inputId === 's3-expert-file') updateStep3AlignModeHint();
      });
    }
  });
  loadStep1SchemaAndTemplates();
  setupFormAutoSave();
  loadModels();
  // 初始化列宽拖动调节
  if (App.initResizableColumns) { setTimeout(App.initResizableColumns, 300); }
});

/* ===== Navigation ===== */
document.querySelectorAll('.step-btn').forEach(btn => {
  btn.addEventListener('click', () => {
    if (!currentPipeline) return;
    const targetStep = parseInt(btn.dataset.step);
    const farthestStep = currentPipeline.current_step || 1;
    if (targetStep < farthestStep) {
      if (confirm(`当前在第 ${farthestStep} 步，回退到第 ${targetStep} 步只会重置后续步骤状态，不会删除已生成数据。确定回退？`)) {
        rollbackToStep(targetStep);
      }
    } else {
      switchPanel(targetStep);
    }
  });
});

function updateStepProgress() {
  if (!currentPipeline) return;
  const status = currentPipeline.step_status || {};
  const cur = Math.min(MAX_STEP, Math.max(1, currentPipeline.current_step || 1));
  document.querySelectorAll('.step-btn').forEach(b => {
    const s = parseInt(b.dataset.step);
    b.classList.remove('active', 'done');
    if (s === currentStep) {
      b.classList.add('active');
    } else if (status[s] === 'done') {
      b.classList.add('done');
    }
  });
  document.querySelectorAll('.step-connector').forEach((c, i) => {
    const stepNum = i + 1;
    c.classList.remove('done', 'reached');
    if (status[stepNum] === 'done') c.classList.add('done');
    else if (stepNum < cur) c.classList.add('reached');
  });
}

function renderPipelineProgressSummary(currentPanelStep) {
  if (!currentPipeline || currentPanelStep === 0) {
    document.querySelectorAll('.pipeline-progress-summary').forEach(function (el) { el.remove(); });
    return;
  }
  const status = currentPipeline.step_status || {};
  const steps = [
    { n: 1, label: '场景锚定' },
    { n: 2, label: '知识萃取' },
    { n: 3, label: '知识对齐' },
    { n: 4, label: '智能转化' },
  ];
  var html = '<div class="pipeline-progress-summary">';
  steps.forEach(function (s, i) {
    if (i > 0) html += '<span class="pp-arrow">&gt;</span>';
    var cls = 'pp-step';
    if (s.n === currentPanelStep) cls += ' current';
    else if (status[s.n] === 'done') cls += ' done';
    html += '<span class="' + cls + '">' + s.label + '</span>';
  });
  html += '</div>';

  var activePanel = document.getElementById('panel-' + currentPanelStep);
  if (!activePanel) return;
  var threeCol = activePanel.querySelector('.three-col, .three-col-wide');
  if (!threeCol) return;
  var existing = activePanel.querySelector('.pipeline-progress-summary');
  if (existing) existing.remove();
  threeCol.insertAdjacentHTML('beforebegin', html);
}

function switchPanel(step) {
  if (step > MAX_STEP) step = MAX_STEP;
  if (step < 0) step = 0;
  currentStep = step;
  document.querySelectorAll('.panel').forEach(p => p.classList.remove('active'));
  document.getElementById('panel-' + step).classList.add('active');
  // 切换面板后重新初始化列宽拖动
  if (App.initResizableColumns) { setTimeout(App.initResizableColumns, 100); }

  renderPipelineProgressSummary(step);

  const navSteps = document.getElementById('nav-steps');
  const brandEl = document.getElementById('nav-brand');

  if (step === 0) {
    // Overview mode: hide step nav
    navSteps.classList.remove('visible');
    brandEl.textContent = '隐性知识显性化 · 四步法';
    loadPipelineOverview();
  } else {
    // Pipeline mode: show step nav with progress
    navSteps.classList.add('visible');
    if (currentPipeline) {
      brandEl.textContent = currentPipeline.name;
      updateStepProgress();
      // Restore form data for this step
      const stepDataKey = 'step' + step + '_form_data';
      const savedData = currentPipeline.step_data?.[stepDataKey];
      if (savedData) restoreStepFormData(step, savedData);
    }
  }

  // Refresh model selects when entering a step with AI
  if ([2, 3, 4].includes(step)) {
    if (!allModels.length) loadModels();
    refreshModelSelects();
  }

  // Load skill selects for steps 2, 3
  if ([2, 3].includes(step)) loadStepSkillSelects(step);
  if (step === 1) {
    const fd = currentPipeline?.step_data?.step1_form_data || {};
    loadStep1SchemaAndTemplates(fd.legacy_template || fd.default_template || '', fd.knowledge_columns);
  }

  // Auto-load previous step output（先刷新流水线再检测上一步产出）
  if (step === 2 && currentPipeline) {
    refreshCurrentPipeline().then(() => {
      loadStep2PrevOutput();
      updateStep2Readiness();
    });
  }
  if (step === 3 && currentPipeline) {
    loadStep3PrevOutput();
    loadStep3RevisionContext();
    updateStep3AlignModeHint();
  }
  if (step === 4 && currentPipeline) loadStep5PrevOutput();
  refreshCachedUploadLabels(step);
}

/* ===== Back to Overview ===== */
function goBackToOverview() {
  currentPipeline = null;
  switchPanel(0);
}

async function rollbackToStep(step) {
  if (!currentPipeline) return;
  try {
    const resp = await fetch(API_BASE + '/api/pipelines/' + currentPipeline.id + '/rollback/' + step, { method: 'POST' });
    const data = await resp.json();
    if (data.status === 'ok') {
      currentPipeline = data.pipeline;
      currentPipeline.step_data = currentPipeline.step_data || {};
      switchPanel(step);
      showToast('已回退到第 ' + step + ' 步');
    } else {
      showToast(data.error || '回退失败', 'error');
    }
  } catch (e) {
    showToast('回退失败: ' + e.message, 'error');
  }
}

/* ===== Server Status（委托到 index.html 内联脚本，含 30s 自动重试） ===== */
checkServer = function() { /* no-op: 由 index.html 内联脚本处理 */ };
checkServer();

/* ===== Utility ===== */
function renderOutput(containerId, html) {
  document.getElementById(containerId).innerHTML = '<div class="output-result">' + html + '</div>';
}
function renderLoading(containerId) {
  document.getElementById(containerId).innerHTML = '<div class="loading"><div class="spinner"></div>处理中...</div>';
}
function escapeHtml(str) {
  if (str == null) return '';
  return String(str).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}

async function copyTextToClipboard(text) {
  const value = text == null ? '' : String(text);
  if (!value) return false;
  try {
    if (navigator.clipboard && window.isSecureContext) {
      await navigator.clipboard.writeText(value);
      return true;
    }
  } catch (_) { /* fallback below */ }
  const ta = document.createElement('textarea');
  ta.value = value;
  ta.setAttribute('readonly', '');
  ta.style.position = 'fixed';
  ta.style.left = '-9999px';
  document.body.appendChild(ta);
  ta.select();
  let ok = false;
  try {
    ok = document.execCommand('copy');
  } catch (_) {
    ok = false;
  }
  document.body.removeChild(ta);
  return ok;
}
function qualityBar(pct) {
  const cls = pct >= 80 ? 'green' : pct >= 60 ? 'orange' : 'red';
  return '<div class="quality-bar"><div class="quality-fill ' + cls + '" style="width:' + pct + '%"></div></div>';
}
function statRow(label, value, cls) {
  return '<div class="stat-row"><span class="stat-label">' + label + '</span><span class="stat-value ' + (cls||'') + '">' + value + '</span></div>';
}

async function apiCall(endpoint, formData, timeoutMs = 120000) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);
  try {
    const resp = await fetch(API_BASE + endpoint, { method: 'POST', body: formData, signal: controller.signal });
    const text = await resp.text();
    try { return JSON.parse(text); } catch { return { raw: text }; }
  } catch (e) {
    if (e.name === 'AbortError') {
      return { status: 'error', error: '请求超时，请检查网络后重试' };
    }
    throw e;
  } finally {
    clearTimeout(timer);
  }
}

/** 包裹异步操作：自动禁用/恢复按钮，防止重复点击 */
async function withButtonLock(btnEl, fn) {
  if (!btnEl) return fn();
  if (btnEl._locked) return;
  btnEl._locked = true;
  btnEl.disabled = true;
  const origText = btnEl.innerHTML;
  try {
    return await fn();
  } finally {
    btnEl.disabled = false;
    btnEl.innerHTML = origText;
    btnEl._locked = false;
  }
}

async function apiCallJSON(endpoint, body, method = 'POST', timeoutMs = 0) {
  const controller = new AbortController();
  let timer = null;
  if (timeoutMs > 0) {
    timer = setTimeout(() => controller.abort(), timeoutMs);
  }
  try {
    const resp = await fetch(API_BASE + endpoint, {
      method: method,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body),
      signal: controller.signal
    });
    const text = await resp.text();
    try { return JSON.parse(text); } catch { return { raw: text }; }
  } catch (e) {
    if (e.name === 'AbortError') {
      return { status: 'error', error: '请求超时，请检查 API 地址与网络连通性' };
    }
    throw e;
  } finally {
    if (timer) clearTimeout(timer);
  }
}

async function markStepDone(step) {
  if (!currentPipeline) return;
  const allForm = collectAllStepsFormData();
  currentPipeline.step_data = mergeStepDataPreserveOutputs(currentPipeline.step_data, allForm);

  currentPipeline.step_status[String(step)] = 'done';
  if (step < MAX_STEP) {
    currentPipeline.step_status[String(step + 1)] = currentPipeline.step_status[String(step + 1)] || 'active';
    currentPipeline.current_step = Math.max(currentPipeline.current_step || 1, step + 1);
  } else {
    currentPipeline.current_step = MAX_STEP;
  }
  try {
    await persistPipeline(currentPipeline.step_data, {
      current_step: currentPipeline.current_step,
      step_status: currentPipeline.step_status,
    });
  } catch (e) {
    console.error('Failed to update pipeline:', e);
  }
  updateStepProgress();
}

async function saveCurrentPipeline() {
  if (!currentPipeline) {
    showToast('请先创建或进入一条流水线', 'error');
    return;
  }
  try {
    const allForm = collectAllStepsFormData();
    currentPipeline.step_data = currentPipeline.step_data || {};
    Object.assign(currentPipeline.step_data, allForm);

    const resp = await fetch(API_BASE + '/api/pipelines/' + currentPipeline.id);
    const fresh = await resp.json();
    if (fresh.pipeline) {
      currentPipeline.current_step = Math.min(MAX_STEP, Math.max(1, fresh.pipeline.current_step || 1));
      currentPipeline.step_status = fresh.pipeline.step_status;
      currentPipeline.step_data = mergeStepDataPreserveOutputs(
        fresh.pipeline.step_data,
        currentPipeline.step_data
      );
    }
    await persistPipeline(allForm, {
      current_step: currentPipeline.current_step,
      step_status: currentPipeline.step_status,
    });
    showToast('流水线已保存（含各步骤填写内容）');
  } catch (e) {
    console.error('Save failed:', e);
    showToast('保存失败', 'error');
  }
}

function clearCurrentPipeline() {
  if (!currentPipeline) return;
  if (!confirm('确定要清空当前流水线所有数据吗？此操作不可撤销。')) return;
  fetch(API_BASE + '/api/pipelines/' + currentPipeline.id + '/clear', {
    method: 'POST'
  }).then(r => r.json()).then(r => {
    if (r.status !== 'ok') throw new Error(r.error || '清空失败');
    const pipelineId = currentPipeline.id;
    clearTimeout(_formSaveTimer);

    currentPipeline.current_step = 1;
    currentPipeline.step_status = { '1': 'pending', '2': 'pending', '3': 'pending', '4': 'pending' };
    currentPipeline.step_data = {};

    // Remove per-pipeline browser cache to avoid stale UI restoration.
    sessionStorage.removeItem('step1_output:' + pipelineId);

    // Reset all step form inputs.
    const resetValue = (id, value = '') => {
      const el = document.getElementById(id);
      if (el) el.value = value;
    };
    const resetSelect = (id, preferred = '') => {
      const el = document.getElementById(id);
      if (!el) return;
      if (preferred && Array.from(el.options || []).some(o => o.value === preferred)) {
        el.value = preferred;
      } else {
        el.selectedIndex = 0;
      }
    };
    const resetText = (id, text = '') => {
      const el = document.getElementById(id);
      if (el) el.textContent = text;
    };

    resetValue('s1-scenario-name', '');
    resetValue('s1-scenario-content', '');
    resetValue('s2-doc-text', '');
    resetValue('s3-expert-text', '');

    resetValue('s1-template-file', '');
    resetValue('s2-source-file', '');
    resetValue('s3-expert-file', '');

    resetSelect('s2-extract-style', '标准萃取');
    resetSelect('s3-revision-style', '标准修订');
    resetSelect('s1-output-format', 'excel');
    resetSelect('s1-legacy-template', '');
    step1RenderKnowledgeColumns([]);
    resetSelect('s2-model', '');
    resetSelect('s3-model', '');
    resetSelect('s2-skill-select', '');
    resetSelect('s3-skill-select', '');

    resetText('s1-template-file-name', '未选择');
    resetText('s2-file-name', '');
    resetText('s3-file-name', '');

    updateStepProgress();

    // Clear output display areas
    document.querySelectorAll(
      '#s1-output, #s2-output, #s3-output, #s4-output'
    ).forEach(el => {
      if (el) el.innerHTML = '';
    });
    const s2Prev = document.getElementById('s2-prev-output-area');
    if (s2Prev) s2Prev.innerHTML = '<div class="s2-prev-empty">尚未检测到上一步输出，请先完成场景锚定</div>';
    const s3Info = document.getElementById('s3-prev-info');
    const s3Tags = document.getElementById('s3-prev-tags');
    if (s3Info) s3Info.innerHTML = '';
    if (s3Tags) s3Tags.innerHTML = '';

    const s4Info = document.getElementById('s4-info');
    if (s4Info) s4Info.innerHTML = '';
    const s5Prev = document.getElementById('s5-prev-draft');
    if (s5Prev) s5Prev.innerHTML = '';
    _lastStep2ExtractedText = '';
    closeExcelEditor();
    // Reset sub-scenarios
    const subList = document.getElementById('s1-sub-scenarios');
    if (subList) subList.innerHTML = '';
    _s1SubScenarioCount = 0;

    // Reset step3/step4 preview card visibility states.
    const s3Draft = document.getElementById('s3-prev-draft');
    const s3Empty = document.getElementById('s3-prev-empty');
    if (s3Draft) s3Draft.style.display = 'none';
    if (s3Empty) s3Empty.style.display = '';

    const s4Box = document.getElementById('s4-box');
    const s4Empty = document.getElementById('s4-empty');
    const s4Edit = document.getElementById('s4-edit');
    if (s4Box) s4Box.style.display = 'none';
    if (s4Empty) s4Empty.style.display = 'flex';
    if (s4Edit) s4Edit.style.display = 'none';

    loadStep1SchemaAndTemplates();
    showToast('流水线已清空');
    // Reload current panel
    switchPanel(currentPipeline.current_step);
  }).catch(e => {
    console.error('Clear failed:', e);
    showToast('清空失败', 'error');
  });
}

function showToast(msg, type = 'ok') {
  const existing = document.querySelector('.s-toast');
  if (existing) existing.remove();
  const toast = document.createElement('div');
  toast.className = 's-toast' + (type === 'error' ? ' s-toast-error' : '');
  toast.textContent = msg;
  document.body.appendChild(toast);
  setTimeout(() => toast.classList.add('s-toast-show'), 10);
  setTimeout(() => {
    toast.classList.remove('s-toast-show');
    setTimeout(() => toast.remove(), 300);
  }, 2000);
}

// ═══════════════════════════════════════════════════════════════════
// LLM Model Management
// ═══════════════════════════════════════════════════════════════════

function openModelPanel() {
  closeSkillPanel();
  document.getElementById('model-panel').classList.add('open');
  document.getElementById('model-overlay').classList.remove('hidden');
  document.getElementById('model-nav-btn').classList.add('active');
  loadModels();
}
function closeModelPanel() {
  document.getElementById('model-panel').classList.remove('open');
  document.getElementById('model-overlay').classList.add('hidden');
  document.getElementById('model-nav-btn').classList.remove('active');
}

// ─── Skill Panel ─────────────────────────────────────────────────

async function openSkillPanel() {
  closeModelPanel();
  document.getElementById('skill-panel').classList.add('open');
  document.getElementById('model-overlay').classList.remove('hidden');
  document.getElementById('skill-nav-btn').classList.add('active');
  await loadSkills();
}
function closeSkillPanel() {
  document.getElementById('skill-panel').classList.remove('open');
  document.getElementById('model-overlay').classList.add('hidden');
  document.getElementById('skill-nav-btn').classList.remove('active');
}

let allSkills = [];

const SKILL_META = {
  'knowledge-extraction': { icon: '🔍', iconCls: 'icon-purple', step: 2 },
  'knowledge-revision': { icon: '📝', iconCls: 'icon-orange', step: 3 },
  'knowledge-pattern-mining': { icon: '🔬', iconCls: 'icon-teal', step: 2 },
  'knowledge-gap-analysis': { icon: '🎯', iconCls: 'icon-blue', step: 2 },
  'knowledge-freshness-audit': { icon: '🔄', iconCls: 'icon-green', step: 5 },
};

async function loadSkills() {
  const body = document.getElementById('skill-panel-body');
  body.innerHTML = '<div class="skill-loading">加载中...</div>';
  try {
    const resp = await fetch(API_BASE + '/api/skills');
    const data = await resp.json();
    if (data.status === 'ok') {
      allSkills = data.skills || [];
      renderSkills(body);
    } else {
      body.innerHTML = '<div class="skill-error">加载失败</div>';
    }
  } catch (e) {
    body.innerHTML = '<div class="skill-error">网络错误</div>';
  }
}

function renderSkills(container) {
  if (allSkills.length === 0) {
    container.innerHTML = '<div class="skill-empty">暂无已注册的 Skill</div>';
    return;
  }
  let html = '<div class="skill-section"><div class="skill-section-header"><span>已注册技能</span><span style="font-weight:400;color:#aaa">' + allSkills.length + ' 个</span></div>';
  for (const skill of allSkills) {
    const enabled = skill.enabled !== false;
    const meta = SKILL_META[skill.id] || { icon: '⚡', iconCls: 'icon-blue' };
    html += `
      <div class="skill-card ${enabled ? '' : 'skill-card-disabled'}" id="skill-item-${skill.id}" data-skill-id="${skill.id}">
        <div class="skill-card-row" onclick="toggleSkillDetail('${skill.id}')">
          <div class="skill-card-icon ${meta.iconCls}">${meta.icon}</div>
          <div class="skill-card-info">
            <div class="skill-card-name">${escapeHtml(skill.name)}</div>
            <div class="skill-card-brief">${enabled ? '已启用' : '已禁用'}</div>
          </div>
          <div class="skill-card-controls">
            <div class="skill-toggle ${enabled ? 'on' : ''}" onclick="event.stopPropagation(); toggleSkill('${skill.id}', ${!enabled})">
              <div class="skill-toggle-knob"></div>
            </div>
            <span class="skill-card-arrow" id="skill-arrow-${skill.id}">▾</span>
          </div>
        </div>
        <div class="skill-card-detail hidden" id="skill-detail-${skill.id}">
          <div class="skill-loading">加载详情...</div>
        </div>
      </div>
    `;
  }
  html += '</div>';
  container.innerHTML = html;
}

async function toggleSkillDetail(skillId) {
  const detail = document.getElementById('skill-detail-' + skillId);
  const card = document.getElementById('skill-item-' + skillId);

  if (!detail.classList.contains('hidden')) {
    detail.classList.add('hidden');
    card.classList.remove('skill-card-expanded');
    return;
  }

  detail.classList.remove('hidden');
  card.classList.add('skill-card-expanded');

  if (detail.querySelector('.skill-loading')) {
    try {
      const resp = await fetch(API_BASE + '/api/skills/' + skillId);
      const data = await resp.json();
      if (data.status === 'ok') {
        renderSkillDetail(detail, data.skill);
      } else {
        detail.innerHTML = '<div class="skill-error">加载失败</div>';
      }
    } catch (e) {
      detail.innerHTML = '<div class="skill-error">网络错误</div>';
    }
  }
}

function renderSkillDetail(container, s) {
  const tagGroup = (label, items, cls) => {
    if (!items || !items.length) return '';
    return '<div class="skill-detail-group"><div class="skill-detail-label">' + label + '</div><div class="skill-detail-tags">' + items.map(function(i) { return '<span class="skill-tag ' + (cls || '') + '">' + escapeHtml(i) + '</span>'; }).join('') + '</div></div>';
  };
  const section = (title, content) => {
    if (!content) return '';
    return '<div class="skill-detail-section"><div class="skill-detail-section-title">' + title + '</div><div class="skill-detail-section-body">' + content + '</div></div>';
  };
  const listSection = (title, items) => {
    if (!items || !items.length) return '';
    return section(title, '<ul class="skill-detail-list">' + items.map(function(i) { return '<li>' + escapeHtml(i) + '</li>'; }).join('') + '</ul>');
  };
  const formatText = function(text) {
    return escapeHtml(text).replace(/\n\n/g, '</p><p>').replace(/\n/g, '<br>');
  };

  let html = '';

  // 详细描述
  if (s.detailed_description) {
    html += section('📖 详细说明', '<p>' + formatText(s.detailed_description) + '</p>');
  }

  // 业务价值
  if (s.business_value) {
    html += section('💡 业务价值', '<p>' + formatText(s.business_value) + '</p>');
  }

  // 使用指南
  if (s.usage_guide) {
    html += section('📋 使用步骤', '<p>' + formatText(s.usage_guide) + '</p>');
  }

  // 输入输出示例
  if (s.input_example || s.output_example) {
    let ioHtml = '';
    if (s.input_example) {
      ioHtml += '<div class="skill-io-item"><div class="skill-io-label skill-io-label-in">📥 输入示例</div><div class="skill-io-content">' + escapeHtml(s.input_example) + '</div></div>';
    }
    if (s.output_example) {
      ioHtml += '<div class="skill-io-item"><div class="skill-io-label skill-io-label-out">📤 输出示例</div><div class="skill-io-content">' + formatText(s.output_example) + '</div></div>';
    }
    html += section('🔧 输入 / 输出', ioHtml);
  }

  // 适用场景
  html += listSection('✅ 适用场景', s.applicable_scenarios);

  // 能力标签
  html += tagGroup('🏷️ 核心能力', s.capabilities, 'capability');
  html += tagGroup('📂 支持格式', s.supported_formats, '');
  html += tagGroup('🎨 输出风格', s.output_styles, '');

  // 触发条件
  html += tagGroup('🔍 触发条件', s.triggers, '');

  // 局限性
  html += listSection('⚠️ 局限性', s.limitations);

  // 文件限制 + 版本
  var metaHtml = '';
  if (s.max_file_size_mb) {
    metaHtml += '<div class="skill-detail-meta-item">📦 最大文件：<strong>' + s.max_file_size_mb + ' MB</strong></div>';
  }
  if (s.version) {
    metaHtml += '<div class="skill-detail-meta-item">🔖 版本：<strong>' + escapeHtml(s.version) + '</strong></div>';
  }
  if (s.related_step) {
    metaHtml += '<div class="skill-detail-meta-item">📌 关联步骤：<strong>Step ' + s.related_step + '</strong></div>';
  }
  if (metaHtml) {
    html += '<div class="skill-detail-meta">' + metaHtml + '</div>';
  }

  container.innerHTML = html;
}

async function toggleSkill(skillId, enable) {
  try {
    const resp = await fetch(API_BASE + '/api/skills/' + skillId, {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ enabled: enable }),
    });
    const data = await resp.json();
    if (data.status === 'ok') {
      await loadSkills();
    } else {
      alert('操作失败: ' + (data.error || '未知错误'));
    }
  } catch (e) {
    alert('网络错误');
  }
}

let editingModelName = null;

function toggleCcbModelFields() {
  const apiType = document.getElementById('new-model-api-type')?.value || 'openai';
  const ccbBlock = document.getElementById('ccb-model-fields');
  const openaiTuning = document.getElementById('openai-model-tuning');
  if (ccbBlock) ccbBlock.classList.toggle('hidden', apiType !== 'ccb_ainlplm');
  if (openaiTuning) openaiTuning.classList.toggle('hidden', apiType === 'ccb_ainlplm');
}

function showAddModelForm() {
  editingModelName = null;
  const titleEl = document.getElementById('model-form-title');
  const saveBtn = document.getElementById('model-form-save-btn');
  if (titleEl) titleEl.textContent = '添加自定义模型';
  if (saveBtn) saveBtn.textContent = '添加';
  document.getElementById('new-model-name').readOnly = false;
  ['new-model-name','new-model-model','new-model-url','new-model-apikey','new-model-desc','new-model-tx-code','new-model-sec-node'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.value = '';
  });
  const apiTypeEl = document.getElementById('new-model-api-type');
  if (apiTypeEl) apiTypeEl.value = 'openai';
  document.getElementById('new-model-maxtokens').value = '4096';
  document.getElementById('new-model-temp').value = '0.7';
  toggleCcbModelFields();
  document.getElementById('add-model-form').classList.remove('hidden');
}

function hideAddModelForm() {
  editingModelName = null;
  document.getElementById('add-model-form').classList.add('hidden');
}

async function editModel(name) {
  try {
    const resp = await fetch(API_BASE + '/api/llm/models/' + encodeURIComponent(name));
    const data = await resp.json();
    if (data.status !== 'ok' || !data.model) {
      alert(data.error || '加载模型失败');
      return;
    }
    const m = data.model;
    editingModelName = name;
    document.getElementById('model-form-title').textContent = m.is_preset ? '编辑预设模型' : '编辑自定义模型';
    document.getElementById('model-form-save-btn').textContent = '保存';
    document.getElementById('new-model-name').value = m.name || '';
    document.getElementById('new-model-name').readOnly = true;
    document.getElementById('new-model-model').value = m.model || '';
    document.getElementById('new-model-url').value = m.url || '';
    document.getElementById('new-model-apikey').value = m.api_key || '';
    document.getElementById('new-model-maxtokens').value = m.max_tokens || 4096;
    document.getElementById('new-model-temp').value = m.temperature != null ? m.temperature : 0.7;
    document.getElementById('new-model-desc').value = m.description || '';
    const apiTypeEl = document.getElementById('new-model-api-type');
    if (apiTypeEl) apiTypeEl.value = m.api_type || 'openai';
    const txEl = document.getElementById('new-model-tx-code');
    const secEl = document.getElementById('new-model-sec-node');
    if (txEl) txEl.value = m.tx_code || '';
    if (secEl) secEl.value = m.sec_node_no || '';
    toggleCcbModelFields();
    document.getElementById('add-model-form').classList.remove('hidden');
  } catch (e) {
    alert('加载模型失败: ' + e.message);
  }
}

async function loadModels() {
  try {
    const resp = await fetch(API_BASE + '/api/llm/models');
    const data = await resp.json();
    if (data.status === 'ok') {
      allModels = data.models || [];
      renderModelList();
      refreshModelSelects();
    }
  } catch (e) {
    console.error('Failed to load models:', e);
  }
}

function renderModelList() {
  const container = document.getElementById('model-list');
  if (!allModels.length) {
    container.innerHTML = '<div style="color:var(--text-muted);font-size:12px;padding:12px">暂无模型配置</div>';
    return;
  }
  let html = '';
  allModels.forEach((m, idx) => {
    const cls = m.is_preset ? 'preset' : 'custom';
    const badge = m.is_preset ? '<span style="font-size:10px;padding:1px 5px;background:var(--red);color:#fff;border-radius:2px">预设</span>' : '<span style="font-size:10px;padding:1px 5px;background:#3491fa;color:#fff;border-radius:2px">自定义</span>';
    const apiBadge = (m.api_type === 'ccb_ainlplm')
      ? '<span style="font-size:10px;padding:1px 5px;background:#6b4;border-radius:2px;color:#fff">建行</span>'
      : '<span style="font-size:10px;padding:1px 5px;background:#888;border-radius:2px;color:#fff">OpenAI</span>';
    html += '<div class="model-card ' + cls + '">';
    html += '<div class="model-card-name">' + escapeHtml(m.name) + ' ' + badge + ' ' + apiBadge + '</div>';
    html += '<div class="model-card-model">' + escapeHtml(m.model) + '</div>';
    if (m.description) html += '<div class="model-card-desc">' + escapeHtml(m.description) + '</div>';
    html += '<div class="model-card-url">' + escapeHtml(m.url) + '</div>';
    html += '<div class="model-card-key">Key: ' + escapeHtml(m.api_key || m.api_key_masked || '') + '</div>';
    html += '<div class="model-card-actions">';
    html += '<button type="button" class="mini-btn" data-action="edit" data-model-index="' + idx + '">编辑</button>';
    html += '<button type="button" class="mini-btn" data-action="test" data-model-index="' + idx + '">测试连接</button>';
    html += '<button type="button" class="mini-btn" data-action="stream" data-model-index="' + idx + '">流式测试</button>';
    if (!m.is_preset) {
      html += '<button type="button" class="mini-btn danger" data-action="delete" data-model-index="' + idx + '">删除</button>';
    }
    html += '</div>';
    html += '<div class="model-card-status" id="model-status-' + idx + '"></div>';
    html += '<div class="model-card-stream hidden" id="model-stream-' + idx + '"></div>';
    html += '</div>';
  });
  container.innerHTML = html;
}

function initModelListEvents() {
  const container = document.getElementById('model-list');
  if (!container || container._modelEventsBound) return;
  container._modelEventsBound = true;
  container.addEventListener('click', (e) => {
    const btn = e.target.closest('[data-action]');
    if (!btn) return;
    e.preventDefault();
    e.stopPropagation();
    const idx = parseInt(btn.dataset.modelIndex, 10);
    const m = allModels[idx];
    if (!m) return;
    const action = btn.dataset.action;
    if (action === 'test') testModel(m.name, idx, btn);
    else if (action === 'stream') testModelStream(m.name, idx, btn);
    else if (action === 'edit') editModel(m.name);
    else if (action === 'delete') deleteModel(m.name);
  });
}

function resolveModelName(selectId) {
  const el = document.getElementById(selectId);
  if (el && el.value) return el.value;
  if (allModels && allModels.length) return allModels[0].name;
  return '';
}

function refreshModelSelects() {
  const selects = ['s2-model', 's3-model', 's5-model', 's5-validate-model'];
  selects.forEach(id => {
    const el = document.getElementById(id);
    if (!el) return;
    const cur = el.value;
    el.innerHTML = '<option value="">-- 选择模型 --</option>';
    allModels.forEach(m => {
      const opt = document.createElement('option');
      opt.value = m.name;
      opt.textContent = m.name + (m.is_preset ? ' (预设)' : '');
      el.appendChild(opt);
    });
    if (cur) el.value = cur;
  });
}

// Skill step mapping: which skills are relevant to which step
const SKILL_STEP_MAP = {
  2: ['knowledge-extraction', 'knowledge-pattern-mining', 'knowledge-gap-analysis'],
  3: ['knowledge-revision'],
  5: ['knowledge-freshness-audit'],
};

async function loadStepSkillSelects(step) {
  const selectId = 's' + step + '-skill-select';
  const el = document.getElementById(selectId);
  if (!el) return;

  el.innerHTML = '<option value="">-- 选择 Skill --</option>';

  try {
    const resp = await fetch(API_BASE + '/api/skills');
    const data = await resp.json();
    if (data.status !== 'ok') return;

    const skills = data.skills || [];
    const relevantIds = SKILL_STEP_MAP[step] || [];

    skills.forEach(s => {
      if (s.enabled === false) return;
      // If step has specific skill mapping, filter by it
      if (relevantIds.length > 0 && !relevantIds.includes(s.id)) return;

      const opt = document.createElement('option');
      opt.value = s.id;
      opt.textContent = s.name + ' - ' + (s.description || '').substring(0, 30);
      el.appendChild(opt);
    });
  } catch (e) {
    console.error('Failed to load skills:', e);
  }
}

function setModelTestStatus(statusEl, state, message) {
  if (!statusEl) return;
  statusEl.className = 'model-card-status ' + (state || '');
  statusEl.textContent = message || '';
}

async function testModelStream(name, modelIndex, btnEl) {
  const statusEl = document.getElementById('model-status-' + modelIndex);
  const streamEl = document.getElementById('model-stream-' + modelIndex);
  if (streamEl) {
    streamEl.classList.remove('hidden');
    streamEl.textContent = '';
  }
  if (btnEl) {
    btnEl.disabled = true;
    btnEl.textContent = '流式中...';
  }
  setModelTestStatus(statusEl, 'testing', '流式连接 ' + name + ' ...');
  let fullText = '';
  try {
    const resp = await fetch(API_BASE + '/api/llm/stream-test', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ name: name, prompt: '你好，请用一句话介绍你自己。' }),
    });
    if (!resp.ok || !resp.body) {
      throw new Error('流式请求失败 HTTP ' + resp.status);
    }
    const reader = resp.body.getReader();
    const decoder = new TextDecoder();
    let buffer = '';
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      buffer += decoder.decode(value, { stream: true });
      const parts = buffer.split('\n\n');
      buffer = parts.pop() || '';
      for (const part of parts) {
        const line = part.split('\n').find(l => l.startsWith('data:'));
        if (!line) continue;
        try {
          const payload = JSON.parse(line.replace(/^data:\s*/, ''));
          if (payload.error) throw new Error(payload.error);
          if (payload.delta) {
            fullText += payload.delta;
            if (streamEl) streamEl.textContent = fullText;
          }
          if (payload.done) break;
        } catch (parseErr) {
          if (parseErr.message && parseErr.message !== 'Unexpected end of JSON input') throw parseErr;
        }
      }
    }
    setModelTestStatus(statusEl, 'ok', '流式完成: ' + (fullText.slice(0, 80) || '(空)'));
  } catch (e) {
    setModelTestStatus(statusEl, 'fail', e.message || '流式失败');
    if (streamEl) streamEl.textContent = '错误: ' + (e.message || '流式失败');
  }
  if (btnEl) {
    btnEl.disabled = false;
    btnEl.textContent = '流式测试';
  }
}

async function testModel(name, modelIndex, btnEl) {
  const statusEl = document.getElementById('model-status-' + modelIndex);
  if (btnEl) {
    btnEl.disabled = true;
    btnEl.classList.add('testing');
    btnEl.textContent = '测试中...';
  }
  setModelTestStatus(statusEl, 'testing', '正在连接 ' + name + ' ...');
  try {
    const result = await apiCallJSON('/api/llm/test', { name: name }, 'POST', 25000);
    if (result.status === 'ok') {
      setModelTestStatus(statusEl, 'ok', result.message || '连接成功');
    } else {
      setModelTestStatus(statusEl, 'fail', result.error || '连接失败');
    }
  } catch (e) {
    setModelTestStatus(statusEl, 'fail', e.message || '连接失败');
  }
  if (btnEl) {
    btnEl.disabled = false;
    btnEl.classList.remove('testing');
    btnEl.textContent = '测试连接';
  }
}

async function saveModel() {
  const name = document.getElementById('new-model-name').value.trim();
  const model = document.getElementById('new-model-model').value.trim();
  const url = document.getElementById('new-model-url').value.trim();
  const apiKey = document.getElementById('new-model-apikey').value.trim();
  const apiType = document.getElementById('new-model-api-type')?.value || 'openai';
  const maxTokens = parseInt(document.getElementById('new-model-maxtokens').value) || 4096;
  const temp = parseFloat(document.getElementById('new-model-temp').value) || 0.7;
  const desc = document.getElementById('new-model-desc').value.trim();
  const txCode = document.getElementById('new-model-tx-code')?.value.trim() || '';
  const secNode = document.getElementById('new-model-sec-node')?.value.trim() || '';

  if (!name || !model || !url) {
    alert('名称、模型标识、API 地址均为必填');
    return;
  }
  if (!editingModelName && !apiKey) {
    alert('添加模型时 API Key 为必填');
    return;
  }
  if (apiType === 'ccb_ainlplm' && (!txCode || !secNode)) {
    alert('建行接口需填写 Tx-Code 与 Sec-Node-No');
    return;
  }

  const payload = {
    name, model, url, api_type: apiType,
    max_tokens: maxTokens, temperature: temp, description: desc
  };
  if (apiKey) payload.api_key = apiKey;
  if (apiType === 'ccb_ainlplm') {
    payload.tx_code = txCode;
    payload.sec_node_no = secNode;
  }

  try {
    let result;
    if (editingModelName) {
      result = await apiCallJSON(
        '/api/llm/models/' + encodeURIComponent(editingModelName),
        payload,
        'PUT'
      );
    } else {
      result = await apiCallJSON('/api/llm/models', payload, 'POST');
    }
    if (result.status === 'ok') {
      hideAddModelForm();
      loadModels();
    } else {
      alert(result.error || (editingModelName ? '保存失败' : '添加失败'));
    }
  } catch (e) {
    alert((editingModelName ? '保存失败: ' : '添加失败: ') + e.message);
  }
}

async function deleteModel(name) {
  if (!confirm('确定删除模型 "' + name + '"？')) return;
  try {
    const resp = await fetch(API_BASE + '/api/llm/models/' + encodeURIComponent(name), { method: 'DELETE' });
    const result = await resp.json();
    if (result.status === 'ok') {
      loadModels();
    } else {
      alert(result.error || '删除失败');
    }
  } catch (e) {
    alert('删除失败: ' + e.message);
  }
}

// Load models on startup
initModelListEvents();
loadModels();

// ═══════════════════════════════════════════════════════════════════
// Step 1: Scenario Anchoring
// ═══════════════════════════════════════════════════════════════════

let _s1SubScenarioCount = 0;

function step1AddSubScenario() {
  _s1SubScenarioCount++;
  const idx = _s1SubScenarioCount;
  const container = document.getElementById('s1-sub-scenarios');
  const div = document.createElement('div');
  div.className = 's1-sub-item';
  div.id = 's1-sub-' + idx;
  div.innerHTML = '<div class="s1-sub-row">' +
    '<input type="text" class="s1-sub-name" placeholder="子场景名称" data-idx="' + idx + '">' +
    '<button class="s1-sub-remove" onclick="step1RemoveSubScenario(' + idx + ')">✕</button>' +
    '</div>' +
    '<textarea class="s1-sub-content" rows="2" placeholder="子场景内容描述" data-idx="' + idx + '"></textarea>';
  container.appendChild(div);
}

function step1RemoveSubScenario(idx) {
  const el = document.getElementById('s1-sub-' + idx);
  if (el) el.remove();
}

function step1GetSubScenarios() {
  const subs = [];
  document.querySelectorAll('.s1-sub-item').forEach(el => {
    const name = el.querySelector('.s1-sub-name').value.trim();
    const content = el.querySelector('.s1-sub-content').value.trim();
    if (name || content) subs.push({ name, content });
  });
  return subs;
}

async function step1Generate() {
  const btn = document.getElementById('s1-generate');
  if (!btn || btn._locked) return;
  btn.disabled = true;
  btn._locked = true;
  try {
  const scenarioName = document.getElementById('s1-scenario-name').value.trim();
  const scenarioContent = document.getElementById('s1-scenario-content').value.trim();
  const templateFile = document.getElementById('s1-template-file').files[0];
  const outputFormat = document.getElementById('s1-output-format')?.value || 'excel';
  const legacyTemplate = document.getElementById('s1-legacy-template')?.value || '';
  const knowledgeColumns = step1GetKnowledgeColumns();
  const hasCustomColumns = knowledgeColumns.length > 0;

  if (!scenarioName) { alert('请填写场景名称'); return; }
  if (!currentPipeline) {
    alert('请先从总览页「新建流水线」或「继续」进入一条流水线，再生成场景骨架');
    return;
  }
  if (!templateFile && !legacyTemplate && knowledgeColumns.length === 0) {
    alert('请至少添加一列知识字段，或上传/选用 Excel 模板');
    return;
  }

  const subScenarios = step1GetSubScenarios();
  renderLoading('s1-output');

  const fd = new FormData();
  fd.append('scenario_name', scenarioName);
  fd.append('scenario_content', scenarioContent);
  fd.append('sub_scenarios', JSON.stringify(subScenarios));
  fd.append('knowledge_columns', JSON.stringify(knowledgeColumns));
  if (templateFile) {
    fd.append('template', templateFile);
    fd.append('output_format', outputFormat);
  } else {
    fd.append('output_format', outputFormat);
    // 有自定义知识列时，默认按自定义列生成；不再隐式落回 legacy 模板
    if (legacyTemplate && !hasCustomColumns) {
      fd.append('template_mode', 'legacy');
      fd.append('default_template', legacyTemplate);
    }
  }
  fd.append('pipeline_id', currentPipeline.id);

  try {
    const result = await apiCall('/api/step1/generate', fd);
    let html = '';
    if (result.status === 'ok') {
      applyStep1GenerateResult(result);
      syncPipelineFromStep1Response(result);
      const allForm = collectAllStepsFormData();
      currentPipeline.step_data = mergeStepDataPreserveOutputs(currentPipeline.step_data, allForm);
      const excelFile = result.excel_file || result.file_name;
      await persistPipeline({
        ...allForm,
        step1_output_file: excelFile,
        step1_download_url: result.excel_download_url || ('/downloads/' + excelFile),
        step1_knowledge_columns: result.knowledge_columns || knowledgeColumns,
        step1_output_format: result.output_format || outputFormat,
        ...(result.markdown_file ? {
          step1_md_file: result.markdown_file,
          step1_md_download_url: result.markdown_download_url,
        } : {}),
      });
      html += '<div class="s1-result-box">';
      html += '<div class="s1-result-title">场景骨架生成成功</div>';
      html += '<div class="s1-result-stats">';
      html += '<div class="s2-stat"><span class="s2-stat-num">' + (result.fields_info ? result.fields_info.length : 0) + '</span><span class="s2-stat-label">工作表</span></div>';
      html += '<div class="s2-stat"><span class="s2-stat-num">' + (result.sub_scenario_count || 0) + '</span><span class="s2-stat-label">子场景</span></div>';
      html += '</div>';
      const templateSourceMap = {
        schema: '自定义列 · Excel',
        schema_markdown: '自定义列 · Markdown+Excel',
        legacy: '部门 Excel 模板',
        legacy_markdown: '部门 Excel 模板 · Markdown+Excel',
        upload: '上传 Excel 模板',
        upload_markdown: '上传 Excel 模板 · Markdown+Excel',
      };
      const templateSourceLabel = templateSourceMap[result.template_source] || '模板';
      const templateName = result.template_name || '未命名';
      html += '<div class="s1-result-template">来源：' + escapeHtml(templateSourceLabel) + ' · ' + escapeHtml(templateName) + '</div>';
      if (result.knowledge_columns && result.knowledge_columns.length) {
        html += '<div class="s1-result-template">知识列：' + escapeHtml(result.knowledge_columns.join('、')) + '</div>';
      }
      if (result.columns_enriched) {
        html += '<div class="s1-result-template file-hint">已按 Markdown 模式自动补齐富语义列，Step2 将按完整字段深度萃取。</div>';
        step1RenderKnowledgeColumns(result.knowledge_columns);
      }
      if (result.fields_info && result.fields_info.length) {
        html += '<div class="s1-result-sheets">';
        result.fields_info.forEach(function(fi) {
          html += '<div class="s1-sheet-item"><span class="s1-sheet-name">' + escapeHtml(fi.sheet) + '</span><span class="s1-sheet-meta">' + fi.headers.length + ' 列 · ' + (fi.data_rows || 0) + ' 行</span></div>';
        });
        html += '</div>';
      }
      html += '<div class="s1-result-actions">';
      const mdFlow = (result.output_format === 'markdown') || prefersMarkdownFlow();
      const step1MdFile = result.markdown_file || '';
      if (result.download_url) {
        const dlLabel = mdFlow ? '下载 Markdown 骨架' : '下载 Excel 骨架';
        html += '<a class="action-btn small-btn" href="' + API_BASE + result.download_url + '" download>' + dlLabel + '</a>';
      }
      if (mdFlow && step1MdFile) {
        html += `<button class="action-btn small-btn secondary-btn" onclick="previewStep5File('${escapeHtml(step1MdFile)}','Step1 骨架 Markdown 预览')">预览/编辑 Markdown</button>`;
      }
      if (!mdFlow && result.excel_download_url) {
        html += '<button class="action-btn small-btn secondary-btn" onclick="step1PreviewExcel(\'' + (result.excel_file || result.file_name) + '\')">预览 Excel</button>';
      }
      html += '</div>';
      html += '</div>';
      await markStepDone(1);
      if (currentStep === 2) loadStep2PrevOutput();
    } else {
      html = '<div class="error-list"><div class="error-item">' + escapeHtml(result.error || '未知错误') + '</div></div>';
    }
    renderOutput('s1-output', html);
  } catch (e) {
    renderOutput('s1-output', '<div class="error-list"><div class="error-item">' + escapeHtml(e.message) + '</div></div>');
  }
  } finally {
    btn.disabled = false;
    btn._locked = false;
  }
}

async function step1PreviewExcel(fileName) {
  const modal = document.getElementById('excel-editor-modal');
  if (modal) modal.classList.add('active');
  renderExcelEditorLoading();
  if (typeof ExcelEditor !== 'undefined' && ExcelEditor.preloadLuckysheet) {
    ExcelEditor.preloadLuckysheet();
  }

  const fd = new FormData();
  fd.append('file_name', fileName);
  if (currentPipeline) fd.append('pipeline_id', currentPipeline.id);
  fd.append('step', '1');
  try {
    const result = await apiCall('/api/excel/read', fd);
    if (result.status === 'ok') {
      openExcelEditorWithSheets(result.sheets, fileName, result.file_path || fileName);
    } else {
      if (modal) modal.classList.remove('active');
      alert(result.error || '读取失败');
    }
  } catch (e) {
    if (modal) modal.classList.remove('active');
    alert('读取失败: ' + e.message);
  }
}

function openExcelEditorWithSheets(sheets, fileName, filePath) {
  const modal = document.getElementById('excel-editor-modal');
  const stepLabel = document.getElementById('excel-editor-step-label');
  if (stepLabel) stepLabel.textContent = '预览: ' + fileName;
  _excelEditorData.sheets = ExcelEditor.normalizeSheetsFromApi(
    typeof sheets === 'object' && !Array.isArray(sheets) ? sheets : { Sheet1: sheets }
  );
  _excelEditorData.file_path = filePath || fileName;
  _excelEditorData.active_sheet = Object.keys(_excelEditorData.sheets)[0] || '';
  _excelEditorData.modified = false;
  _excelEditorData.step = 1;
  if (modal) modal.classList.add('active');
  requestAnimationFrame(function () {
    requestAnimationFrame(function () {
      renderExcelEditorContent();
    });
  });
}

// ═══════════════════════════════════════════════════════════════════
// Step 2: Knowledge Extraction
// ═══════════════════════════════════════════════════════════════════

function renderStep2PrevOutputCard(data) {
  const area = document.getElementById('s2-prev-output-area');
  if (!area) return;
  let fieldsHtml = '';
  if (data.fields_info && data.fields_info.length) {
    const allHeaders = data.fields_info.flatMap(f => f.headers);
    fieldsHtml = '<div class="s2-prev-card-fields">' +
      allHeaders.slice(0, 12).map(h => '<span class="s2-prev-field-tag">' + escapeHtml(h) + '</span>').join('') +
      (allHeaders.length > 12 ? '<span class="s2-prev-field-tag">+' + (allHeaders.length - 12) + '</span>' : '') +
      '</div>';
  }
  area.innerHTML = `
    <div class="s2-prev-card">
      <div class="s2-prev-card-top">
        <div class="s2-prev-card-icon">📋</div>
        <div class="s2-prev-card-name">${escapeHtml(data.scenario || '场景模板')}</div>
        <span class="s2-prev-card-badge">已就绪</span>
      </div>
      <div class="s2-prev-card-info">
        领域：<span>${escapeHtml(data.domain || '-')}</span> · 文件：<span>${escapeHtml(data.file_name || '-')}</span>
      </div>
      ${fieldsHtml}
      <div class="s2-result-actions" style="margin-top:8px;">
        ${data.markdown_download_url ? `<a class="s2-result-btn s2-result-btn-primary" href="${API_BASE + data.markdown_download_url}" download>下载 Markdown</a>` : ''}
        ${data.markdown_file ? `<button class="s2-result-btn" onclick="previewStep5File('${escapeHtml(data.markdown_file)}','Step1 骨架 Markdown 预览')">预览/编辑 Markdown</button>` : ''}
      </div>
    </div>
  `;
}

async function ensureStep1OutputLinked() {
  if (!currentPipeline?.step_data?.step1_output_file) {
    const recalled = recallStep1Output(currentPipeline.id);
    if (recalled?.file_name) {
      currentPipeline.step_data = currentPipeline.step_data || {};
      currentPipeline.step_data.step1_output_file = recalled.file_name;
      currentPipeline.step_data.step1_download_url = recalled.download_url;
      await persistPipeline({
        step1_output_file: recalled.file_name,
        step1_download_url: recalled.download_url,
      });
    }
  }
}

async function loadStep2PrevOutput() {
  const area = document.getElementById('s2-prev-output-area');
  if (!currentPipeline) {
    area.innerHTML = '<div class="s2-prev-empty">当前无流水线</div>';
    return;
  }

  area.innerHTML = '<div class="s2-prev-empty" style="color:var(--text-secondary)">检测中...</div>';

  try {
    await ensureStep1OutputLinked();

    const resp = await fetch(API_BASE + '/api/step2/prev_output?pipeline_id=' + currentPipeline.id);
    const data = await resp.json();

    if (data.status !== 'ok' || !data.has_output) {
      const hint = data.hint || '请先在「场景锚定」点击「生成场景骨架」（需已从总览进入当前流水线）';
      area.innerHTML = '<div class="s2-prev-empty">' + escapeHtml(hint) + '</div>';
      return;
    }

    renderStep2PrevOutputCard(data);
  } catch (e) {
    area.innerHTML = '<div class="s2-prev-empty">检测失败，可手动上传</div>';
  }
}


function isStep2PreextractFile(fileName) {
  const n = String(fileName || '').toLowerCase();
  if (!n.endsWith('.xlsx')) return false;
  return n.startsWith('preextract_') || n.startsWith('edited_step2_');
}

function resolveStep2DownloadInfo(result) {
  const apiName = result.download_name || '';
  const apiUrl = result.download_url || '';

  let dlName = '';
  let dlUrl = '';

  if (isStep2PreextractFile(apiName)) {
    dlName = apiName;
    dlUrl = apiUrl || ('/downloads/' + dlName);
  } else if (apiUrl) {
    const m = String(apiUrl).match(/\/([^/?#]+\.xlsx)(?:\?|$)/i);
    if (m && isStep2PreextractFile(m[1])) {
      dlName = m[1];
      dlUrl = apiUrl;
    }
  }

  if (dlName && !dlUrl) dlUrl = '/downloads/' + dlName;
  return { dlName, dlUrl };
}
async function step2SkillExtract() {
  const btn = document.getElementById('s2-skill-extract');
  if (!btn || btn._locked) return;
  const skillId = _step2ActiveSkill || 'knowledge-extraction';
  const model = resolveModelName('s2-model');
  const style = document.getElementById('s2-extract-style').value;

  // ── 预校验 ──
  if (!model || !currentPipeline) { updateStep2Readiness(); showToast('请先补全执行条件', 'error'); return; }
  if (skillId === 'knowledge-pattern-mining') {
    var pf = document.getElementById('s2-pattern-files');
    var pt = document.getElementById('s2-pattern-text');
    if ((!pf || !pf.files || pf.files.length < 2) && (!pt || (pt.value || '').trim().length < 50)) {
      showToast('模式发现需要至少2个案例文件或多段案例文本', 'error'); return;
    }
  }

  btn._locked = true;
  btn.disabled = true;
  btn.classList.add('loading');
  btn.innerHTML = '执行中<span class="btn-estimate">· 通常 10-60s</span>';
  renderLoading('s2-output');

  const fd = new FormData();
  fd.append('skill_id', skillId);
  fd.append('model', model);
  fd.append('style', style);
  fd.append('pipeline_id', currentPipeline.id);

  // ── 按 Skill 构建请求 ──
  if (skillId === 'knowledge-gap-analysis') {
    // 盲区检测：只需 pipeline_id，后端自动读取 Schema 和 Excel
  } else if (skillId === 'knowledge-pattern-mining') {
    // 模式发现：发送多文件或粘贴文本
    var patternFiles = document.getElementById('s2-pattern-files');
    var patternText = document.getElementById('s2-pattern-text');
    if (patternFiles && patternFiles.files) {
      for (var i = 0; i < patternFiles.files.length; i++) { fd.append('files', patternFiles.files[i]); }
    }
    if (patternText && patternText.value.trim()) { fd.append('content', patternText.value.trim()); }
  } else {
    // 知识萃取：文档或案例复盘
    if (_step2InputMode === 'case') {
      var caseText = _buildCaseReviewText();
      if (caseText) { fd.append('content', caseText); fd.append('content_type', 'case_review'); }
    } else {
      var docText = document.getElementById('s2-doc-text').value;
      var sourceFile = document.getElementById('s2-source-file').files[0];
      var cachedFile = currentPipeline?.step_data?.step2_cached_file || '';
      if (docText.trim()) fd.append('content', docText);
      if (sourceFile) fd.append('file', sourceFile);
      if (!sourceFile && cachedFile) fd.append('cached_file', cachedFile);
    }
  }

  try {
    const resp = await fetch(API_BASE + '/api/skills/execute', { method: 'POST', body: fd });
    const text = await resp.text();
    let result;
    try { result = JSON.parse(text); } catch { result = { raw: text }; }

    if (result.status === 'ok') {
      // ── 模式发现 / 盲区检测：渲染报告下载卡片 ──
      if (skillId === 'knowledge-pattern-mining' || skillId === 'knowledge-gap-analysis' || skillId === 'knowledge-freshness-audit') {
        var rptName = result.report_name || '';
        var rptUrl = result.download_url || ('/downloads/' + rptName);
        var summary = '';
        if (result.case_count) summary += '<div class="s2-stat"><span class="s2-stat-num">' + result.case_count + '</span><span class="s2-stat-label">案例数</span></div>';
        if (result.gaps_found != null) summary += '<div class="s2-stat"><span class="s2-stat-num">' + result.gaps_found + '</span><span class="s2-stat-label">盲区数</span></div>';
        if (result.total_items) summary += '<div class="s2-stat"><span class="s2-stat-num">' + result.total_items + '</span><span class="s2-stat-label">条目数</span></div>';
        var rptHtml = '<div class="s2-result-success"><div class="s2-result-header">' + escapeHtml(result.skill_name || '分析完成') + '</div>';
        if (summary) rptHtml += '<div class="s1-result-stats">' + summary + '</div>';
        rptHtml += '<div class="s2-result-actions" style="margin-top:12px;">';
        if (rptUrl) rptHtml += '<a class="s2-result-btn s2-result-btn-primary" href="' + API_BASE + rptUrl + '" download>📥 下载分析报告</a>';
        rptHtml += '</div></div>';
        renderOutput('s2-output', rptHtml);
        markStepDone(2);
        return;
      }

      // ── 知识萃取：处理 Excel 下载 ──
      clearDownstreamOutputs(2);
      const { dlName, dlUrl } = resolveStep2DownloadInfo(result);
      const hasExcel = dlUrl && isStep2PreextractFile(dlName);
      if (!hasExcel) {
        const hint = result.error
          ? escapeHtml(result.error)
          : '未返回有效的萃取 Excel（文件名须为 preextract_*.xlsx）。可能下载到了场景骨架，请 Ctrl+F5 强刷后重试，并确认容器已更新到最新前端/后端。';
        renderOutput('s2-output', '<h4>萃取异常</h4><div class="error-list"><div class="error-item">' + hint + '</div></div>');
        return;
      }
      result.download_name = dlName;
      result.download_url = dlUrl;

      _lastStep2ExtractedText = result.extracted || '';

      if (currentPipeline) {
        currentPipeline.step_data = currentPipeline.step_data || {};
        currentPipeline.step_data.step2_output_file = dlName;
        currentPipeline.step_data.step2_download_url = dlUrl;
        currentPipeline.step_data.step2_md_file = result.markdown_file || '';
        currentPipeline.step_data.step2_md_download_url = result.markdown_download_url || '';
        currentPipeline.step_data.step2_extracted_count = result.extracted_count || 0;
        currentPipeline.step_data.skill_extract_result = _lastStep2ExtractedText;
        currentPipeline.step_data.skill_extract_style = result.style || style;
        delete currentPipeline.step_data.step2_preview_name;
        delete currentPipeline.step_data.step2_preview_url;
        await persistPipeline({
          step2_output_file: dlName,
          step2_download_url: dlUrl,
          step2_md_file: result.markdown_file || '',
          step2_md_download_url: result.markdown_download_url || '',
          step2_extracted_count: result.extracted_count || 0,
          skill_extract_result: result.extracted || '',
          skill_extract_style: result.style || style,
        });
        await refreshCurrentPipeline();
      }

      const tplHint = result.used_step1_template ? '已按 Step1 场景模板列回填' : '已生成标准萃取表';
      const ruleHint = result.style_rule
        ? ` · 规则过滤 ${result.style_rule.raw_count || 0} → ${result.style_rule.processed_count || 0}（目标 ${result.style_rule.min_items || '-'}-${result.style_rule.max_items || '-'}）`
        : '';
      const mdFlow = prefersMarkdownFlow();
      const mdFile = result.markdown_file || '';
      const mdUrl = result.markdown_download_url || '';
      const html = `
          <div class="s2-result-success">
            <div class="s2-result-header">知识萃取完成</div>
            <div class="s2-result-meta">模型：${escapeHtml(result.model || model)} · 风格：${escapeHtml(result.style || style)} · 共提取 <strong>${result.extracted_count || 0}</strong> 条 · ${escapeHtml(tplHint)}${escapeHtml(ruleHint)}</div>
            <div class="s2-prev-card" style="margin-top:12px;">
              <div class="s2-prev-card-top">
                <div class="s2-prev-card-icon">&#128196;</div>
                <div class="s2-prev-card-name">${escapeHtml(mdFlow && mdFile ? mdFile : dlName)}</div>
                <span class="s2-prev-card-badge">${mdFlow ? 'Markdown' : 'Excel'}</span>
              </div>
            </div>
            <div class="s2-result-actions" style="margin-top:12px;">
              ${mdUrl ? `<a class="s2-result-btn s2-result-btn-primary" href="${API_BASE + mdUrl}" download="${escapeHtml(mdFile || 'preextract.md')}">&#11015; 下载萃取 Markdown</a>` : ''}
              ${(!mdFlow) ? `<a class="s2-result-btn ${mdUrl ? '' : 's2-result-btn-primary'}" href="${API_BASE + dlUrl}" download="${escapeHtml(dlName)}">&#11015; 下载萃取 Excel</a>` : ''}
              ${mdFile ? `<button class="s2-result-btn" onclick="previewStep5File('${escapeHtml(mdFile)}','Step2 萃取 Markdown 预览')">&#128065; 预览/编辑 Markdown</button>` : ''}
              ${(!mdFlow) ? '<button class="s2-result-btn" onclick="editStep2Preextract()">&#9998; 在线编辑</button>' : ''}
              <button type="button" class="s2-result-btn copy" id="s2-copy-extract-btn" onclick="copyExtractResult()">&#128203; 复制 LLM 原文</button>
              <button type="button" class="s2-result-btn" onclick="openInterviewFromExtract()" style="background:#f0fdf4;color:#16a34a;border-color:#86efac;">&#128269; 深挖隐性知识</button>
            </div>
            <div class="s2-result-detail" id="s2-result-text" style="max-height:200px;overflow-y:auto;margin-top:8px;font-size:12px;color:#666;">
              <pre>${escapeHtml(result.extracted || '').substring(0, 1000)}${(result.extracted || '').length > 1000 ? '...' : ''}</pre>
            </div>
          </div>
        `;
      renderOutput('s2-output', html);
      markStepDone(2);
    } else {
      const preview = result.extracted_preview || result.extracted || result.raw_output || '';
      if (preview) _lastStep2ExtractedText = preview;
      let errHtml = '<h4>萃取失败</h4><div class="error-list"><div class="error-item">' + escapeHtml(result.error || '未知错误') + '</div>';
      if (result.parse_mode) {
        errHtml += '<div class="error-item" style="font-size:12px;color:#666;">解析模式：' + escapeHtml(result.parse_mode);
        if (result.target_column_count != null) {
          errHtml += ' · 模板列数：' + escapeHtml(String(result.target_column_count));
        }
        if (result.style_rule) {
          errHtml += ' · 过滤：' + escapeHtml(String(result.style_rule.raw_count || 0)) + '→' + escapeHtml(String(result.style_rule.processed_count || 0));
        }
        errHtml += '</div>';
      }
      if (preview) {
        errHtml += '<div style="margin-top:8px;font-size:12px;"><button type="button" class="s2-result-btn" onclick="copyExtractResult()">复制模型原文</button></div>';
        errHtml += '<pre style="max-height:180px;overflow:auto;font-size:11px;margin-top:6px;background:#f5f5f5;padding:8px;">' + escapeHtml(String(preview).substring(0, 2000)) + '</pre>';
      }
      errHtml += '</div>';
      renderOutput('s2-output', errHtml);
    }
  } catch (e) {
    renderOutput('s2-output', '<h4>错误</h4><div class="error-list"><div class="error-item">' + escapeHtml(e.message) + '</div></div>');
  } finally {
    btn.disabled = false;
    btn._locked = false;
    btn.innerHTML = '<span class="action-icon">&#9654;</span> 执行';
  }
}

async function copyExtractResult() {
  const text = _lastStep2ExtractedText
    || currentPipeline?.step_data?.skill_extract_result
    || '';
  if (!String(text).trim()) {
    showToast('暂无可复制的 LLM 原文', 'error');
    return;
  }
  const btn = document.getElementById('s2-copy-extract-btn');
  const ok = await copyTextToClipboard(text);
  if (ok) {
    if (btn) {
      const orig = btn.innerHTML;
      btn.textContent = '已复制';
      setTimeout(() => { btn.innerHTML = orig; }, 1500);
    } else {
      showToast('已复制 LLM 原文');
    }
  } else {
    showToast('复制失败，请手动选中文本复制', 'error');
  }
}

// ═══════════════════════════════════════════════════════════════════
// Step 3: Interactive Knowledge Alignment
// ═══════════════════════════════════════════════════════════════════

let _alignNotes = [];
let _alignNoteStates = {};
let _alignEditedValues = {};
let _alignCurrentFilter = 'all';
let _alignChatHistory = [];

const ALIGN_ACTION_LABELS = { modify: '修改', delete: '删除', add: '新增', supplement: '补充' };
const ALIGN_ACTION_COLORS = { modify: '#faad14', delete: '#ff4d4f', add: '#52c41a', supplement: '#1890ff' };

// ── 隐性注释 / 追问卡片 ──
const TACIT_FOLLOWUP_QUESTIONS = {
  modify: '你改这个，是因为遇到过不适用的情况吗？能举一个具体的案例吗？',
  delete: '这条规则在什么情况下反而会误导人？有没有踩过坑？',
  add: '这条新知识你是从哪学到的？是自己的经验还是听说的？',
  supplement: '补充的内容是你最近才意识到的，还是一直知道但没写下来的？'
};

function showTacitFollowup(noteEl, noteId, actionType) {
  // 已有追问卡片则跳过
  if (noteEl.querySelector('.tacit-followup')) return;
  var question = TACIT_FOLLOWUP_QUESTIONS[actionType] || '能分享一下这次修订背后的经验吗？';
  var card = document.createElement('div');
  card.className = 'tacit-followup';
  card.innerHTML =
    '<div class="tacit-followup-label">💡 ' + escapeHtml(question) + '</div>' +
    '<textarea id="tacit-answer-' + noteId + '" placeholder="写几句话就行，哪怕只是「当时感觉不对」也比留空有信息量..."></textarea>' +
    '<div class="tacit-followup-actions">' +
      '<button class="tacit-followup-skip" onclick="dismissTacitFollowup(this)">跳过</button>' +
      '<button class="tacit-followup-save" onclick="saveTacitAnnotation(\'' + noteId + '\', \'' + escapeHtml(actionType) + '\', \'' + escapeHtml(question) + '\')">保存隐性注释</button>' +
    '</div>';
  noteEl.appendChild(card);
}

function dismissTacitFollowup(btn) {
  var card = btn.closest('.tacit-followup');
  if (card) card.remove();
}

function saveTacitAnnotation(noteId, actionType, question) {
  var ta = document.getElementById('tacit-answer-' + noteId);
  var answer = ta ? ta.value.trim() : '';
  if (!answer) { dismissTacitFollowup(ta); return; }
  _alignTacitAnnotations[noteId] = { action: actionType, question: question, answer: answer };
  // 视觉反馈
  var card = ta.closest('.tacit-followup');
  if (card) {
    card.innerHTML = '<div style="color:var(--green);font-size:12px;padding:4px 0">✓ 隐性注释已记录 — 将在生成对齐稿时一并保存</div>';
    setTimeout(function () { if (card.parentNode) card.remove(); }, 2000);
  }
  showToast('隐性注释已保存');
}

function getTacitAnnotationsPayload() {
  var list = [];
  Object.keys(_alignTacitAnnotations).forEach(function (id) {
    var a = _alignTacitAnnotations[id];
    if (a && a.answer) list.push({ note_id: id, action: a.action, question: a.question, answer: a.answer });
  });
  return list;
}

async function loadStep3PrevOutput() {
  const pid = currentPipeline ? currentPipeline.id : null;
  if (!pid) return;
  try {
    const resp = await fetch(API_BASE + `/api/step3/prev_output?pipeline_id=${pid}`);
    const data = await resp.json();
    const card = document.getElementById('s3-prev-draft');
    const empty = document.getElementById('s3-prev-empty');
    const info = document.getElementById('s3-prev-info');
    const tags = document.getElementById('s3-prev-tags');
    _alignChatHistory = Array.isArray(currentPipeline?.step_data?._align_chat_history)
      ? currentPipeline.step_data._align_chat_history
      : [];
    renderStep3ChatHistory();
    if (data.has_output) {
      card.style.display = '';
      empty.style.display = 'none';
      const extractedCount = data.extracted_count || (data.fields_info && data.fields_info[0] ? data.fields_info[0].rows : 0);
      info.textContent = `${data.file_name}` + (data.scenario ? ` · ${data.scenario}` : '') + (data.style ? ` · ${data.style}` : '') + (extractedCount > 0 ? ` · 萃取${extractedCount}条知识` : '');
      if (data.fields_info && data.fields_info.length > 0) {
        const mainSheet = data.fields_info[0];
        const tagHtml = mainSheet.headers.slice(0, 6).map(h => `<span class="s2-prev-tag">${escapeHtml(h)}</span>`).join('');
        tags.innerHTML = tagHtml + `<span class="s2-prev-tag">共${mainSheet.rows}行</span>`;
        const mdFlow = prefersMarkdownFlow();
        if (data.download_url && !mdFlow) {
          tags.innerHTML += ` <button type="button" class="s2-result-btn" style="margin-top:8px;font-size:12px;" onclick="editStep3Revision()">预览/编辑萃取底稿</button>`;
        }
        if (data.markdown_file) {
          tags.innerHTML += ` <button type="button" class="s2-result-btn" style="margin-top:8px;font-size:12px;" onclick="previewStep5File('${escapeHtml(data.markdown_file)}','Step2 萃取 Markdown 预览')">预览/编辑 Markdown</button>`;
        }
      }
    } else {
      card.style.display = 'none';
      empty.style.display = '';
    }
  } catch (e) {
    console.error('加载Step2输出失败:', e);
  }
}

function renderStep3ChatHistory() {
  const box = document.getElementById('s3-chat-history');
  if (!box) return;
  if (!Array.isArray(_alignChatHistory) || _alignChatHistory.length === 0) {
    box.innerHTML = '<div class="align-chat-empty">暂无对话。请先输入一条修订意见并发送。</div>';
    return;
  }
  let html = '';
  _alignChatHistory.forEach((item) => {
    const role = item?.role === 'assistant' ? 'assistant' : 'user';
    const roleText = role === 'assistant' ? '模型' : '专家';
    const ts = item?.ts ? String(item.ts) : '';
    html += `<div class="align-chat-item role-${role}">`;
    html += `<div class="align-chat-meta"><span class="align-chat-role">${roleText}</span><span>${escapeHtml(ts)}</span></div>`;
    html += `<div class="align-chat-content">${escapeHtml(String(item?.content || ''))}</div>`;
    html += '</div>';
  });
  box.innerHTML = html;
  box.scrollTop = box.scrollHeight;
}

/* ===== 修订上下文侧栏 ===== */
async function loadStep3RevisionContext() {
  var ctxEl = document.getElementById('s3-revision-context');
  var bodyEl = document.getElementById('s3-revision-context-body');
  if (!ctxEl || !bodyEl) return;
  var pid = currentPipeline ? currentPipeline.id : null;
  if (!pid) { ctxEl.style.display = 'none'; return; }

  try {
    var resp = await fetch(API_BASE + '/api/step3/revision_context?pipeline_id=' + pid);
    var data = await resp.json();
    if (data.status !== 'ok') { ctxEl.style.display = 'none'; return; }
    var insights = data.insights || [];
    var warnings = data.warnings || [];
    if (!insights.length && !warnings.length) { ctxEl.style.display = 'none'; return; }

    var html = '';
    warnings.forEach(function(w) {
      html += '<div class="rc-warning"><div class="rc-title">' + (w.icon || '') + ' ' + escapeHtml(w.source) + '</div>';
      if (w.text) html += '<div class="rc-text">' + escapeHtml(w.text) + '</div>';
      if (w.details && w.details.length) {
        html += '<ul class="rc-list">';
        w.details.forEach(function(d) { html += '<li>' + escapeHtml(String(d)) + '</li>'; });
        html += '</ul>';
      }
      html += '</div>';
    });
    insights.forEach(function(i) {
      html += '<div class="rc-insight"><div class="rc-title">' + (i.icon || '') + ' ' + escapeHtml(i.source) + '</div>';
      html += '<div class="rc-text">' + escapeHtml(i.text) + '</div>';
      if (i.details && i.details.length) {
        html += '<ul class="rc-list">';
        i.details.forEach(function(d) { html += '<li>' + escapeHtml(String(d)) + '</li>'; });
        html += '</ul>';
      }
      html += '</div>';
    });
    bodyEl.innerHTML = html;
    ctxEl.style.display = '';
  } catch (e) { ctxEl.style.display = 'none'; }
}

// Phase 1: Generate alignment preview (AI suggestions only)
async function step3GeneratePreview() {
  const pid = currentPipeline ? currentPipeline.id : null;
  if (!pid) { alert('请先进入流水线'); return; }
  const expertTextEl = document.getElementById('s3-expert-text');
  const expertText = expertTextEl.value.trim();
  const expertFileEl = document.getElementById('s3-expert-file');
  const cachedExpertFile = currentPipeline?.step_data?.step3_cached_file || '';
  let finalMessage = expertText;

  if (!finalMessage && expertFileEl.files.length > 0) {
    finalMessage = await expertFileEl.files[0].text();
  }
  const style = document.getElementById('s3-revision-style').value;
  const model = document.getElementById('s3-model').value;

  const btn = document.getElementById('s3-revise-btn');
  btn._locked = true;
  btn.disabled = true;
  btn.classList.add('loading');
  btn.innerHTML = '修订中<span class="btn-estimate">· 通常 15-60s</span>';
  btn.innerHTML = '<span class="action-icon">&#9203;</span> 处理中...';
  renderLoading('s3-output');

  const fd = new FormData();
  fd.append('pipeline_id', pid);
  if (finalMessage) fd.append('message', finalMessage);
  fd.append('style', style);
  if (model) fd.append('model', model);
  if (expertFileEl && expertFileEl.files && expertFileEl.files.length > 0) {
    fd.append('expert_file', expertFileEl.files[0]);
  } else if (!finalMessage && cachedExpertFile) {
    fd.append('expert_cached_file', cachedExpertFile);
  }
  // #region agent log
  fetch('http://127.0.0.1:7620/ingest/373a0517-ca00-4a54-8abb-05f50a475736',{method:'POST',headers:{'Content-Type':'application/json','X-Debug-Session-Id':'5b6690'},body:JSON.stringify({sessionId:'5b6690',runId:'run-1',hypothesisId:'H1',location:'frontend/js/app.js:step3GeneratePreview:before_fetch',message:'step3 preview submit',data:{hasPid:!!pid,style:style||'',model:model||'',expertLen:(finalMessage||'').length,hasCached:!!cachedExpertFile},timestamp:Date.now()})}).catch(()=>{});
  // #endregion

  try {
    const resp = await fetch(API_BASE + '/api/step4/align_chat', { method: 'POST', body: fd });
    const result = await resp.json();
    // #region agent log
    fetch('http://127.0.0.1:7620/ingest/373a0517-ca00-4a54-8abb-05f50a475736',{method:'POST',headers:{'Content-Type':'application/json','X-Debug-Session-Id':'5b6690'},body:JSON.stringify({sessionId:'5b6690',runId:'run-1',hypothesisId:'H1',location:'frontend/js/app.js:step3GeneratePreview:after_fetch',message:'step3 preview response',data:{status:result?.status||'',notesCount:Array.isArray(result?.notes)?result.notes.length:-1,error:result?.error||''},timestamp:Date.now()})}).catch(()=>{});
    // #endregion

    if (result.status === 'ok') {
      if (Array.isArray(result.chat_history)) {
        _alignChatHistory = result.chat_history;
      } else {
        _alignChatHistory = _alignChatHistory || [];
        if (finalMessage) _alignChatHistory.push({ role: 'user', content: finalMessage });
        if (result.assistant_reply) _alignChatHistory.push({ role: 'assistant', content: result.assistant_reply });
      }
      renderStep3ChatHistory();
      if (expertTextEl) expertTextEl.value = '';
      _alignNotes = result.notes || [];
      if (!(_alignNotes.length > 0)) {
        const passThrough = result.auto_finalized || result.align_mode === 'pass_through' || result.no_opinion;
        if (passThrough) {
          await showStep3AlignComplete(result, { noRevision: true });
          renderOutput('s3-output', '<div class="s2-result-success"><div class="s2-result-header">' +
            escapeHtml(result.message || '已按预萃稿生成对齐稿（无修订）') + '</div></div>');
          return;
        }
        const msg = result.message || '未从专家意见/上传材料中解析出可执行的修订条目，请补充更明确的修改说明。';
        renderOutput('s3-output', '<div class="error-list"><div class="error-item">' + escapeHtml(msg) + '</div></div>');
        document.getElementById('s3-input-section').style.display = '';
        document.getElementById('s3-review-section').style.display = 'none';
        document.getElementById('s3-result-section').style.display = 'none';
        return;
      }

      _alignNoteStates = {};
      _alignEditedValues = {};
      _alignCurrentFilter = 'all';
      _alignNotes.forEach(n => { _alignNoteStates[n.id] = 'pending'; });

      const ruleHint = result.style_rule
        ? `风格 ${result.style_rule.mode} · 原始 ${result.style_rule.raw_count} 条 → 过滤后 ${result.style_rule.processed_count} 条`
        : '';
      renderOutput('s3-output', `<div class="s2-result-success"><div class="s2-result-header">AI 生成了 ${_alignNotes.length} 条对齐建议</div><div class="s2-result-meta">${escapeHtml(ruleHint)} · 请在下方逐条审核</div></div>`);

      document.getElementById('s3-input-section').style.display = 'none';
      document.getElementById('s3-review-section').style.display = '';
      document.getElementById('s3-result-section').style.display = 'none';

      renderAlignNotesList();
      updateAlignStats();
    } else {
      renderOutput('s3-output', '<div class="error-list"><div class="error-item">' + escapeHtml(result.error || '生成对齐建议失败') + '</div></div>');
    }
  } catch (e) {
    renderOutput('s3-output', '<div class="error-list"><div class="error-item">' + escapeHtml(e.message) + '</div></div>');
  } finally {
    btn.disabled = false;
    updateStep3AlignModeHint();
  }
}

function step3BackToInput() {
  document.getElementById('s3-input-section').style.display = '';
  document.getElementById('s3-review-section').style.display = 'none';
  document.getElementById('s3-result-section').style.display = 'none';
}

function renderAlignNotesList() {
  const container = document.getElementById('s3-notes-list');
  const countEl = document.getElementById('s3-review-count');
  if (countEl) countEl.textContent = `共 ${_alignNotes.length} 条`;

  let html = '';
  _alignNotes.forEach((n, i) => {
    if (_alignCurrentFilter !== 'all' && n.action !== _alignCurrentFilter) return;

    const state = _alignNoteStates[n.id] || 'pending';
    const stateClass = `note-${state}`;
    const actionTag = ALIGN_ACTION_LABELS[n.action] || n.action;
    const actionClass = `action-${n.action}`;
    const editedVal = _alignEditedValues[n.id];

    html += `<div class="align-note-card ${stateClass}" id="align-note-${n.id}">`;
    html += `<div class="align-note-header">`;
    html += `<span class="align-note-idx">#${i + 1}</span>`;
    html += `<span class="align-note-action-tag ${actionClass}">${escapeHtml(actionTag)}</span>`;
    html += `<span class="align-note-location">${escapeHtml(n.sheet || '')} · 行${n.row || '?'} · 列${n.col || '?'}</span>`;
    if (n.note) html += `<span class="align-note-reason" title="${escapeHtml(n.note)}">${escapeHtml(n.note)}</span>`;
    html += `</div>`;

    // Diff display
    html += `<div class="align-note-diff">`;
    if (n.action === 'add' || n.action === 'supplement') {
      html += `<div class="align-diff-add">${escapeHtml(editedVal || n.new_value || '')}</div>`;
    } else if (n.action === 'delete') {
      html += `<div class="align-diff-old">${escapeHtml(n.old_value || '(原值)')}</div>`;
    } else {
      html += `<div class="align-diff-old">${escapeHtml(n.old_value || '(原值)')}</div>`;
      html += `<div class="align-diff-new">${escapeHtml(editedVal || n.new_value || '')}</div>`;
    }
    html += `</div>`;

    // Action buttons
    html += `<div class="align-note-actions">`;
    html += `<button class="align-note-btn align-note-btn-accept ${state === 'accepted' || state === 'edited' ? 'active' : ''}" onclick="alignSetState(${n.id}, 'accepted')">采纳</button>`;
    html += `<button class="align-note-btn align-note-btn-reject ${state === 'rejected' ? 'active' : ''}" onclick="alignSetState(${n.id}, 'rejected')">驳回</button>`;
    if (n.action !== 'delete') {
      html += `<button class="align-note-btn align-note-btn-edit" onclick="alignToggleEdit(${n.id})">编辑</button>`;
    }
    html += `</div>`;

    // Inline editor
    if (n.action !== 'delete') {
      html += `<div class="align-inline-editor" id="align-editor-${n.id}">`;
      html += `<textarea class="align-inline-textarea" id="align-textarea-${n.id}" placeholder="修改后的内容">${escapeHtml(editedVal || n.new_value || '')}</textarea>`;
      html += `<button class="align-inline-save" onclick="alignSaveEdit(${n.id})">保存修改</button>`;
      html += `</div>`;
    }

    html += `</div>`;
  });

  if (!html) {
    html = '<div style="text-align:center;padding:40px;color:var(--text-muted);">无匹配的对齐建议</div>';
  }
  container.innerHTML = html;
}

function alignSetState(id, state) {
  const current = _alignNoteStates[id];
  if (current === state) {
    _alignNoteStates[id] = 'pending';
  } else {
    _alignNoteStates[id] = state;
  }
  renderAlignNotesList();
  updateAlignStats();

  // 当专家「采纳」或「编辑」修订建议时，弹出隐性注释追问卡片
  if (state === 'accepted' || state === 'edited') {
    var note = (_alignNotes || []).find(function (n) { return n.id === id; });
    if (note) {
      setTimeout(function () {
        var noteEl = document.getElementById('align-note-' + id);
        if (noteEl) showTacitFollowup(noteEl, id, note.action || 'modify');
      }, 200);
    }
  }
}

function alignToggleEdit(id) {
  const editor = document.getElementById('align-editor-' + id);
  if (!editor) return;
  editor.classList.toggle('visible');
  if (editor.classList.contains('visible')) {
    const ta = document.getElementById('align-textarea-' + id);
    if (ta) ta.focus();
  }
}

function alignSaveEdit(id) {
  const ta = document.getElementById('align-textarea-' + id);
  if (!ta) return;
  const val = ta.value.trim();
  if (!val) return;
  _alignEditedValues[id] = val;
  _alignNoteStates[id] = 'edited';
  renderAlignNotesList();
  updateAlignStats();
}

function alignFilterNotes(filter) {
  _alignCurrentFilter = filter;
  document.querySelectorAll('.align-filter-btn').forEach(btn => {
    btn.classList.toggle('active', btn.dataset.filter === filter);
  });
  renderAlignNotesList();
}

function alignBatchAcceptAll() {
  _alignNotes.forEach(n => {
    if (_alignCurrentFilter === 'all' || n.action === _alignCurrentFilter) {
      if (_alignNoteStates[n.id] !== 'edited') _alignNoteStates[n.id] = 'accepted';
    }
  });
  renderAlignNotesList();
  updateAlignStats();
}

function alignBatchRejectAll() {
  _alignNotes.forEach(n => {
    if (_alignCurrentFilter === 'all' || n.action === _alignCurrentFilter) {
      _alignNoteStates[n.id] = 'rejected';
    }
  });
  renderAlignNotesList();
  updateAlignStats();
}

function updateAlignStats() {
  let accepted = 0, rejected = 0, pending = 0, edited = 0;
  for (const [id, st] of Object.entries(_alignNoteStates)) {
    if (st === 'accepted') accepted++;
    else if (st === 'rejected') rejected++;
    else if (st === 'edited') { edited++; accepted++; }
    else pending++;
  }
  const total = _alignNotes.length;
  const processed = accepted + rejected;
  document.getElementById('s3-accepted-count').textContent = accepted;
  document.getElementById('s3-rejected-count').textContent = rejected;
  document.getElementById('s3-pending-count').textContent = pending;
  document.getElementById('s3-edited-count').textContent = edited;

  // Update toolbar progress
  var reviewTitle = document.getElementById('s3-review-title');
  if (reviewTitle) reviewTitle.textContent = `审核进度 ${processed}/${total}`;
  var reviewCount = document.getElementById('s3-review-count');
  if (reviewCount) reviewCount.textContent = `共 ${total} 条`;

  const applyBtn = document.getElementById('s3-apply-btn');
  if (applyBtn) applyBtn.disabled = (accepted === 0);
}

async function showStep3AlignComplete(result, options) {
  const opts = options || {};
  clearDownstreamOutputs(3);
  const dlName = result.download_name || result.output_file || '';
  const dlUrl = result.download_url
    ? (result.download_url.startsWith('http') ? result.download_url : API_BASE + result.download_url)
    : '';
  const mdName = result.markdown_file || '';
  const mdUrl = result.markdown_download_url
    ? (result.markdown_download_url.startsWith('http') ? result.markdown_download_url : API_BASE + result.markdown_download_url)
    : '';

  if (currentPipeline && dlName && isStep4FinalFile(dlName)) {
    currentPipeline.step_data = currentPipeline.step_data || {};
    currentPipeline.step_data.step4_final_file = dlName;
    currentPipeline.step_data.step4_download_url = result.download_url || ('/downloads/' + dlName);
    currentPipeline.step_data.step4_md_file = mdName;
    currentPipeline.step_data.step4_md_download_url = result.markdown_download_url || '';
    currentPipeline.step_data.step4_final_count = result.revision_count || 0;
    await persistPipeline({
      step4_final_file: dlName,
      step4_download_url: currentPipeline.step_data.step4_download_url,
      step4_md_file: mdName,
      step4_md_download_url: currentPipeline.step_data.step4_md_download_url,
      step4_final_count: result.revision_count || 0,
    });
    await refreshCurrentPipeline();
  }

  document.getElementById('s3-input-section').style.display = 'none';
  document.getElementById('s3-review-section').style.display = 'none';
  document.getElementById('s3-result-section').style.display = '';

  const mdFlow = prefersMarkdownFlow();
  let html = '';
  if (opts.noRevision) {
    html += '<div class="align-result-header"><span>&#10003;</span> 已确认（无修订）</div>';
    html += '<div class="align-result-meta">' + escapeHtml(result.message || '当前稿已作为对齐稿') + '</div>';
  } else {
    html += '<div class="align-result-header"><span>&#10003;</span> 知识对齐完成</div>';
    html += '<div class="align-result-meta">采纳 <strong>' + (result.accepted_count || 0) + '</strong> / ' +
      (result.total_suggested || _alignNotes.length) + ' 条建议 · 共处理 <strong>' +
      (result.revision_count || 0) + '</strong> 处修订</div>';
  }
  html += '<div class="align-result-actions">';
  if (mdUrl) html += '<a href="' + escapeHtml(mdUrl) + '" class="s2-result-btn s2-result-btn-primary" download>下载对齐稿 Markdown</a>';
  if (dlUrl && !mdFlow) html += '<a href="' + escapeHtml(dlUrl) + '" class="s2-result-btn s2-result-btn-primary" download>下载对齐稿 Excel</a>';
  if (mdName) html += '<button class="s2-result-btn" onclick="previewStep5File(\'' + escapeHtml(mdName) + '\',\'Step3 对齐 Markdown 预览\')">预览/编辑 Markdown</button>';
  if (!mdFlow && !opts.noRevision) html += '<button class="s2-result-btn" onclick="editStep3Revision()">&#9998; 在线编辑对齐稿</button>';
  html += '<button class="s2-result-btn" onclick="step3BackToInput()">重新对齐</button>';
  html += '</div>';
  document.getElementById('s3-result-card').innerHTML = html;
  await markStepDone(3);
}

// Phase 2: Apply selected notes to generate final
async function step3ApplyNotes() {
  const pid = currentPipeline ? currentPipeline.id : null;
  if (!pid) { alert('请先进入流水线'); return; }

  const acceptedIds = [];
  const editedNotes = [];
  for (const [id, st] of Object.entries(_alignNoteStates)) {
    const nid = parseInt(id);
    if (st === 'accepted' || st === 'edited') {
      acceptedIds.push(nid);
      if (st === 'edited' && _alignEditedValues[nid] !== undefined) {
        editedNotes.push({ id: nid, new_value: _alignEditedValues[nid] });
      }
    }
  }

  if (acceptedIds.length === 0) { alert('请至少采纳一条对齐建议'); return; }
  // #region agent log
  fetch('http://127.0.0.1:7620/ingest/373a0517-ca00-4a54-8abb-05f50a475736',{method:'POST',headers:{'Content-Type':'application/json','X-Debug-Session-Id':'5b6690'},body:JSON.stringify({sessionId:'5b6690',runId:'run-1',hypothesisId:'H3',location:'frontend/js/app.js:step3ApplyNotes:before_fetch',message:'step3 apply submit',data:{acceptedCount:acceptedIds.length,editedCount:editedNotes.length,stateSummary:Object.values(_alignNoteStates||{}).reduce((m,s)=>{m[s]=(m[s]||0)+1;return m;},{} )},timestamp:Date.now()})}).catch(()=>{});
  // #endregion

  const btn = document.getElementById('s3-apply-btn');
  btn.disabled = true;
  btn.innerHTML = '<span class="action-icon">&#9203;</span> 生成中...';

  try {
    const resp = await fetch(API_BASE + '/api/step4/apply_notes', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ pipeline_id: pid, accepted_ids: acceptedIds, edited_notes: editedNotes, tacit_annotations: getTacitAnnotationsPayload() }),
    });
    const result = await resp.json();
    // #region agent log
    fetch('http://127.0.0.1:7620/ingest/373a0517-ca00-4a54-8abb-05f50a475736',{method:'POST',headers:{'Content-Type':'application/json','X-Debug-Session-Id':'5b6690'},body:JSON.stringify({sessionId:'5b6690',runId:'run-1',hypothesisId:'H3',location:'frontend/js/app.js:step3ApplyNotes:after_fetch',message:'step3 apply response',data:{status:result?.status||'',revisionCount:result?.revision_count??null,error:result?.error||''},timestamp:Date.now()})}).catch(()=>{});
    // #endregion

    if (result.status === 'ok') {
      result.accepted_count = result.accepted_count || acceptedIds.length;
      result.total_suggested = result.total_suggested || _alignNotes.length;
      await showStep3AlignComplete(result, { noRevision: false });
    } else {
      alert(result.error || '生成对齐稿失败');
    }
  } catch (e) {
    alert('生成对齐稿出错: ' + e.message);
  } finally {
    btn.disabled = false;
    btn.innerHTML = '<span class="action-icon">&#10003;</span> 确认并生成对齐稿';
  }
}

async function step3ConfirmAsIs() {
  const pid = currentPipeline ? currentPipeline.id : null;
  if (!pid) { alert('请先进入流水线'); return; }
  try {
    const resp = await fetch(API_BASE + '/api/step4/confirm_as_is', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ pipeline_id: pid }),
    });
    const result = await resp.json();
    if (result.status !== 'ok') {
      alert(result.error || '确认失败');
      return;
    }
    await showStep3AlignComplete(result, { noRevision: true });
    renderOutput('s3-output', '<div class="s2-result-success"><div class="s2-result-header">' +
      escapeHtml(result.message || '已确认当前稿为对齐稿') + '</div></div>');
  } catch (e) {
    alert('确认对齐稿出错: ' + e.message);
  }
}

// ═══════════════════════════════════════════════════════════════════
// Step 5: Compile & Quality & AI Refine
// ═══════════════════════════════════════════════════════════════════

async function loadStep5PrevOutput() {
  if (!currentPipeline) return;
  const el = document.getElementById('s5-prev-draft');
  const emptyEl = document.getElementById('s5-prev-empty');
  if (!el) return;
  try {
    const resp = await fetch(API_BASE + '/api/step5/prev_output?pipeline_id=' + currentPipeline.id);
    const data = await resp.json();
    if (data.has_output) {
      el.style.display = 'block';
      if (emptyEl) emptyEl.style.display = 'none';
      const infoNameEl = document.getElementById('s5-prev-name');
      const infoMetaEl = document.getElementById('s5-prev-meta');
      const tagsEl = document.getElementById('s5-prev-tags');
      const mdFlow = prefersMarkdownFlow();
      if (infoNameEl) infoNameEl.textContent = (mdFlow && data.markdown_file) ? data.markdown_file : (data.file_name || '知识对齐稿');
      if (infoMetaEl) infoMetaEl.textContent = (data.scenario || '');
      if (tagsEl) {
        let tags = '';
        if (data.fields_info) {
          data.fields_info.forEach(s => {
            tags += '<span class="s2-prev-tag">' + escapeHtml(s.sheet) + ' (' + s.data_rows + '行)</span>';
          });
        }
        if (data.markdown_file) {
          tags += '<button type="button" class="s2-result-btn" style="margin-top:8px;font-size:12px;" onclick="previewStep5File(\'' + escapeHtml(data.markdown_file) + '\', \'Step3 对齐 Markdown 预览\')">预览/编辑 Markdown</button>';
        }
        if (!mdFlow && data.download_url) {
          tags += '<a class="s2-result-btn" style="margin-top:8px;font-size:12px;" href="' + API_BASE + data.download_url + '" download>下载 Excel</a>';
        }
        tagsEl.innerHTML = tags;
      }
    } else {
      el.style.display = 'none';
      if (emptyEl) emptyEl.style.display = 'block';
    }
  } catch (e) { console.error('loadStep5PrevOutput error', e); }
}

function getSelectedStep5Formats() {
  const formats = [];
  if (document.getElementById('s5-fmt-cot')?.checked) formats.push('cot');
  if (document.getElementById('s5-fmt-qa')?.checked) formats.push('qa');
  if (document.getElementById('s5-fmt-skill')?.checked) formats.push('skill');
  return formats;
}

function renderStep5ArtifactCard(key, title, desc, countLabel, downloads, previewFn) {
  let html = '<div class="s5-artifact-card">';
  html += '<div class="s5-artifact-head"><div class="s5-artifact-title">' + escapeHtml(title) + '</div>';
  if (countLabel) html += '<span class="s5-artifact-badge">' + escapeHtml(countLabel) + '</span>';
  html += '</div>';
  html += '<div class="s5-artifact-desc">' + escapeHtml(desc) + '</div>';
  html += '<div class="s5-artifact-actions">';
  (downloads || []).forEach(d => {
    if (d.url) {
      html += '<a class="s5-btn s5-btn-outline" href="' + escapeHtml(API_BASE + d.url) + '" download>' + escapeHtml(d.label) + '</a>';
    }
  });
  if (previewFn) {
    html += '<button type="button" class="s5-btn s5-btn-outline" onclick="' + previewFn + '">&#128065; 预览</button>';
  }
  html += '</div></div>';
  return html;
}

async function step5FreshnessAudit() {
  var btn = document.getElementById('s5-freshness-btn');
  if (!btn || btn._locked || !currentPipeline) return;
  btn._locked = true;
  btn.disabled = true;
  btn.innerHTML = '🔄 审计中...';
  var fd = new FormData();
  fd.append('skill_id', 'knowledge-freshness-audit');
  fd.append('pipeline_id', currentPipeline.id);
  fd.append('model', resolveModelName('s5-model') || (allModels.length ? allModels[0].name : ''));
  try {
    var resp = await fetch(API_BASE + '/api/skills/execute', { method: 'POST', body: fd });
    var result = await resp.json();
    var s5out = document.getElementById('s5-output');
    if (result.status === 'ok') {
      var dl = result.download_url || '';
      s5out.innerHTML =
        '<div class="s2-result-success"><div class="s2-result-header">保鲜度审计完成</div>' +
        '<div class="s2-result-meta">共 ' + (result.total_items || 0) + ' 条知识 · 高置信度占比 ' + (result.high_confidence_pct || 0) + '%</div>' +
        (result.stale_indicators && result.stale_indicators.length ?
          '<div style="margin-top:8px">' + result.stale_indicators.map(function(s){return '<div style="font-size:12px;color:#8b6914;margin:2px 0">⚠️ '+escapeHtml(s)+'</div>';}).join('') + '</div>' : '') +
        (dl ? '<div class="s2-result-actions" style="margin-top:12px"><a class="s2-result-btn s2-result-btn-primary" href="'+API_BASE+dl+'" download>📥 下载审计报告</a></div>' : '') +
        '</div>';
    } else {
      s5out.innerHTML = '<div class="error-list"><div class="error-item">' + escapeHtml(result.error || '审计失败') + '</div></div>';
    }
  } catch (e) {
    document.getElementById('s5-output').innerHTML = '<div class="error-list"><div class="error-item">' + escapeHtml(e.message) + '</div></div>';
  } finally {
    btn._locked = false;
    btn.disabled = false;
    btn.innerHTML = '🔄 知识保鲜度审计';
  }
}

async function step5Compile() {
  if (!currentPipeline) { alert('请先进入流水线'); return; }
  const formats = getSelectedStep5Formats();
  if (!formats.length) { alert('请至少选择一种交付物类型'); return; }
  // #region agent log
  fetch('http://127.0.0.1:7620/ingest/373a0517-ca00-4a54-8abb-05f50a475736',{method:'POST',headers:{'Content-Type':'application/json','X-Debug-Session-Id':'5b6690'},body:JSON.stringify({sessionId:'5b6690',runId:'run-2',hypothesisId:'H6',location:'frontend/js/app.js:step5Compile:before_call',message:'step5 compile submit',data:{pipelineId:currentPipeline?.id||'',formats,step4Final:currentPipeline?.step_data?.step4_final_file||'',hasStepData:!!currentPipeline?.step_data},timestamp:Date.now()})}).catch(()=>{});
  // #endregion

  renderLoading('s5-output');
  try {
    const fd = new FormData();
    fd.append('pipeline_id', currentPipeline.id);
    fd.append('formats', formats.join(','));
    const result = await apiCall('/api/step5/compile', fd);
    let html = '';
    if (result.status === 'ok') {
      const km = result.knowledge_count || 0;
      const metrics = result.quality_metrics || {};
      const dl = result.artifacts_download || {};

      if (currentPipeline) {
        currentPipeline.step_data = currentPipeline.step_data || {};
        if (dl.skill?.file_name) {
          currentPipeline.step_data.step5_skill_file = dl.skill.file_name;
          currentPipeline.step_data.step5_download_url = dl.skill.download_url;
        }
        if (dl.cot?.file_name) {
          currentPipeline.step_data.step5_cot_file = dl.cot.file_name;
          currentPipeline.step_data.step5_cot_download_url = dl.cot.download_url;
        }
        if (dl.qa?.file_name) {
          currentPipeline.step_data.step5_qa_file = dl.qa.file_name;
          currentPipeline.step_data.step5_qa_download_url = dl.qa.download_url;
        }
        if (dl.qa_md?.file_name) {
          currentPipeline.step_data.step5_qa_md_file = dl.qa_md.file_name;
          currentPipeline.step_data.step5_qa_md_download_url = dl.qa_md.download_url;
        }
        if (dl.openclaw_manifest?.file_name) {
          currentPipeline.step_data.step5_openclaw_manifest_file = dl.openclaw_manifest.file_name;
          currentPipeline.step_data.step5_openclaw_manifest_url = dl.openclaw_manifest.download_url;
        }
        await persistPipeline({
          step5_skill_file: currentPipeline.step_data.step5_skill_file,
          step5_download_url: currentPipeline.step_data.step5_download_url,
          step5_cot_file: currentPipeline.step_data.step5_cot_file,
          step5_cot_download_url: currentPipeline.step_data.step5_cot_download_url,
          step5_qa_file: currentPipeline.step_data.step5_qa_file,
          step5_qa_download_url: currentPipeline.step_data.step5_qa_download_url,
          step5_qa_md_file: currentPipeline.step_data.step5_qa_md_file,
          step5_qa_md_download_url: currentPipeline.step_data.step5_qa_md_download_url,
          step5_openclaw_manifest_file: currentPipeline.step_data.step5_openclaw_manifest_file,
          step5_openclaw_manifest_url: currentPipeline.step_data.step5_openclaw_manifest_url,
        });
      }

      html += '<div class="s5-compile-result">';
      html += '<div class="s5-compile-header">';
      html += '<div class="s5-compile-icon">&#9881;</div>';
      html += '<div class="s5-compile-title">智能转化完成</div>';
      html += '<div class="s5-compile-subtitle">已生成 ' + formats.length + ' 类交付物 · 共 ' + km + ' 条知识</div>';
      html += '</div>';

      html += '<div class="s5-artifact-grid">';
      if (formats.includes('cot') && dl.cot) {
        html += renderStep5ArtifactCard(
          'cot', '思维链', '情境识别 → 推理步骤 → 结论校验，适合培训与推理复现',
          (dl.cot.count != null ? dl.cot.count + ' 条' : ''),
          [{ label: '下载 .md', url: dl.cot.download_url }],
          'previewStep5Cot()'
        );
      }
      if (formats.includes('qa') && dl.qa) {
        const qaDownloads = [{ label: '下载 JSON', url: dl.qa.download_url }];
        if (dl.qa_md?.download_url) qaDownloads.push({ label: '下载 .md', url: dl.qa_md.download_url });
        html += renderStep5ArtifactCard(
          'qa', 'QA 对', '问答对格式，可用于 RAG、评测集或微调样本',
          (dl.qa.count != null ? dl.qa.count + ' 组' : ''),
          qaDownloads,
          'previewStep5Qa()'
        );
      }
      if (formats.includes('skill') && dl.skill) {
        const skillDownloads = [{ label: '下载 SKILL.md', url: dl.skill.download_url }];
        if (dl.openclaw_manifest?.download_url) {
          skillDownloads.push({ label: 'OpenClaw 清单', url: dl.openclaw_manifest.download_url });
        }
        html += renderStep5ArtifactCard(
          'skill', 'Skill（OpenClaw）', '含 openclaw.skill.json，可接入 OpenClaw Agent 技能目录',
          'OpenClaw 兼容',
          skillDownloads,
          'previewStep5Skill()'
        );
      }
      html += '</div>';

      if (metrics && Object.keys(metrics).length) {
        const confDist = metrics.confidence_distribution || {};
        html += '<div class="s5-section"><div class="s5-section-title">质量摘要</div>';
        html += '<div class="s5-stats-grid">';
        html += '<div class="s5-stat-card"><div class="s5-stat-value">' + km + '</div><div class="s5-stat-label">知识条目</div></div>';
        html += '<div class="s5-stat-card"><div class="s5-stat-value">' + (metrics.category_count || 0) + '</div><div class="s5-stat-label">分类数</div></div>';
        html += '<div class="s5-stat-card"><div class="s5-stat-value green">' + (confDist.high || 0) + '</div><div class="s5-stat-label">高置信度</div></div>';
        html += '<div class="s5-stat-card"><div class="s5-stat-value">' + (metrics.anti_pattern_count || 0) + '</div><div class="s5-stat-label">反模式</div></div>';
        html += '</div></div>';
      }

      html += '</div>';
      markStepDone(4);
    } else {
      html += '<div class="error-list"><div class="error-item">' + escapeHtml(result.error || result.raw || '智能转化失败') + '</div></div>';
    }
    renderOutput('s5-output', html);
  } catch (e) {
    renderOutput('s5-output', '<div class="error-list"><div class="error-item">' + escapeHtml(e.message) + '</div></div>');
  }
}

async function previewStep5File(fileName, title) {
  if (!fileName) { alert('暂无可预览文件'); return; }
  try {
    const resp = await fetch(API_BASE + '/api/files/read?file_name=' + encodeURIComponent(fileName));
    const data = await resp.json();
    if (data.status === 'ok') {
      openMarkdownEditor(fileName, data.content);
      document.getElementById('markdown-editor-title').textContent = title || fileName;
    } else {
      alert('加载失败: ' + (data.error || '未知错误'));
    }
  } catch (e) {
    alert('加载失败: ' + e.message);
  }
}

function previewStep5Cot() {
  previewStep5File(currentPipeline?.step_data?.step5_cot_file, '思维链预览');
}
function previewStep5Qa() {
  const sd = currentPipeline?.step_data || {};
  previewStep5File(sd.step5_qa_md_file || sd.step5_qa_file, 'QA 对预览');
}
function previewStep5Skill() {
  previewStep5File(currentPipeline?.step_data?.step5_skill_file, 'Skill 预览');
}

async function step5Quality() {
  if (!currentPipeline) { alert('请先进入流水线'); return; }
  renderLoading('s5-output');
  try {
    const fd = new FormData();
    fd.append('pipeline_id', currentPipeline.id);
    const result = await apiCall('/api/step5/quality', fd);
    let html = '';
    if (result.overall_score !== undefined) {
      const grade = result.grade || '-';
      const gradeCls = grade === 'A' ? 'grade-a' : grade === 'B' ? 'grade-b' : grade === 'C' ? 'grade-c' : 'grade-d';

      html += '<div class="s5-compile-result">';
      html += '<div class="s5-compile-header">';
      html += '<div class="s5-compile-icon">&#11088;</div>';
      html += '<div class="s5-compile-title">质量报告</div>';
      html += '<div class="s5-compile-subtitle">五维度质量评估</div>';
      html += '</div>';

      html += '<div class="s5-stats-grid">';
      html += '<div class="s5-stat-card"><div class="s5-stat-value">' + result.overall_score + '</div><div class="s5-stat-label">综合评分</div></div>';
      html += '<div class="s5-stat-card"><div class="s5-stat-value"><span class="grade-badge ' + gradeCls + '">' + grade + '</span></div><div class="s5-stat-label">质量等级</div></div>';
      html += '</div>';

      if (result.dimensions) {
        html += '<div class="s5-section"><div class="s5-section-title">维度详情</div>';
        html += '<div class="s5-coverage-grid">';
        Object.entries(result.dimensions).forEach(([k, v]) => {
          const label = { completeness: '完整性', accuracy: '准确性', actionability: '可操作性', antipattern: '反模式覆盖', traceability: '来源可溯' }[k] || k;
          const pct = typeof v === 'object' ? (v.score || v.value || 0) : v;
          const cls = pct >= 80 ? 'green' : pct >= 60 ? 'orange' : 'red';
          html += '<div class="s5-coverage-item"><div class="s5-coverage-info"><div class="s5-coverage-label">' + label + '</div>';
          html += '<div class="s5-coverage-bar"><div class="s5-coverage-fill ' + cls + '" style="width:' + pct + '%"></div></div></div>';
          html += '<div class="s5-coverage-value">' + pct + '%</div></div>';
        });
        html += '</div></div>';
      }

      if (result.download_url) {
        html += '<div class="s5-actions">';
        html += '<a class="s5-btn s5-btn-primary" href="' + API_BASE + result.download_url + '" download>&#11015; 下载质量报告</a>';
        html += '</div>';
      }
      html += '</div>';
    } else {
      html += '<pre style="font-size:12px;overflow:auto;max-height:300px">' + escapeHtml(JSON.stringify(result, null, 2)) + '</pre>';
    }
    renderOutput('s5-output', html);
  } catch (e) {
    renderOutput('s5-output', '<div class="error-list"><div class="error-item">' + escapeHtml(e.message) + '</div></div>');
  }
}

// Markdown Editor for SKILL.md
function openMarkdownEditor(fileName, content) {
  const modal = document.getElementById('markdown-editor-modal');
  if (!modal) return;
  document.getElementById('markdown-editor-title').textContent = fileName;
  document.getElementById('markdown-editor-content').value = content;
  renderMarkdownPreview(content);
  modal.classList.add('active');
}

function closeMarkdownEditor() {
  const modal = document.getElementById('markdown-editor-modal');
  if (modal) modal.classList.remove('active');
}

function renderMarkdownPreview(content) {
  const preview = document.getElementById('markdown-preview');
  if (!preview) return;
  // Simple markdown rendering
  let html = escapeHtml(content);
  // Headers
  html = html.replace(/^### (.+)$/gm, '<h3>$1</h3>');
  html = html.replace(/^## (.+)$/gm, '<h2>$1</h2>');
  html = html.replace(/^# (.+)$/gm, '<h1>$1</h1>');
  // Bold
  html = html.replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>');
  // Italic
  html = html.replace(/\*(.+?)\*/g, '<em>$1</em>');
  // Code blocks
  html = html.replace(/```([\s\S]*?)```/g, '<pre><code>$1</code></pre>');
  // Inline code
  html = html.replace(/`(.+?)`/g, '<code>$1</code>');
  // Lists
  html = html.replace(/^- (.+)$/gm, '<li>$1</li>');
  html = html.replace(/(<li>.*<\/li>)/s, '<ul>$1</ul>');
  // Line breaks
  html = html.replace(/\n\n/g, '</p><p>');
  html = html.replace(/\n/g, '<br>');
  html = '<p>' + html + '</p>';
  preview.innerHTML = html;
}

function onMarkdownEditorInput() {
  const content = document.getElementById('markdown-editor-content').value;
  renderMarkdownPreview(content);
}

function copyMarkdownContent() {
  const content = document.getElementById('markdown-editor-content').value;
  navigator.clipboard.writeText(content).then(() => {
    const btn = document.querySelector('.md-btn-copy');
    if (btn) {
      const orig = btn.innerHTML;
      btn.innerHTML = '&#10003; 已复制';
      setTimeout(() => { btn.innerHTML = orig; }, 1500);
    }
  });
}

async function saveMarkdownContent() {
  const fileName = document.getElementById('markdown-editor-title').textContent;
  const content = document.getElementById('markdown-editor-content').value;
  const saveBtn = document.getElementById('md-editor-save-btn');

  if (saveBtn) { saveBtn.disabled = true; saveBtn.innerHTML = '&#9203; 保存中...'; }

  try {
    const resp = await fetch(API_BASE + '/api/files/save', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ file_name: fileName, content: content })
    });
    const data = await resp.json();
    if (data.status === 'ok') {
      if (saveBtn) { saveBtn.innerHTML = '&#10003; 已保存'; }
      showToast('文件已保存');
      setTimeout(() => { if (saveBtn) saveBtn.innerHTML = '&#128190; 保存'; }, 2000);
    } else {
      alert('保存失败: ' + (data.error || '未知错误'));
      if (saveBtn) saveBtn.innerHTML = '&#128190; 保存';
    }
  } catch (e) {
    alert('保存失败: ' + e.message);
    if (saveBtn) saveBtn.innerHTML = '&#128190; 保存';
  } finally {
    if (saveBtn) saveBtn.disabled = false;
  }
}

function downloadMarkdownContent() {
  const fileName = document.getElementById('markdown-editor-title').textContent;
  const content = document.getElementById('markdown-editor-content').value;
  const blob = new Blob([content], { type: 'text/markdown;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = fileName;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

// ═══════════════════════════════════════════════════════════════════
// Shared Validation Renderer
// ═══════════════════════════════════════════════════════════════════

// ═══════════════════════════════════════════════════════════════════
// Pipeline Management
// ═══════════════════════════════════════════════════════════════════

async function loadPipelineOverview() {
  const container = document.getElementById('pipeline-list-container');
  if (!container) return;

  container.innerHTML = '<div class="pipeline-overview"><div style="text-align:center;padding:40px"><div class="spinner"></div>加载中...</div></div>';

  try {
    const resp = await fetch(API_BASE + '/api/pipelines');
    const data = await resp.json();
    if (data.status !== 'ok') throw new Error(data.error || '加载失败');

    const pipelines = data.pipelines || [];

    let html = '<div class="pipeline-overview">';

    // Banner
    html += '<div class="overview-banner">';
    html += '<div class="overview-banner-inner">';
    html += '<div class="overview-banner-text">';
    html += '<div class="overview-banner-title">隐性知识显性化 · 四步法萃取流水线</div>';
    html += '<div class="overview-banner-desc">将领域专家的隐性经验系统性显性化为 AI 可加载的结构化知识，通过四步法流水线，从场景定义到智能转化，层层递进、步步可追溯。</div>';
    html += '</div>';
    html += '<div class="overview-banner-action" onclick="showNewPipelineForm()">';
    html += '<span class="banner-action-icon">+</span>';
    html += '<span>新建流水线</span>';
    html += '</div>';
    html += '</div>';
    html += '</div>';

    // Steps Roadmap - vertical layout with timeline
    const stepCards = [
      { num: '01', title: '场景锚定', desc: '定义知识模板骨架，确定领域边界与字段规范，生成场景配置文件', icon: '&#9776;' },
      { num: '02', title: '知识萃取', desc: '从已有文档中提取知识条目，AI 辅助生成待审稿', icon: '&#9997;' },
      { num: '03', title: '知识对齐', desc: '融合修订与确认，完成专家意见对齐并生成最终可发布稿', icon: '&#10003;' },
      { num: '04', title: '智能转化', desc: '生成思维链、QA 对、OpenClaw Skill 三类交付物', icon: '&#9881;' },
    ];
    html += '<div class="overview-roadmap">';
    html += '<div class="roadmap-header"><span class="roadmap-header-line"></span><span class="roadmap-header-text">四步法流程概览</span><span class="roadmap-header-line"></span></div>';
    html += '<div class="roadmap-cards">';
    stepCards.forEach((s, i) => {
      html += '<div class="roadmap-step" data-step="' + s.num + '">';
      html += '<div class="roadmap-step-track">';
      html += '<div class="roadmap-step-node">' + s.icon + '</div>';
      if (i < stepCards.length - 1) html += '<div class="roadmap-step-connector"></div>';
      html += '</div>';
      html += '<div class="roadmap-step-body">';
      html += '<div class="roadmap-step-num">STEP ' + s.num + '</div>';
      html += '<div class="roadmap-step-title">' + s.title + '</div>';
      html += '<div class="roadmap-step-desc">' + s.desc + '</div>';
      html += '</div>';
      html += '</div>';
    });
    html += '</div>';
    html += '</div>';

    // New Pipeline Form (hidden)
    html += '<div class="new-pipeline-form" id="new-pipeline-form" style="display:none">';
    html += '<h3>新建流水线</h3>';
    html += '<div class="form-row">';
    html += '<div class="form-group half"><label>流水线名称</label><input type="text" id="np-name" placeholder="如：信贷审批知识萃取"></div>';
    html += '<div class="form-group half"><label>场景名称</label>';
    html += '<select id="np-scenario" onchange="if(this.value===\'custom\'){document.getElementById(\'np-scenario-custom\').classList.remove(\'hidden\')}else{document.getElementById(\'np-scenario-custom\').classList.add(\'hidden\')}">';
    html += '<option value="">-- 选择预设场景 --</option>';
    html += '<option value="信贷审批">信贷审批</option>';
    html += '<option value="风控">风控</option>';
    html += '<option value="营销">营销</option>';
    html += '<option value="custom">自定义</option>';
    html += '</select>';
    html += '<input type="text" id="np-scenario-custom" class="hidden" placeholder="输入自定义场景名称">';
    html += '</div>';
    html += '</div>';
    html += '<div class="form-group"><label>业务领域</label><input type="text" id="np-domain" placeholder="如：信贷、风控、营销（默认同场景名称）"></div>';
    html += '<div class="form-actions">';
    html += '<button class="action-btn secondary small" onclick="hideNewPipelineForm()">取消</button>';
    html += '<button class="action-btn small" onclick="createPipeline()">创建并开始</button>';
    html += '</div>';
    html += '</div>';

    // Pipeline History
    html += '<div class="overview-history-header">';
    html += '<div class="overview-history-title">历史流水线</div>';
    html += '<div class="overview-history-count">共 ' + pipelines.length + ' 条</div>';
    html += '</div>';

    if (pipelines.length > 0) {
      html += '<div class="pipeline-list">';
      pipelines.forEach(p => {
        html += renderPipelineItem(p);
      });
      html += '</div>';
    } else {
      html += '<div class="pipeline-empty">';
      html += '<div class="pipeline-empty-icon">📋</div>';
      html += '<div class="pipeline-empty-text">暂无历史流水线，点击上方"新建流水线"开始</div>';
      html += '</div>';
    }

    html += '</div>';
    container.innerHTML = html;
  } catch (e) {
    container.innerHTML = '<div class="pipeline-overview"><div class="error-list"><div class="error-item">加载失败: ' + escapeHtml(e.message) + '</div></div></div>';
  }
}

function renderPipelineItem(p) {
  const ss = p.step_status || {};
  const doneCount = Object.values(ss).filter(v => v === 'done').length;
  const shownDoneCount = Math.min(doneCount, MAX_STEP);
  const isComplete = shownDoneCount >= MAX_STEP;
  const activeStep = p.current_step || 1;
  const updatedAt = p.updated_at ? p.updated_at.slice(0, 16).replace('T', ' ') : '-';

  const statusClass = isComplete ? 'completed' : 'in-progress';
  const statusLabel = isComplete ? '已完成' : '进行中';
  const statusIcon = isComplete ? '&#10003;' : '&#9679;';

  let html = '<div class="pipeline-item ' + statusClass + '" onclick="enterPipeline(\'' + p.id + '\')">';
  html += '<div class="pipeline-item-top">';
  html += '<div class="pipeline-item-left">';
  html += '<span class="pipeline-item-status ' + statusClass + '">' + statusIcon + ' ' + statusLabel + '</span>';
  html += '<span class="pipeline-item-name">' + escapeHtml(p.name) + '</span>';
  html += '</div>';
  html += '<div class="pipeline-item-right">';
  if (!isComplete) {
    html += '<span class="pipeline-item-continue">继续 &#8250;</span>';
  }
  html += '<button class="pipeline-delete-btn" onclick="event.stopPropagation();deletePipeline(\'' + p.id + '\', this)" title="删除">&#10005;</button>';
  html += '</div>';
  html += '</div>';
  html += '<div class="pipeline-item-meta">';
  html += '<span class="pipeline-item-scenario">' + escapeHtml(p.scenario || '-') + '</span>';
  html += '<span>进度 ' + shownDoneCount + '/' + MAX_STEP + '</span>';
  html += '<span class="pipeline-item-time">' + updatedAt + '</span>';
  html += '</div>';

  // Progress bar
  html += '<div class="pipeline-progress">';
  for (let i = 1; i <= MAX_STEP; i++) {
    const status = ss[String(i)] || 'pending';
    html += '<div class="pipeline-progress-step ' + status + '" title="' + STEP_NAMES[i] + '"></div>';
  }
  html += '</div>';

  // Step mini labels
  html += '<div class="pipeline-step-badges">';
  for (let i = 1; i <= MAX_STEP; i++) {
    const status = ss[String(i)] || 'pending';
    const label = String(i).padStart(2, '0') + ' ' + STEP_NAMES[i];
    const icon = status === 'done' ? '&#10003;' : status === 'active' ? '&#9654;' : '&#9675;';
    html += '<span class="pipeline-step-badge ' + status + '">' + icon + ' ' + label + '</span>';
  }
  html += '</div>';

  html += '</div>';
  return html;
}

function showNewPipelineForm() {
  const form = document.getElementById('new-pipeline-form');
  const roadmap = document.querySelector('.overview-roadmap');
  if (form) { form.style.display = 'block'; }
  if (roadmap) { roadmap.style.display = 'none'; }
}

function hideNewPipelineForm() {
  const form = document.getElementById('new-pipeline-form');
  const roadmap = document.querySelector('.overview-roadmap');
  if (form) { form.style.display = 'none'; }
  if (roadmap) { roadmap.style.display = ''; }
  // Clear form
  const nameEl = document.getElementById('np-name');
  const domainEl = document.getElementById('np-domain');
  if (nameEl) nameEl.value = '';
  if (domainEl) domainEl.value = '';
}

async function createPipeline() {
  const name = (document.getElementById('np-name') || {}).value || '';
  let scenario = (document.getElementById('np-scenario') || {}).value || '';
  if (scenario === 'custom') {
    scenario = (document.getElementById('np-scenario-custom') || {}).value || '';
  }
  const domain = (document.getElementById('np-domain') || {}).value || scenario;

  if (!name.trim()) { alert('请输入流水线名称'); return; }
  if (!scenario.trim()) { alert('请选择或输入场景名称'); return; }

  try {
    const result = await apiCallJSON('/api/pipelines', {
      name: name.trim(),
      scenario: scenario.trim(),
      domain: domain.trim() || scenario.trim()
    });
    if (result.status === 'ok' && result.pipeline) {
      currentPipeline = result.pipeline;
      // Pre-fill Step 1 form for new pipeline
      const nameEl = document.getElementById('s1-scenario-name');
      if (nameEl) nameEl.value = scenario;
      switchPanel(1);
    } else {
      alert(result.error || '创建失败');
    }
  } catch (e) {
    alert('创建失败: ' + e.message);
  }
}

async function enterPipeline(pipelineId) {
  try {
    const resp = await fetch(API_BASE + '/api/pipelines/' + pipelineId);
    const data = await resp.json();
    if (data.status === 'ok' && data.pipeline) {
      currentPipeline = data.pipeline;
      currentPipeline.step_data = currentPipeline.step_data || {};
      const recalled = recallStep1Output(currentPipeline.id);
      if (recalled?.file_name && !currentPipeline.step_data.step1_output_file) {
        currentPipeline.step_data.step1_output_file = recalled.file_name;
        currentPipeline.step_data.step1_download_url = recalled.download_url;
      }
      restoreAllStepsFormData();
      const step = Math.min(MAX_STEP, Math.max(1, currentPipeline.current_step || 1));
      switchPanel(step);
    } else {
      alert(data.error || '加载流水线失败');
    }
  } catch (e) {
    alert('加载失败: ' + e.message);
  }
}

async function deletePipeline(pipelineId, btnEl) {
  if (!confirm('确定删除该流水线？删除后不可恢复。')) return;
  if (btnEl) { btnEl.textContent = '...'; btnEl.disabled = true; }
  try {
    const resp = await fetch(API_BASE + '/api/pipelines/' + pipelineId, { method: 'DELETE' });
    const data = await resp.json();
    if (data.status === 'ok') {
      loadPipelineOverview();
    } else {
      alert(data.error || '删除失败');
    }
  } catch (e) {
    alert('删除失败: ' + e.message);
  }
}

// ═══════════════════════════════════════════════════════════════════
// Excel Online Editor
// ═══════════════════════════════════════════════════════════════════

let _excelEditorData = {
  sheets: {},
  file_path: '',
  active_sheet: '',
  step: 3,
  isRevision: false,
  modified: false
};

function openExcelEditor(step) {
  _excelEditorData.step = step;
  const modal = document.getElementById('excel-editor-modal');
  if (modal) modal.classList.add('active');
}

function closeExcelEditor() {
  if (typeof ExcelEditor !== 'undefined') ExcelEditor.destroy();
  const modal = document.getElementById('excel-editor-modal');
  if (modal) modal.classList.remove('active');
}

async function editStep2Preextract() {
  if (!currentPipeline) { alert('请先进入流水线'); return; }

  const fileName = currentPipeline.step_data?.step2_output_file;
  if (!fileName || !isStep2PreextractFile(fileName)) {
    alert('未找到有效的萃取 Excel（preextract_*.xlsx），请先执行知识萃取');
    return;
  }

  _excelEditorData.step = 2;
  renderExcelEditorLoading();

  try {
    const resp = await fetch(API_BASE + '/api/excel/read', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        file_name: fileName,
        pipeline_id: currentPipeline.id,
        step: 2,
      }),
    });
    const data = await resp.json();
    if (data.status === 'ok') {
      _excelEditorData.sheets = ExcelEditor.normalizeSheetsFromApi(data.sheets);
      _excelEditorData.file_path = data.file_path || fileName;
      _excelEditorData.active_sheet = Object.keys(_excelEditorData.sheets)[0] || '';
      _excelEditorData.modified = false;
      openExcelEditor(2);
      requestAnimationFrame(function () {
        requestAnimationFrame(function () { renderExcelEditorContent(); });
      });
    } else {
      alert('读取萃取 Excel 失败: ' + (data.error || '未知错误'));
    }
  } catch (e) {
    alert('加载萃取 Excel 出错: ' + e.message);
  }
}


function resolveStep3ExcelFile(stepData) {
  const sd = stepData || {};
  if (isStep4FinalFile(sd.step4_final_file)) {
    return {
      fileName: sd.step4_final_file,
      downloadUrl: sd.step4_download_url || ('/downloads/' + sd.step4_final_file),
      isRevision: true,
      editorStep: 4,
    };
  }
  if (isStep2PreextractFile(sd.step2_output_file)) {
    return {
      fileName: sd.step2_output_file,
      downloadUrl: sd.step2_download_url || ('/downloads/' + sd.step2_output_file),
      isRevision: false,
      editorStep: 2,
    };
  }
  return null;
}

async function editStep3Revision() {
  if (!currentPipeline) { alert('请先进入流水线'); return; }

  await refreshCurrentPipeline();
  const resolved = resolveStep3ExcelFile(currentPipeline.step_data);
  if (!resolved || !resolved.fileName) {
    alert('未找到可编辑的 Excel，请先完成知识萃取（Step2）');
    return;
  }

  if (!resolved.isRevision) {
    const go = confirm('尚未生成知识对齐稿。将先打开 Step2 萃取 Excel 供查看/编辑；执行「知识对齐」后将生成对齐稿。是否继续？');
    if (!go) return;
  }

  _excelEditorData.step = resolved.editorStep || (resolved.isRevision ? 3 : 2);
  _excelEditorData.isRevision = !!resolved.isRevision;
  renderExcelEditorLoading();

  try {
    const resp = await fetch(API_BASE + '/api/excel/read', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        file_name: resolved.fileName,
        pipeline_id: currentPipeline.id,
        step: String(_excelEditorData.step),
      }),
    });
    const data = await resp.json();
    if (data.status === 'ok') {
      _excelEditorData.sheets = ExcelEditor.normalizeSheetsFromApi(data.sheets);
      _excelEditorData.file_path = data.file_path || resolved.fileName;
      _excelEditorData.active_sheet = Object.keys(_excelEditorData.sheets)[0] || '';
      _excelEditorData.modified = false;
      const stepLabel = document.getElementById('excel-editor-step-label');
      if (stepLabel) {
        stepLabel.textContent = resolved.isRevision ? '在线编辑：知识对齐稿' : '在线编辑：萃取底稿（执行后生成知识对齐稿）';
      }
      openExcelEditor(_excelEditorData.step);
      requestAnimationFrame(function () {
        requestAnimationFrame(function () { renderExcelEditorContent(); });
      });
    } else {
      alert('读取 Excel 失败: ' + (data.error || '未知错误'));
    }
  } catch (e) {
    alert('加载 Excel 出错: ' + e.message);
  }
}

async function loadExcelToEditor(fileInputId, step) {
  const fileInput = document.getElementById(fileInputId);
  if (!fileInput || !fileInput.files[0]) {
    alert('请先上传 Excel 文件');
    return;
  }

  _excelEditorData.step = step;
  renderExcelEditorLoading();

  const fd = new FormData();
  fd.append('excel', fileInput.files[0]);
  if (currentPipeline) fd.append('pipeline_id', currentPipeline.id);
  fd.append('step', String(step));

  try {
    const resp = await fetch(API_BASE + '/api/excel/read', { method: 'POST', body: fd });
    const data = await resp.json();
    if (data.status === 'ok') {
      _excelEditorData.sheets = ExcelEditor.normalizeSheetsFromApi(data.sheets);
      _excelEditorData.file_path = data.file_path;
      _excelEditorData.active_sheet = Object.keys(_excelEditorData.sheets)[0] || '';
      _excelEditorData.modified = false;
      openExcelEditor(step);
      requestAnimationFrame(function () {
        requestAnimationFrame(function () { renderExcelEditorContent(); });
      });
    } else {
      alert(data.error || '读取 Excel 失败');
    }
  } catch (e) {
    alert('读取失败: ' + e.message);
  }
}

function renderExcelEditorLoading() {
  const modal = document.getElementById('excel-editor-modal');
  if (!modal) return;
  modal.classList.add('active');
  const content = document.getElementById('excel-editor-content');
  if (content) content.innerHTML = '<div class="excel-editor-loading">正在加载 Excel 数据...</div>';
}

function renderExcelEditorContent() {
  const content = document.getElementById('excel-editor-content');
  if (!content) return;

  content.innerHTML =
    '<div class="excel-luckysheet-wrap">' +
    '<div id="excel-luckysheet-mount" class="excel-luckysheet-mount"></div>' +
    '<div class="excel-status-bar">' +
    '<span id="excel-modified-indicator" style="display:none" class="modified-dot"></span>' +
    '<span id="excel-modified-text">未修改</span>' +
    '</div></div>';

  const mountEl = document.getElementById('excel-luckysheet-mount');
  if (mountEl && typeof ExcelEditor !== 'undefined') {
    ExcelEditor.mount(mountEl, _excelEditorData.sheets, _excelEditorData.active_sheet).catch(function (e) {
      console.error('ExcelEditor.mount failed:', e);
      mountEl.innerHTML = '<div class="excel-editor-loading">表格编辑器加载异常：' + escapeHtml(e.message || String(e)) + '</div>';
    });
  }
}

function syncExcelDataFromDOM() {
  if (typeof ExcelEditor === 'undefined') return;
  const synced = ExcelEditor.syncBeforeSave(_excelEditorData.sheets);
  if (synced) _excelEditorData.sheets = synced;
}

function markExcelModified() {
  _excelEditorData.modified = true;
  const indicator = document.getElementById('excel-modified-indicator');
  const text = document.getElementById('excel-modified-text');
  if (indicator) indicator.style.display = 'inline-block';
  if (text) text.textContent = '已修改（未保存）';
}

async function saveExcelEditor() {
  syncExcelDataFromDOM();

  if (!_excelEditorData.file_path) {
    alert('无源文件路径，请重新上传');
    return;
  }

  const saveBtn = document.getElementById('excel-editor-save-btn');
  if (saveBtn) { saveBtn.disabled = true; saveBtn.textContent = '保存中...'; }

  try {
    const resp = await fetch(API_BASE + '/api/excel/save', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        file_path: _excelEditorData.file_path,
        sheets: _excelEditorData.sheets,
        pipeline_id: currentPipeline ? currentPipeline.id : '',
        step: String(_excelEditorData.step)
      })
    });
    const data = await resp.json();

    if (data.status === 'ok') {
      _excelEditorData.modified = false;
      _excelEditorData.file_path = data.file_path;

      const indicator = document.getElementById('excel-modified-indicator');
      const text = document.getElementById('excel-modified-text');
      if (indicator) indicator.style.display = 'none';
      if (text) text.textContent = '已保存';

      const step = _excelEditorData.step;
      if (String(step) === '1' && currentPipeline && data.file_path) {
        const base = data.file_path.replace(/^.*[\\/]/, '');
        applyStep1GenerateResult({
          file_name: base,
          download_url: data.download_url || ('/downloads/' + base),
        });
        persistPipeline({
          step1_output_file: base,
          step1_download_url: data.download_url || ('/downloads/' + base),
        }).catch(e => console.error('Link step1 output failed:', e));
        if (currentStep === 2) loadStep2PrevOutput();
      }

      if (String(step) === '2' && currentPipeline && data.file_path) {
        const base = data.file_path.replace(/^.*[\\/]/, '');
        if (isStep2PreextractFile(base)) {
          currentPipeline.step_data.step2_output_file = base;
          currentPipeline.step_data.step2_download_url = data.download_url || ('/downloads/' + base);
          persistPipeline({
            step2_output_file: base,
            step2_download_url: currentPipeline.step_data.step2_download_url,
          }).catch(function (e) { console.error('Link step2 output failed:', e); });
        }
      }

      if (String(step) === '3' && currentPipeline && data.file_path && _excelEditorData.isRevision) {
        const base = data.file_path.replace(/^.*[\\/]/, '');
        if (isStep3RevisionFile(base)) {
          currentPipeline.step_data.step3_revision_file = base;
          currentPipeline.step_data.step3_download_url = data.download_url || ('/downloads/' + base);
          persistPipeline({
            step3_revision_file: base,
            step3_download_url: currentPipeline.step_data.step3_download_url,
          }).catch(function (e) { console.error('Link step3 output failed:', e); });
        }
      }
      if (String(step) === '4' && currentPipeline && data.file_path) {
        const base = data.file_path.replace(/^.*[\\/]/, '');
        const snapUrl = data.download_url || '';
        if (isStep4FinalFile(base)) {
          currentPipeline.step_data.step4_final_file = base;
          currentPipeline.step_data.step4_download_url = snapUrl || ('/downloads/' + base);
          persistPipeline({
            step4_final_file: base,
            step4_download_url: currentPipeline.step_data.step4_download_url,
          }).catch(function (e) { console.error('Link step4 output failed:', e); });
        } else if (isStep3RevisionFile(base) || _excelEditorData.isRevision) {
          currentPipeline.step_data.step3_revision_file = base;
          currentPipeline.step_data.step3_download_url = snapUrl || ('/downloads/' + base);
          persistPipeline({
            step3_revision_file: base,
            step3_download_url: currentPipeline.step_data.step3_download_url,
          }).catch(function (e) { console.error('Link step3 revision from step4 editor failed:', e); });
        }
      }


      const outputId = 's' + step + '-output';
      const outputEl = document.getElementById(outputId);
      if (outputEl && data.download_url) {
        let existing = outputEl.querySelector('.excel-download-link');
        if (!existing) {
          const label = step === 3 ? '修订稿' : step === 4 ? '最终稿' : step === 1 ? '场景骨架' : 'Excel';
          const div = document.createElement('div');
          div.className = 'excel-download-link';
          div.innerHTML = '<span class="download-icon">📥</span> <a href="' + escapeHtml(API_BASE + data.download_url) + '" download>下载' + label + ' Excel</a>';
          outputEl.appendChild(div);
        }
      }
    } else {
      alert(data.error || '保存失败');
    }
  } catch (e) {
    alert('保存失败: ' + e.message);
  } finally {
    if (saveBtn) { saveBtn.disabled = false; saveBtn.textContent = '保存'; }
  }
}

// ── 委托公共函数到 utils.js / state.js 模块 ──
;(function() {
  var A = window.App;
  if (!A) return;
  // 工具函数委托（utils.js 在 app.js 之前加载）
  if (A.escapeHtml)    { escapeHtml = A.escapeHtml; showToast = A.showToast; copyTextToClipboard = A.copyTextToClipboard; }
  if (A.qualityBar)    { qualityBar = A.qualityBar; statRow = A.statRow; }
  if (A.apiCall)       { apiCall = A.apiCall; apiCallJSON = A.apiCallJSON; withButtonLock = A.withButtonLock; }
  if (A.renderOutput)  { renderOutput = A.renderOutput; renderLoading = A.renderLoading; }
  // 状态管理器委托
  if (A.PipelineState) {
    // scheduleFormSave 委托到 PipelineState（防抖已改为 2 秒）
    var _origScheduleFormSave = scheduleFormSave;
    scheduleFormSave = function(step) { A.PipelineState.scheduleFormSave(step, collectAllStepsFormData); };
    // persistPipeline 委托
    persistPipeline = function(extraStepData, extraFields) { return A.PipelineState.persist(extraStepData, extraFields); };
  }
})();

// ═══════════════════════════════════════════════════════════════════
// 方案三: 结构化访谈模块
// ═══════════════════════════════════════════════════════════════════

let _interviewCtx = { knowledge: null, method: 'case_reverse', probes: [] };

function openInterviewFromExtract() {
  // 从 Step2 萃取结果中读一条知识作为深挖上下文
  var scenarioName = document.getElementById('s1-scenario-name')?.value || currentPipeline?.scenario || '';
  var scenarioContent = document.getElementById('s1-scenario-content')?.value || '';
  var knowledgeItem = {
    knowledge_id: 'KN-' + (currentPipeline?.id || 'p').substring(0, 8) + '-001',
    '知识描述': scenarioContent || scenarioName || '当前流水线场景知识',
    '知识分类': '判断规则',
    '适用条件': '',
    '判断逻辑': '',
  };
  openInterviewProbe(knowledgeItem);
}

async function openInterviewProbe(knowledgeItem) {
  // knowledgeItem: { 知识描述, 知识分类, 适用条件, 判断逻辑, knowledge_id, ... }
  _interviewCtx.knowledge = knowledgeItem;
  _interviewCtx.method = 'case_reverse';
  _interviewCtx.probes = [];
  document.getElementById('interview-modal').classList.add('active');
  document.getElementById('interview-body').innerHTML = '<div class="interview-loading">选择访谈方法后点击「生成追问」</div>';
  renderInterviewMethodTabs();
  await generateInterviewProbes();
}

function renderInterviewMethodTabs() {
  var body = document.getElementById('interview-body');
  var methods = [
    { id: 'case_reverse', name: '案例反推', desc: '构造反例场景，追问专家会如何判断' },
    { id: 'contrast_probe', name: '对比追问', desc: '构造相似但不同的场景，找规则真正边界' },
    { id: 'limit_hypothesis', name: '极限假设', desc: '推到极限条件，找出规则失效边界' },
  ];
  var tabsHtml = '<div class="interview-method-tabs">';
  methods.forEach(function (m) {
    tabsHtml += '<button class="interview-method-tab' + (_interviewCtx.method === m.id ? ' active' : '') + '" onclick="switchInterviewMethod(\'' + m.id + '\')">' + m.name + '</button>';
  });
  tabsHtml += '</div><div id="interview-probes-container"></div>';
  body.innerHTML = tabsHtml;
}

async function switchInterviewMethod(method) {
  _interviewCtx.method = method;
  renderInterviewMethodTabs();
  await generateInterviewProbes();
}

async function generateInterviewProbes() {
  var container = document.getElementById('interview-probes-container');
  if (!container) return;
  container.innerHTML = '<div class="interview-loading"><div class="spinner"></div>生成追问中...</div>';

  var model = resolveModelName('s2-model') || resolveModelName('s3-model');
  if (!model) { container.innerHTML = '<div class="interview-loading" style="color:var(--error)">请先配置模型</div>'; return; }

  try {
    var resp = await fetch(API_BASE + '/api/interview/probe', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ method: _interviewCtx.method, knowledge: _interviewCtx.knowledge, model: model }),
    });
    var data = await resp.json();
    if (data.status === 'ok' && data.probes) {
      _interviewCtx.probes = data.probes;
      renderInterviewProbes();
    } else {
      container.innerHTML = '<div class="interview-loading" style="color:var(--error)">' + escapeHtml(data.error || '生成失败') + '</div>';
    }
  } catch (e) {
    container.innerHTML = '<div class="interview-loading" style="color:var(--error)">网络错误: ' + escapeHtml(e.message) + '</div>';
  }
}

function renderInterviewProbes() {
  var container = document.getElementById('interview-probes-container');
  if (!container) return;
  var html = '<div class="interview-probe-list">';
  _interviewCtx.probes.forEach(function (p, i) {
    html += '<div class="interview-probe-card">';
    html += '<div class="interview-probe-q">' + escapeHtml(p.question) + '<span class="interview-probe-cat">' + escapeHtml(p.category || '经验判断') + '</span></div>';
    if (p.hint) html += '<div class="interview-probe-hint">💡 ' + escapeHtml(p.hint) + '</div>';
    html += '<textarea class="interview-probe-answer" id="interview-answer-' + i + '" placeholder="输入你的回答..."></textarea>';
    html += '</div>';
  });
  html += '</div>';
  container.innerHTML = html;
}

function saveInterviewAnswers() {
  if (!_interviewCtx.probes.length) { showToast('请先生成追问', 'error'); return; }
  if (!currentPipeline) { showToast('请先从总览进入一条流水线', 'error'); return; }

  var annotations = [];
  _interviewCtx.probes.forEach(function (p, i) {
    var ans = document.getElementById('interview-answer-' + i);
    if (ans && ans.value.trim()) {
      annotations.push({
        note_id: 'interview_' + Date.now() + '_' + i,
        action: 'supplement',
        knowledge_id: _interviewCtx.knowledge.knowledge_id || _interviewCtx.knowledge['知识编号'] || '',
        category: p.category || '经验判断',
        question: p.question,
        answer: ans.value.trim(),
      });
    }
  });

  if (!annotations.length) { showToast('请至少回答一条追问', 'error'); return; }

  // 保存到当前流水线的 step_data.tacit_annotations
  currentPipeline.step_data = currentPipeline.step_data || {};
  var existing = currentPipeline.step_data.step4_tacit_annotations || [];
  currentPipeline.step_data.step4_tacit_annotations = existing.concat(annotations);
  persistPipeline({ step4_tacit_annotations: currentPipeline.step_data.step4_tacit_annotations }).then(function () {
    showToast('已保存 ' + annotations.length + ' 条隐性注释');
    closeInterviewModal();
  }).catch(function () {
    showToast('保存失败', 'error');
  });
}

function closeInterviewModal() {
  document.getElementById('interview-modal').classList.remove('active');
}

// ═══════════════════════════════════════════════════════════════════
// 方案四: 显性化校验回放
// ═══════════════════════════════════════════════════════════════════

function openValidatePanel() {
  var panel = document.getElementById('s5-validate-panel');
  panel.classList.toggle('hidden');
  if (!panel.classList.contains('hidden')) {
    if (allModels.length) refreshModelSelects();
  }
}

async function runValidateReplay() {
  var btn = document.getElementById('s5-validate-run-btn');
  var resultEl = document.getElementById('s5-validate-result');
  var text = document.getElementById('s5-validate-cases').value.trim();
  if (!text) { showToast('请输入历史案例', 'error'); return; }
  if (!currentPipeline) { showToast('请先进入一条流水线', 'error'); return; }

  // 解析案例文本
  var lines = text.split('\n').filter(function (l) { return l.trim(); });
  var cases = [];
  lines.forEach(function (line) {
    var parts = line.split(/[,，]/);
    if (parts.length >= 3) {
      cases.push({
        case_id: parts[0].trim(),
        description: parts[1].trim(),
        conclusion: parts[2].trim(),
      });
    }
  });
  if (cases.length === 0) { showToast('案例格式错误，每行为: 案例ID,场景描述,专家结论', 'error'); return; }

  var model = document.getElementById('s5-validate-model')?.value || resolveModelName('s5-model');
  if (!model) { showToast('请选择模型', 'error'); return; }

  btn.disabled = true; btn.textContent = '校验中...';
  resultEl.innerHTML = '<div class="loading"><div class="spinner"></div>正在用知识库判断 ' + cases.length + ' 个案例...</div>';

  try {
    var fd = new FormData();
    fd.append('pipeline_id', currentPipeline.id);
    fd.append('model', model);
    fd.append('cases', JSON.stringify(cases));
    var resp = await fetch(API_BASE + '/api/validate/replay', { method: 'POST', body: fd });
    var data = await resp.json();
    if (data.status === 'ok') {
      var pct = Math.round(data.hit_rate * 100);
      var fillColor = pct >= 80 ? '#16a34a' : pct >= 60 ? '#f59e0b' : '#ef4444';
      var html = '<div class="validate-result">';
      html += '<h4>校验结果</h4>';
      html += '<div style="font-size:24px;font-weight:700;color:' + fillColor + '">' + pct + '% 命中率</div>';
      html += '<div style="font-size:12px;color:var(--text-muted)">' + data.hits + '/' + data.total + ' 一致 · ' + data.mismatch_count + ' 分歧</div>';
      html += '<div class="validate-hit-bar"><div class="validate-hit-fill" style="width:' + pct + '%;background:' + fillColor + '"></div></div>';
      if (data.mismatches && data.mismatches.length) {
        html += '<h4 style="margin-top:12px;">分歧案例</h4>';
        data.mismatches.forEach(function (m) {
          html += '<div class="validate-mismatch"><strong>' + escapeHtml(m.case_id) + '</strong>: LLM判「' + escapeHtml(m.prediction) + '」→ 专家判「' + escapeHtml(m.expert_conclusion) + '」<br><span style="color:var(--text-muted);font-size:11px">推理: ' + escapeHtml((m.reasoning || '').substring(0, 120)) + '</span></div>';
        });
        html += '<div class="file-hint" style="margin-top:8px">💡 这些分歧条目可反推为 Step3 修订建议来源</div>';
      }
      html += '</div>';
      resultEl.innerHTML = html;
    } else {
      resultEl.innerHTML = '<div style="color:var(--error);font-size:12px;margin-top:8px">' + escapeHtml(data.error || '校验失败') + '</div>';
    }
  } catch (e) {
    resultEl.innerHTML = '<div style="color:var(--error);font-size:12px;margin-top:8px">网络错误: ' + escapeHtml(e.message) + '</div>';
  }
  btn.disabled = false; btn.textContent = '执行校验';
}

// Load pipeline overview on startup
loadPipelineOverview();
