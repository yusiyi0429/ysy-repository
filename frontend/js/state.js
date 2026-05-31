/**
 * 流水线状态管理器
 * 封装 currentPipeline 的读写、表单自动保存、状态变更通知
 */
(function (global) {
  'use strict';

  const API_BASE = global.location.origin;
  const MAX_STEP = 4;
  const MAX_FORM_STEP = 3;
  const SAVE_DEBOUNCE_MS = 2000;

  /** 各步骤产出物字段 — 保存表单时不可覆盖丢失 */
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

  // ── Internal State ──
  let _pipeline = null;
  let _formSaveTimer = null;
  let _listeners = [];

  function _notify(event, data) {
    _listeners.forEach(function (fn) {
      try { fn(event, data); } catch (_) { /* ignore */ }
    });
  }

  // ── Public API ──

  const PipelineState = {
    /** 获取当前流水线（深拷贝引用，避免外部直接修改内部状态） */
    get current() {
      return _pipeline;
    },

    /** 设置当前流水线 */
    set current(p) {
      _pipeline = p;
      _notify('changed', _pipeline);
    },

    /** 是否有活跃的流水线 */
    get active() {
      return !!(_pipeline && _pipeline.id);
    },

    /** 当前步骤号 */
    get currentStep() {
      return _pipeline ? Math.min(MAX_STEP, Math.max(1, _pipeline.current_step || 1)) : 0;
    },

    /** 获取 step_data（安全返回 {}） */
    get stepData() {
      return (_pipeline && _pipeline.step_data) ? _pipeline.step_data : {};
    },

    /** 获取 step_status（安全返回 {}） */
    get stepStatus() {
      return (_pipeline && _pipeline.step_status) ? _pipeline.step_status : {};
    },

    /** 获取指定步骤的表单数据 */
    getFormData(step) {
      var sd = this.stepData;
      return sd['step' + step + '_form_data'] || {};
    },

    /** 注册状态变更监听器 */
    on(fn) {
      _listeners.push(fn);
    },

    /** 移除监听器 */
    off(fn) {
      _listeners = _listeners.filter(function (f) { return f !== fn; });
    },

    /** 合并 step_data：保留产出物字段不被覆盖 */
    mergeStepDataPreserveOutputs(serverData, localData, preferServer) {
      preferServer = preferServer === true;
      var server = serverData || {};
      var local = localData || {};
      var merged = preferServer ? Object.assign({}, local, server) : Object.assign({}, server, local);
      for (var i = 0; i < PIPELINE_OUTPUT_KEYS.length; i++) {
        var key = PIPELINE_OUTPUT_KEYS[i];
        var v = preferServer ? (server[key] || local[key]) : (local[key] || server[key]);
        if (v) merged[key] = v;
      }
      return merged;
    },

    /** 清除指定步骤及其下游的输出字段 */
    clearDownstreamOutputs(fromStep) {
      if (!_pipeline || !_pipeline.step_data) return;
      var start = fromStep <= 1 ? 2 : (fromStep + 1);
      var keysByStep = {
        2: DOWNSTREAM_OUTPUT_KEYS.filter(function (k) { return k.indexOf('step2_') === 0 || k.indexOf('skill_') === 0; }),
        3: DOWNSTREAM_OUTPUT_KEYS.filter(function (k) { return k.indexOf('step3_') === 0; }),
        4: DOWNSTREAM_OUTPUT_KEYS.filter(function (k) { return k.indexOf('step4_') === 0; }),
        5: DOWNSTREAM_OUTPUT_KEYS.filter(function (k) { return k.indexOf('step5_') === 0; }),
      };
      for (var s = start; s <= 5; s++) {
        (keysByStep[s] || []).forEach(function (k) { delete _pipeline.step_data[k]; });
      }
      _notify('outputsCleared', fromStep);
    },

    // ── 自动保存 ──

    /** 调度表单自动保存（防抖 2 秒） */
    scheduleFormSave(step, collectFn) {
      if (!_pipeline) return;
      clearTimeout(_formSaveTimer);
      _formSaveTimer = setTimeout(async function () {
        if (!collectFn) return;
        var allForm = collectFn();
        _pipeline.step_data = _pipeline.step_data || {};
        var changed = false;
        var keys = Object.keys(allForm);
        for (var i = 0; i < keys.length; i++) {
          var key = keys[i];
          if (JSON.stringify(_pipeline.step_data[key]) !== JSON.stringify(allForm[key])) {
            _pipeline.step_data[key] = allForm[key];
            changed = true;
          }
        }
        if (!changed) return;
        try {
          await PipelineState.persist(allForm);
        } catch (e) {
          console.error('Auto-save failed:', e);
        }
      }, SAVE_DEBOUNCE_MS);
    },

    /** 持久化到服务端 */
    async persist(extraStepData, extraFields) {
      if (!_pipeline) return null;
      var payload = {
        current_step: _pipeline.current_step,
        step_status: _pipeline.step_status,
        step_data: PipelineState.mergeStepDataPreserveOutputs(_pipeline.step_data, extraStepData || {}),
      };
      if (extraFields) Object.assign(payload, extraFields);
      _pipeline.step_data = payload.step_data;
      var result = await App.apiCallJSON('/api/pipelines/' + _pipeline.id, payload, 'PUT');
      if (result && result.pipeline) {
        _pipeline.step_data = PipelineState.mergeStepDataPreserveOutputs(result.pipeline.step_data, _pipeline.step_data);
        if (result.pipeline.current_step != null) {
          _pipeline.current_step = Math.min(MAX_STEP, Math.max(1, result.pipeline.current_step || 1));
        }
        if (result.pipeline.step_status) _pipeline.step_status = result.pipeline.step_status;
        _notify('persisted', _pipeline);
      }
      return result;
    },

    /** 从服务端刷新当前流水线 */
    async refresh() {
      if (!_pipeline || !_pipeline.id) return null;
      try {
        var resp = await fetch(API_BASE + '/api/pipelines/' + _pipeline.id);
        var data = await resp.json();
        if (data.status === 'ok' && data.pipeline) {
          _pipeline.current_step = Math.min(MAX_STEP, Math.max(1, data.pipeline.current_step || 1));
          _pipeline.step_status = data.pipeline.step_status;
          _pipeline.step_data = PipelineState.mergeStepDataPreserveOutputs(data.pipeline.step_data, _pipeline.step_data, true);
          if (data.pipeline.scenario) _pipeline.scenario = data.pipeline.scenario;
          if (data.pipeline.domain) _pipeline.domain = data.pipeline.domain;
          _notify('refreshed', _pipeline);
        }
        return _pipeline;
      } catch (e) {
        console.warn('refreshCurrentPipeline failed:', e);
        return _pipeline;
      }
    },

    /** 标记步骤完成 */
    async markStepDone(step, collectFn) {
      if (!_pipeline) return;
      if (collectFn) {
        var allForm = collectFn();
        _pipeline.step_data = PipelineState.mergeStepDataPreserveOutputs(_pipeline.step_data, allForm);
      }
      _pipeline.step_status[String(step)] = 'done';
      if (step < MAX_STEP) {
        _pipeline.step_status[String(step + 1)] = _pipeline.step_status[String(step + 1)] || 'active';
        _pipeline.current_step = Math.max(_pipeline.current_step || 1, step + 1);
      } else {
        _pipeline.current_step = MAX_STEP;
      }
      try {
        await PipelineState.persist(_pipeline.step_data, {
          current_step: _pipeline.current_step,
          step_status: _pipeline.step_status,
        });
      } catch (e) {
        console.error('Failed to update pipeline:', e);
      }
      _notify('stepDone', step);
    },
  };

  // ── 暴露到全局 ──
  global.App = global.App || {};
  global.App.PipelineState = PipelineState;
  global.App.PIPELINE_OUTPUT_KEYS = PIPELINE_OUTPUT_KEYS;
  global.App.DOWNSTREAM_OUTPUT_KEYS = DOWNSTREAM_OUTPUT_KEYS;

})(window);
