# 生产前逻辑审计（2026-05-25）

本文档记录端到端流水线中已加固的不变量与曾出现的生产级缺陷根因。

## 核心不变量

| 步骤 | `step_data` 字段 | 合法文件名前缀 |
|------|------------------|----------------|
| Step1 | `step1_output_file` | `template_*.xlsx` |
| Step2 | `step2_output_file` | `preextract_*` 或 `edited_step2_*` |
| Step3 | `step3_revision_file` | `revision_*` 或 `edited_step3_*` |
| Step4 | `step4_final_file` | `final_*` 或 `edited_step4_*` |
| Step5 | `step5_skill_file` | `SKILL_*.md` |

**禁止**：Step3/4 修订或下载使用 Step1 的 `template_*` 作为 Step2 萃取底稿。

## 已修复的生产缺陷

1. **下载萃取却得到场景骨架**：前端 `resolveStep2DownloadInfo` 曾用本地缓存 `step1_output_file` 兜底；已改为仅接受 API 返回的 `preextract_*`。
2. **Step3 误用 Step1 模板**：后端 `_resolve_step2_excel_path` 与前端 `resolveStep3ExcelFile` 已去掉 Step1 回退。
3. **Excel 保存错链步骤产出**：`saveExcelEditor` 按 `step` + 文件名前缀写回，Step3 编辑萃取表（step=2）不再写入 `step3_revision_file`。
4. **刷新后本地覆盖服务端**：`refreshCurrentPipeline` 合并时 `preferServer: true`，服务端 `step_data` 优先。
5. **Step1 重新生成未清空下游**：服务端 `downstream_output_keys(1)` + 前端 `clearDownstreamOutputs(1)`。
6. **回滚只改状态不清产出**：`rollback` 清除 `downstream_output_keys(step)`。
7. **空萃取表误导用户**：模型有文本但 0 条时不再写入空 `preextract_*.xlsx`。

## 服务端防护（`backend/pipeline_artifacts.py`）

- `/downloads/<file>`：白名单前缀 + 禁止 `pipelines.json` 等系统文件
- `PUT /api/pipelines/<id>`：`validate_step_data_patch` 校验产出字段
- `/api/excel/read|save`：`safe_workspace_path` / `resolve_client_excel_path`，响应仅 basename
- `/api/files/save`：禁止写入受保护文件

## 前端防护（`frontend/js/app.js`）

- `resolveStep2DownloadInfo` / `resolveArtifact` + `isStep2/3/4*File` 校验
- `mergeStepDataPreserveOutputs(..., { preferServer: true })`  on refresh
- `clearCurrentPipeline` 不销毁 `#s3-prev-info` 等固定 DOM 节点
- Step2 萃取失败时展示 `parse_mode` / `style_rule`（若 API 返回）

## 二次审计修复（2026-05-25 续）

- 修复 `_persist_step2_excel_pipeline` 中 `_is_step2_preextract_filename` 未定义（会导致 Step2 成功后服务端不落库）
- Excel 保存不再无条件污染 `step{N}_download_url`；Step4 编辑修订稿时回写 `step3_revision_file`
- Step4/5 输入与 Step4 prev 增加 `revision_*` / `final_*` 前缀校验
- 前端：重跑 Step2/3/4 时清除下游产出；Step3 无有效 `revision_*` 不标完成；`step2_extracted_count` 纳入合并保护

## 修订行号对齐（2026-05-25 续）

- 新增 `backend/workbook_layout.py`：预览带 `excel_row` / `data_row`，Step3/4 prompt 统一使用物理行号
- `normalize_revision_notes`：数据行序号自动换算；跳过对表头行的 modify/delete
- `revision_processor`：修订元数据列表头与模板 `header_rows` 纵向合并对齐
- 回滚 / Step1 重生成：`keys_to_clear_from_step` 同时清理 `*_cached_file`、`*_excel_path` 等
- `/api/files/read`：与 downloads 一致，禁止读取 `pipelines.json` 等非白名单文件

## 仍须关注

- 复杂合并单元格模板下，列号 col 仍依赖模型判断，极端版式可能偏差
- `api_clear` 仍不删除 workspace 磁盘上的历史 xlsx（仅清 JSON）

## 验证命令

```bash
# 需先启动 backend（端口 5000）
python backend/scripts/test_artifacts.py

# 全链路（依赖 LLM，耗时较长）
set E2E_BASE=http://127.0.0.1:5000
python backend/scripts/e2e_pipeline_test.py
```

## 容器部署注意

1. 必须使用包含本次改动的镜像标签重新 `docker build` / `docker run`。
2. 浏览器 **Ctrl+F5** 强刷，确认 `app.js?v=20260525audit` 已加载。
3. 映射端口示例：`-p 5001:5000`，访问 `http://<host>:5001`。
