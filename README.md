# 隐性知识显性化 · 四步法萃取流水线

将银行信贷专家的隐性经验（制度文件、案例复盘、会议纪要）转化为结构化、可复用、可审计的知识资产。

---

## 核心管线

```
场景锚定 → 知识萃取 → 知识对齐 → 智能转化
  Step1      Step2       Step3       Step5
```

| 步骤 | 做什么 | 输入 | 输出 |
|------|--------|------|------|
| **Step1 场景锚定** | 定义知识结构，生成 Excel 骨架 | 场景名称、场景说明、子场景、知识列定义 | `template_*.xlsx` |
| **Step2 知识萃取** | 从文档/案例中提取结构化知识 | 制度文件 / 操作手册 / 案例复盘 | `preextract_*.xlsx` |
| **Step3 知识对齐** | 专家审核修订 + 沉淀隐性注释 | 专家修订意见 / 会议纪要 | `final_*.xlsx` |
| **Step5 智能转化** | 生成可交付的知识资产 | 对齐后的知识 Excel | SKILL.md / QA对 / 思维链 |

---

## 5 个 AI Skill

| Skill | 步骤 | 能力 |
|-------|:---:|------|
| 🔍 **知识萃取** | Step2 | 文档→结构化知识；支持 📄文档模式 和 📋案例复盘模式 |
| 🔬 **跨案例模式发现** | Step2 | 多案例交叉分析，发现反复出现的隐性信号和系统性风险盲区 |
| 🎯 **知识盲区检测** | Step2 | 对比 Schema 与实际填充率，识别「应该知道但还不知道」的内容 |
| 📝 **知识对齐** | Step3 | 专家修订 + 隐性注释追问卡片（自动捕获修订背后的经验判断） |
| 🔄 **知识保鲜度审计** | Step5 | 审计知识时效性，检测被案例突破的规则和置信度衰减 |

---

## 快速启动

```bash
cd backend
pip install -r ../requirements.txt
python app_server.py --host 127.0.0.1 --port 5000
```

浏览器访问 `http://127.0.0.1:5000`

---

## 测试数据

`data/samples/` 按步骤组织，包含科技企业普惠贷款全流程测试数据。

[→ 测试数据使用指南](data/samples/README.md)

---

## 项目结构

```
├── backend/              # Flask API + 业务模块
│   ├── app_server.py     # 主服务（路由 + Skill 执行器）
│   ├── llm_client.py     # LLM 调用适配（OpenAI / CCB 网关 + 重试）
│   ├── pipeline_artifacts.py  # 文件命名与安全策略
│   ├── step1_*.py        # 场景锚定模块
│   ├── step2_preextract.py    # 知识萃取 Excel 生成
│   ├── revision_processor.py  # 知识修订处理器
│   ├── knowledge_delivery.py  # 智能转化（Skill/QA/COT）
│   └── scripts/          # 测试脚本
├── frontend/             # 前端（Vanilla JS + Luckysheet）
│   ├── js/
│   │   ├── state.js      # 流水线状态管理器
│   │   ├── utils.js      # 公共工具函数
│   │   ├── app.js        # 主逻辑
│   │   └── excel-luckysheet.js  # Excel 在线编辑器
│   └── vendor/           # 离线静态资源（内网部署）
├── config/               # LLM 配置 + 场景 Schema
├── data/samples/         # 测试数据（按 Step 组织）
└── docker/               # Docker 构建文件
```

---

## LLM 模型配置

支持两种 `api_type`：

| api_type | 说明 |
|----------|------|
| `openai`（默认） | OpenAI 兼容 `/v1/chat/completions`，Bearer 鉴权 |
| `ccb_ainlplm` | 建行内部网关，需配置 `tx_code`、`sec_node_no` |

配置文件：`config/llm-config.yaml`，密钥写入 `config/llm-config.local.yaml`（已 gitignore）。

---

## Docker（ARM64）

```powershell
.\scripts\build-docker-arm64.ps1
```

生成 `tacit-knowledge-externalization-arm64.tar`，详见 [docker/README.md](docker/README.md)。
