# Docker 部署（ARM64）

## 版本规范

- 应用版本：根目录 `VERSION`（当前 `2.0.0`）
- 镜像标签：`{app_version}-arm64-{yyyyMMdd-HHmm}`，例如 `2.0.0-arm64-202605291753`
- 构建后查看 `*.manifest.json` 中的 `compose_image_line`
- 运行后查询：`curl http://localhost:5000/api/version`

详见 [docs/VERSIONING.md](../docs/VERSIONING.md)。

## 一、在本机构建 ARM64 镜像

### Windows（PowerShell）

```powershell
cd projects
.\scripts\build-docker-arm64.ps1
```

生成：

- `tacit-knowledge-externalization-2.0.0-arm64-<时间>.tar`
- `tacit-knowledge-externalization-arm64.tar`（最新构建副本）
- `tacit-knowledge-externalization-2.0.0-arm64-<时间>.manifest.json`

### Linux / macOS

```bash
cd projects
docker buildx build --platform linux/arm64 \
  -t tacit-knowledge-externalization:latest \
  -f docker/Dockerfile --load .
docker save tacit-knowledge-externalization:latest \
  -o tacit-knowledge-externalization-arm64.tar
```

## 二、ARM64 服务器离线部署

```bash
# 1. 传输 tar 与 docker-compose.yml、config/、workspace/、logs/ 目录
docker load -i tacit-knowledge-externalization-arm64.tar

# 2. 启动（需 docker compose v2）
docker compose up -d

# 3. 健康检查
curl http://localhost:5000/api/version
# 或 curl http://localhost:5000/api/health
```

浏览器访问：`http://<服务器IP>:5000`

## 三、数据持久化（重要）

应用把所有「可恢复状态」写在 **`WORKSPACE_DIR`（默认 `/app/workspace`）**：

| 文件/目录 | 内容 |
|-----------|------|
| `pipelines.json` | 流水线列表与各步 `step_data`（进度、下载链接） |
| `custom_models.json` / `preset_overrides.json` | 自定义模型配置 |
| `template_*.xlsx` / `preextract_*.xlsx` / `revision_*.xlsx` / `final_*.xlsx` | 各步产出 Excel |
| `SKILL_*.md` / `quality_report_*` | Step5 产出 |
| `upload_*` | 上传的模板与附件 |

**若不挂载 `workspace`，容器删除或重建后上述数据全部丢失，需要从头跑流水线。**

### docker compose（推荐）

```bash
mkdir -p workspace logs
docker compose up -d
```

`docker-compose.yml` 已默认挂载 `./workspace` 与 `./logs`。

### 仅用 docker run 时（必须加 -v）

```bash
mkdir -p /data/tacit-knowledge/workspace /data/tacit-knowledge/logs

docker run -d --name tacit-knowledge-externalization \
  --platform linux/arm64 \
  -p 5001:5000 \
  -v /data/tacit-knowledge/workspace:/app/workspace \
  -v /data/tacit-knowledge/logs:/app/logs \
  -v /data/tacit-knowledge/llm-config.yaml:/app/config/llm-config.yaml:ro \
  -e WORKSPACE_DIR=/app/workspace \
  tacit-knowledge-externalization:arm64-v2026.05.25-XXXX
```

替换镜像 tag 为 `docker load` 后 `docker images` 中显示的标签。

### 从旧容器迁移数据

若旧容器未挂卷、数据在容器内：

```bash
docker cp tacit-knowledge-externalization:/app/workspace ./workspace
docker cp tacit-knowledge-externalization:/app/logs ./logs
# 停删旧容器后，用上面带 -v 的命令重新启动
```

| 挂载 | 说明 |
|------|------|
| `./workspace` → `/app/workspace` | **必挂**：流水线 JSON + 全部 Excel/SKILL |
| `./logs` → `/app/logs` | 应用日志 |
| `./config/llm-config.yaml` | LLM 模型配置（可热更新后重启） |

## 四、环境变量

| 变量 | 默认值 | 说明 |
|------|--------|------|
| `WORKSPACE_DIR` | `/app/workspace` | 工作区路径 |
| `TZ` | `Asia/Shanghai` | 时区 |

## 五、内网 LLM

编辑 `config/llm-config.yaml` 中的模型 URL 为内网可达地址，重启容器：

```bash
docker compose restart
```
