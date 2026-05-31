# 版本规范（内网部署识别）

## 版本号来源

| 字段 | 来源 | 示例 |
|------|------|------|
| `app_version` | 项目根目录 `VERSION` 文件（语义化版本） | `2.0.0` |
| `image_version` / `build_id` | 构建时注入，格式 `{app_version}-arm64-{yyyyMMdd-HHmm}` | `2.0.0-arm64-202605291753` |
| `build` | Step2 Excel 兼容标识（历史字段） | `20260525-optional-sub-scenario` |
| `platform` | 构建参数 `APP_PLATFORM` | `linux/arm64` |

## 查询接口

```bash
curl http://<host>:5000/api/version
curl http://<host>:5000/api/health
curl http://<host>:5000/api/build_info
```

返回示例：

```json
{
  "status": "ok",
  "app_name": "tacit-knowledge-externalization",
  "app_version": "2.0.0",
  "image_version": "2.0.0-arm64-202605291753",
  "build_id": "2.0.0-arm64-202605291753",
  "platform": "linux/arm64",
  "build_date": "2026-05-29T09:53:00Z",
  "build": "20260525-optional-sub-scenario"
}
```

## 发布流程

1. 修改根目录 `VERSION`（如 `2.0.1`）
2. Windows 构建：`.\scripts\build-docker-arm64.ps1`
3. 得到：
   - `tacit-knowledge-externalization-2.0.1-arm64-*.tar`
   - `*.manifest.json`（含 `compose_image_line`）
4. 内网：`docker load -i <tar>`，按 manifest 更新 `docker-compose.yml` 的 `image`
5. 升级后对比：`curl /api/version` 中 `app_version` 与 `image_version`

## 镜像标签约定

- **推荐（可追踪）**：`tacit-knowledge-externalization:2.0.0-arm64-202605291753`
- **次要别名**：`tacit-knowledge-externalization:2.0.0-arm64`（同次构建会覆盖，勿用于多版本并存）
