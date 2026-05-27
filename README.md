# ysy-repository

一个包含基础前后端架构的网页应用示例。

## 项目结构

```text
.
├── backend
│   ├── package.json
│   └── src
│       └── server.js
├── frontend
│   └── public
│       ├── app.js
│       ├── index.html
│       └── styles.css
└── package.json
```

## 快速开始

1. 安装依赖：

   ```bash
   npm install
   ```

2. 启动应用：

   ```bash
   npm run dev
   ```

3. 打开浏览器访问：

   ```text
   http://localhost:3000
   ```

## 已包含能力

- Express 后端服务
- 健康检查接口：`GET /api/health`
- 业务示例接口：`GET /api/message`
- 静态前端页面（HTML/CSS/JavaScript）
- 前端调用后端 API 并渲染结果
