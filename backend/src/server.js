const express = require("express");
const cors = require("cors");
const path = require("path");

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.json());

app.get("/api/health", (_req, res) => {
  res.json({
    ok: true,
    service: "backend",
    timestamp: new Date().toISOString()
  });
});

app.get("/api/message", (_req, res) => {
  res.json({
    title: "基础前后端架构已就绪",
    description: "前端通过 fetch 调用后端 API，并动态渲染结果。"
  });
});

app.use(express.static(path.resolve(__dirname, "../../frontend/public")));

app.get("*", (_req, res) => {
  res.sendFile(path.resolve(__dirname, "../../frontend/public/index.html"));
});

app.listen(PORT, () => {
  console.log(`Server is running on http://localhost:${PORT}`);
});
