async function fetchJson(url) {
  const response = await fetch(url);
  if (!response.ok) {
    throw new Error(`请求失败: ${response.status}`);
  }
  return response.json();
}

async function initPage() {
  const healthOutput = document.getElementById("health-output");
  const messageOutput = document.getElementById("message-output");

  try {
    const [health, message] = await Promise.all([
      fetchJson("/api/health"),
      fetchJson("/api/message")
    ]);

    healthOutput.textContent = JSON.stringify(health, null, 2);
    messageOutput.textContent = JSON.stringify(message, null, 2);
  } catch (error) {
    healthOutput.textContent = "健康检查失败";
    messageOutput.textContent = error instanceof Error ? error.message : "未知错误";
  }
}

initPage();
