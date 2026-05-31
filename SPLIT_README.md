# 隐性知识显性化 — 内网传输分片包

## 文件说明
- tacit-knowledge-deploy-2.0.0-arm64.zip.part0000 ... : 压缩包分片（每片 <= 50 KB）
- tacit-knowledge-deploy-2.0.0-arm64.zip.manifest.json : 分片清单与校验
- REASSEMBLE.sh : Linux 内网合并脚本
- REASSEMBLE.ps1 : Windows 合并脚本

## Linux 内网合并
cd /data/隐性知识显性化
chmod +x REASSEMBLE.sh
./REASSEMBLE.sh

## 合并后部署
unzip tacit-knowledge-deploy-2.0.0-arm64.zip -d deploy
cd deploy
docker load -i tacit-knowledge-externalization-*.tar   # 若有镜像文件
docker compose -p tacit-knowledge up -d
curl http://127.0.0.1:5000/api/version
