#!/bin/bash
# 隐性知识提取系统启动脚本
# 用法: ./start.sh -up   启动服务
#       ./start.sh -down 停止服务

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_DIR="$(cd "$SCRIPT_DIR/.." && pwd)"
APP_SCRIPT="$PROJECT_DIR/backend/app_server.py"
PID_FILE="$PROJECT_DIR/.app_server.pid"
LOG_FILE="$PROJECT_DIR/logs/app_server.log"
PORT=5000

start_server() {
    # 检查是否已运行
    if [ -f "$PID_FILE" ]; then
        PID=$(cat "$PID_FILE")
        if ps -p "$PID" > /dev/null 2>&1; then
            echo "服务已在运行中 (PID: $PID)"
            exit 1
        fi
        rm -f "$PID_FILE"
    fi
    
    # 检查端口是否被占用
    if ss -tuln 2>/dev/null | grep -q ":$PORT "; then
        echo "端口 $PORT 已被占用，请先停止占用进程"
        exit 1
    fi
    
    # 检查依赖
    if ! python3 -c "import flask, openpyxl, yaml, requests" 2>/dev/null; then
        echo "正在安装依赖..."
        pip install flask openpyxl pyyaml requests -q
    fi
    
    mkdir -p "$PROJECT_DIR/logs"
    echo "正在启动服务..."
    nohup python3 "$APP_SCRIPT" --host 0.0.0.0 --port $PORT > "$LOG_FILE" 2>&1 &
    PID=$!
    echo $PID > "$PID_FILE"
    
    # 等待服务启动
    sleep 2
    
    # 验证服务是否启动成功
    if curl -s -o /dev/null -w "%{http_code}" "http://localhost:$PORT/api/health" 2>/dev/null | grep -q "200"; then
        echo "服务启动成功!"
        echo "  PID: $PID"
        echo "  端口: $PORT"
        echo "  日志: $LOG_FILE"
        echo "  访问: http://localhost:$PORT"
    else
        echo "服务启动失败，请检查日志: $LOG_FILE"
        tail -n 20 "$LOG_FILE"
        exit 1
    fi
}

stop_server() {
    if [ ! -f "$PID_FILE" ]; then
        echo "未找到 PID 文件，尝试通过端口查找进程..."
        PID=$(ss -lptn "sport = :$PORT" 2>/dev/null | grep -oP 'pid=\K[0-9]+' | head -1)
        if [ -z "$PID" ]; then
            echo "服务未运行"
            exit 0
        fi
    else
        PID=$(cat "$PID_FILE")
    fi
    
    if ! ps -p "$PID" > /dev/null 2>&1; then
        echo "服务未运行 (PID: $PID 已不存在)"
        rm -f "$PID_FILE"
        exit 0
    fi
    
    echo "正在停止服务 (PID: $PID)..."
    kill "$PID" 2>/dev/null
    
    # 等待进程结束
    for i in {1..10}; do
        if ! ps -p "$PID" > /dev/null 2>&1; then
            break
        fi
        sleep 1
    done
    
    # 强制杀死
    if ps -p "$PID" > /dev/null 2>&1; then
        echo "强制停止服务..."
        kill -9 "$PID" 2>/dev/null
    fi
    
    rm -f "$PID_FILE"
    echo "服务已停止"
}

status_server() {
    if [ -f "$PID_FILE" ]; then
        PID=$(cat "$PID_FILE")
        if ps -p "$PID" > /dev/null 2>&1; then
            echo "服务运行中 (PID: $PID, 端口: $PORT)"
            curl -s "http://localhost:$PORT/api/health" 2>/dev/null && echo ""
            return 0
        fi
    fi
    echo "服务未运行"
    return 1
}

case "$1" in
    -up|--up|up)
        start_server
        ;;
    -down|--down|down)
        stop_server
        ;;
    -status|--status|status)
        status_server
        ;;
    -restart|--restart|restart)
        stop_server
        sleep 1
        start_server
        ;;
    *)
        echo "用法: $0 { -up | -down | -status | -restart }"
        echo ""
        echo "  -up       启动服务"
        echo "  -down     停止服务"
        echo "  -status   查看状态"
        echo "  -restart  重启服务"
        exit 1
        ;;
esac
