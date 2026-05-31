#!/bin/bash
# =============================================================================
# 隐性知识提取系统 - Docker 镜像构建脚本
# 支持 amd64 / arm64 架构选择
# =============================================================================

set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_DIR="$(cd "$SCRIPT_DIR/.." && pwd)"
cd "$PROJECT_DIR"

# 镜像信息
IMAGE_NAME="tacit-knowledge-externalization"
IMAGE_TAG="latest"

# 颜色输出
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
CYAN='\033[0;36m'
NC='\033[0m' # No Color

print_banner() {
    echo ""
    echo -e "${CYAN}╔══════════════════════════════════════════════════════════════╗${NC}"
    echo -e "${CYAN}║         隐性知识提取系统 - Docker 镜像构建工具               ║${NC}"
    echo -e "${CYAN}╚══════════════════════════════════════════════════════════════╝${NC}"
    echo ""
}

print_step() {
    echo -e "${BLUE}[INFO]${NC} $1"
}

print_success() {
    echo -e "${GREEN}[OK]${NC} $1"
}

print_warn() {
    echo -e "${YELLOW}[WARN]${NC} $1"
}

print_error() {
    echo -e "${RED}[ERROR]${NC} $1"
}

# 检查 Docker 是否安装
check_docker() {
    if ! command -v docker &> /dev/null; then
        print_error "Docker 未安装，请先安装 Docker"
        exit 1
    fi
    
    if ! docker info &> /dev/null; then
        print_error "Docker 服务未运行，请启动 Docker 服务"
        exit 1
    fi
    
    print_success "Docker 环境检查通过"
}

# 检查 Buildx 是否可用
check_buildx() {
    if docker buildx version &> /dev/null; then
        print_success "Docker Buildx 已安装"
        return 0
    else
        print_warn "Docker Buildx 未安装，将无法构建多架构镜像"
        return 1
    fi
}

# 交互式选择架构
select_arch() {
    echo ""
    echo -e "${CYAN}请选择目标架构：${NC}"
    echo ""
    echo -e "  ${GREEN}1) amd64${NC}   - x86_64 架构（Intel/AMD 处理器，服务器主流）"
    echo -e "  ${GREEN}2) arm64${NC}   - ARM 架构（Apple Silicon M1/M2/M3、树莓派、ARM 服务器）"
    echo ""
    echo -e "  ${YELLOW}提示：如果不确定，请选择 amd64（大多数服务器环境）${NC}"
    echo ""
    
    while true; do
        read -r -p "请输入选项 [1/2]: " choice
        case "$choice" in
            1|amd64|x86_64|x86)
                ARCH="linux/amd64"
                ARCH_SHORT="amd64"
                print_step "已选择架构: ${GREEN}amd64${NC}"
                break
                ;;
            2|arm64|aarch64|arm)
                ARCH="linux/arm64"
                ARCH_SHORT="arm64"
                print_step "已选择架构: ${GREEN}arm64${NC}"
                break
                ;;
            *)
                print_warn "无效选项，请重新输入"
                ;;
        esac
    done
}

# 构建镜像
build_image() {
    echo ""
    print_step "开始构建 Docker 镜像..."
    print_step "镜像名称: ${IMAGE_NAME}:${IMAGE_TAG}"
    print_step "目标架构: ${ARCH_SHORT}"
    print_step "构建上下文: ${SCRIPT_DIR}"
    echo ""
    
    # 清理旧镜像
    if docker images --format "{{.Repository}}:{{.Tag}}" | grep -q "^${IMAGE_NAME}:${IMAGE_TAG}$"; then
        print_step "发现旧镜像，正在清理..."
        docker rmi "${IMAGE_NAME}:${IMAGE_TAG}" &> /dev/null || true
    fi
    
    # 执行构建
    if check_buildx; then
        # 使用 Buildx 构建指定架构
        docker buildx build \
            --platform "${ARCH}" \
            --tag "${IMAGE_NAME}:${IMAGE_TAG}" \
            --load \
            --progress=plain \
            -f docker/Dockerfile \
            . 2>&1 | tee "build_${ARCH_SHORT}.log"
    else
        # 回退到普通 docker build
        docker build \
            --tag "${IMAGE_NAME}:${IMAGE_TAG}" \
            -f docker/Dockerfile \
            . 2>&1 | tee "build_${ARCH_SHORT}.log"
    fi
    
    if [ "${PIPESTATUS[0]}" -eq 0 ]; then
        print_success "镜像构建成功！"
        return 0
    else
        print_error "镜像构建失败，日志已保存到 build_${ARCH_SHORT}.log"
        return 1
    fi
}

# 验证镜像
verify_image() {
    echo ""
    print_step "验证镜像..."
    
    if ! docker images --format "{{.Repository}}:{{.Tag}}" | grep -q "^${IMAGE_NAME}:${IMAGE_TAG}$"; then
        print_error "镜像未找到"
        return 1
    fi
    
    # 获取镜像信息
    IMAGE_ID=$(docker images --format "{{.ID}}" "${IMAGE_NAME}:${IMAGE_TAG}")
    IMAGE_SIZE=$(docker images --format "{{.Size}}" "${IMAGE_NAME}:${IMAGE_TAG}")
    CREATED=$(docker images --format "{{.CreatedAt}}" "${IMAGE_NAME}:${IMAGE_TAG}")
    
    print_success "镜像验证通过"
    echo ""
    echo -e "  ${CYAN}镜像名称:${NC} ${IMAGE_NAME}:${IMAGE_TAG}"
    echo -e "  ${CYAN}镜像 ID:${NC} ${IMAGE_ID}"
    echo -e "  ${CYAN}镜像大小:${NC} ${IMAGE_SIZE}"
    echo -e "  ${CYAN}创建时间:${NC} ${CREATED}"
    echo -e "  ${CYAN}目标架构:${NC} ${ARCH_SHORT}"
    echo ""
}

# 导出镜像（离线使用）
export_image() {
    echo ""
    print_step "正在导出镜像为 tar 包（用于离线传输）..."
    
    TAR_FILE="${IMAGE_NAME}_${ARCH_SHORT}_${IMAGE_TAG}.tar"
    
    docker save "${IMAGE_NAME}:${IMAGE_TAG}" -o "${TAR_FILE}"
    
    if [ -f "${TAR_FILE}" ]; then
        FILE_SIZE=$(du -h "${TAR_FILE}" | cut -f1)
        print_success "镜像导出成功！"
        echo ""
        echo -e "  ${CYAN}导出文件:${NC} ${TAR_FILE}"
        echo -e "  ${CYAN}文件大小:${NC} ${FILE_SIZE}"
        echo -e "  ${CYAN}存放路径:${NC} ${SCRIPT_DIR}/${TAR_FILE}"
        echo ""
    fi
}

# 询问是否导出
ask_export() {
    echo ""
    read -r -p "是否导出镜像为 tar 包以便离线传输? [y/N]: " answer
    if [[ "$answer" =~ ^[Yy]$ ]]; then
        export_image
    fi
}

# 显示使用说明
show_usage_info() {
    echo ""
    echo -e "${CYAN}═══════════════════════════════════════════════════════════════${NC}"
    echo -e "${CYAN}                       部署使用说明                            ${NC}"
    echo -e "${CYAN}═══════════════════════════════════════════════════════════════${NC}"
    echo ""
    echo -e "${GREEN}在线环境（有 Docker 网络）：${NC}"
    echo "  docker-compose up -d"
    echo ""
    echo -e "${GREEN}离线环境（无 Docker 网络）：${NC}"
    echo "  1. 将镜像 tar 包传输到目标服务器"
    echo "  2. 在目标服务器上执行:"
    echo "     docker load -i ${IMAGE_NAME}_${ARCH_SHORT}_${IMAGE_TAG}.tar"
    echo "  3. 将 docker-compose.yml 复制到目标服务器"
    echo "  4. 修改 docker-compose.yml 中 image 字段为已加载的镜像"
    echo "  5. 执行: docker-compose up -d"
    echo ""
    echo -e "${GREEN}验证服务：${NC}"
    echo "  curl http://localhost:5000/api/health"
    echo ""
    echo -e "${GREEN}查看日志：${NC}"
    echo "  docker logs -f tacit-knowledge-app"
    echo ""
    echo -e "${CYAN}═══════════════════════════════════════════════════════════════${NC}"
    echo ""
}

# 主流程
main() {
    print_banner
    
    # 检查 Docker
    check_docker
    
    # 选择架构
    select_arch
    
    # 构建镜像
    if ! build_image; then
        exit 1
    fi
    
    # 验证镜像
    verify_image
    
    # 询问导出
    ask_export
    
    # 显示使用说明
    show_usage_info
    
    print_success "构建完成！"
}

# 执行
main
