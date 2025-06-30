#!/bin/bash

# 获取脚本所在目录
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

# 进入脚本所在目录
cd "$SCRIPT_DIR"

echo "🚀 启动图片转Word工具..."
echo "📁 当前目录: $PWD"
echo ""

# 检查依赖是否安装
if ! command -v node &> /dev/null; then
    echo "❌ Node.js 未安装，请先安装 Node.js"
    read -p "按任意键退出..."
    exit 1
fi

# 运行程序
echo "🎯 开始处理图片..."
npm start

echo ""
echo "✅ 处理完成！"
echo "📄 生成的Word文档: output.docx"
echo ""

# 打开生成的Word文档
open "output.docx"

echo ""
echo "🔄 3秒后自动关闭终端窗口..."
sleep 3

# 自动关闭当前终端窗口
osascript -e 'tell application "Terminal" to close first window' > /dev/null 2>&1
        