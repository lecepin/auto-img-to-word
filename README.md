# 🖼️ 图片转Word工具

一个自动将图片批量转换为Word文档的Node.js工具。

## ✨ 功能特点

- 📁 批量处理 `images` 文件夹中的所有图片
- 📏 智能缩放适配页面尺寸（支持横图和竖图）
- 📄 每张图片占用一页，自动分页
- 🎯 图片居中对齐，保持原始比例
- 🔄 支持多种图片格式：jpg、jpeg、png、gif、bmp、webp

## 🚀 快速开始

### 环境要求
- Node.js (建议 v14+)

### 安装依赖
```bash
pnpm install
```

### 使用方法

1. **准备图片**：将图片文件放入 `images` 文件夹
2. **运行程序**：
   - **Windows**: 双击 `start.bat`
   - **Mac**: 双击 `start.command`
   - **命令行**: `npm start`
3. **获取结果**：生成的Word文档为 `output.docx`

## 📋 项目结构

```
auto-cut-img-to-word/
├── images/          # 放置待处理的图片文件
├── index.js         # 主程序代码
├── start.bat        # Windows启动脚本
├── start.command    # Mac启动脚本
├── package.json     # 项目配置
└── output.docx      # 生成的Word文档
```

## 🔧 技术细节

- **页面设置**: A4纸张，窄边距（0.5英寸）
- **缩放逻辑**: 根据图片宽高比智能选择适配策略
- **依赖库**: docx（Word文档生成）、image-size（图片尺寸获取）

## 📝 License

MIT 