const fs = require('fs');
const path = require('path');
const sizeOf = require('image-size');
const { Document, Packer, Paragraph, ImageRun, PageBreak, AlignmentType, Media } = require('docx');

async function createWordFromImages() {
    try {
        // 读取images目录中的所有文件
        const imagesDir = path.join(__dirname, 'images');
        console.log('正在读取images目录...');
        
        if (!fs.existsSync(imagesDir)) {
            console.error('images目录不存在！');
            return;
        }

        const files = fs.readdirSync(imagesDir);
        
        // 过滤出图片文件（支持常见的图片格式）
        const imageExtensions = ['.jpg', '.jpeg', '.png', '.gif', '.bmp', '.webp'];
        const imageFiles = files.filter(file => {
            const ext = path.extname(file).toLowerCase();
            return imageExtensions.includes(ext);
        });

        if (imageFiles.length === 0) {
            console.log('images目录中没有找到图片文件！');
            return;
        }

        // 按名称升序排序
        imageFiles.sort((a, b) => a.localeCompare(b));
        console.log(`找到 ${imageFiles.length} 个图片文件，按名称升序排序：`);
        imageFiles.forEach(file => console.log(`  - ${file}`));

        // 创建文档内容数组
        const documentChildren = [];

        // 为每张图片创建页面
        for (let i = 0; i < imageFiles.length; i++) {
            const imageFile = imageFiles[i];
            const imagePath = path.join(imagesDir, imageFile);
            
            console.log(`正在处理图片: ${imageFile}`);

            try {
                // 读取图片文件
                const imageBuffer = fs.readFileSync(imagePath);
                
                // 获取图片原始尺寸
                const dimensions = sizeOf(imagePath);
                const originalWidth = dimensions.width;
                const originalHeight = dimensions.height;
                
                // 根据docx库文档，transformation使用96 DPI的像素系统
                // A4页面: 8.27" x 11.69" 
                // 窄边距: 0.5英寸，所以可用宽度 = 8.27 - 1.0 = 7.27英寸，可用高度 = 11.69 - 1.0 = 10.69英寸
                // 按96 DPI计算
                const pageAvailableWidthPixels = Math.round(7.27 * 96);   // 698像素
                const pageAvailableHeightPixels = Math.round(10.69 * 96); // 1026像素
                
                // 智能缩放逻辑：比较图片和页面的宽高比，选择合适的缩放方式
                const imageAspectRatio = originalWidth / originalHeight;
                const pageAspectRatio = pageAvailableWidthPixels / pageAvailableHeightPixels;
                
                let newWidth, newHeight;
                
                if (imageAspectRatio > pageAspectRatio) {
                    // 图片相对更宽，以宽度为准缩放
                    newWidth = pageAvailableWidthPixels;
                    newHeight = originalHeight * (pageAvailableWidthPixels / originalWidth);
                    console.log(`  图片较宽，以页面宽度为准缩放`);
                } else {
                    // 图片相对更高，以高度为准缩放
                    newHeight = pageAvailableHeightPixels;
                    newWidth = originalWidth * (pageAvailableHeightPixels / originalHeight);
                    console.log(`  图片较高，以页面高度为准缩放`);
                }
                
                console.log(`  原始尺寸: ${originalWidth}x${originalHeight} 像素`);
                console.log(`  页面可用区域: 7.27英寸(宽) x 10.69英寸(高) = ${pageAvailableWidthPixels}x${pageAvailableHeightPixels}像素 (96 DPI)`);
                console.log(`  图片宽高比: ${imageAspectRatio.toFixed(2)}, 页面宽高比: ${pageAspectRatio.toFixed(2)}`);
                console.log(`  最终尺寸: ${Math.round(newWidth)} x ${Math.round(newHeight)} 像素`);

                // 创建图片段落（居中对齐）
                const imageParagraph = new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                        new ImageRun({
                            data: imageBuffer,
                            transformation: {
                                width: Math.round(newWidth),
                                height: Math.round(newHeight)
                            }
                        })
                    ]
                });

                // 添加图片段落
                documentChildren.push(imageParagraph);

                // 如果不是最后一张图片，添加分页符
                if (i < imageFiles.length - 1) {
                    documentChildren.push(
                        new Paragraph({
                            children: [new PageBreak()]
                        })
                    );
                }

            } catch (error) {
                console.error(`处理图片 ${imageFile} 时出错:`, error.message);
                continue;
            }
        }

        // 创建Word文档
        const doc = new Document({
            sections: [{
                properties: {
                    page: {
                        // A4尺寸：8.27 x 11.69 inches
                        size: {
                            orientation: "portrait"
                        },
                        // 窄边距：0.5英寸 = 720缇
                        margin: {
                            top: 720,
                            right: 720,
                            bottom: 720,
                            left: 720
                        }
                    }
                },
                children: documentChildren
            }]
        });

        // 生成Word文档
        console.log('正在生成Word文档...');
        const buffer = await Packer.toBuffer(doc);
        
        // 保存文件
        const outputPath = path.join(__dirname, 'output.docx');
        fs.writeFileSync(outputPath, buffer);
        
        console.log(`✅ Word文档已成功生成！`);
        console.log(`📄 文件路径: ${outputPath}`);
        console.log(`📊 共处理 ${imageFiles.length} 张图片`);

    } catch (error) {
        console.error('❌ 生成Word文档时出错:', error);
    }
}

// 运行函数
console.log('🚀 开始生成Word文档...');
createWordFromImages(); 