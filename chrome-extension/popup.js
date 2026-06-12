// 飞书文档转 Markdown - Popup 转换器
// 使用 mammoth.js (docx→HTML) + turndown.js (HTML→Markdown) 进行转换

document.addEventListener('DOMContentLoaded', function() {
    const dropZone = document.getElementById('dropZone');
    const fileInput = document.getElementById('fileInput');
    const fileInfo = document.getElementById('fileInfo');
    const fileName = document.getElementById('fileName');
    const fileSize = document.getElementById('fileSize');
    const convertBtn = document.getElementById('convertBtn');
    const statusEl = document.getElementById('status');

    // 配置选项
    const configMarkdown = document.getElementById('config-markdown');
    const configWord = document.getElementById('config-word');
    const configPDF = document.getElementById('config-pdf');
    const imageTypeOption = document.getElementById('image-type-option');
    const imageLocal = document.getElementById('image-local');
    const imageBase64 = document.getElementById('image-base64');

    let selectedFile = null;

    // 默认配置
    const defaultConfig = {
        showMarkdown: true,
        showWord: true,
        showPDF: true,
        imageType: 'local'
    };

    // 加载配置
    chrome.storage.local.get(defaultConfig, (items) => {
        configMarkdown.checked = items.showMarkdown;
        configWord.checked = items.showWord;
        configPDF.checked = items.showPDF;

        if (items.imageType === 'base64') {
            imageBase64.checked = true;
        } else {
            imageLocal.checked = true;
        }

        updateImageTypeVisibility();
    });

    // 更新图片类型选项可见性
    function updateImageTypeVisibility() {
        if (configMarkdown.checked) {
            imageTypeOption.classList.add('show');
        } else {
            imageTypeOption.classList.remove('show');
        }
    }

    // 保存配置
    function saveConfig() {
        const config = {
            showMarkdown: configMarkdown.checked,
            showWord: configWord.checked,
            showPDF: configPDF.checked,
            imageType: imageBase64.checked ? 'base64' : 'local'
        };
        chrome.storage.local.set(config, () => {
            console.log('配置已保存:', config);
        });
        updateImageTypeVisibility();
    }

    configMarkdown.addEventListener('change', saveConfig);
    configWord.addEventListener('change', saveConfig);
    configPDF.addEventListener('change', saveConfig);
    imageLocal.addEventListener('change', saveConfig);
    imageBase64.addEventListener('change', saveConfig);

    // 检查依赖库是否加载
    function checkDependencies() {
        const missing = [];
        if (typeof mammoth === 'undefined') missing.push('mammoth.browser.min.js');
        if (typeof JSZip === 'undefined') missing.push('jszip.min.js');
        if (typeof TurndownService === 'undefined') missing.push('turndown.min.js');
        return missing;
    }

    // 点击区域选择文件
    dropZone.addEventListener('click', () => fileInput.click());

    fileInput.addEventListener('change', (e) => {
        if (e.target.files[0]) handleFile(e.target.files[0]);
    });

    // 拖拽支持
    dropZone.addEventListener('dragenter', (e) => {
        e.preventDefault();
        dropZone.classList.add('drag-over');
    });

    dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
    });

    dropZone.addEventListener('dragleave', (e) => {
        if (!dropZone.contains(e.relatedTarget)) {
            dropZone.classList.remove('drag-over');
        }
    });

    dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropZone.classList.remove('drag-over');
        const file = e.dataTransfer.files[0];
        if (file && (file.name.toLowerCase().endsWith('.docx') ||
            file.type === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')) {
            handleFile(file);
        } else {
            showStatus('❌ 请选择 .docx 格式的 Word 文件', 'error');
        }
    });

    function handleFile(file) {
        selectedFile = file;
        fileInfo.classList.add('show');
        fileName.textContent = '📄 ' + file.name;
        fileSize.textContent = formatFileSize(file.size);
        convertBtn.disabled = false;
        convertBtn.textContent = '开始转换';
        hideStatus();
    }

    function formatFileSize(bytes) {
        if (bytes < 1024) return bytes + ' B';
        if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
        return (bytes / 1024 / 1024).toFixed(1) + ' MB';
    }

    convertBtn.addEventListener('click', async () => {
        if (!selectedFile) return;

        // 检查依赖
        const missing = checkDependencies();
        if (missing.length > 0) {
            showStatus(
                '❌ 缺少依赖库，请下载并放到 libs/ 目录：\n' + missing.join(', '),
                'error'
            );
            return;
        }

        convertBtn.disabled = true;
        convertBtn.textContent = '转换中...';

        try {
            const outputName = selectedFile.name.replace(/\.docx$/i, '');
            const imageType = imageBase64.checked ? 'base64' : 'local';
            showStatus('⏳ 解析文档...', 'loading');
            const { markdown, images, warnings } = await convertDocxToMarkdown(selectedFile, imageType);

            const isBase64Mode = imageType === 'base64';
            const imgCountText = isBase64Mode
                ? '图片已嵌入 Markdown'
                : `${images.length} 张图片`;
            showStatus(`⏳ 打包文件 (${imgCountText})...`, 'loading');
            const blob = await packageOutput(outputName, markdown, images, isBase64Mode);

            triggerDownload(blob, `${outputName}.zip`);

            const baseMsg = isBase64Mode
                ? `✅ 转换完成！图片已 Base64 嵌入 Markdown`
                : `✅ 转换完成！${images.length} 张图片`;
            const msg = warnings.length > 0
                ? `${baseMsg}\n⚠ 有 ${warnings.length} 条提示，请查看控制台`
                : baseMsg;
            showStatus(msg, 'success');
            convertBtn.textContent = '再次转换';
        } catch (error) {
            console.error('Conversion failed:', error);
            showStatus('❌ 转换失败：' + error.message, 'error');
            convertBtn.textContent = '重试';
        } finally {
            convertBtn.disabled = false;
        }
    });

    async function convertDocxToMarkdown(file, imageType) {
        const arrayBuffer = await file.arrayBuffer();
        const images = [];
        const useBase64 = imageType === 'base64';

        // 本地图片模式：先从 docx 的 word/media 中直接提取所有图片，
        // 这样不依赖 mammoth 的图片回退行为，避免漏图。
        if (!useBase64) {
            try {
                const zip = await JSZip.loadAsync(arrayBuffer);
                const mediaFolder = zip.folder('word/media');
                if (mediaFolder) {
                    let index = 1;
                    for (const [filename, entry] of Object.entries(mediaFolder.files)) {
                        if (entry.dir) continue;
                        const ext = (filename.split('.').pop() || '').toLowerCase();
                        if (!['png', 'jpg', 'jpeg', 'gif', 'webp', 'svg', 'bmp'].includes(ext)) continue;
                        const base64Data = await entry.async('base64');
                        const imgFilename = `image_${String(index).padStart(3, '0')}.${ext}`;
                        images.push({
                            filename: imgFilename,
                            base64: base64Data,
                            contentType: `image/${ext === 'jpg' ? 'jpeg' : ext}`
                        });
                        index++;
                    }
                    console.log(`从 docx 中提取了 ${images.length} 张图片`);
                }
            } catch (error) {
                console.warn('从 docx 预提取图片失败，将回退到 mammoth 提取:', error);
            }
        }

        // 把 base64 数据落盘为本地图片，返回相对路径（用于本地模式 / data URI 兜底）
        function registerLocalImage(base64Data, contentType) {
            const existing = images.findIndex(img => img.base64 === base64Data);
            if (existing >= 0) {
                return `./images/${images[existing].filename}`;
            }
            const ext = ((contentType || 'image/png').split('/')[1] || 'png').split('+')[0];
            const validExt = ['png', 'jpg', 'jpeg', 'gif', 'webp', 'svg'].includes(ext) ? ext : 'png';
            const imgFilename = `image_${String(images.length + 1).padStart(3, '0')}.${validExt}`;
            images.push({ filename: imgFilename, base64: base64Data, contentType: contentType || 'image/png' });
            return `./images/${imgFilename}`;
        }

        // Step 1: 使用 mammoth.js 将 docx 转换为 HTML，同时处理图片
        const result = await mammoth.convertToHtml({
            arrayBuffer: arrayBuffer,
            convertImage: mammoth.images.imgElement(function(image) {
                return image.read('base64').then(function(base64Data) {
                    const contentType = image.contentType || 'image/png';
                    if (useBase64) {
                        // Base64 模式：直接嵌入 markdown
                        return { src: `data:${contentType};base64,${base64Data}` };
                    }
                    // 本地模式：落盘并使用相对路径
                    return { src: registerLocalImage(base64Data, contentType) };
                });
            })
        });

        const warnings = result.messages.filter(m => m.type === 'warning');
        if (warnings.length > 0) {
            console.warn('Mammoth conversion warnings:', warnings);
        }

        // Step 2: 使用 turndown.js 将 HTML 转换为 Markdown
        const turndownService = new TurndownService({
            headingStyle: 'atx',
            codeBlockStyle: 'fenced',
            bulletListMarker: '-',
            emDelimiter: '*'
        });

        // 图片规则：本地模式下对残留的 data URI 做兜底，转为本地图片
        const dataUriPattern = /^data:([^;,]+)?(?:;[^,]*)*;base64,(.*)$/i;
        turndownService.addRule('images', {
            filter: 'img',
            replacement: function(content, node) {
                const alt = node.getAttribute('alt') || '';
                let src = node.getAttribute('src') || '';
                if (!src) return '';

                if (!useBase64) {
                    const m = src.match(dataUriPattern);
                    if (m) {
                        // 选了本地图片但 src 仍是 base64：落盘并改写为相对路径
                        const contentType = m[1] || 'image/png';
                        const base64Data = m[2];
                        src = registerLocalImage(base64Data, contentType);
                    }
                }
                return `\n![${alt}](${src})\n`;
            }
        });

        // 代码块规则
        turndownService.addRule('preCode', {
            filter: function(node) {
                return node.nodeName === 'PRE' && node.querySelector('code');
            },
            replacement: function(content, node) {
                const code = node.querySelector('code');
                const lang = (code.className.match(/language-(\w+)/) || [])[1] || '';
                return `\n\`\`\`${lang}\n${code.textContent.trim()}\n\`\`\`\n`;
            }
        });

        const markdown = turndownService.turndown(result.value);

        return { markdown, images, warnings };
    }

    async function packageOutput(docName, markdown, images, isBase64Mode) {
        const zip = new JSZip();
        zip.file(`${docName}.md`, markdown);

        // Base64 模式下图片已内嵌在 markdown 中，不再生成 images 目录
        if (!isBase64Mode && images.length > 0) {
            const imgFolder = zip.folder('images');
            for (const img of images) {
                // base64 转 Uint8Array
                const binaryStr = atob(img.base64);
                const bytes = new Uint8Array(binaryStr.length);
                for (let i = 0; i < binaryStr.length; i++) {
                    bytes[i] = binaryStr.charCodeAt(i);
                }
                imgFolder.file(img.filename, bytes);
            }
        }

        return zip.generateAsync({
            type: 'blob',
            compression: 'DEFLATE',
            compressionOptions: { level: 6 }
        });
    }

    function triggerDownload(blob, filename) {
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        setTimeout(() => URL.revokeObjectURL(url), 1000);
    }

    function showStatus(message, type) {
        statusEl.textContent = message;
        statusEl.className = `status ${type} show`;
    }

    function hideStatus() {
        statusEl.className = 'status';
    }
});
