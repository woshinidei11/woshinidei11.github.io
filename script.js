let uploadedFiles = [];

document.addEventListener('DOMContentLoaded', function() {
    if (typeof pdfjsLib !== 'undefined') {
        pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
    }
    
    const uploadArea = document.getElementById('uploadArea');
    const fileInput = document.getElementById('fileInput');
    const fileList = document.getElementById('files');
    const clearBtn = document.getElementById('clearBtn');
    
    if (uploadArea) {
        uploadArea.addEventListener('click', function() {
            fileInput.click();
        });

        uploadArea.addEventListener('dragover', function(e) {
            e.preventDefault();
            uploadArea.classList.add('dragover');
        });

        uploadArea.addEventListener('dragleave', function(e) {
            e.preventDefault();
            uploadArea.classList.remove('dragover');
        });

        uploadArea.addEventListener('drop', function(e) {
            e.preventDefault();
            uploadArea.classList.remove('dragover');
            const files = e.dataTransfer.files;
            handleFiles(files);
        });

        fileInput.addEventListener('change', function(e) {
            const files = e.target.files;
            handleFiles(files);
        });
    }

    if (clearBtn) {
        clearBtn.addEventListener('click', clearFiles);
    }

    const convertBtn = document.getElementById('convertBtn');
    const compressBtn = document.getElementById('compressBtn');
    const mergeBtn = document.getElementById('mergeBtn');
    const wordToPdfBtn = document.getElementById('wordToPdfBtn');

    if (convertBtn) {
        convertBtn.addEventListener('click', convertPdfToWord);
    }

    if (compressBtn) {
        compressBtn.addEventListener('click', compressPdf);
    }

    if (mergeBtn) {
        mergeBtn.addEventListener('click', mergePdfs);
    }

    if (wordToPdfBtn) {
        wordToPdfBtn.addEventListener('click', convertWordToPdf);
    }
});

function handleFiles(files) {
    const isWordPage = document.getElementById('wordToPdfBtn') !== null;
    for (let i = 0; i < files.length; i++) {
        const file = files[i];
        const isValid = isWordPage 
            ? (file.name.endsWith('.docx') || file.name.endsWith('.doc'))
            : file.type === 'application/pdf';
        if (isValid) {
            const fileInfo = {
                id: Date.now() + i,
                file: file,
                name: file.name,
                size: formatSize(file.size)
            };
            uploadedFiles.push(fileInfo);
            addFileToList(fileInfo);
        }
    }
}

function formatSize(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

function addFileToList(fileInfo) {
    const fileList = document.getElementById('files');
    if (!fileList) return;

    const isWordPage = document.getElementById('wordToPdfBtn') !== null;
    const iconClass = isWordPage ? 'fas fa-file-word' : 'fas fa-file-pdf';
    const li = document.createElement('li');
    li.dataset.id = fileInfo.id;
    li.innerHTML = '<div class="file-name"><i class="' + iconClass + '"></i><span>' + fileInfo.name + '</span><span class="file-size">(' + fileInfo.size + ')</span></div><button class="remove-btn" onclick="removeFile(' + fileInfo.id + ')"><i class="fas fa-times"></i></button>';
    fileList.appendChild(li);
}

function removeFile(id) {
    const index = uploadedFiles.findIndex(f => f.id === id);
    if (index !== -1) {
        uploadedFiles.splice(index, 1);
        const li = document.querySelector('li[data-id="' + id + '"]');
        if (li) {
            li.remove();
        }
    }
}

function clearFiles() {
    uploadedFiles.length = 0;
    const fileList = document.getElementById('files');
    if (fileList) {
        fileList.innerHTML = '';
    }
}

function showProgress() {
    const progress = document.getElementById('progress');
    if (progress) {
        progress.classList.add('show');
        const progressText = progress.querySelector('.progress-text');
        let percent = 0;
        const interval = setInterval(function() {
            percent += Math.random() * 15;
            if (percent >= 100) {
                percent = 100;
                clearInterval(interval);
            }
            progressText.textContent = Math.floor(percent) + '%';
        }, 200);
    }
}

function hideProgress() {
    const progress = document.getElementById('progress');
    if (progress) {
        progress.classList.remove('show');
        progress.querySelector('.progress-text').textContent = '0%';
    }
}

function showResults(results) {
    const resultsDiv = document.getElementById('results');
    const resultList = document.getElementById('resultList');
    
    if (resultsDiv && resultList) {
        resultsDiv.classList.add('show');
        resultList.innerHTML = '';
        
        results.forEach(result => {
            const item = document.createElement('div');
            item.className = 'result-item';
            item.innerHTML = '<div><i class="fas fa-check-circle success-icon"></i><span>' + result.name + '</span></div><a href="' + result.url + '" download="' + result.downloadName + '" class="download-btn"><i class="fas fa-download"></i> 下载</a>';
            resultList.appendChild(item);
        });
    }
}

async function convertPdfToWord() {
    if (uploadedFiles.length === 0) {
        alert('请先上传PDF文件');
        return;
    }

    showProgress();

    setTimeout(async function() {
        const results = [];
        let firstError = null;

        for (const fileInfo of uploadedFiles) {
            try {
                const text = await extractTextFromPdf(fileInfo.file);
                const blob = createWordDocument(text);
                const url = URL.createObjectURL(blob);
                const downloadName = fileInfo.name.replace('.pdf', '.docx');
                
                results.push({
                    name: fileInfo.name,
                    url: url,
                    downloadName: downloadName
                });
            } catch (error) {
                console.error('转换失败:', error);
                if (!firstError) {
                    firstError = error;
                }
            }
        }

        hideProgress();
        
        if (results.length > 0) {
            showResults(results);
        } else {
            let errorMsg = '转换失败，请重试';
            if (firstError) {
                errorMsg = '转换失败: ' + (firstError.message || firstError);
            }
            alert(errorMsg);
        }
    }, 2000);
}

function createWordDocument(text) {
    const content = '<!DOCTYPE html><html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word"><head><meta charset="UTF-8"><title>PDF转换结果</title></head><body><p>' + text.replace(/\n/g, '</p><p>') + '</p></body></html>';
    const blob = new Blob(['\ufeff' + content], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
    return blob;
}

async function extractTextFromPdf(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = async function(e) {
            try {
                const arrayBuffer = e.target.result;
                const pdf = await pdfjsLib.getDocument({ 
                    data: arrayBuffer,
                    disableWorker: true
                }).promise;
                let text = '';
                
                for (let i = 1; i <= pdf.numPages; i++) {
                    const page = await pdf.getPage(i);
                    const textContent = await page.getTextContent();
                    textContent.items.forEach(item => {
                        text += item.str + '\n';
                    });
                }
                
                resolve(text);
            } catch (error) {
                reject(error);
            }
        };
        reader.onerror = reject;
        reader.readAsArrayBuffer(file);
    });
}

async function compressPdf() {
    if (uploadedFiles.length === 0) {
        alert('请先上传PDF文件');
        return;
    }

    showProgress();

    setTimeout(async function() {
        const quality = document.querySelector('input[name="quality"]:checked')?.value || 'medium';
        const results = [];

        for (const fileInfo of uploadedFiles) {
            try {
                const compressedBlob = await compressPdfFile(fileInfo.file, quality);
                const url = URL.createObjectURL(compressedBlob);
                const downloadName = fileInfo.name.replace('.pdf', '_compressed.pdf');
                
                results.push({
                    name: fileInfo.name,
                    url: url,
                    downloadName: downloadName
                });
            } catch (error) {
                console.error('压缩失败:', error);
            }
        }

        hideProgress();
        
        if (results.length > 0) {
            showResults(results);
        } else {
            alert('压缩失败，请重试');
        }
    }, 2000);
}

async function compressPdfFile(file, quality) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = async function(e) {
            try {
                const arrayBuffer = e.target.result;
                const pdf = await pdfjsLib.getDocument({ 
                    data: arrayBuffer,
                    disableWorker: true
                }).promise;
                const pages = [];
                
                for (let i = 1; i <= pdf.numPages; i++) {
                    const page = await pdf.getPage(i);
                    const viewport = page.getViewport({ scale: 1 });
                    const canvas = document.createElement('canvas');
                    const context = canvas.getContext('2d');
                    
                    canvas.width = viewport.width;
                    canvas.height = viewport.height;
                    
                    await page.render({
                        canvasContext: context,
                        viewport: viewport
                    }).promise;
                    
                    const imageData = canvas.toDataURL('image/jpeg', quality === 'low' ? 0.5 : quality === 'medium' ? 0.7 : 0.9);
                    pages.push(imageData);
                }
                
                const jsPDF = window.jspdf.jsPDF;
                const doc = new jsPDF({
                    orientation: pages.length > 0 ? 'portrait' : 'portrait',
                    unit: 'px',
                    format: [pages[0] ? 612 : 612, pages[0] ? 792 : 792]
                });
                
                for (let i = 0; i < pages.length; i++) {
                    if (i > 0) doc.addPage();
                    doc.addImage(pages[i], 'JPEG', 0, 0, 612, 792);
                }
                
                const blob = doc.output('blob');
                resolve(blob);
            } catch (error) {
                reject(error);
            }
        };
        reader.onerror = reject;
        reader.readAsArrayBuffer(file);
    });
}

async function mergePdfs() {
    if (uploadedFiles.length < 2) {
        alert('请至少上传2个PDF文件');
        return;
    }

    showProgress();

    setTimeout(async function() {
        try {
            const mergedBlob = await mergePdfFiles(uploadedFiles);
            const url = URL.createObjectURL(mergedBlob);
            const downloadName = 'merged.pdf';
            
            showResults([{
                name: '合并后的文件',
                url: url,
                downloadName: downloadName
            }]);
        } catch (error) {
            console.error('合并失败:', error);
            alert('合并失败，请重试');
        }

        hideProgress();
    }, 2000);
}

async function mergePdfFiles(fileInfos) {
    return new Promise((resolve, reject) => {
        const promises = fileInfos.map(fileInfo => {
            return new Promise((res, rej) => {
                const reader = new FileReader();
                reader.onload = function(e) {
                    res(e.target.result);
                };
                reader.onerror = rej;
                reader.readAsArrayBuffer(fileInfo.file);
            });
        });

        Promise.all(promises).then(async function(arrayBuffers) {
            try {
                const jsPDF = window.jspdf.jsPDF;
                const doc = new jsPDF();
                
                for (let i = 0; i < arrayBuffers.length; i++) {
                    const pdf = await pdfjsLib.getDocument({ 
                        data: arrayBuffers[i],
                        disableWorker: true
                    }).promise;
                    
                    for (let j = 1; j <= pdf.numPages; j++) {
                        if (i === 0 && j === 1) {
                            const page = await pdf.getPage(j);
                            const viewport = page.getViewport({ scale: 1 });
                            const canvas = document.createElement('canvas');
                            const context = canvas.getContext('2d');
                            
                            canvas.width = viewport.width;
                            canvas.height = viewport.height;
                            
                            await page.render({
                                canvasContext: context,
                                viewport: viewport
                            }).promise;
                            
                            const imageData = canvas.toDataURL('image/jpeg', 0.9);
                            doc.addImage(imageData, 'JPEG', 0, 0, 210, 297);
                        } else {
                            doc.addPage();
                            const page = await pdf.getPage(j);
                            const viewport = page.getViewport({ scale: 1 });
                            const canvas = document.createElement('canvas');
                            const context = canvas.getContext('2d');
                            
                            canvas.width = viewport.width;
                            canvas.height = viewport.height;
                            
                            await page.render({
                                canvasContext: context,
                                viewport: viewport
                            }).promise;
                            
                            const imageData = canvas.toDataURL('image/jpeg', 0.9);
                            doc.addImage(imageData, 'JPEG', 0, 0, 210, 297);
                        }
                    }
                }
                
                const blob = doc.output('blob');
                resolve(blob);
            } catch (error) {
                reject(error);
            }
        }).catch(reject);
    });
}

async function convertWordToPdf() {
    if (uploadedFiles.length === 0) {
        alert('请先上传Word文件');
        return;
    }

    showProgress();

    setTimeout(async function() {
        const results = [];
        let firstError = null;

        for (const fileInfo of uploadedFiles) {
            try {
                const blob = await convertWordFileToPdf(fileInfo.file);
                const url = URL.createObjectURL(blob);
                const downloadName = fileInfo.name.replace(/\.docx?$/, '.pdf');

                results.push({
                    name: fileInfo.name,
                    url: url,
                    downloadName: downloadName
                });
            } catch (error) {
                console.error('转换失败:', fileInfo.name, error);
                if (!firstError) {
                    firstError = error;
                }
            }
        }

        hideProgress();

        if (results.length > 0) {
            showResults(results);
        } else {
            let errorMsg = '转换失败，请重试';
            if (firstError) {
                errorMsg = '转换失败: ' + (firstError.message || firstError);
            }
            alert(errorMsg);
        }
    }, 2000);
}

async function convertWordFileToPdf(file) {
    const arrayBuffer = await file.arrayBuffer();
    const result = await mammoth.extractRawText({ arrayBuffer: arrayBuffer });
    const text = result.value;

    if (!text || text.trim().length === 0) {
        throw new Error('文档内容为空');
    }

    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({
        orientation: 'portrait',
        unit: 'mm',
        format: 'a4'
    });

    const pageWidth = 210;
    const pageHeight = 297;
    const margin = 20;
    const maxWidth = pageWidth - margin * 2;
    const lineHeight = 7;
    let y = margin;

    const lines = text.split('\n');

    for (const line of lines) {
        if (line.trim() === '') {
            y += lineHeight / 2;
            continue;
        }

        const splitLines = doc.splitTextToSize(line, maxWidth);
        for (const splitLine of splitLines) {
            if (y > pageHeight - margin) {
                doc.addPage();
                y = margin;
            }
            doc.text(splitLine, margin, y);
            y += lineHeight;
        }
    }

    return doc.output('blob');
}
