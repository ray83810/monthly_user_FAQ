document.addEventListener('DOMContentLoaded', () => {
    const dropzone = document.getElementById('dropzone');
    const fileInput = document.getElementById('file-input');
    const fileInfo = document.getElementById('file-info');
    const fileName = document.getElementById('file-name');
    const fileSize = document.getElementById('file-size');
    const cancelBtn = document.getElementById('cancel-btn');
    const processBtn = document.getElementById('process-btn');
    const progressContainer = document.getElementById('progress-container');
    const progressStatus = document.getElementById('progress-status');
    const progressPercentage = document.getElementById('progress-percentage');
    const progressBarFill = document.getElementById('progress-bar-fill');
    const progressDetail = document.getElementById('progress-detail');
    const errorContainer = document.getElementById('error-container');
    const errorMessage = document.getElementById('error-message');
    const retryBtn = document.getElementById('retry-btn');
    const resultContainer = document.getElementById('result-container');
    const resultSummary = document.getElementById('result-summary');
    const downloadLink = document.getElementById('download-link');
    const resetBtn = document.getElementById('reset-btn');

    let selectedFile = null;
    let processedBlob = null;
    let outputFilename = '每月案件問題分享.xlsx';

    // Helper: format file size
    function formatBytes(bytes, decimals = 2) {
        if (bytes === 0) return '0 Bytes';
        const k = 1024;
        const dm = decimals < 0 ? 0 : decimals;
        const sizes = ['Bytes', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(dm)) + ' ' + sizes[i];
    }

    // Dropzone drag-and-drop events
    ['dragenter', 'dragover'].forEach(eventName => {
        dropzone.addEventListener(eventName, (e) => {
            e.preventDefault();
            e.stopPropagation();
            dropzone.classList.add('dragover');
        }, false);
    });

    ['dragleave', 'drop'].forEach(eventName => {
        dropzone.addEventListener(eventName, (e) => {
            e.preventDefault();
            e.stopPropagation();
            dropzone.classList.remove('dragover');
        }, false);
    });

    dropzone.addEventListener('drop', (e) => {
        const dt = e.dataTransfer;
        const files = dt.files;
        if (files.length) {
            handleFileSelect(files[0]);
        }
    });

    // Clicking dropzone
    dropzone.addEventListener('click', () => {
        fileInput.click();
    });

    fileInput.addEventListener('change', (e) => {
        if (e.target.files.length) {
            handleFileSelect(e.target.files[0]);
        }
    });

    // Handle selected file
    function handleFileSelect(file) {
        if (!file.name.endsWith('.xlsx')) {
            alert('只支援 Excel 報表檔案 (.xlsx)');
            return;
        }
        selectedFile = file;
        fileName.textContent = file.name;
        fileSize.textContent = formatBytes(file.size);

        dropzone.classList.add('hidden');
        fileInfo.classList.remove('hidden');
        errorContainer.classList.add('hidden');
        resultContainer.classList.add('hidden');
    }

    // Cancel selection
    cancelBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        resetUploadState();
    });

    function resetUploadState() {
        selectedFile = null;
        processedBlob = null;
        fileInput.value = '';
        dropzone.classList.remove('hidden');
        fileInfo.classList.add('hidden');
        progressContainer.classList.add('hidden');
        errorContainer.classList.add('hidden');
        resultContainer.classList.add('hidden');
    }

    // Symptoms Parser
    function parseSymptoms(text) {
        if (!text || typeof text !== 'string') return ["", ""];
        text = text.trim();
        
        const reasonRegex = /Reason\s*:\s*/i;
        const actionRegex = /Action\s*:\s*/i;
        const toolsRegex = /Tools\s*:\s*/i;
        const endResultRegex = /End\s+Result\s*:\s*/i;
        
        const reasonMatch = text.match(reasonRegex);
        const actionMatch = text.match(actionRegex);
        
        let reasonVal = "";
        let actionVal = "";
        
        if (reasonMatch && actionMatch) {
            const reasonStart = reasonMatch.index + reasonMatch[0].length;
            const reasonEnd = actionMatch.index;
            reasonVal = text.substring(reasonStart, reasonEnd).trim();
            
            const actionStart = actionMatch.index + actionMatch[0].length;
            const nextDelims = [];
            
            const toolsMatch = text.substring(actionStart).match(toolsRegex);
            if (toolsMatch) {
                nextDelims.push(actionStart + toolsMatch.index);
            }
            const endResultMatch = text.substring(actionStart).match(endResultRegex);
            if (endResultMatch) {
                nextDelims.push(actionStart + endResultMatch.index);
            }
            
            const actionEnd = nextDelims.length > 0 ? Math.min(...nextDelims) : text.length;
            actionVal = text.substring(actionStart, actionEnd).trim();
        } else {
            reasonVal = text;
            actionVal = "";
        }
        return [reasonVal, actionVal];
    }

    // Device Validator
    function isValidDevice(device) {
        if (!device) return false;
        const d = device.toString().replace(/\t/g, ' ').trim().toUpperCase();
        if (["", "N/A", "NA", "NAN", "NULL", "NONE"].includes(d)) return false;
        if (d.includes("PREMIUM") || d.includes("方案") || d.includes("合約") || d.includes("EWS")) return false;
        return true;
    }

    // Process file client-side
    processBtn.addEventListener('click', async () => {
        if (!selectedFile) return;

        fileInfo.classList.add('hidden');
        progressContainer.classList.remove('hidden');
        updateProgress(10, '正在讀取報表檔案...', '正在加載 Excel 活頁簿...');

        try {
            const reader = new FileReader();
            reader.onload = async function(e) {
                try {
                    const buffer = e.target.result;
                    const workbook = new ExcelJS.Workbook();
                    await workbook.xlsx.load(buffer);
                    
                    updateProgress(30, '正在解析與過濾資料...', '正在套用篩選條件...');
                    
                    const worksheet = workbook.worksheets[0];
                    if (!worksheet) {
                        throw new Error('Excel 檔案中找不到任何工作表。');
                    }
                    
                    // Read Headers
                    const headerIndices = {};
                    const headerRow = worksheet.getRow(1);
                    headerRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
                        const val = (cell.value || "").toString().trim();
                        if (val) {
                            headerIndices[val] = colNumber;
                        }
                    });
                    
                    // Check required headers
                    const required = ['WRAP_UP_CODE', 'ENGAGEMENT_TYPE', 'session_reason', 'session_sub_reason', 'symptoms'];
                    for (const r of required) {
                        if (!headerIndices[r]) {
                            throw new Error(`找不到必要的欄位：${r}`);
                        }
                    }
                    
                    // Helper to get values
                    const getVal = (row, headerName) => {
                        const idx = headerIndices[headerName];
                        if (!idx) return "";
                        const val = row.getCell(idx).value;
                        if (val && typeof val === 'object') {
                            if (val.richText) {
                                return val.richText.map(t => t.text).join("");
                            }
                            if (val.result !== undefined) {
                                return val.result;
                            }
                            if (val.text) {
                                return val.text;
                            }
                        }
                        return val === null || val === undefined ? "" : val;
                    };
                    
                    // Parse Date/Month from session_start_time
                    let yearMonth = "202605"; // fallback
                    const timeIdx = headerIndices['session_start_time'];
                    if (timeIdx) {
                        for (let r = 2; r <= worksheet.maxRow; r++) {
                            const timeVal = worksheet.getRow(r).getCell(timeIdx).value;
                            if (timeVal) {
                                const d = new Date(timeVal);
                                if (!isNaN(d.getTime())) {
                                    const y = d.getFullYear();
                                    const m = (d.getMonth() + 1).toString().padStart(2, '0');
                                    yearMonth = `${y}${m}`;
                                    break;
                                }
                            }
                        }
                    }
                    
                    const solutoRows = [];
                    const ewsRows = [];
                    
                    // Process Data Rows
                    for (let r = 2; r <= worksheet.maxRow; r++) {
                        const row = worksheet.getRow(r);
                        
                        // Check if row has data
                        const hasVal = row.values.some(v => v !== null && v !== undefined && v !== "");
                        if (!hasVal) continue;
                        
                        const wrapUp = (getVal(row, 'WRAP_UP_CODE') || "").toString().trim().toUpperCase();
                        if (wrapUp !== 'RESOLVED') continue;
                        
                        const engTypeRaw = (getVal(row, 'ENGAGEMENT_TYPE') || "").toString().trim();
                        if (engTypeRaw.toUpperCase().includes('SUR')) continue;
                        
                        const reason = (getVal(row, 'session_reason') || "").toString().trim();
                        const subReason = (getVal(row, 'session_sub_reason') || "").toString().trim();
                        
                        // Exclude primary == 未提供服務, sub == 服務內容相關問題 / 個案外撥
                        if (reason === '未提供服務' || subReason === '服務內容相關問題' || subReason === '個案外撥') {
                            continue;
                        }
                        
                        // Classify EWS vs SOLUTO
                        const carrier = (getVal(row, 'Carrier') || "").toString().trim().toUpperCase();
                        const isEwsCarrier = carrier.includes('CHT_HOME') || carrier.includes('SENAO');
                        const isSoluto = engTypeRaw.toUpperCase().includes('SOLUTO') && !isEwsCarrier;
                        
                        // Parsing symptoms
                        const symptomsText = (getVal(row, 'symptoms') || "").toString().trim();
                        const [reasonVal, actionVal] = parseSymptoms(symptomsText);
                        
                        // Device Cleansing
                        const rawDevice = getVal(row, 'Device') || "";
                        const deviceDisplay = isValidDevice(rawDevice) ? rawDevice.toString().replace(/\t/g, ' ').trim() : "";
                        
                        const item = {
                            '主分類': reason,
                            '次分類': subReason,
                            '裝置': deviceDisplay,
                            '用戶問題': reasonVal,
                            '問題說明': actionVal,
                            '相關補充': ""
                        };
                        
                        if (isSoluto) {
                            solutoRows.push(item);
                        } else {
                            ewsRows.push(item);
                        }
                    }
                    
                    updateProgress(70, '正在產生新工作表與樣式...', '正在寫入並套用 Excel 格式與設計系統...');
                    
                    // Create output workbook
                    const outWorkbook = new ExcelJS.Workbook();
                    const wsSoluto = outWorkbook.addWorksheet(`${yearMonth}_SOLUTO`, { views: [{ showGridLines: true }] });
                    const wsEws = outWorkbook.addWorksheet(`${yearMonth}_EWS相關`, { views: [{ showGridLines: true }] });
                    
                    const writeSheet = (ws, data) => {
                        ws.columns = [
                            { header: '主分類', key: '主分類', width: 18 },
                            { header: '次分類', key: '次分類', width: 24 },
                            { header: '裝置', key: '裝置', width: 20 },
                            { header: '用戶問題', key: '用戶問題', width: 45 },
                            { header: '問題說明', key: '問題說明', width: 45 },
                            { header: '相關補充', key: '相關補充', width: 45 }
                        ];
                        
                        // Header Styling
                        const headerRow = ws.getRow(1);
                        headerRow.height = 25;
                        
                        const headerFills = [
                            'FFEAD1DC', // Pink
                            'FFB7B7B7', // Grey
                            'FFD0E0E3', // Cyan
                            'FFFCE5CD', // Orange
                            'FFB6D7A8', // Green
                            'FFFFFFFF'  // White
                        ];
                        
                        for (let c = 1; c <= 6; c++) {
                            const cell = headerRow.getCell(c);
                            cell.font = { name: 'Noto Sans TC Medium', size: 12, bold: true, color: { argb: 'FF000000' } };
                            cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
                            cell.fill = {
                                type: 'pattern',
                                pattern: 'solid',
                                fgColor: { argb: headerFills[c - 1] }
                            };
                            cell.border = {
                                top: { style: 'thin', color: { argb: 'FFD3D3D3' } },
                                left: { style: 'thin', color: { argb: 'FFD3D3D3' } },
                                bottom: { style: 'thin', color: { argb: 'FFD3D3D3' } },
                                right: { style: 'thin', color: { argb: 'FFD3D3D3' } }
                            };
                        }
                        
                        // Data Rows
                        data.forEach(item => {
                            const row = ws.addRow([
                                item['主分類'],
                                item['次分類'],
                                item['裝置'],
                                item['用戶問題'],
                                item['問題說明'],
                                item['相關補充']
                            ]);
                            
                            row.height = 40;
                            
                            for (let c = 1; c <= 6; c++) {
                                const cell = row.getCell(c);
                                cell.font = { name: 'Noto Sans TC Medium', size: 12, bold: false, color: { argb: 'FF000000' } };
                                cell.alignment = { horizontal: 'left', vertical: 'center', wrapText: true };
                                cell.border = {
                                    top: { style: 'thin', color: { argb: 'FFD3D3D3' } },
                                    left: { style: 'thin', color: { argb: 'FFD3D3D3' } },
                                    bottom: { style: 'thin', color: { argb: 'FFD3D3D3' } },
                                    right: { style: 'thin', color: { argb: 'FFD3D3D3' } }
                                };
                            }
                        });
                    };
                    
                    writeSheet(wsSoluto, solutoRows);
                    writeSheet(wsEws, ewsRows);
                    
                    updateProgress(90, '正在生成 Excel 成品...', '準備進行下載下載...');
                    
                    const outBuffer = await outWorkbook.xlsx.writeBuffer();
                    processedBlob = new Blob([outBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
                    outputFilename = `每月案件問題分享 - ${yearMonth}.xlsx`;
                    
                    const totalProcessed = solutoRows.length + ewsRows.length;
                    
                    updateProgress(100, '處理完成！', '檔案已就緒，準備下載。');
                    setTimeout(() => {
                        showSuccess(totalProcessed);
                    }, 500);
                    
                } catch (err) {
                    showError(err.message || '處理 Excel 資料時發生錯誤。');
                }
            };
            
            reader.onerror = () => {
                showError('讀起檔案失敗，請檢查檔案是否損毀。');
            };
            
            reader.readAsArrayBuffer(selectedFile);
            
        } catch (err) {
            showError('初始化處理失敗。');
        }
    });

    // Update progress elements
    function updateProgress(percent, status, detail) {
        progressPercentage.textContent = percent + '%';
        progressBarFill.style.width = percent + '%';
        progressStatus.textContent = status;
        progressDetail.textContent = detail;
    }

    // Show Success State
    function showSuccess(totalCases) {
        progressContainer.classList.add('hidden');
        resultContainer.classList.remove('hidden');
        resultSummary.textContent = `成功篩選出 ${totalCases} 筆符合條件的案件，自動解析用戶問題與問題說明（已排除服務內容相關問題與個案外撥），已完成分頁劃分與樣式套用。`;
        
        // Setup download link
        const url = URL.createObjectURL(processedBlob);
        downloadLink.href = url;
        downloadLink.download = outputFilename;
        
        // Auto trigger download for convenience
        downloadLink.click();
    }

    // Show Error State
    function showError(msg) {
        progressContainer.classList.add('hidden');
        errorContainer.classList.remove('hidden');
        errorMessage.textContent = msg;
    }

    // Retry or Reset
    retryBtn.addEventListener('click', resetUploadState);
    resetBtn.addEventListener('click', resetUploadState);
});
