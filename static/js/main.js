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
    let pollInterval = null;

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
        fileInput.value = '';
        dropzone.classList.remove('hidden');
        fileInfo.classList.add('hidden');
        progressContainer.classList.add('hidden');
        errorContainer.classList.add('hidden');
        resultContainer.classList.add('hidden');
        if (pollInterval) {
            clearInterval(pollInterval);
            pollInterval = null;
        }
    }

    // Process file
    processBtn.addEventListener('click', () => {
        if (!selectedFile) return;

        fileInfo.classList.add('hidden');
        progressContainer.classList.remove('hidden');
        updateProgress(0, '正在上傳報表檔案...', '準備開始處理...');

        const formData = new FormData();
        formData.append('file', selectedFile);

        fetch('/upload', {
            method: 'POST',
            body: formData
        })
        .then(response => response.json())
        .then(data => {
            if (data.error) {
                showError(data.error);
            } else {
                const taskId = data.task_id;
                startPolling(taskId);
            }
        })
        .catch(err => {
            showError('與伺服器連線失敗，請檢查網路。');
        });
    });

    // Update progress elements
    function updateProgress(percent, status, detail) {
        progressPercentage.textContent = percent + '%';
        progressBarFill.style.width = percent + '%';
        progressStatus.textContent = status;
        progressDetail.textContent = detail;
    }

    // Start Polling for Status
    function startPolling(taskId) {
        if (pollInterval) clearInterval(pollInterval);

        pollInterval = setInterval(() => {
            fetch(`/status/${taskId}`)
            .then(response => response.json())
            .then(data => {
                if (data.status === 'processing') {
                    updateProgress(data.progress, '正在整理案件資料...', data.detail);
                } else if (data.status === 'success') {
                    clearInterval(pollInterval);
                    showSuccess(data.result_file, data.total_cases);
                } else if (data.status === 'error') {
                    clearInterval(pollInterval);
                    showError(data.error_msg);
                }
            })
            .catch(err => {
                // If it fails occasionally, do not clear, let it retry in next poll
                console.error('Status poll failed:', err);
            });
        }, 1000);
    }

    // Show Success State
    function showSuccess(resultFile, totalCases) {
        progressContainer.classList.add('hidden');
        resultContainer.classList.remove('hidden');
        resultSummary.textContent = `成功篩選出 ${totalCases} 筆符合條件的案件，自動解析用戶問題、處理說明並於網路搜尋相關補充連結，已完成套用報表樣式。`;
        downloadLink.href = `/download/${resultFile}`;
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
