// PDF to DOCX Converter JavaScript

let selectedFile = null;
let isConverting = false;

// DOM elements
const uploadArea = document.getElementById('uploadArea');
const fileInput = document.getElementById('fileInput');
const fileInfo = document.getElementById('fileInfo');
const fileName = document.getElementById('fileName');
const fileSize = document.getElementById('fileSize');
const convertBtn = document.getElementById('convertBtn');
const progressContainer = document.getElementById('progressContainer');
const progressBar = document.getElementById('progressBar');
const progressText = document.getElementById('progressText');
const progressPercent = document.getElementById('progressPercent');
const successMessage = document.getElementById('successMessage');
const errorMessage = document.getElementById('errorMessage');
const errorText = document.getElementById('errorText');
const downloadBtn = document.getElementById('downloadBtn');

// File size formatter
function formatFileSize(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

// Validate file
function validateFile(file) {
    // Check file type
    if (!file.type.includes('pdf') && !file.name.toLowerCase().endsWith('.pdf')) {
        throw new Error('Please select a PDF file');
    }
    
    // Check file size (50MB limit)
    const maxSize = 50 * 1024 * 1024;
    if (file.size > maxSize) {
        throw new Error('File size must be less than 50MB');
    }
    
    // Check if file has content
    if (file.size === 0) {
        throw new Error('File appears to be empty');
    }
    
    return true;
}

// Handle file selection
function handleFileSelect(file) {
    try {
        validateFile(file);
        selectedFile = file;
        
        // Update UI
        fileName.textContent = file.name;
        fileSize.textContent = formatFileSize(file.size);
        
        fileInfo.style.display = 'block';
        convertBtn.disabled = false;
        
        // Hide messages
        clearMessages();
        
        // Update upload area
        uploadArea.style.display = 'none';
        
    } catch (error) {
        showError(error.message);
        clearFile();
    }
}

// Clear selected file
function clearFile() {
    selectedFile = null;
    fileInput.value = '';
    fileInfo.style.display = 'none';
    uploadArea.style.display = 'block';
    convertBtn.disabled = true;
    clearMessages();
    hideProgress();
}

// Clear messages
function clearMessages() {
    successMessage.style.display = 'none';
    errorMessage.style.display = 'none';
}

// Show error message
function showError(message) {
    errorText.textContent = message;
    errorMessage.style.display = 'block';
    successMessage.style.display = 'none';
}

// Show success message
function showSuccess(message, downloadUrl, filename) {
    successMessage.style.display = 'block';
    errorMessage.style.display = 'none';
    downloadBtn.href = downloadUrl;
    downloadBtn.download = filename;
}

// Show progress
function showProgress(text = 'Processing...', percent = 0) {
    progressContainer.style.display = 'block';
    progressText.textContent = text;
    progressPercent.textContent = percent + '%';
    progressBar.style.width = percent + '%';
    progressBar.setAttribute('aria-valuenow', percent);
}

// Hide progress
function hideProgress() {
    progressContainer.style.display = 'none';
}

// Update progress
function updateProgress(text, percent) {
    showProgress(text, percent);
}

// Convert file
async function convertFile() {
    if (!selectedFile || isConverting) {
        return;
    }
    
    isConverting = true;
    convertBtn.disabled = true;
    clearMessages();
    
    try {
        // Show initial progress
        showProgress('Uploading file...', 10);
        document.body.classList.add('converting');
        
        // Create form data
        const formData = new FormData();
        formData.append('file', selectedFile);
        
        // Upload and convert
        updateProgress('Converting PDF to DOCX...', 50);
        
        const response = await fetch('/upload', {
            method: 'POST',
            body: formData
        });
        
        updateProgress('Processing conversion...', 80);
        
        const result = await response.json();
        
        updateProgress('Finalizing...', 95);
        
        if (result.success) {
            updateProgress('Complete!', 100);
            setTimeout(() => {
                hideProgress();
                showSuccess(result.message, result.download_url, result.filename);
            }, 500);
        } else {
            throw new Error(result.error || 'Conversion failed');
        }
        
    } catch (error) {
        console.error('Conversion error:', error);
        hideProgress();
        showError(error.message || 'An error occurred during conversion');
    } finally {
        isConverting = false;
        convertBtn.disabled = false;
        document.body.classList.remove('converting');
    }
}

// Event Listeners

// File input change
fileInput.addEventListener('change', function(e) {
    if (e.target.files.length > 0) {
        handleFileSelect(e.target.files[0]);
    }
});

// Drag and drop
uploadArea.addEventListener('click', function() {
    if (!isConverting) {
        fileInput.click();
    }
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
    
    if (!isConverting && e.dataTransfer.files.length > 0) {
        handleFileSelect(e.dataTransfer.files[0]);
    }
});

// Prevent default drag behaviors
['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
    document.addEventListener(eventName, function(e) {
        e.preventDefault();
        e.stopPropagation();
    });
});

// Convert button
convertBtn.addEventListener('click', convertFile);

// Allow Enter key to trigger conversion
document.addEventListener('keydown', function(e) {
    if (e.key === 'Enter' && selectedFile && !isConverting) {
        convertFile();
    }
});

// Download button click tracking
downloadBtn.addEventListener('click', function() {
    // Clear file after download starts
    setTimeout(() => {
        clearFile();
    }, 1000);
});
