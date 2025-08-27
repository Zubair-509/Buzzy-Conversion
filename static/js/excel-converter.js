// PDF to Excel Converter JavaScript

let selectedFileExcel = null;
let isConvertingExcel = false;

// DOM elements for Excel converter
const uploadAreaExcel = document.getElementById('uploadAreaExcel');
const fileInputExcel = document.getElementById('fileInputExcel');
const fileInfoExcel = document.getElementById('fileInfoExcel');
const fileNameExcel = document.getElementById('fileNameExcel');
const fileSizeExcel = document.getElementById('fileSizeExcel');
const convertBtnExcel = document.getElementById('convertBtnExcel');
const progressContainerExcel = document.getElementById('progressContainerExcel');
const progressBarExcel = document.getElementById('progressBarExcel');
const progressTextExcel = document.getElementById('progressTextExcel');
const progressPercentExcel = document.getElementById('progressPercentExcel');
const successMessageExcel = document.getElementById('successMessageExcel');
const errorMessageExcel = document.getElementById('errorMessageExcel');
const errorTextExcel = document.getElementById('errorTextExcel');
const downloadBtnExcel = document.getElementById('downloadBtnExcel');

// File size formatter (reuse from main.js)
function formatFileSizeExcel(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

// Validate file for Excel conversion
function validateFileExcel(file) {
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

// Handle file selection for Excel conversion
function handleFileSelectExcel(file) {
    try {
        validateFileExcel(file);
        selectedFileExcel = file;
        
        // Update UI
        fileNameExcel.textContent = file.name;
        fileSizeExcel.textContent = formatFileSizeExcel(file.size);
        
        fileInfoExcel.style.display = 'block';
        convertBtnExcel.disabled = false;
        
        // Hide messages
        clearMessagesExcel();
        
        // Update upload area
        uploadAreaExcel.style.display = 'none';
        
    } catch (error) {
        showErrorExcel(error.message);
        clearFileExcel();
    }
}

// Clear selected file for Excel conversion
function clearFileExcel() {
    selectedFileExcel = null;
    fileInputExcel.value = '';
    fileInfoExcel.style.display = 'none';
    uploadAreaExcel.style.display = 'block';
    convertBtnExcel.disabled = true;
    clearMessagesExcel();
    hideProgressExcel();
}

// Clear messages for Excel conversion
function clearMessagesExcel() {
    successMessageExcel.style.display = 'none';
    errorMessageExcel.style.display = 'none';
}

// Show error message for Excel conversion
function showErrorExcel(message) {
    errorTextExcel.textContent = message;
    errorMessageExcel.style.display = 'block';
    successMessageExcel.style.display = 'none';
}

// Show success message for Excel conversion
function showSuccessExcel(message, downloadUrl, filename) {
    successMessageExcel.style.display = 'block';
    errorMessageExcel.style.display = 'none';
    downloadBtnExcel.href = downloadUrl;
    downloadBtnExcel.download = filename;
}

// Show progress for Excel conversion
function showProgressExcel(text = 'Processing...', percent = 0) {
    progressContainerExcel.style.display = 'block';
    progressTextExcel.textContent = text;
    progressPercentExcel.textContent = percent + '%';
    progressBarExcel.style.width = percent + '%';
    progressBarExcel.setAttribute('aria-valuenow', percent);
}

// Hide progress for Excel conversion
function hideProgressExcel() {
    progressContainerExcel.style.display = 'none';
}

// Update progress for Excel conversion
function updateProgressExcel(text, percent) {
    showProgressExcel(text, percent);
}

// Convert file to Excel
async function convertFileExcel() {
    if (!selectedFileExcel || isConvertingExcel) {
        return;
    }
    
    isConvertingExcel = true;
    convertBtnExcel.disabled = true;
    clearMessagesExcel();
    
    try {
        // Show initial progress
        showProgressExcel('Uploading file...', 10);
        document.body.classList.add('converting');
        
        // Create form data
        const formData = new FormData();
        formData.append('file', selectedFileExcel);
        
        // Upload and convert
        updateProgressExcel('Analyzing PDF structure...', 30);
        
        const response = await fetch('/upload-excel', {
            method: 'POST',
            body: formData
        });
        
        updateProgressExcel('Extracting tables...', 60);
        
        const result = await response.json();
        
        updateProgressExcel('Creating Excel file...', 85);
        
        if (result.success) {
            updateProgressExcel('Complete!', 100);
            setTimeout(() => {
                hideProgressExcel();
                showSuccessExcel(result.message, result.download_url, result.filename);
            }, 500);
        } else {
            throw new Error(result.error || 'Excel conversion failed');
        }
        
    } catch (error) {
        console.error('Excel conversion error:', error);
        hideProgressExcel();
        showErrorExcel(error.message || 'An error occurred during Excel conversion');
    } finally {
        isConvertingExcel = false;
        convertBtnExcel.disabled = false;
        document.body.classList.remove('converting');
    }
}

// Event Listeners for Excel Converter

// File input change for Excel
if (fileInputExcel) {
    fileInputExcel.addEventListener('change', function(e) {
        if (e.target.files.length > 0) {
            handleFileSelectExcel(e.target.files[0]);
        }
    });
}

// Drag and drop for Excel converter
if (uploadAreaExcel) {
    uploadAreaExcel.addEventListener('click', function() {
        if (!isConvertingExcel) {
            fileInputExcel.click();
        }
    });

    uploadAreaExcel.addEventListener('dragover', function(e) {
        e.preventDefault();
        uploadAreaExcel.classList.add('dragover');
    });

    uploadAreaExcel.addEventListener('dragleave', function(e) {
        e.preventDefault();
        uploadAreaExcel.classList.remove('dragover');
    });

    uploadAreaExcel.addEventListener('drop', function(e) {
        e.preventDefault();
        uploadAreaExcel.classList.remove('dragover');
        
        if (!isConvertingExcel && e.dataTransfer.files.length > 0) {
            handleFileSelectExcel(e.dataTransfer.files[0]);
        }
    });
}

// Convert button for Excel
if (convertBtnExcel) {
    convertBtnExcel.addEventListener('click', convertFileExcel);
}

// Download button click tracking for Excel
if (downloadBtnExcel) {
    downloadBtnExcel.addEventListener('click', function() {
        // Clear file after download starts
        setTimeout(() => {
            clearFileExcel();
        }, 1000);
    });
}

// Allow Enter key to trigger Excel conversion
document.addEventListener('keydown', function(e) {
    if (e.key === 'Enter' && selectedFileExcel && !isConvertingExcel && document.activeElement !== fileInputExcel) {
        convertFileExcel();
    }
});

console.log('PDF to Excel converter initialized successfully');