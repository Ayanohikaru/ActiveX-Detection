// NAB ActiveX Scanner Portal - Main Script

// Constants
const ACTIVEX_KEYWORDS = [
    "CreateObject(",
    "GetObject(",
    "MSComctlLib",
    "Forms.CommandButton",
    "ClassId={",
    "Object=",
    "VBComponent",
    "ActiveX"
];

const MAX_FILE_SIZE = 10 * 1024 * 1024; // 10MB limit

// DOM Elements
const dropZone = document.getElementById('dropZone');
const fileInput = document.getElementById('fileInput');
const scanButton = document.getElementById('scanButton');
const resultsDiv = document.getElementById('results');
const spinner = document.getElementById('spinner');

// Event Listeners
document.addEventListener('DOMContentLoaded', initializeApp);
dropZone.addEventListener('click', () => fileInput.click());
fileInput.addEventListener('change', handleFileSelection);
scanButton.addEventListener('click', handleScanButtonClick);

// Drag and Drop Events
['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
    dropZone.addEventListener(eventName, preventDefaults);
});

['dragenter', 'dragover'].forEach(eventName => {
    dropZone.addEventListener(eventName, highlight);
});

['dragleave', 'drop'].forEach(eventName => {
    dropZone.addEventListener(eventName, unhighlight);
});

dropZone.addEventListener('drop', handleDrop);

// Initialize Application
function initializeApp() {
    console.log('NAB ActiveX Scanner initialized');
    clearResults();
}

// Event Handlers
function preventDefaults(e) {
    e.preventDefault();
    e.stopPropagation();
}

function highlight() {
    dropZone.classList.add('drag-over');
}

function unhighlight() {
    dropZone.classList.remove('drag-over');
}

function handleDrop(e) {
    const dt = e.dataTransfer;
    const files = dt.files;
    handleFiles(files);
}

function handleFileSelection(e) {
    const files = e.target.files;
    handleFiles(files);
}

function handleFiles(files) {
    if (files.length > 0) {
        scanButton.disabled = false;
        const fileList = Array.from(files).map(f => f.name).join(', ');
        console.log(`Files selected: ${fileList}`);
    } else {
        scanButton.disabled = true;
    }
}

async function handleScanButtonClick() {
    const files = fileInput.files;
    if (!files.length) return;

    clearResults();
    showSpinner();
    
    try {
        const results = await scanFiles(files);
        displayResults(results);
    } catch (error) {
        console.error('Scan error:', error);
        displayError(error.message);
    } finally {
        hideSpinner();
    }
}

// Scanning Functions
async function scanFiles(files) {
    const results = [];
    
    for (const file of files) {
        if (file.size > MAX_FILE_SIZE) {
            console.warn(`File ${file.name} exceeds size limit`);
            results.push({
                fileName: file.name,
                error: 'File exceeds 10MB size limit'
            });
            continue;
        }

        try {
            const fileResult = await scanSingleFile(file);
            results.push({
                fileName: file.name,
                ...fileResult
            });
        } catch (error) {
            results.push({
                fileName: file.name,
                error: error.message
            });
        }
    }

    return results;
}

async function scanSingleFile(file) {
    return new Promise((resolve) => {
        const reader = new FileReader();
        
        reader.onload = (e) => {
            const content = e.target.result;
            const findings = scanFileContent(content);
            resolve({ findings });
        };

        reader.onerror = () => {
            resolve({
                error: 'Failed to read file content'
            });
        };

        reader.readAsText(file);
    });
}

function scanFileContent(text) {
    const findings = [];
    
    for (const keyword of ACTIVEX_KEYWORDS) {
        const regex = new RegExp(keyword.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'gi');
        let match;
        
        while ((match = regex.exec(text)) !== null) {
            const start = Math.max(0, match.index - 40);
            const end = Math.min(text.length, match.index + keyword.length + 40);
            const snippet = text.substring(start, end);
            
            findings.push({
                keyword,
                snippet: sanitizeSnippet(snippet),
                position: match.index
            });
        }
    }
    
    return findings;
}

// UI Functions
function displayResults(results) {
    clearResults();
    
    const totalFindings = results.reduce((sum, result) => 
        sum + (result.findings?.length || 0), 0);

    if (totalFindings === 0) {
        displaySafeResult(results);
    } else {
        displayWarningResult(results);
    }

    // Add clear button
    const clearButton = document.createElement('button');
    clearButton.className = 'primary-button';
    clearButton.textContent = 'Clear';
    clearButton.onclick = resetForm;
    resultsDiv.appendChild(clearButton);
}

function displaySafeResult(results) {
    const banner = document.createElement('div');
    banner.className = 'result-banner safe';
    banner.innerHTML = `
        <i class="fas fa-check-circle"></i>
        <h3>No ActiveX indicators found</h3>
        <p>Scanned ${results.length} file(s) successfully.</p>
    `;
    resultsDiv.appendChild(banner);
}

function displayWarningResult(results) {
    const banner = document.createElement('div');
    banner.className = 'result-banner warning';
    banner.innerHTML = `
        <i class="fas fa-exclamation-triangle"></i>
        <h3>Possible ActiveX indicators detected</h3>
    `;
    resultsDiv.appendChild(banner);

    results.forEach(result => {
        if (result.error) {
            displayFileError(result);
            return;
        }

        if (result.findings && result.findings.length > 0) {
            displayFileFindings(result);
        }
    });
}

function displayFileError(result) {
    const errorDiv = document.createElement('div');
    errorDiv.className = 'found-item';
    errorDiv.innerHTML = `
        <strong>${result.fileName}:</strong> 
        <span class="error">${result.error}</span>
    `;
    resultsDiv.appendChild(errorDiv);
}

function displayFileFindings(result) {
    const fileDiv = document.createElement('div');
    fileDiv.className = 'found-item';
    
    const findings = result.findings
        .map(f => `
            <div class="finding">
                <strong>Found "${f.keyword}"</strong>
                <pre>${f.snippet}</pre>
            </div>
        `)
        .join('');

    fileDiv.innerHTML = `
        <h4>${result.fileName}</h4>
        ${findings}
    `;
    
    resultsDiv.appendChild(fileDiv);
}

function displayError(message) {
    const errorDiv = document.createElement('div');
    errorDiv.className = 'result-banner warning';
    errorDiv.innerHTML = `
        <i class="fas fa-exclamation-circle"></i>
        <h3>Error</h3>
        <p>${message}</p>
    `;
    resultsDiv.appendChild(errorDiv);
}

// Utility Functions
function sanitizeSnippet(text) {
    return text
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .trim();
}

function showSpinner() {
    spinner.classList.remove('hidden');
}

function hideSpinner() {
    spinner.classList.add('hidden');
}

function clearResults() {
    resultsDiv.innerHTML = '';
}

function resetForm() {
    fileInput.value = '';
    clearResults();
    scanButton.disabled = true;
}

// Console Helper
function logScanSummary(results) {
    console.group('Scan Summary');
    results.forEach(result => {
        console.log(`File: ${result.fileName}`);
        if (result.error) {
            console.warn(`Error: ${result.error}`);
        } else {
            console.log(`Findings: ${result.findings.length}`);
        }
    });
    console.groupEnd();
}