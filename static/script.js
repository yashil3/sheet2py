// BAH RuleBuilder Demo - Main JavaScript
document.addEventListener('DOMContentLoaded', function() {
    // Initialize the demo
    initializeDemo();
});

function initializeDemo() {
    // Set up event listeners
    setupFileUpload();
    setupTabs();
    setupDragAndDrop();
    setupJsonInput();
}

// File Upload Handling
function setupFileUpload() {
    const uploadForm = document.getElementById('uploadForm');
    const fileInput = document.getElementById('excelFile');
    const uploadBtn = document.getElementById('uploadBtn');

    uploadForm.addEventListener('submit', handleFileUpload);
    
    // File input change handler
    fileInput.addEventListener('change', function(e) {
        const file = e.target.files[0];
        if (file) {
            updateFileLabel(file.name);
        }
    });
}

function setupDragAndDrop() {
    const fileLabel = document.querySelector('.file-label');
    const fileInput = document.getElementById('excelFile');

    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
        fileLabel.addEventListener(eventName, preventDefaults, false);
    });

    function preventDefaults(e) {
        e.preventDefault();
        e.stopPropagation();
    }

    ['dragenter', 'dragover'].forEach(eventName => {
        fileLabel.addEventListener(eventName, highlight, false);
    });

    ['dragleave', 'drop'].forEach(eventName => {
        fileLabel.addEventListener(eventName, unhighlight, false);
    });

    function highlight(e) {
        fileLabel.classList.add('highlight');
    }

    function unhighlight(e) {
        fileLabel.classList.remove('highlight');
    }

    fileLabel.addEventListener('drop', handleDrop, false);

    function handleDrop(e) {
        const dt = e.dataTransfer;
        const files = dt.files;
        
        if (files.length > 0) {
            fileInput.files = files;
            updateFileLabel(files[0].name);
        }
    }
}

function updateFileLabel(fileName) {
    const fileText = document.querySelector('.file-text');
    const fileIcon = document.querySelector('.file-icon');
    
    if (fileName) {
        fileText.textContent = `Selected: ${fileName}`;
        fileIcon.textContent = '📁';
        fileText.style.color = '#667eea';
    } else {
        fileText.textContent = 'Choose Excel file or drag & drop';
        fileIcon.textContent = '📊';
        fileText.style.color = '#666';
    }
}

async function handleFileUpload(e) {
    e.preventDefault();
    
    const formData = new FormData();
    const fileInput = document.getElementById('excelFile');
    const includeCode = document.getElementById('includeCode').checked;
    const strictMode = document.getElementById('strictMode').checked;
    
    if (!fileInput.files[0]) {
        showNotification('Please select a file to upload', 'error');
        return;
    }
    
    formData.append('file', fileInput.files[0]);
    
    // Show loading state
    setLoadingState(true);
    
    try {
        // Build query parameters
        const params = new URLSearchParams();
        if (includeCode) params.append('include_code', '1');
        if (strictMode) params.append('strict', '1');
        
        const response = await fetch(`/api/convert?${params.toString()}`, {
            method: 'POST',
            body: formData
        });
        
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }
        
        const result = await response.json();
        displayResults(result);
        showNotification('File converted successfully!', 'success');
        
    } catch (error) {
        console.error('Upload error:', error);
        showNotification('Error converting file: ' + error.message, 'error');
    } finally {
        setLoadingState(false);
    }
}

function setLoadingState(loading) {
    const uploadBtn = document.getElementById('uploadBtn');
    const btnText = uploadBtn.querySelector('.btn-text');
    const btnLoading = uploadBtn.querySelector('.btn-loading');
    
    if (loading) {
        uploadBtn.disabled = true;
        btnText.style.display = 'none';
        btnLoading.style.display = 'inline';
        uploadBtn.classList.add('loading');
    } else {
        uploadBtn.disabled = false;
        btnText.style.display = 'inline';
        btnLoading.style.display = 'none';
        uploadBtn.classList.remove('loading');
    }
}

// Results Display
function displayResults(result) {
    // Store the result for JSON evaluation
    window.lastConversionResult = result;
    
    const resultsSection = document.getElementById('resultsSection');
    resultsSection.style.display = 'block';
    
    // Update summary cards
    updateSummaryCards(result);
    
    // Update tab content
    updateRulesTab(result);
    updateDependenciesTab(result);
    updateEvaluationTab(result);
    updateCodeTab(result);
    
    // Scroll to results
    resultsSection.scrollIntoView({ behavior: 'smooth' });
}

function updateSummaryCards(result) {
    document.getElementById('fileName').textContent = result.file || 'Unknown';
    document.getElementById('convertedCount').textContent = result.converted_count || 0;
    
    const totalFormulas = (result.converted_count || 0) + (result.errors?.length || 0);
    const successRate = totalFormulas > 0 ? 
        Math.round(((result.converted_count || 0) / totalFormulas) * 100) : 0;
    document.getElementById('successRate').textContent = `${successRate}%`;
}

function updateRulesTab(result) {
    const rulesList = document.getElementById('rulesList');
    
    if (!result.summary || result.summary.length === 0) {
        rulesList.innerHTML = '<p class="no-data">No rules converted. Check the file format and try again.</p>';
        return;
    }
    
    rulesList.innerHTML = result.summary.map(rule => `
        <div class="rule-item">
            <div class="rule-header">
                <span class="rule-cell">${rule.sheet}!${rule.cell}</span>
                <span class="rule-type">${rule.rule_type || 'formula'}</span>
            </div>
            <div class="rule-description">${rule.description || 'No description available'}</div>
            <div class="rule-formula">Original: ${rule.original_formula}</div>
            <div class="rule-python">Python: ${rule.python_expression}</div>
            ${rule.inputs && rule.inputs.length > 0 ? 
                `<div class="rule-inputs"><strong>Inputs:</strong> ${rule.inputs.join(', ')}</div>` : ''}
        </div>
    `).join('');
}

function updateDependenciesTab(result) {
    const dependencyList = document.getElementById('dependencyList');
    
    if (!result.sorted_cells || result.sorted_cells.length === 0) {
        dependencyList.innerHTML = '<p class="no-data">No dependency information available.</p>';
        return;
    }
    
    dependencyList.innerHTML = `
        <h4>Execution Order (Topological Sort)</h4>
        <div class="dependency-order">
            ${result.sorted_cells.map((cell, index) => `
                <div class="dependency-item">
                    <span class="step-number">${index + 1}</span>
                    <span class="cell-reference">${cell}</span>
                </div>
            `).join('')}
        </div>
    `;
}

function updateEvaluationTab(result) {
    const evaluationList = document.getElementById('evaluationList');
    
    if (!result.summary || result.summary.length === 0) {
        evaluationList.innerHTML = '<p class="no-data">No evaluation results available.</p>';
        return;
    }
    
    evaluationList.innerHTML = result.summary.map(rule => `
        <div class="evaluation-item">
            <div class="evaluation-header">
                <strong>${rule.sheet}!${rule.cell}</strong>
            </div>
            <div class="evaluation-result">
                Result: ${formatEvaluationResult(rule.evaluation_result)}
            </div>
            ${rule.dependencies && rule.dependencies.length > 0 ? 
                `<div class="evaluation-deps"><strong>Dependencies:</strong> ${rule.dependencies.join(', ')}</div>` : ''}
        </div>
    `).join('');
}

function updateCodeTab(result) {
    const codeOutput = document.getElementById('codeOutput');
    
    if (result.generated_code) {
        codeOutput.innerHTML = `<code>${escapeHtml(result.generated_code)}</code>`;
    } else {
        codeOutput.innerHTML = '<code>Upload a file and check "Include Generated Python Code" to see the generated rules.</code>';
    }
}

function formatEvaluationResult(result) {
    if (result === null || result === undefined) {
        return 'Not evaluated';
    }
    if (typeof result === 'number') {
        return result.toFixed(4);
    }
    if (typeof result === 'string') {
        return result;
    }
    return JSON.stringify(result);
}

function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

// Tab Functionality
function setupTabs() {
    const tabBtns = document.querySelectorAll('.tab-btn');
    const tabPanes = document.querySelectorAll('.tab-pane');
    
    tabBtns.forEach(btn => {
        btn.addEventListener('click', () => {
            const targetTab = btn.getAttribute('data-tab');
            
            // Update active tab button
            tabBtns.forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            
            // Update active tab pane
            tabPanes.forEach(pane => pane.classList.remove('active'));
            document.getElementById(targetTab).classList.add('active');
        });
    });
}

// Notifications
function showNotification(message, type = 'info') {
    // Create notification element
    const notification = document.createElement('div');
    notification.className = `notification notification-${type}`;
    notification.innerHTML = `
        <span class="notification-message">${message}</span>
        <button class="notification-close">&times;</button>
    `;
    
    // Add styles
    notification.style.cssText = `
        position: fixed;
        top: 20px;
        right: 20px;
        background: ${type === 'error' ? '#dc3545' : type === 'success' ? '#28a745' : '#17a2b8'};
        color: white;
        padding: 15px 20px;
        border-radius: 8px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.15);
        z-index: 1000;
        display: flex;
        align-items: center;
        gap: 15px;
        max-width: 400px;
        animation: slideIn 0.3s ease;
    `;
    
    // Add close button functionality
    const closeBtn = notification.querySelector('.notification-close');
    closeBtn.style.cssText = `
        background: none;
        border: none;
        color: white;
        font-size: 20px;
        cursor: pointer;
        padding: 0;
        line-height: 1;
    `;
    
    closeBtn.addEventListener('click', () => {
        notification.remove();
    });
    
    // Add to page
    document.body.appendChild(notification);
    
    // Auto-remove after 5 seconds
    setTimeout(() => {
        if (notification.parentNode) {
            notification.remove();
        }
    }, 5000);
}

// JSON Input Functionality
function setupJsonInput() {
    const loadSampleBtn = document.getElementById('loadSampleData');
    const validateBtn = document.getElementById('validateJson');
    const evaluateBtn = document.getElementById('evaluateRules');
    const jsonInput = document.getElementById('jsonInput');

    if (loadSampleBtn) loadSampleBtn.addEventListener('click', loadSampleData);
    if (validateBtn) validateBtn.addEventListener('click', validateJsonInput);
    if (evaluateBtn) evaluateBtn.addEventListener('click', evaluateRulesWithCustomData);
    if (jsonInput) jsonInput.addEventListener('input', clearJsonValidation);
}

function loadSampleData() {
    const sampleData = {
        "Formulas": {
            "Country of origin": 15,
            "Color": 50,
            "Mileage": 100,
            "Year": 75,
            "Options": 50,
            "Engine size": 30,
            "Transmission": 80,
            "Features": 50
        },
        "Source": {
            "Make": "Mercedes-Benz",
            "Model": "E-350",
            "Mileage": 57351,
            "Year": 2016,
            "Trim": "4Matic",
            "Option1": "Air suspension",
            "Option2": "Panoramic Roof",
            "Color": "Orange",
            "Make:GER": "GER",
            "Made in": "Germany"
        }
    };

    const jsonInput = document.getElementById('jsonInput');
    jsonInput.value = JSON.stringify(sampleData, null, 2);
    jsonInput.classList.remove('error', 'valid');
    
    showJsonValidation('Sample data loaded successfully! This matches your Excel spreadsheet data.', 'success');
}

function validateJsonInput() {
    const jsonInput = document.getElementById('jsonInput');
    const jsonText = jsonInput.value.trim();
    
    if (!jsonText) {
        showJsonValidation('Please enter some JSON data to validate.', 'info');
        return;
    }
    
    try {
        const parsed = JSON.parse(jsonText);
        jsonInput.classList.remove('error');
        jsonInput.classList.add('valid');
        showJsonValidation('JSON is valid! Ready for rule evaluation.', 'success');
        
        // Store the parsed data for later use
        window.currentJsonData = parsed;
        
    } catch (error) {
        jsonInput.classList.remove('valid');
        jsonInput.classList.add('error');
        showJsonValidation(`Invalid JSON: ${error.message}`, 'error');
        window.currentJsonData = null;
    }
}

function clearJsonValidation() {
    const jsonInput = document.getElementById('jsonInput');
    jsonInput.classList.remove('error', 'valid');
    
    const validation = document.getElementById('jsonValidation');
    if (validation) {
        validation.innerHTML = '';
        validation.className = 'json-validation';
    }
}

function showJsonValidation(message, type) {
    const validation = document.getElementById('jsonValidation');
    validation.innerHTML = message;
    validation.className = `json-validation ${type}`;
}

async function evaluateRulesWithCustomData() {
    if (!window.currentJsonData) {
        showNotification('Please validate your JSON data first.', 'error');
        return;
    }
    
    if (!window.lastConversionResult) {
        showNotification('Please convert an Excel file first to get rules for evaluation.', 'error');
        return;
    }
    
    try {
        const response = await fetch('/api/evaluate', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({
                rules: window.lastConversionResult.summary,
                data: window.currentJsonData
            })
        });
        
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }
        
        const result = await response.json();
        displayCustomEvaluationResults(result);
        showNotification('Rules evaluated successfully with custom data!', 'success');
        
    } catch (error) {
        console.error('Evaluation error:', error);
        showNotification('Error evaluating rules: ' + error.message, 'error');
    }
}

function displayCustomEvaluationResults(result) {
    const resultsDiv = document.getElementById('customEvaluationResults');
    const resultsList = document.getElementById('customResultsList');
    
    if (!result.results || Object.keys(result.results).length === 0) {
        resultsList.innerHTML = '<p class="no-data">No evaluation results available.</p>';
        resultsDiv.style.display = 'block';
        return;
    }
    
    resultsList.innerHTML = Object.entries(result.results).map(([cell, value]) => `
        <div class="custom-result-item">
            <div class="custom-result-header">
                <span class="custom-result-cell">${cell}</span>
                <span class="custom-result-value">${formatEvaluationResult(value)}</span>
            </div>
        </div>
    `).join('');
    
    resultsDiv.style.display = 'block';
}

// Add CSS animation
const style = document.createElement('style');
style.textContent = `
    @keyframes slideIn {
        from {
            transform: translateX(100%);
            opacity: 0;
        }
        to {
            transform: translateX(0);
            opacity: 1;
        }
    }
`;
document.head.appendChild(style);