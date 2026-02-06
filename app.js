// Global variables
let workbookData = null;
let parsedData = null;
let previewContent = null; // Store preview content for editing
let docxLib = null; // Store docx library reference

// DOM elements
const fileInput = document.getElementById('fileInput');
const fileName = document.getElementById('fileName');
const downloadBtn = document.getElementById('downloadBtn');
const statusDiv = document.getElementById('status');
const previewSection = document.getElementById('previewSection');
const dataPreview = document.getElementById('dataPreview');
const wordPreviewSection = document.getElementById('wordPreviewSection');
const wordPreview = document.getElementById('wordPreview');
const refreshPreviewBtn = document.getElementById('refreshPreviewBtn');

// Wait for libraries to load
window.addEventListener('load', function() {
    // Check if docx library is loaded
    if (typeof docx !== 'undefined') {
        docxLib = docx;
        console.log('docx library loaded successfully');
    } else if (window.docx) {
        docxLib = window.docx;
        console.log('docx library loaded from window');
    } else {
        console.error('docx library not found');
    }
});

// Event listeners
fileInput.addEventListener('change', handleFileSelect);
downloadBtn.addEventListener('click', generateWordDocument);
refreshPreviewBtn.addEventListener('click', refreshWordPreview);

// Handle file selection
function handleFileSelect(event) {
    const file = event.target.files[0];
    
    if (!file) {
        return;
    }

    // Validate file type
    const validTypes = ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'application/vnd.ms-excel'];
    if (!validTypes.includes(file.type) && !file.name.match(/\.(xlsx|xls)$/i)) {
        showStatus('Please select a valid XLSX or XLS file.', 'error');
        return;
    }

    fileName.textContent = `Selected: ${file.name}`;
    showStatus('Reading file...', 'info');

    // Read the file
    const reader = new FileReader();
    
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            workbookData = XLSX.read(data, { type: 'array' });
            
            // Parse and display data
            parseWorkbook(workbookData);
            showStatus('File loaded successfully!', 'success');
            downloadBtn.disabled = false;
        } catch (error) {
            showStatus('Error reading file: ' + error.message, 'error');
            downloadBtn.disabled = true;
        }
    };

    reader.onerror = function() {
        showStatus('Error reading file.', 'error');
        downloadBtn.disabled = true;
    };

    reader.readAsArrayBuffer(file);
}

// Parse workbook and extract data
function parseWorkbook(workbook) {
    try {
        // Get the first sheet
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        // Convert to JSON
        parsedData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        
        // Display preview
        displayPreview(parsedData);
        
        // Generate and display Word preview
        generateWordPreview(parsedData);
    } catch (error) {
        showStatus('Error parsing workbook: ' + error.message, 'error');
    }
}

// Display data preview
function displayPreview(data) {
    if (!data || data.length === 0) {
        dataPreview.innerHTML = '<p>No data found in the spreadsheet.</p>';
        previewSection.style.display = 'block';
        return;
    }

    // Create a table for preview (show first 10 rows)
    let html = '<table><thead><tr>';
    
    // Add headers (first row)
    const headers = data[0] || [];
    headers.forEach(header => {
        html += `<th>${escapeHtml(String(header || ''))}</th>`;
    });
    html += '</tr></thead><tbody>';

    // Add data rows (up to 10 rows)
    const previewRows = data.slice(1, 11);
    previewRows.forEach(row => {
        html += '<tr>';
        headers.forEach((_, index) => {
            const cellValue = row[index] !== undefined ? row[index] : '';
            html += `<td>${escapeHtml(String(cellValue))}</td>`;
        });
        html += '</tr>';
    });

    html += '</tbody></table>';
    
    if (data.length > 11) {
        html += `<p style="margin-top: 10px; color: #718096;">Showing 10 of ${data.length - 1} rows</p>`;
    }

    dataPreview.innerHTML = html;
    previewSection.style.display = 'block';
}

// Generate Word document preview
function generateWordPreview(data) {
    if (!data || data.length === 0) {
        wordPreview.innerHTML = '<p class="text-muted">No data to preview.</p>';
        wordPreviewSection.style.display = 'block';
        return;
    }

    const headers = data[0] || [];
    const rows = data.slice(1);

    // Find column indices
    const ptNoIndex = headers.findIndex(h => String(h).toLowerCase().includes('pt no'));
    const patientNameIndex = headers.findIndex(h => String(h).toLowerCase().includes('patient name'));
    const visitDateIndex = headers.findIndex(h => String(h).toLowerCase().includes('visit date'));
    const doctorIndex = headers.findIndex(h => String(h).toLowerCase().includes('doctor'));
    const personalRemindersIndex = headers.findIndex(h => String(h).toLowerCase().includes('personal reminders'));

    // Build HTML preview
    let html = '<div class="document-preview">';
    html += '<h3 class="mb-3">Patient Reports</h3>';
    html += `<p class="text-muted mb-4">Generated on: ${new Date().toLocaleString()}</p>`;

    rows.forEach((row, index) => {
        if (index > 0) {
            html += '<hr class="my-4">';
        }

        const ptNo = row[ptNoIndex] !== undefined ? String(row[ptNoIndex]) : '';
        const patientName = row[patientNameIndex] !== undefined ? String(row[patientNameIndex]) : '';
        const visitDate = row[visitDateIndex] !== undefined ? formatDate(row[visitDateIndex]) : '';
        const doctor = row[doctorIndex] !== undefined ? String(row[doctorIndex]) : '';
        const personalReminders = row[personalRemindersIndex] !== undefined ? row[personalRemindersIndex] : '';
        const remarks = getRemarks(personalReminders);

        html += `<div class="patient-record mb-3">`;
        html += `<p class="mb-1"><strong>Date:</strong> ${escapeHtml(visitDate)}</p>`;
        html += `<p class="mb-1 ms-2"><strong>File Number:</strong> ${escapeHtml(ptNo)}</p>`;
        html += `<p class="mb-1 ms-2"><strong>Patient Name:</strong> ${escapeHtml(patientName)}</p>`;
        html += `<p class="mb-1 ms-2"><strong>Doctor Name:</strong> ${escapeHtml(doctor)}</p>`;
        html += `<p class="mb-1 ms-2"><strong>Remarks:</strong> ${escapeHtml(remarks)}</p>`;
        html += `</div>`;
    });

    html += '</div>';
    
    wordPreview.innerHTML = html;
    wordPreviewSection.style.display = 'block';
}

// Refresh Word preview from current data
function refreshWordPreview() {
    if (parsedData) {
        generateWordPreview(parsedData);
        showStatus('Preview refreshed!', 'success');
    }
}

// Generate Word document
async function generateWordDocument() {
    if (!parsedData || parsedData.length === 0) {
        showStatus('No data to export.', 'error');
        return;
    }

    try {
        showStatus('Generating Word document...', 'info');
        downloadBtn.disabled = true;

        // Check if docx library is available
        const lib = docxLib || window.docx || docx;
        if (!lib) {
            throw new Error('docx library is not loaded. Please refresh the page and try again.');
        }

        // Get edited content from preview or use original data
        const contentToUse = getEditedContent(lib);

        // Create a new document using docx
        const doc = new lib.Document({
            sections: [{
                properties: {},
                children: contentToUse
            }]
        });

        // Generate and download the document
        const blob = await lib.Packer.toBlob(doc);
        saveAs(blob, 'patient-report.docx');
        
        showStatus('Word document generated successfully!', 'success');
        downloadBtn.disabled = false;
    } catch (error) {
        showStatus('Error generating document: ' + error.message, 'error');
        console.error('Error details:', error);
        downloadBtn.disabled = false;
    }
}

// Get edited content from preview or generate from data
function getEditedContent(lib) {
    // Parse the edited HTML preview to extract text
    const previewDiv = wordPreview;
    const paragraphs = [];

    if (previewDiv.textContent.trim()) {
        // Extract text from the editable preview
        const lines = previewDiv.innerText.split('\n').filter(line => line.trim());
        
        lines.forEach((line, index) => {
            const trimmedLine = line.trim();
            if (!trimmedLine) return;

            // Check if it's a title
            if (trimmedLine === 'Patient Reports') {
                paragraphs.push(
                    new lib.Paragraph({
                        text: trimmedLine,
                        heading: lib.HeadingLevel.HEADING_1,
                        spacing: { after: 300 }
                    })
                );
            }
            // Check if it's a separator
            else if (trimmedLine.includes('___') || trimmedLine === '---') {
                paragraphs.push(
                    new lib.Paragraph({
                        text: '_____________________',
                        spacing: { before: 200, after: 200 }
                    })
                );
            }
            // Regular content
            else {
                paragraphs.push(
                    new lib.Paragraph({
                        text: trimmedLine,
                        spacing: { after: 100 }
                    })
                );
            }
        });
    } else {
        // Fallback to original data
        return createDocumentContent(parsedData, lib);
    }

    return paragraphs.length > 0 ? paragraphs : createDocumentContent(parsedData, lib);
}

// Helper function to generate remarks from Personal Reminders field
function getRemarks(personalReminders) {
    if (!personalReminders) {
        return '';
    }
    const remindersStr = String(personalReminders).toUpperCase();
    if (remindersStr.includes('OPG')) {
        return 'Patient with new OPG';
    }
    return '';
}

// Helper function to convert Excel serial date to readable format
function excelDateToJSDate(serial) {
    // Check if it's already a string date
    if (typeof serial === 'string' && isNaN(serial)) {
        return serial;
    }
    
    // Check if it's a number (Excel serial date)
    if (typeof serial === 'number' || !isNaN(serial)) {
        const utc_days = Math.floor(serial - 25569);
        const utc_value = utc_days * 86400;
        const date_info = new Date(utc_value * 1000);

        const fractional_day = serial - Math.floor(serial) + 0.0000001;

        let total_seconds = Math.floor(86400 * fractional_day);

        const seconds = total_seconds % 60;

        total_seconds -= seconds;

        const hours = Math.floor(total_seconds / (60 * 60));
        const minutes = Math.floor(total_seconds / 60) % 60;

        const date = new Date(date_info.getFullYear(), date_info.getMonth(), date_info.getDate(), hours, minutes, seconds);
        
        // Format as "DD Month YYYY"
        const day = date.getDate();
        const monthNames = ["January", "February", "March", "April", "May", "June",
                           "July", "August", "September", "October", "November", "December"];
        const month = monthNames[date.getMonth()];
        const year = date.getFullYear();
        
        return `${day} ${month} ${year}`;
    }
    
    return String(serial);
}

// Helper function to format any date value
function formatDate(dateValue) {
    if (!dateValue) return '';
    
    // If it's already formatted nicely, return it
    const str = String(dateValue);
    if (str.match(/\d{1,2}\s+\w+\s+\d{4}/)) {
        return str;
    }
    
    // Otherwise convert from Excel serial
    return excelDateToJSDate(dateValue);
}

// Create document content
function createDocumentContent(data, lib) {
    // Use the provided lib or try to get it from global scope
    const docxLib = lib || docxLib || window.docx || docx;
    
    const children = [];

    // Add title
    children.push(
        new docxLib.Paragraph({
            text: 'Patient Reports',
            heading: docxLib.HeadingLevel.HEADING_1,
            spacing: {
                after: 300
            }
        })
    );

    // Add generation date
    children.push(
        new docxLib.Paragraph({
            text: `Generated on: ${new Date().toLocaleString()}`,
            spacing: {
                after: 400
            }
        })
    );

    // Process data if exists
    if (data.length > 0) {
        const headers = data[0] || [];
        const rows = data.slice(1);

        // Find column indices
        const ptNoIndex = headers.findIndex(h => String(h).toLowerCase().includes('pt no'));
        const patientNameIndex = headers.findIndex(h => String(h).toLowerCase().includes('patient name'));
        const visitDateIndex = headers.findIndex(h => String(h).toLowerCase().includes('visit date'));
        const doctorIndex = headers.findIndex(h => String(h).toLowerCase().includes('doctor'));
        const personalRemindersIndex = headers.findIndex(h => String(h).toLowerCase().includes('personal reminders'));

        // Process each patient record
        rows.forEach((row, index) => {
            if (index > 0) {
                // Add separator line between records
                children.push(
                    new docxLib.Paragraph({
                        text: '_____________________',
                        spacing: {
                            before: 200,
                            after: 200
                        }
                    })
                );
            }

            // Extract data from row
            const ptNo = row[ptNoIndex] !== undefined ? String(row[ptNoIndex]) : '';
            const patientName = row[patientNameIndex] !== undefined ? String(row[patientNameIndex]) : '';
            const visitDate = row[visitDateIndex] !== undefined ? formatDate(row[visitDateIndex]) : '';
            const doctor = row[doctorIndex] !== undefined ? String(row[doctorIndex]) : '';
            const personalReminders = row[personalRemindersIndex] !== undefined ? row[personalRemindersIndex] : '';

            // Determine remarks
            const remarks = getRemarks(personalReminders);

            // Add formatted patient record
            children.push(
                new docxLib.Paragraph({
                    text: `Date: ${visitDate}`,
                    spacing: { after: 100 }
                }),
                new docxLib.Paragraph({
                    text: ` File Number: ${ptNo}`,
                    spacing: { after: 100 }
                }),
                new docxLib.Paragraph({
                    text: ` Patient Name: ${patientName}`,
                    spacing: { after: 100 }
                }),
                new docxLib.Paragraph({
                    text: ` Doctor Name: ${doctor}`,
                    spacing: { after: 100 }
                }),
                new docxLib.Paragraph({
                    text: ` Remarks: ${remarks}`,
                    spacing: { after: 100 }
                })
            );
        });
    } else {
        children.push(
            new docxLib.Paragraph({
                text: 'No data available.',
                spacing: {
                    before: 200
                }
            })
        );
    }

    return children;
}

// Show status message
function showStatus(message, type) {
    statusDiv.textContent = message;
    statusDiv.className = 'status-message ' + type;
}

// Escape HTML to prevent XSS
function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}
