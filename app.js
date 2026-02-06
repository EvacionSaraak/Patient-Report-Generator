// Global variables
let workbookData = null;
let parsedData = null;

// DOM elements
const fileInput = document.getElementById('fileInput');
const fileName = document.getElementById('fileName');
const downloadBtn = document.getElementById('downloadBtn');
const statusDiv = document.getElementById('status');
const previewSection = document.getElementById('previewSection');
const dataPreview = document.getElementById('dataPreview');

// Event listeners
fileInput.addEventListener('change', handleFileSelect);
downloadBtn.addEventListener('click', generateWordDocument);

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

// Generate Word document
async function generateWordDocument() {
    if (!parsedData || parsedData.length === 0) {
        showStatus('No data to export.', 'error');
        return;
    }

    try {
        showStatus('Generating Word document...', 'info');
        downloadBtn.disabled = true;

        // Create a new document using docx
        const doc = new docx.Document({
            sections: [{
                properties: {},
                children: createDocumentContent(parsedData)
            }]
        });

        // Generate and download the document
        const blob = await docx.Packer.toBlob(doc);
        saveAs(blob, 'patient-report.docx');
        
        showStatus('Word document generated successfully!', 'success');
        downloadBtn.disabled = false;
    } catch (error) {
        showStatus('Error generating document: ' + error.message, 'error');
        downloadBtn.disabled = false;
    }
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

// Create document content
function createDocumentContent(data) {
    const children = [];

    // Add title
    children.push(
        new docx.Paragraph({
            text: 'Patient Reports',
            heading: docx.HeadingLevel.HEADING_1,
            spacing: {
                after: 300
            }
        })
    );

    // Add generation date
    children.push(
        new docx.Paragraph({
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
                    new docx.Paragraph({
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
            const visitDate = row[visitDateIndex] !== undefined ? String(row[visitDateIndex]) : '';
            const doctor = row[doctorIndex] !== undefined ? String(row[doctorIndex]) : '';
            const personalReminders = row[personalRemindersIndex] !== undefined ? row[personalRemindersIndex] : '';

            // Determine remarks
            const remarks = getRemarks(personalReminders);

            // Add formatted patient record
            children.push(
                new docx.Paragraph({
                    text: `Date: ${visitDate}`,
                    spacing: { after: 100 }
                }),
                new docx.Paragraph({
                    text: ` File Number: ${ptNo}`,
                    spacing: { after: 100 }
                }),
                new docx.Paragraph({
                    text: ` Patient Name: ${patientName}`,
                    spacing: { after: 100 }
                }),
                new docx.Paragraph({
                    text: ` Doctor Name: ${doctor}`,
                    spacing: { after: 100 }
                }),
                new docx.Paragraph({
                    text: ` Remarks: ${remarks}`,
                    spacing: { after: 100 }
                })
            );
        });
    } else {
        children.push(
            new docx.Paragraph({
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
