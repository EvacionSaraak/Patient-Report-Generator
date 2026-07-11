// Global variables
let workbookData = null;
let parsedData = null;
let previewContent = null; // Store preview content for editing
let docxLib = null; // Store docx library reference
let selectedDownloadFormat = 'docx';

// DOM elements
const fileInput = document.getElementById('fileInput');
const fileName = document.getElementById('fileName');
const downloadBtn = document.getElementById('downloadBtn');
const downloadBtnText = document.getElementById('downloadBtnText');
const statusDiv = document.getElementById('status');
const previewSection = document.getElementById('previewSection');
const dataPreview = document.getElementById('dataPreview');
const reportPreviewSection = document.getElementById('reportPreviewSection');
const wordPreview = document.getElementById('wordPreview');
const textPreview = document.getElementById('textPreview');
const wordPreviewPanel = document.getElementById('wordPreviewPanel');
const textPreviewPanel = document.getElementById('textPreviewPanel');
const wordTabBtn = document.getElementById('wordTabBtn');
const textTabBtn = document.getElementById('textTabBtn');
const refreshPreviewBtn = document.getElementById('refreshPreviewBtn');

// Wait for libraries to load
window.addEventListener('load', function() {
    // Wait a bit for external scripts to fully initialize
    setTimeout(function() {
        // Check if docx library is loaded
        if (typeof docx !== 'undefined') {
            docxLib = docx;
            console.log('docx library loaded successfully');
        } else if (window.docx) {
            docxLib = window.docx;
            console.log('docx library loaded from window');
        } else {
            console.error('docx library not found - check CDN connection');
            showStatus('Warning: Document library not loaded. Refresh the page if download fails.', 'error');
        }
    }, 100);
});

// Event listeners
fileInput.addEventListener('change', handleFileSelect);
downloadBtn.addEventListener('click', generateReport);
refreshPreviewBtn.addEventListener('click', refreshWordPreview);
wordTabBtn.addEventListener('click', () => setDownloadFormat('docx'));
textTabBtn.addEventListener('click', () => setDownloadFormat('txt'));
setDownloadFormat('docx');

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
        
        // Generate and display report previews
        generateWordPreview(parsedData);
        generateTextPreview(parsedData);
        reportPreviewSection.style.display = 'block';
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

// Generate Word document preview – mirrors the 2-column table format used in the
// downloaded .docx (matching the Resources example report).
function generateWordPreview(data) {
    if (!data || data.length === 0) {
        wordPreview.innerHTML = '<p class="text-muted">No data to preview.</p>';
        return;
    }

    const headers = data[0] || [];
    const rows = filterEmptyRows(data.slice(1));

    const ptNoIndex = headers.findIndex(h => String(h).toLowerCase().includes('pt no'));
    const patientNameIndex = headers.findIndex(h => String(h).toLowerCase().includes('patient name'));
    const visitDateIndex = headers.findIndex(h => String(h).toLowerCase().includes('visit date'));
    const doctorIndex = headers.findIndex(h => String(h).toLowerCase().includes('doctor'));
    const personalRemindersIndex = headers.findIndex(h => String(h).toLowerCase().includes('personal reminders'));

    let html = '<div class="document-preview">';

    rows.forEach((row, index) => {
        if (index > 0) {
            html += '<div style="margin: 8px 0;"></div>';
        }

        const ptNo = row[ptNoIndex] !== undefined ? String(row[ptNoIndex]) : '';
        const patientName = row[patientNameIndex] !== undefined ? String(row[patientNameIndex]) : '';
        const visitDate = row[visitDateIndex] !== undefined ? formatDate(row[visitDateIndex]) : '';
        const doctor = row[doctorIndex] !== undefined ? String(row[doctorIndex]).trim() : '';
        const personalReminders = row[personalRemindersIndex] !== undefined ? String(row[personalRemindersIndex]).trim() : '';

        // Highlight colour: yellow for "LAST VISIT …", green for "NEW PATIENT …"
        let reminderStyle = '';
        const remUpper = personalReminders.toUpperCase();
        if (remUpper.startsWith('NEW PATIENT')) {
            reminderStyle = 'background-color: #90EE90;';
        } else if (remUpper.startsWith('LAST VISIT')) {
            reminderStyle = 'background-color: yellow;';
        }

        html += '<table style="border-collapse:collapse;width:100%;font-family:Arial,sans-serif;font-size:12pt;font-weight:bold;">';

        // Row 0 – Personal Reminders
        html += '<tr>';
        html += '<td style="border:1px solid #000;padding:2px 6px;width:30%;"></td>';
        if (personalReminders) {
            html += `<td style="border:1px solid #000;padding:2px 6px;${reminderStyle}">${escapeHtml(personalReminders)}</td>`;
        } else {
            html += '<td style="border:1px solid #000;padding:2px 6px;"></td>';
        }
        html += '</tr>';

        // Row 1 – Date
        html += `<tr><td style="border:1px solid #000;padding:2px 6px;">Date:</td><td style="border:1px solid #000;padding:2px 6px;">${escapeHtml(visitDate)}</td></tr>`;

        // Row 2 – File Number
        html += `<tr><td style="border:1px solid #000;padding:2px 6px;">File Number:</td><td style="border:1px solid #000;padding:2px 6px;">${escapeHtml(ptNo)}</td></tr>`;

        // Row 3 – Patient name
        html += `<tr><td style="border:1px solid #000;padding:2px 6px;">Patient name:</td><td style="border:1px solid #000;padding:2px 6px;">${escapeHtml(patientName)}</td></tr>`;

        // Row 4 – Doctor Name (omit when empty)
        if (doctor) {
            html += `<tr><td style="border:1px solid #000;padding:2px 6px;">Doctor Name:</td><td style="border:1px solid #000;padding:2px 6px;">${escapeHtml(doctor)}</td></tr>`;
        }

        html += '</table>';
    });

    html += '</div>';

    wordPreview.innerHTML = html;
}

// Generate text report preview
function generateTextPreview(data) {
    if (!data || data.length === 0) {
        textPreview.textContent = 'No data to preview.';
        return;
    }

    const headers = data[0] || [];
    const rows = filterEmptyRows(data.slice(1));

    const ptNoIndex = headers.findIndex(h => String(h).toLowerCase().includes('pt no'));
    const patientNameIndex = headers.findIndex(h => String(h).toLowerCase().includes('patient name'));
    const visitDateIndex = headers.findIndex(h => String(h).toLowerCase().includes('visit date'));
    const doctorIndex = headers.findIndex(h => String(h).toLowerCase().includes('doctor'));
    const personalRemindersIndex = headers.findIndex(h => String(h).toLowerCase().includes('personal reminders'));

    const dateRange = getDateRange(data);
    const headerText = dateRange.min && dateRange.max
        ? `PATIENT REPORT | ${dateRange.min} - ${dateRange.max}`
        : 'PATIENT REPORT';

    const lines = [headerText, ''];

    rows.forEach((row, index) => {
        const ptNo = row[ptNoIndex] !== undefined ? String(row[ptNoIndex]) : '';
        const patientName = row[patientNameIndex] !== undefined ? String(row[patientNameIndex]) : '';
        const visitDate = row[visitDateIndex] !== undefined ? formatDate(row[visitDateIndex]) : '';
        const doctor = row[doctorIndex] !== undefined ? String(row[doctorIndex]) : '';
        const personalReminders = row[personalRemindersIndex] !== undefined ? row[personalRemindersIndex] : '';
        const remarks = getRemarks(personalReminders);

        lines.push(`Date: ${visitDate}`);
        lines.push(` File Number: ${ptNo}`);
        lines.push(` Patient Name: ${patientName}`);
        lines.push(` Doctor Name: ${doctor}`);
        if (remarks && remarks.trim()) {
            lines.push(` Remarks: ${remarks}`);
        }

        if (index < rows.length - 1) {
            lines.push('', '----------------------------------------', '');
        }
    });

    textPreview.textContent = lines.join('\n');
}

// Refresh Word preview from current data
function refreshWordPreview() {
    if (parsedData) {
        generateWordPreview(parsedData);
        generateTextPreview(parsedData);
        showStatus('Preview refreshed!', 'success');
    }
}

// Set active download format
function setDownloadFormat(format) {
    selectedDownloadFormat = format;
    const isWordFormat = format === 'docx';

    wordTabBtn.classList.toggle('active', isWordFormat);
    textTabBtn.classList.toggle('active', !isWordFormat);
    wordPreviewPanel.style.display = isWordFormat ? 'block' : 'none';
    textPreviewPanel.style.display = isWordFormat ? 'none' : 'block';
    downloadBtnText.textContent = isWordFormat ? 'Download Word Report' : 'Download Text Report';
}

async function generateReport() {
    if (selectedDownloadFormat === 'txt') {
        generateTextDocument();
        return;
    }

    await generateWordDocument();
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

        // Check if docx library is available with multiple fallback attempts
        let lib = docxLib || window.docx;
        
        // Try accessing it directly as a last resort
        if (!lib && typeof docx !== 'undefined') {
            lib = docx;
            docxLib = docx; // Cache it for future use
        }
        
        if (!lib) {
            throw new Error('docx library is not loaded. Please refresh the page and try again.');
        }

        // Always use createDocumentContent to preserve formatting
        // Don't use getEditedContent which strips formatting
        const contentToUse = createDocumentContent(parsedData, lib);

        // Create a new document using docx
        const doc = new lib.Document({
            sections: [{
                properties: {},
                children: contentToUse
            }]
        });

        // Generate dynamic filename based on date range
        const dateRange = getDateRange(parsedData);
        let filename = 'PATIENT REPORT _ DATED ';
        if (dateRange.min && dateRange.max) {
            filename += `${dateRange.min} - ${dateRange.max}.docx`;
        } else {
            filename += 'Unknown.docx';
        }

        // Generate and download the document
        const blob = await lib.Packer.toBlob(doc);
        saveAs(blob, filename);
        
        showStatus('Word document generated successfully!', 'success');
        downloadBtn.disabled = false;
    } catch (error) {
        showStatus('Error generating document: ' + error.message, 'error');
        console.error('Error details:', error);
        downloadBtn.disabled = false;
    }
}

// Generate text document
function generateTextDocument() {
    if (!parsedData || parsedData.length === 0) {
        showStatus('No data to export.', 'error');
        return;
    }

    try {
        showStatus('Generating text document...', 'info');
        downloadBtn.disabled = true;

        const dateRange = getDateRange(parsedData);
        let filename = 'PATIENT REPORT _ DATED ';
        if (dateRange.min && dateRange.max) {
            filename += `${dateRange.min} - ${dateRange.max}.txt`;
        } else {
            filename += 'Unknown.txt';
        }

        const content = textPreview.textContent || '';
        const blob = new Blob([content], { type: 'text/plain;charset=utf-8' });
        saveAs(blob, filename);

        showStatus('Text document generated successfully!', 'success');
        downloadBtn.disabled = false;
    } catch (error) {
        showStatus('Error generating text document: ' + error.message, 'error');
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

// Helper function to format date for header (e.g., "Jan 21")
function formatHeaderDate(dateValue) {
    if (!dateValue) return '';
    
    let date;
    
    // If it's an Excel serial number, convert it
    if (typeof dateValue === 'number' || !isNaN(dateValue)) {
        const utc_days = Math.floor(dateValue - 25569);
        const utc_value = utc_days * 86400;
        date = new Date(utc_value * 1000);
    } else {
        // Try to parse as date string
        date = new Date(dateValue);
    }
    
    // If invalid date, return empty
    if (isNaN(date.getTime())) {
        return String(dateValue);
    }
    
    // Format as "Mon DD" (e.g., "Jan 21")
    const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
                       "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
    const month = monthNames[date.getMonth()];
    const day = date.getDate();
    
    return `${month} ${day}`;
}

// Helper function to get min and max dates from data
function getDateRange(data) {
    if (!data || data.length <= 1) return { min: '', max: '' };
    
    const headers = data[0] || [];
    const rows = data.slice(1);
    const visitDateIndex = headers.findIndex(h => String(h).toLowerCase().includes('visit date'));
    
    if (visitDateIndex === -1) return { min: '', max: '' };
    
    const dates = rows
        .map(row => row[visitDateIndex])
        .filter(date => date !== undefined && date !== null && date !== '');
    
    if (dates.length === 0) return { min: '', max: '' };
    
    // Convert all to comparable format
    const comparableDates = dates.map(d => {
        if (typeof d === 'number') return d;
        const parsed = new Date(d);
        return isNaN(parsed.getTime()) ? 0 : parsed.getTime();
    });
    
    const minValue = Math.min(...comparableDates);
    const maxValue = Math.max(...comparableDates);
    
    // Find original values
    const minIndex = comparableDates.indexOf(minValue);
    const maxIndex = comparableDates.indexOf(maxValue);
    
    return {
        min: formatHeaderDate(dates[minIndex]),
        max: formatHeaderDate(dates[maxIndex])
    };
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

// Create document content matching the Resources report format:
// each patient is a 2-column table (label | value) with Arial 12pt bold,
// personal reminders in the first row with yellow/green highlight.
function createDocumentContent(data, lib) {
    const docxLib = lib || window.docx || docx;

    const children = [];

    if (!data || data.length <= 1) {
        children.push(new docxLib.Paragraph({ text: 'No data available.' }));
        return children;
    }

    const headers = data[0] || [];
    const rows = filterEmptyRows(data.slice(1));

    const ptNoIndex = headers.findIndex(h => String(h).toLowerCase().includes('pt no'));
    const patientNameIndex = headers.findIndex(h => String(h).toLowerCase().includes('patient name'));
    const visitDateIndex = headers.findIndex(h => String(h).toLowerCase().includes('visit date'));
    const doctorIndex = headers.findIndex(h => String(h).toLowerCase().includes('doctor'));
    const personalRemindersIndex = headers.findIndex(h => String(h).toLowerCase().includes('personal reminders'));

    const font = 'Arial';
    const sz = 24; // 12pt in half-points

    const makeRun = (text, highlightColor) => {
        const opts = {
            text: String(text || ''),
            font,
            size: sz,
            bold: true,
            color: '000000'
        };
        if (highlightColor) opts.highlight = highlightColor;
        return new docxLib.TextRun(opts);
    };

    const borderDef = { style: 'single', size: 1, color: '000000' };
    const tableBorders = {
        top: borderDef,
        bottom: borderDef,
        left: borderDef,
        right: borderDef,
        insideHorizontal: borderDef,
        insideVertical: borderDef
    };

    const makeCell = (runs, widthDxa) => new docxLib.TableCell({
        children: [new docxLib.Paragraph({ children: runs })],
        width: { size: widthDxa, type: 'dxa' },
        margins: { top: 0, bottom: 0, left: 108, right: 108 }
    });

    rows.forEach((row, index) => {
        if (index > 0) {
            children.push(new docxLib.Paragraph({ text: '' }));
        }

        const ptNo = row[ptNoIndex] !== undefined ? String(row[ptNoIndex]) : '';
        const patientName = row[patientNameIndex] !== undefined ? String(row[patientNameIndex]) : '';
        const visitDate = row[visitDateIndex] !== undefined ? formatDate(row[visitDateIndex]) : '';
        const doctor = row[doctorIndex] !== undefined ? String(row[doctorIndex]).trim() : '';
        const personalReminders = row[personalRemindersIndex] !== undefined ? String(row[personalRemindersIndex]).trim() : '';

        // Highlight colour: yellow for "LAST VISIT …", green for "NEW PATIENT …"
        let reminderHighlight = null;
        const remUpper = personalReminders.toUpperCase();
        if (remUpper.startsWith('NEW PATIENT')) {
            reminderHighlight = 'green';
        } else if (remUpper.startsWith('LAST VISIT')) {
            reminderHighlight = 'yellow';
        }

        const tableRows = [];

        // Row 0 – Personal Reminders (left cell empty, right cell = reminders text)
        tableRows.push(new docxLib.TableRow({
            children: [
                makeCell([], 2405),
                makeCell(personalReminders ? [makeRun(personalReminders, reminderHighlight)] : [], 5670)
            ]
        }));

        // Row 1 – Date
        tableRows.push(new docxLib.TableRow({
            children: [
                makeCell([makeRun('Date:')], 2405),
                makeCell([makeRun(visitDate)], 5670)
            ]
        }));

        // Row 2 – File Number
        tableRows.push(new docxLib.TableRow({
            children: [
                makeCell([makeRun('File Number:')], 2405),
                makeCell([makeRun(ptNo)], 5670)
            ]
        }));

        // Row 3 – Patient name
        tableRows.push(new docxLib.TableRow({
            children: [
                makeCell([makeRun('Patient name:')], 2405),
                makeCell([makeRun(patientName)], 5670)
            ]
        }));

        // Row 4 – Doctor Name (omit row when empty, matching the reference format)
        if (doctor) {
            tableRows.push(new docxLib.TableRow({
                children: [
                    makeCell([makeRun('Doctor Name:')], 2405),
                    makeCell([makeRun(doctor)], 5670)
                ]
            }));
        }

        children.push(new docxLib.Table({
            rows: tableRows,
            width: { size: 8075, type: 'dxa' },
            borders: tableBorders
        }));
    });

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

// Filter out empty rows from data
function filterEmptyRows(rows) {
    return rows.filter(row => {
        // A row is empty if all cells are null, undefined, empty, or whitespace-only
        // Using != null to check for both null and undefined (nullish coalescing)
        return row && row.some(cell => cell != null && String(cell).trim() !== '');
    });
}
