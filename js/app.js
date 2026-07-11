// Global variables
let workbookData = null;
let parsedData = null;
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

// Generate Word document using docxtemplater and the patient_report_template.docx
async function generateWordDocument() {
    if (!parsedData || parsedData.length === 0) {
        showStatus('No data to export.', 'error');
        return;
    }

    try {
        showStatus('Generating Word document...', 'info');
        downloadBtn.disabled = true;

        if (typeof PizZip === 'undefined' || typeof docxtemplater === 'undefined') {
            throw new Error('Templating libraries not loaded. Please refresh the page.');
        }
        if (typeof PATIENT_REPORT_TEMPLATE_B64 === 'undefined') {
            throw new Error('Report template not found. Please refresh the page.');
        }

        const headers = parsedData[0] || [];
        const rows = filterEmptyRows(parsedData.slice(1));

        const ptNoIdx       = headers.findIndex(h => String(h).toLowerCase().includes('pt no'));
        const nameIdx       = headers.findIndex(h => String(h).toLowerCase().includes('patient name'));
        const dateIdx       = headers.findIndex(h => String(h).toLowerCase().includes('visit date'));
        const drIdx         = headers.findIndex(h => String(h).toLowerCase().includes('doctor'));
        const remindersIdx  = headers.findIndex(h => String(h).toLowerCase().includes('personal reminders'));

        const patients = rows.map(row => {
            const reminder = remindersIdx >= 0 && row[remindersIdx] !== undefined
                ? String(row[remindersIdx]).trim() : '';
            const remUpper = reminder.toUpperCase();
            return {
                file_no:    ptNoIdx  >= 0 && row[ptNoIdx]  !== undefined ? String(row[ptNoIdx])  : '',
                pt_name:    nameIdx  >= 0 && row[nameIdx]  !== undefined ? String(row[nameIdx])  : '',
                visit_date: dateIdx  >= 0 && row[dateIdx]  !== undefined ? formatDate(row[dateIdx]) : '',
                dr:         drIdx    >= 0 && row[drIdx]    !== undefined ? String(row[drIdx]).trim() : '',
                reminder,
                is_yellow: remUpper.startsWith('LAST VISIT'),
                is_green:  remUpper.startsWith('NEW PATIENT'),
            };
        });

        const zip = new PizZip(PATIENT_REPORT_TEMPLATE_B64, { base64: true });
        const doc = new docxtemplater(zip, { paragraphLoop: true, linebreaks: true });
        doc.render({ patients });

        const dateRange = getDateRange(parsedData);
        let filename = 'PATIENT REPORT DATED ';
        filename += dateRange.min ? dateRange.min : 'Unknown';
        filename += '.docx';

        const blob = doc.getZip().generate({
            type: 'blob',
            mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        });
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
