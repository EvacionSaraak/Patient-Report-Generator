// Example page logic
// Uses utility functions defined in app.js (escapeHtml, formatDate, getRemarks,
// filterEmptyRows, getDateRange, excelDateToJSDate, formatHeaderDate)

// Excel serial numbers: July 7 = 46210, July 8 = 46211, July 9 = 46212
const EXAMPLE_DATA = [
    ['PT NO.', 'Patient Name', 'Visit Date', 'Doctor', 'Personal Reminders'],
    [2401, 'Ahmad Al-Hassan',   46212, 'Dr. Samir Nasser',  'OPG taken today'],
    [2402, 'Fatima Al-Rashid',  46212, 'Dr. Layla Karimi',  ''],
    [2403, 'Mohammed Al-Khatib',46211, 'Dr. Samir Nasser',  'New OPG required'],
    [2404, 'Sara Abdullah',     46211, 'Dr. Layla Karimi',  ''],
    [2405, 'Omar Al-Farsi',     46210, 'Dr. Samir Nasser',  ''],
    [2406, 'Nour Al-Amin',      46210, 'Dr. Layla Karimi',  'OPG imaging done'],
    [2407, 'Khalid Al-Mansour', 46212, 'Dr. Samir Nasser',  ''],
    [2408, 'Rania Al-Zahra',    46211, 'Dr. Layla Karimi',  ''],
];

let exampleSelectedFormat = 'docx';
let exampleDocxLib = null;

// DOM element references for the example section
function getExampleEls() {
    return {
        dataPreview:      document.getElementById('exampleDataPreview'),
        wordPreview:      document.getElementById('exampleWordPreview'),
        textPreview:      document.getElementById('exampleTextPreview'),
        wordPreviewPanel: document.getElementById('exampleWordPreviewPanel'),
        textPreviewPanel: document.getElementById('exampleTextPreviewPanel'),
        wordTabBtn:       document.getElementById('exampleWordTabBtn'),
        textTabBtn:       document.getElementById('exampleTextTabBtn'),
        downloadBtn:      document.getElementById('exampleDownloadBtn'),
        downloadBtnText:  document.getElementById('exampleDownloadBtnText'),
        statusDiv:        document.getElementById('exampleStatus'),
    };
}

// Called when the Example section is shown for the first time
function initExample() {
    // Resolve docx library
    exampleDocxLib = (typeof docxLib !== 'undefined' && docxLib)
        || window.docx
        || (typeof docx !== 'undefined' ? docx : null);

    const els = getExampleEls();

    // Wire tab buttons
    els.wordTabBtn.addEventListener('click', () => setExampleFormat('docx'));
    els.textTabBtn.addEventListener('click', () => setExampleFormat('txt'));

    // Wire download button
    els.downloadBtn.addEventListener('click', downloadExampleReport);

    // Populate previews
    displayExampleDataPreview(EXAMPLE_DATA);
    generateExampleWordPreview(EXAMPLE_DATA);
    generateExampleTextPreview(EXAMPLE_DATA);

    setExampleFormat('docx');
}

// ── Data preview table ──────────────────────────────────────────────────────

function displayExampleDataPreview(data) {
    const container = document.getElementById('exampleDataPreview');
    if (!data || data.length === 0) {
        container.innerHTML = '<p>No example data.</p>';
        return;
    }

    const headers = data[0] || [];
    let html = '<table><thead><tr>';
    headers.forEach(h => { html += `<th>${escapeHtml(String(h || ''))}</th>`; });
    html += '</tr></thead><tbody>';

    data.slice(1).forEach(row => {
        html += '<tr>';
        headers.forEach((_, i) => {
            const cell = row[i] !== undefined ? row[i] : '';
            // Render dates as human-readable
            const display = (i === 2) ? formatDate(cell) : escapeHtml(String(cell));
            html += `<td>${display}</td>`;
        });
        html += '</tr>';
    });

    html += '</tbody></table>';
    container.innerHTML = html;
}

// ── Word preview ────────────────────────────────────────────────────────────

function generateExampleWordPreview(data) {
    const container = document.getElementById('exampleWordPreview');
    if (!data || data.length === 0) {
        container.innerHTML = '<p class="text-muted">No data to preview.</p>';
        return;
    }

    const headers = data[0] || [];
    const rows = filterEmptyRows(data.slice(1));

    const ptNoIndex           = headers.findIndex(h => String(h).toLowerCase().includes('pt no'));
    const patientNameIndex    = headers.findIndex(h => String(h).toLowerCase().includes('patient name'));
    const visitDateIndex      = headers.findIndex(h => String(h).toLowerCase().includes('visit date'));
    const doctorIndex         = headers.findIndex(h => String(h).toLowerCase().includes('doctor'));
    const personalRemindersIndex = headers.findIndex(h => String(h).toLowerCase().includes('personal reminders'));

    const dateRange = getDateRange(data);
    const headerText = dateRange.min && dateRange.max
        ? `PATIENT REPORT | ${dateRange.min} - ${dateRange.max}`
        : 'PATIENT REPORT';

    let html = '<div class="document-preview">';
    html += `<h3 class="mb-3">${escapeHtml(headerText)}</h3>`;

    rows.forEach((row, index) => {
        if (index > 0) html += '<hr class="my-4">';

        const ptNo            = row[ptNoIndex] !== undefined ? String(row[ptNoIndex]) : '';
        const patientName     = row[patientNameIndex] !== undefined ? String(row[patientNameIndex]) : '';
        const visitDate       = row[visitDateIndex] !== undefined ? formatDate(row[visitDateIndex]) : '';
        const doctor          = row[doctorIndex] !== undefined ? String(row[doctorIndex]) : '';
        const personalReminders = row[personalRemindersIndex] !== undefined ? row[personalRemindersIndex] : '';
        const remarks         = getRemarks(personalReminders);

        html += `<div class="patient-record mb-3">`;
        html += `<p class="mb-1"><strong>Date:</strong> ${escapeHtml(visitDate)}</p>`;
        html += `<p class="mb-1 ms-2"><strong>File Number:</strong> ${escapeHtml(ptNo)}</p>`;
        html += `<p class="mb-1 ms-2"><strong>Patient Name:</strong> ${escapeHtml(patientName)}</p>`;
        html += `<p class="mb-1 ms-2"><strong>Doctor Name:</strong> ${escapeHtml(doctor)}</p>`;

        if (remarks && remarks.trim()) {
            if (remarks === 'Patient with new OPG') {
                html += `<p class="mb-1 ms-2"><strong>Remarks:</strong> <span style="background-color: yellow;">${escapeHtml(remarks)}</span></p>`;
            } else {
                html += `<p class="mb-1 ms-2"><strong>Remarks:</strong> ${escapeHtml(remarks)}</p>`;
            }
        }

        html += `</div>`;
    });

    html += '</div>';
    container.innerHTML = html;
}

// ── Text preview ────────────────────────────────────────────────────────────

function generateExampleTextPreview(data) {
    const container = document.getElementById('exampleTextPreview');
    if (!data || data.length === 0) {
        container.textContent = 'No data to preview.';
        return;
    }

    const headers = data[0] || [];
    const rows    = filterEmptyRows(data.slice(1));

    const ptNoIndex              = headers.findIndex(h => String(h).toLowerCase().includes('pt no'));
    const patientNameIndex       = headers.findIndex(h => String(h).toLowerCase().includes('patient name'));
    const visitDateIndex         = headers.findIndex(h => String(h).toLowerCase().includes('visit date'));
    const doctorIndex            = headers.findIndex(h => String(h).toLowerCase().includes('doctor'));
    const personalRemindersIndex = headers.findIndex(h => String(h).toLowerCase().includes('personal reminders'));

    const dateRange = getDateRange(data);
    const headerText = dateRange.min && dateRange.max
        ? `PATIENT REPORT | ${dateRange.min} - ${dateRange.max}`
        : 'PATIENT REPORT';

    const lines = [headerText, ''];

    rows.forEach((row, index) => {
        const ptNo            = row[ptNoIndex] !== undefined ? String(row[ptNoIndex]) : '';
        const patientName     = row[patientNameIndex] !== undefined ? String(row[patientNameIndex]) : '';
        const visitDate       = row[visitDateIndex] !== undefined ? formatDate(row[visitDateIndex]) : '';
        const doctor          = row[doctorIndex] !== undefined ? String(row[doctorIndex]) : '';
        const personalReminders = row[personalRemindersIndex] !== undefined ? row[personalRemindersIndex] : '';
        const remarks         = getRemarks(personalReminders);

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

    container.textContent = lines.join('\n');
}

// ── Format toggle ───────────────────────────────────────────────────────────

function setExampleFormat(format) {
    const els = getExampleEls();
    exampleSelectedFormat = format;
    const isWord = format === 'docx';

    els.wordTabBtn.classList.toggle('active', isWord);
    els.textTabBtn.classList.toggle('active', !isWord);
    els.wordPreviewPanel.style.display = isWord ? 'block' : 'none';
    els.textPreviewPanel.style.display = isWord ? 'none'  : 'block';
    els.downloadBtnText.textContent = isWord
        ? 'Download Example Word Report'
        : 'Download Example Text Report';
}

// ── Download ────────────────────────────────────────────────────────────────

async function downloadExampleReport() {
    if (exampleSelectedFormat === 'txt') {
        downloadExampleText();
        return;
    }
    await downloadExampleWord();
}

async function downloadExampleWord() {
    const els = getExampleEls();
    try {
        showExampleStatus('Generating Word document…', 'info');
        els.downloadBtn.disabled = true;

        const lib = exampleDocxLib
            || window.docx
            || (typeof docx !== 'undefined' ? docx : null);

        if (!lib) throw new Error('docx library not loaded. Please refresh and try again.');

        const content = createDocumentContent(EXAMPLE_DATA, lib);
        const doc = new lib.Document({ sections: [{ properties: {}, children: content }] });

        const dateRange = getDateRange(EXAMPLE_DATA);
        let filename = 'EXAMPLE PATIENT REPORT _ DATED ';
        filename += dateRange.min && dateRange.max
            ? `${dateRange.min} - ${dateRange.max}.docx`
            : 'Unknown.docx';

        const blob = await lib.Packer.toBlob(doc);
        saveAs(blob, filename);

        showExampleStatus('Example Word document downloaded!', 'success');
    } catch (err) {
        showExampleStatus('Error generating document: ' + err.message, 'error');
        console.error(err);
    } finally {
        els.downloadBtn.disabled = false;
    }
}

function downloadExampleText() {
    const els = getExampleEls();
    try {
        showExampleStatus('Generating text document…', 'info');
        els.downloadBtn.disabled = true;

        const dateRange = getDateRange(EXAMPLE_DATA);
        let filename = 'EXAMPLE PATIENT REPORT _ DATED ';
        filename += dateRange.min && dateRange.max
            ? `${dateRange.min} - ${dateRange.max}.txt`
            : 'Unknown.txt';

        const content = els.textPreview.textContent || '';
        const blob = new Blob([content], { type: 'text/plain;charset=utf-8' });
        saveAs(blob, filename);

        showExampleStatus('Example text document downloaded!', 'success');
    } catch (err) {
        showExampleStatus('Error generating text document: ' + err.message, 'error');
        console.error(err);
    } finally {
        els.downloadBtn.disabled = false;
    }
}

function showExampleStatus(message, type) {
    const statusDiv = document.getElementById('exampleStatus');
    statusDiv.textContent = message;
    statusDiv.className   = 'status-message ' + type;
}
