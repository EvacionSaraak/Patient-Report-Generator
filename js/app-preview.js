// Set active download format and toggle visible panels
function setDownloadFormat(format) {
    selectedDownloadFormat = format;
    const isWordFormat = format === 'docx';

    wordTabBtn.classList.toggle('active', isWordFormat);
    textTabBtn.classList.toggle('active', !isWordFormat);
    wordPreviewPanel.style.display = isWordFormat ? 'block' : 'none';
    textPreviewPanel.style.display = isWordFormat ? 'none' : 'block';
    downloadBtnText.textContent = isWordFormat ? 'Download Word Report' : 'Download Text Report';
}

// Display raw XLSX data preview table.
// Receives rawRows.slice(headerRowIndex) so data[0] is always the header row.
function displayPreview(data) {
    if (!data || data.length === 0) {
        dataPreview.innerHTML = '<p>No data found in the spreadsheet.</p>';
        previewSection.style.display = 'block';
        return;
    }

    let html = '<table><thead><tr>';

    const headers = data[0] || [];
    headers.forEach(header => {
        html += `<th>${escapeHtml(String(header || ''))}</th>`;
    });
    html += '</tr></thead><tbody>';

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

// Generate Word document preview from canonical parsed result.
// Mirrors the 2-column table format of the downloaded .docx.
function generateWordPreview(parsed) {
    if (!parsed || !parsed.records || parsed.records.length === 0) {
        wordPreview.innerHTML = '<p class="text-muted">No data to preview.</p>';
        return;
    }

    let html = '<div class="document-preview">';

    parsed.records.forEach((record, index) => {
        if (index > 0) {
            html += '<div style="margin: 8px 0;"></div>';
        }

        const ptNo              = record.fileNumber;
        const patientName       = record.patientName;
        const visitDate         = formatDate(record.visitDate);
        const doctor            = record.doctor;
        const personalReminders = record.personalReminders;

        let reminderStyle = '';
        const remUpper = personalReminders.toUpperCase();
        if (remUpper.startsWith('NEW PATIENT')) {
            reminderStyle = 'background-color: #90EE90;';
        } else if (remUpper.startsWith('LAST VISIT')) {
            reminderStyle = 'background-color: yellow;';
        }

        html += '<table style="border-collapse:collapse;width:100%;font-family:Arial,sans-serif;font-size:12pt;font-weight:bold;">';

        // Personal Reminders row (highlighted when applicable)
        html += '<tr>';
        html += '<td style="border:1px solid #000;padding:2px 6px;width:30%;"></td>';
        if (personalReminders) {
            html += `<td style="border:1px solid #000;padding:2px 6px;${reminderStyle}">${escapeHtml(personalReminders)}</td>`;
        } else {
            html += '<td style="border:1px solid #000;padding:2px 6px;"></td>';
        }
        html += '</tr>';

        html += `<tr><td style="border:1px solid #000;padding:2px 6px;">Date:</td><td style="border:1px solid #000;padding:2px 6px;">${escapeHtml(visitDate)}</td></tr>`;
        html += `<tr><td style="border:1px solid #000;padding:2px 6px;">File Number:</td><td style="border:1px solid #000;padding:2px 6px;">${escapeHtml(ptNo)}</td></tr>`;
        html += `<tr><td style="border:1px solid #000;padding:2px 6px;">Patient name:</td><td style="border:1px solid #000;padding:2px 6px;">${escapeHtml(patientName)}</td></tr>`;

        if (doctor) {
            html += `<tr><td style="border:1px solid #000;padding:2px 6px;">Doctor Name:</td><td style="border:1px solid #000;padding:2px 6px;">${escapeHtml(doctor)}</td></tr>`;
        }

        html += '</table>';
    });

    html += '</div>';
    wordPreview.innerHTML = html;
}

// Generate text report preview from canonical parsed result.
function generateTextPreview(parsed) {
    if (!parsed || !parsed.records || parsed.records.length === 0) {
        textPreview.textContent = 'No data to preview.';
        return;
    }

    const dateRange = getDateRange(parsed.records);
    const headerText = dateRange.min && dateRange.max
        ? `PATIENT REPORT | ${dateRange.min} - ${dateRange.max}`
        : 'PATIENT REPORT';

    const lines = [headerText, ''];

    parsed.records.forEach((record, index) => {
        const visitDate         = formatDate(record.visitDate);
        const personalReminders = record.personalReminders;

        lines.push(`Date: ${visitDate}`);
        lines.push(` File Number: ${record.fileNumber}`);
        lines.push(` Patient Name: ${record.patientName}`);
        lines.push(` Doctor Name: ${record.doctor}`);
        if (personalReminders && personalReminders.trim()) {
            lines.push(` Remarks: ${personalReminders}`);
        }

        if (index < parsed.records.length - 1) {
            lines.push('', '----------------------------------------', '');
        }
    });

    textPreview.textContent = lines.join('\n');
}

// Refresh both previews from the current canonical parsed result.
function refreshWordPreview() {
    if (parsedData) {
        generateWordPreview(parsedData);
        generateTextPreview(parsedData);
        showStatus('Preview refreshed!', 'success');
    }
}
