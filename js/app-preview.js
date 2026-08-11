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

// Display XLSX-parsed records preview table.
// Receives the records array from parsePatientReportXlsx.
function displayPreview(records) {
    if (!records || records.length === 0) {
        dataPreview.innerHTML = '<p>No patient records found in the file.</p>';
        previewSection.style.display = 'block';
        return;
    }

    const headers = ['PT ID.', 'VISIT ID.', 'Patient Name', 'Visit Date', 'Doctor', 'Personal Reminders', 'Query', 'Status'];
    let html = '<table><thead><tr>';
    headers.forEach(h => { html += `<th>${escapeHtml(h)}</th>`; });
    html += '</tr></thead><tbody>';

    const previewRows = records.slice(0, 10);
    previewRows.forEach(rec => {
        html += '<tr>';
        html += `<td>${escapeHtml(String(rec.fileNumber || ''))}</td>`;
        html += `<td>${escapeHtml(String(rec.visitId || ''))}</td>`;
        html += `<td>${escapeHtml(String(rec.patientName || ''))}</td>`;
        html += `<td>${escapeHtml(String(rec.visitDate || ''))}</td>`;
        html += `<td>${escapeHtml(String(rec.doctor || ''))}</td>`;
        html += `<td>${escapeHtml(String(rec.personalReminders || ''))}</td>`;
        html += `<td>${escapeHtml(String(rec.query || ''))}</td>`;
        html += `<td>${escapeHtml(String(rec.status || ''))}</td>`;
        html += '</tr>';
    });

    html += '</tbody></table>';

    if (records.length > 10) {
        html += `<p style="margin-top: 10px; color: #718096;">Showing 10 of ${records.length} records</p>`;
    }

    dataPreview.innerHTML = html;
    previewSection.style.display = 'block';
}

function generateWordPreview(parsed) {
    if (!parsed || !parsed.records || parsed.records.length === 0) {
        wordPreview.innerHTML = '<p class="text-muted">No data to preview.</p>';
        return;
    }

    wordPreview.innerHTML = generateNormalWordPreviewHtml(parsed);
}

function generateTextPreview(parsed) {
    if (!parsed || !parsed.records || parsed.records.length === 0) {
        textPreview.textContent = 'No data to preview.';
        return;
    }

    textPreview.textContent = generateNormalTextReport(parsed);
}

// Refresh both previews from the current canonical parsed result.
function refreshWordPreview() {
    if (parsedData) {
        generateWordPreview(parsedData);
        generateTextPreview(parsedData);
        showStatus('Preview refreshed!', 'success');
    }
}
