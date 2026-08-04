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
