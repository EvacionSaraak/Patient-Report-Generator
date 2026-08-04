// Handle file selection and initiate XLSX parsing
async function handleFileSelect(event) {
    const file = event.target.files[0];

    if (!file) {
        return;
    }

    if (!/\.xlsx$/i.test(file.name)) {
        event.target.value = '';
        showStatus('Please select a .xlsx file (Excel 2007+). DOCX, XLS and CSV files are not accepted.', 'error');
        return;
    }

    fileName.textContent = `Selected: ${file.name}`;
    showStatus('Reading file…', 'info');

    try {
        const arrayBuffer = await file.arrayBuffer();
        await parseXlsxInput(arrayBuffer);
        showStatus('File loaded successfully!', 'success');
        downloadBtn.disabled = false;
    } catch (error) {
        showStatus('Error reading file: ' + error.message, 'error');
        downloadBtn.disabled = true;
    }
}

// Parse Patient Report XLSX and set up all previews.
// parsedData is set to the canonical object returned by parsePatientReportXlsx:
//   { format, records: [{fileNumber, visitId, patientName, visitDate, doctor, personalReminders, query, status}], warnings }
async function parseXlsxInput(arrayBuffer) {
    const canonical = parsePatientReportXlsx(arrayBuffer);

    parsedData = canonical;

    // Show row-skip warnings
    if (canonical.warnings && canonical.warnings.length > 0) {
        showStatus(canonical.warnings.join('\n'), 'warning');
    }

    displayPreview(canonical.records);
    generateWordPreview(parsedData);
    generateTextPreview(parsedData);
    reportPreviewSection.style.display = 'block';
}
