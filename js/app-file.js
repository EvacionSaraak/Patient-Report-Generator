// Handle file selection and initiate DOCX parsing
async function handleFileSelect(event) {
    const file = event.target.files[0];

    if (!file) {
        return;
    }

    if (!/\.docx$/i.test(file.name)) {
        event.target.value = '';
        showStatus('Please select a .docx file (Word 2007+). Excel and legacy .doc files are not accepted.', 'error');
        return;
    }

    fileName.textContent = `Selected: ${file.name}`;
    showStatus('Reading file…', 'info');

    try {
        const arrayBuffer = await file.arrayBuffer();
        await parseDocxInput(arrayBuffer);
        showStatus('File loaded successfully!', 'success');
        downloadBtn.disabled = false;
    } catch (error) {
        showStatus('Error reading file: ' + error.message, 'error');
        downloadBtn.disabled = true;
    }
}

// Parse Patient Report DOCX and set up all previews.
// parsedData is set to the canonical object returned by parsePatientReportInputDocx:
//   { format, records: [{visitDate, fileNumber, patientName, doctor, personalReminders}] }
async function parseDocxInput(arrayBuffer) {
    const canonical = await parsePatientReportInputDocx(arrayBuffer);

    parsedData = canonical;

    displayPreview(canonical.records);
    generateWordPreview(parsedData);
    generateTextPreview(parsedData);
    reportPreviewSection.style.display = 'block';
}
