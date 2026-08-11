// OPG module state
let opgRecords = [];         // canonical records from parsePatientReportDocx
let opgDocxModulePromise = null;

// Initialise the OPG section (called once when the section is first shown)
function initExample() {
    const input  = document.getElementById('opgFileInput');
    const button = document.getElementById('exampleDownloadBtn');

    if (!input || !button) {
        return;
    }

    input.addEventListener('change', handleOPGFileSelect);
    button.addEventListener('click', downloadOPGReport);
    button.disabled = true;

    setOPGEmptyState();
}

// Reset the OPG UI to its default empty state
function setOPGEmptyState() {
    document.getElementById('exampleDataPreview').innerHTML =
        '<p class="text-muted mb-0">Upload a patient-report DOCX file to view parsed records.</p>';

    document.getElementById('exampleOutputPreview').innerHTML =
        '<p class="text-muted mb-0">The OPG report preview will appear here after a file is loaded.</p>';

    document.getElementById('opgWarning').innerHTML = '';
    document.getElementById('exampleStatus').textContent = '';
    document.getElementById('exampleStatus').className = 'status-message mt-2';
}

// Handle file selection for the OPG DOCX input
async function handleOPGFileSelect(event) {
    const file   = event.target.files[0];
    const button = document.getElementById('exampleDownloadBtn');
    const name   = document.getElementById('opgFileName');

    opgRecords = [];
    button.disabled = true;
    setOPGEmptyState();

    if (!file) {
        name.textContent = '';
        return;
    }

    // Accept only .docx (reject legacy .doc and other formats)
    if (!/\.docx$/i.test(file.name)) {
        name.textContent = '';
        event.target.value = '';
        showExampleStatus(
            'Please select a .docx file (Word 2007+). Legacy .doc files are not supported.',
            'error'
        );
        return;
    }

    name.textContent = `Selected: ${file.name}`;
    showExampleStatus('Parsing patient report…', 'info');

    try {
        const arrayBuffer = await file.arrayBuffer();
        const result = await parseOPGInputDocx(arrayBuffer);

        opgRecords = result.records;

        document.getElementById('opgWarning').innerHTML = '';
        renderOPGDataPreview(opgRecords);
        renderOPGReportPreview(opgRecords, result.date);

        button.disabled = false;

        showExampleStatus(
            `${opgRecords.length} patient record${opgRecords.length === 1 ? '' : 's'} parsed successfully.`,
            'success'
        );
    } catch (err) {
        console.error('OPG DOCX parse error:', err);
        opgRecords = [];
        button.disabled = true;
        document.getElementById('opgWarning').innerHTML =
            `<div class="alert alert-danger mb-0">${opgEscapeHtml(err.message)}</div>`;
        showExampleStatus('Could not parse the selected file.', 'error');
    }
}
