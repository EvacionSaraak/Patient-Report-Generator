// OPG module state
let opgRows = [];
let opgDocxModulePromise = null;

const OPG_REQUIRED_HEADERS = [
    'Performing Clinician',
    'Patient',
    'Last Modified By',
    'Date',
    'Claim ID'
];

// Initialise the OPG section (called once when the section is first shown)
function initExample() {
    const input = document.getElementById('opgFileInput');
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
        '<p class="text-muted mb-0">Upload an XLSX or XLS file to view accepted rows.</p>';

    document.getElementById('exampleOutputPreview').innerHTML =
        '<p class="text-muted mb-0">The OPG report preview will appear here after a file is loaded.</p>';

    document.getElementById('opgWarning').innerHTML = '';
}

// Handle file selection for the OPG claim report input
async function handleOPGFileSelect(event) {
    const file = event.target.files[0];
    const button = document.getElementById('exampleDownloadBtn');
    const name = document.getElementById('opgFileName');

    opgRows = [];
    button.disabled = true;
    setOPGEmptyState();

    if (!file) {
        name.textContent = '';
        return;
    }

    if (!/\.xls[xmb]?$/i.test(file.name)) {
        name.textContent = '';
        event.target.value = '';
        showExampleStatus('Please select a valid XLSX or XLS file.', 'error');
        return;
    }

    if (typeof XLSX === 'undefined') {
        showExampleStatus('SheetJS did not load. Refresh the page and try again.', 'error');
        return;
    }

    name.textContent = `Selected: ${file.name}`;
    showExampleStatus('Reading claim report...', 'info');

    try {
        const workbook = XLSX.read(await file.arrayBuffer(), { type: 'array', cellDates: true });

        const result = parseOPGWorkbook(workbook);

        opgRows = result.accepted;

        renderOPGWarnings(result.invalid);
        renderOPGDataPreview(opgRows);
        renderOPGReportPreview(opgRows);

        button.disabled = !opgRows.length;

        const skipped = result.invalid.length;

        showExampleStatus(
            `${opgRows.length} row${opgRows.length === 1 ? '' : 's'} accepted${skipped ? `; ${skipped} skipped with warnings` : ''}.`,
            opgRows.length ? 'success' : 'warning'
        );
    } catch (error) {
        console.error('OPG input error:', error);
        opgRows = [];
        button.disabled = true;
        showExampleStatus(`Error reading claim report: ${error.message}`, 'error');
    }
}
