const EXAMPLE_OUTPUT_PATH = 'Resources/REPORT FOR OPG (EXAMPLE OUTPUT).docx';
const EXAMPLE_OUTPUT_FILENAME = 'REPORT FOR OPG (EXAMPLE OUTPUT).docx';

// Example page logic
const EXAMPLE_DATA = [
    ['PT NO.', 'Patient Name', 'Visit Date', 'Doctor', 'Personal Reminders'],
    [2401, 'Ahmad Al-Hassan', 46212, 'Dr. Samir Nasser', 'OPG taken today'],
    [2402, 'Fatima Al-Rashid', 46212, 'Dr. Layla Karimi', ''],
    [2403, 'Mohammed Al-Khatib', 46211, 'Dr. Samir Nasser', 'New OPG required'],
    [2404, 'Sara Abdullah', 46211, 'Dr. Layla Karimi', ''],
    [2405, 'Omar Al-Farsi', 46210, 'Dr. Samir Nasser', ''],
    [2406, 'Nour Al-Amin', 46210, 'Dr. Layla Karimi', 'OPG imaging done'],
    [2407, 'Khalid Al-Mansour', 46212, 'Dr. Samir Nasser', ''],
    [2408, 'Rania Al-Zahra', 46211, 'Dr. Layla Karimi', '']
];

function initExample() {
    document.getElementById('exampleDownloadBtn').addEventListener('click', downloadExactExampleOutput);
    displayExampleDataPreview(EXAMPLE_DATA);
    renderExactExampleOutputPreview();
}

function displayExampleDataPreview(data) {
    const container = document.getElementById('exampleDataPreview');
    if (!data || data.length === 0) {
        container.innerHTML = '<p>No example data.</p>';
        return;
    }

    const headers = data[0] || [];
    let html = '<table><thead><tr>';
    headers.forEach(header => {
        html += `<th>${escapeHtml(String(header || ''))}</th>`;
    });
    html += '</tr></thead><tbody>';

    data.slice(1).forEach(row => {
        html += '<tr>';
        headers.forEach((_, index) => {
            const cellValue = row[index] !== undefined ? row[index] : '';
            const displayValue = index === 2 ? formatDate(cellValue) : escapeHtml(String(cellValue));
            html += `<td>${displayValue}</td>`;
        });
        html += '</tr>';
    });

    html += '</tbody></table>';
    container.innerHTML = html;
}

function renderExactExampleOutputPreview() {
    const preview = document.getElementById('exampleOutputPreview');
    preview.innerHTML = `
        <div class="document-preview">
            <h3 class="mb-3">Exact sample output</h3>
            <p class="mb-2">This page now returns the provided example Word document directly.</p>
            <p class="mb-2"><strong>Output file:</strong> ${escapeHtml(EXAMPLE_OUTPUT_FILENAME)}</p>
            <p class="mb-0 text-muted">Download the file below to see the exact table layout and formatting from the sample output in <code>Resources</code>.</p>
        </div>
    `;
}

function downloadExactExampleOutput() {
    try {
        showExampleStatus('Downloading exact example output...', 'info');

        const link = document.createElement('a');
        link.href = encodeURI(EXAMPLE_OUTPUT_PATH);
        link.download = EXAMPLE_OUTPUT_FILENAME;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);

        showExampleStatus('Exact example output downloaded.', 'success');
    } catch (error) {
        showExampleStatus('Error downloading example output: ' + error.message, 'error');
        console.error(error);
    }
}

function showExampleStatus(message, type) {
    const statusDiv = document.getElementById('exampleStatus');
    statusDiv.textContent = message;
    statusDiv.className = 'status-message ' + type;
}
