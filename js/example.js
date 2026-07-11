const EXAMPLE_OUTPUT_PATH = 'Resources/REPORT FOR OPG (EXAMPLE OUTPUT).docx';
const EXAMPLE_OUTPUT_FILENAME = 'REPORT FOR OPG (EXAMPLE OUTPUT).docx';

// Example page logic
const EXAMPLE_DATA = [
    ['PT NO.', 'Patient Name', 'Visit Date', 'Doctor', 'Personal Reminders'],
    ['TVIP00384762', 'Rauda hasan ismail yousef alblooshi', 46212, 'Dr. ALTAYEB Saeed Taher Abu Asbeh', 'LAST VISIT DEC. 11, 2025'],
    ['TVIP00370129', 'KHALFAN MOHAMMED ALI BUTI ALDHAHERI', 46212, 'Dr. Ahmad Hamdan', 'LAST VISIT FEB. 25, 2025'],
    ['TVIP01014122', 'HODA AZIZ SHAHIN DEZH', 46212, 'Dr. Ahmad Hamdan', 'NEW PATIENT (CASH)'],
    ['TVIP00384914', 'Reed Salem Saif Masi Alkaabi', 46212, 'Dr. Kais Altahan', 'LAST VISIT APRIL. 24, 2025'],
    ['TVIP01014497', 'MAYED KHEDHIR EISSA ABBAS MOOSA', 46212, 'Dr. FATIMA ALZHRA ALFAOUR', 'NEW PATIENT'],
    ['TVIP01014496', 'MAHRA KHEDHIR EISSA ABBAS MOOSA', 46212, 'Dr. FATIMA ALZHRA ALFAOUR', 'NEW PATIENT'],
    ['TVIP00391312', 'Sahad Khalifa Ali Muadad Almazrouei', 46212, 'Dr. Basil Mohamed Elsadig Elhag Ahmed', 'LAST VISIT AUG. 13, 2025'],
    ['TVIP00362278', 'HAMMDA SULAIMAN KHALFAN AL ALAWI', 46212, 'Dr. Basil Mohamed Elsadig Elhag Ahmed', 'LAST VISIT DEC. 06, 2025'],
    ['TVIP00357112', 'EISA DARWISH KHALIFA SALEM ALKAABI', 46212, 'Dr. Kais Altahan', 'LAST VISIT JULY 06, 2023'],
    ['TVIP01014576', 'ALI HAMAD DARWISH AHMED ALREMEITHI', 46212, 'Dr. Kais Altahan', 'NEW PATIENT'],
    ['TVIP00390512', 'Saeed Rashed Ahmed Alderei', 46212, 'Dr. Basil Mohamed Elsadig Elhag Ahmed', 'NEW PATIENT'],
    ['TVIP00345476', 'ABDULLA GHUMRAN AL DHAHERI', 46212, '', 'LAST VISIT SEPT. 23, 2025'],
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
