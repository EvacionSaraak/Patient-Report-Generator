// OPG Report page logic
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
    document.getElementById('exampleDownloadBtn').addEventListener('click', downloadOPGReport);
    displayExampleDataPreview(EXAMPLE_DATA);
    renderOPGReportPreview(EXAMPLE_DATA);
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

function renderOPGReportPreview(data) {
    const container = document.getElementById('exampleOutputPreview');
    if (!data || data.length <= 1) {
        container.innerHTML = '<p>No data.</p>';
        return;
    }

    const headers = data[0] || [];
    const rows = data.slice(1);
    const ptNoIdx = headers.findIndex(h => String(h).toLowerCase().includes('pt no'));
    const nameIdx = headers.findIndex(h => String(h).toLowerCase().includes('patient name'));
    const drIdx = headers.findIndex(h => String(h).toLowerCase().includes('doctor'));
    const remindersIdx = headers.findIndex(h => String(h).toLowerCase().includes('personal reminders'));
    const dateIdx = headers.findIndex(h => String(h).toLowerCase().includes('visit date'));

    const visitDate = rows[0] && rows[0][dateIdx] !== undefined
        ? formatDate(rows[0][dateIdx]).toUpperCase()
        : '';

    let html = '<div class="document-preview">';
    html += '<h3 class="mb-2">REPORT FOR OPG</h3>';
    if (visitDate) {
        html += `<p class="mb-3"><strong>Date:</strong> ${escapeHtml(visitDate)}</p>`;
    }

    html += '<table><thead><tr>';
    html += '<th>File #</th><th>Pt. Name -</th><th>Dr.</th>';
    html += '</tr></thead><tbody>';

    rows.forEach(row => {
        const fileNo = String(row[ptNoIdx] !== undefined ? row[ptNoIdx] : '');
        const name = String(row[nameIdx] !== undefined ? row[nameIdx] : '');
        const reminders = String(row[remindersIdx] !== undefined ? row[remindersIdx] : '');
        const ptNameCell = reminders ? `${name} - ${reminders}` : name;
        const dr = String(row[drIdx] !== undefined ? row[drIdx] : '');

        html += '<tr>';
        html += `<td>${escapeHtml(fileNo)}</td>`;
        html += `<td>${escapeHtml(ptNameCell)}</td>`;
        html += `<td>${escapeHtml(dr)}</td>`;
        html += '</tr>';
    });

    html += '</tbody></table></div>';
    container.innerHTML = html;
}

async function downloadOPGReport() {
    try {
        showExampleStatus('Generating OPG Report...', 'info');

        let lib = (typeof docxLib !== 'undefined' && docxLib) || window.docx;
        if (!lib && typeof docx !== 'undefined') {
            lib = docx;
        }
        if (!lib) {
            throw new Error('docx library not loaded. Please refresh the page and try again.');
        }

        const children = generateOPGReportContent(EXAMPLE_DATA, lib);
        const doc = new lib.Document({
            sections: [{ properties: {}, children }]
        });

        const blob = await lib.Packer.toBlob(doc);

        const headers = EXAMPLE_DATA[0];
        const dateIdx = headers.findIndex(h => String(h).toLowerCase().includes('visit date'));
        const dateVal = EXAMPLE_DATA[1] && EXAMPLE_DATA[1][dateIdx];
        const dateStr = dateVal !== undefined ? formatDate(dateVal).toUpperCase() : 'OPG';

        saveAs(blob, `REPORT FOR OPG - ${dateStr}.docx`);
        showExampleStatus('OPG Report downloaded.', 'success');
    } catch (error) {
        showExampleStatus('Error generating OPG Report: ' + error.message, 'error');
        console.error(error);
    }
}

function generateOPGReportContent(data, lib) {
    const headers = data[0] || [];
    const rows = data.slice(1);

    const ptNoIdx = headers.findIndex(h => String(h).toLowerCase().includes('pt no'));
    const nameIdx = headers.findIndex(h => String(h).toLowerCase().includes('patient name'));
    const drIdx = headers.findIndex(h => String(h).toLowerCase().includes('doctor'));
    const remindersIdx = headers.findIndex(h => String(h).toLowerCase().includes('personal reminders'));
    const dateIdx = headers.findIndex(h => String(h).toLowerCase().includes('visit date'));

    const font = 'Arial';
    const sz = 24; // 12pt (half-points)

    const makeRun = (text, bold) => new lib.TextRun({
        text: String(text || ''),
        font,
        size: sz,
        bold: !!bold,
        color: '000000'
    });

    const makeCell = (text, bold, fill) => {
        const cellOpts = {
            children: [new lib.Paragraph({ children: [makeRun(text, bold)] })],
            margins: { top: 80, bottom: 80, left: 120, right: 120 }
        };
        if (fill) {
            cellOpts.shading = { fill, type: 'solid', color: 'auto' };
        }
        return new lib.TableCell(cellOpts);
    };

    const children = [];

    // Title
    children.push(new lib.Paragraph({
        children: [makeRun('REPORT FOR OPG', true)],
        spacing: { after: 200 }
    }));

    // Date line
    const visitDate = rows[0] && rows[0][dateIdx] !== undefined
        ? formatDate(rows[0][dateIdx]).toUpperCase()
        : '';
    if (visitDate) {
        children.push(new lib.Paragraph({
            children: [makeRun(`Date: ${visitDate}`, false)],
            spacing: { after: 300 }
        }));
    }

    // Table rows
    const tableRows = [];

    // Header row (shaded)
    tableRows.push(new lib.TableRow({
        children: [
            makeCell('File #  ', true, 'D0D0D0'),
            makeCell('Pt. Name - ', true, 'D0D0D0'),
            makeCell('Dr. ', true, 'D0D0D0'),
        ]
    }));

    // Data rows
    rows.forEach(row => {
        const fileNo = String(row[ptNoIdx] !== undefined ? row[ptNoIdx] : '');
        const name = String(row[nameIdx] !== undefined ? row[nameIdx] : '');
        const reminders = String(row[remindersIdx] !== undefined ? row[remindersIdx] : '');
        const ptNameCell = reminders ? `${name} - ${reminders}` : name;
        const dr = String(row[drIdx] !== undefined ? row[drIdx] : '');

        tableRows.push(new lib.TableRow({
            children: [
                makeCell(fileNo, false, null),
                makeCell(ptNameCell, false, null),
                makeCell(dr, false, null),
            ]
        }));
    });

    children.push(new lib.Table({
        rows: tableRows,
        columnWidths: [2117, 5386, 2413],
        width: { size: 9916, type: 'dxa' }
    }));

    return children;
}

function showExampleStatus(message, type) {
    const statusDiv = document.getElementById('exampleStatus');
    statusDiv.textContent = message;
    statusDiv.className = 'status-message ' + type;
}
