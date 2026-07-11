// OPG Report page logic

// Remove a leading "Dr." prefix (case-insensitive) so that the template label "Dr. {dr}"
// does not produce "Dr. Dr. XXXX" when the source data already starts with "Dr.".
function stripDrPrefix(name) {
    return name.replace(/^dr\.?\s+/i, '').trim();
}
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

    const ptNoIdx = headers.findIndex(header =>
        String(header).toLowerCase().includes('pt no')
    );

    const nameIdx = headers.findIndex(header =>
        String(header).toLowerCase().includes('patient name')
    );

    const drIdx = headers.findIndex(header =>
        String(header).toLowerCase().includes('doctor')
    );

    const remindersIdx = headers.findIndex(header =>
        String(header).toLowerCase().includes('personal reminders')
    );

    const pageStyle = [
        'border:2px solid #ccc',
        'padding:12px',
        'margin-bottom:16px',
        'background:#fff',
        'page-break-after:always'
    ].join(';');

    const tableStyle = [
        'border-collapse:collapse',
        'width:100%',
        'font-family:Arial,sans-serif',
        'font-weight:bold',
        'font-size:11pt'
    ].join(';');

    const cellStyle = [
        'border:2px solid #000',
        'padding:2px 8px',
        'vertical-align:middle',
        'white-space:nowrap'
    ].join(';');

    const imageAreaStyle = [
        'border:1px solid #000',
        'width:100%',
        'height:420px',
        'box-sizing:border-box',
        'margin-top:16px'
    ].join(';');

    let html = '';

    rows.forEach(row => {
        const fileNo = String(
            row[ptNoIdx] !== undefined
                ? row[ptNoIdx]
                : ''
        ).trim();

        const name = String(
            row[nameIdx] !== undefined
                ? row[nameIdx]
                : ''
        ).trim();

        const doctor = stripDrPrefix(
            String(
                row[drIdx] !== undefined
                    ? row[drIdx]
                    : ''
            ).trim()
        );

        const reminder = String(
            row[remindersIdx] !== undefined
                ? row[remindersIdx]
                : ''
        ).trim();

        const reminderUpper = reminder.toUpperCase();

        let reminderBackground = 'transparent';

        if (reminderUpper.includes('NEW PATIENT')) {
            reminderBackground = '#00ff00';
        } else if (reminderUpper.includes('LAST VISIT')) {
            reminderBackground = '#ffff00';
        }

        html += `<div style="${pageStyle}">`;

        html += `
            <table style="${tableStyle}">
                <colgroup>
                    <col style="width:22%;">
                    <col style="width:55%;">
                    <col style="width:23%;">
                </colgroup>

                <tr>
                    <td style="${cellStyle}">
                        File #&nbsp;&nbsp;${escapeHtml(fileNo)}
                    </td>

                    <td style="${cellStyle}">
                        Pt. Name -&nbsp;&nbsp;${escapeHtml(name)}
                    </td>

                    <td style="${cellStyle}">
                        Dr. ${escapeHtml(doctor)}
                    </td>
                </tr>
            </table>
        `;

        if (reminder) {
            html += `
                <div
                    style="
                        display:table;
                        margin-top:2px;
                        padding:0 2px;
                        background:${reminderBackground};
                        font-family:Arial,sans-serif;
                        font-size:11pt;
                        line-height:1.05;
                        font-weight:bold;
                    "
                >
                    ${escapeHtml(reminder)}
                </div>
            `;
        }

        html += `
            <div style="${imageAreaStyle}"></div>
        `;

        html += '</div>';
    });

    container.innerHTML = html;
}

async function downloadOPGReport() {
    try {
        showExampleStatus(
            'Generating OPG Report...',
            'info'
        );

        if (
            typeof PizZip === 'undefined' ||
            typeof docxtemplater === 'undefined'
        ) {
            throw new Error(
                'Templating libraries not loaded. Please refresh the page.'
            );
        }

        if (typeof OPG_REPORT_TEMPLATE_B64 === 'undefined') {
            throw new Error(
                'OPG report template not found. Please refresh the page.'
            );
        }

        const headers = EXAMPLE_DATA[0];
        const rows = EXAMPLE_DATA.slice(1);

        const ptNoIdx = headers.findIndex(header =>
            String(header)
                .toLowerCase()
                .includes('pt no')
        );

        const nameIdx = headers.findIndex(header =>
            String(header)
                .toLowerCase()
                .includes('patient name')
        );

        const drIdx = headers.findIndex(header =>
            String(header)
                .toLowerCase()
                .includes('doctor')
        );

        const remindersIdx = headers.findIndex(header =>
            String(header)
                .toLowerCase()
                .includes('personal reminders')
        );

        const dateIdx = headers.findIndex(header =>
            String(header)
                .toLowerCase()
                .includes('visit date')
        );

        const patients = rows.map(row => {
            const reminder = String(
                row[remindersIdx] !== undefined
                    ? row[remindersIdx]
                    : ''
            ).trim();

            return {
                file_no: String(
                    row[ptNoIdx] !== undefined
                        ? row[ptNoIdx]
                        : ''
                ).trim(),

                pt_name: String(
                    row[nameIdx] !== undefined
                        ? row[nameIdx]
                        : ''
                ).trim(),

                dr: stripDrPrefix(
                    String(
                        row[drIdx] !== undefined
                            ? row[drIdx]
                            : ''
                    ).trim()
                ),

                reminder: reminder,

                reminder_type: /^NEW\s+PATIENT/i.test(reminder)
                    ? 'new'
                    : /^LAST\s+VISIT/i.test(reminder)
                        ? 'last'
                        : ''
            };
        });

        const zip = new PizZip(
            OPG_REPORT_TEMPLATE_B64,
            {
                base64: true
            }
        );

        const doc = new docxtemplater(
            zip,
            {
                paragraphLoop: true,
                linebreaks: true
            }
        );

        doc.render({
            patients
        });

        const dateValue =
            rows[0] &&
            rows[0][dateIdx] !== undefined
                ? rows[0][dateIdx]
                : null;

        const dateString =
            dateValue !== null
                ? formatDate(dateValue).toUpperCase()
                : 'OPG';

        const filename =
            `REPORT FOR OPG - ${dateString}.docx`;

        const blob = doc
            .getZip()
            .generate({
                type: 'blob',
                mimeType:
                    'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            });

        saveAs(
            blob,
            filename
        );

        showExampleStatus(
            'OPG Report downloaded.',
            'success'
        );

    } catch (error) {
        showExampleStatus(
            'Error generating OPG Report: ' +
            error.message,
            'error'
        );

        console.error(error);
    }
}

function showExampleStatus(message, type) {
    const statusDiv = document.getElementById('exampleStatus');
    statusDiv.textContent = message;
    statusDiv.className = 'status-message ' + type;
}
