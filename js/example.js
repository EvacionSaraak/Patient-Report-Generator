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

    let html = '<div class="document-preview">';

    rows.forEach(row => {
        const fileNo = String(
            row[ptNoIdx] !== undefined
                ? row[ptNoIdx]
                : ''
        ).trim();

        const patientName = String(
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

        let reminderColour = 'transparent';

        if (
            reminderUpper.startsWith('NEW PATIENT') ||
            reminderUpper.startsWith('NEW VISIT')
        ) {
            reminderColour = '#00ff00';
        } else if (
            reminderUpper.startsWith('LAST VISIT')
        ) {
            reminderColour = '#ffff00';
        }

        html += `
            <div
                style="
                    margin-bottom:20px;
                    font-family:Arial,sans-serif;
                "
            >
                <table
                    style="
                        width:100%;
                        table-layout:fixed;
                        border-collapse:collapse;
                        font-size:11pt;
                        font-weight:bold;
                    "
                >
                    <colgroup>
                        <col style="width:22%;">
                        <col style="width:54%;">
                        <col style="width:24%;">
                    </colgroup>

                    <tbody>
                        <tr>
                            <td
                                style="
                                    border:1px solid #000;
                                    padding:3px 6px;
                                "
                            >
                                File #&nbsp;&nbsp;${escapeHtml(fileNo)}
                            </td>

                            <td
                                style="
                                    border:1px solid #000;
                                    padding:3px 6px;
                                "
                            >
                                Pt. Name -&nbsp;&nbsp;${escapeHtml(patientName)}
                            </td>

                            <td
                                style="
                                    border:1px solid #000;
                                    padding:3px 6px;
                                "
                            >
                                Dr. ${escapeHtml(doctor)}
                            </td>
                        </tr>
                    </tbody>
                </table>
        `;

        if (reminder) {
            html += `
                <div
                    style="
                        margin:0 0 3px 0;
                        padding:0;
                        min-height:18px;
                        font-size:11pt;
                        font-weight:bold;
                    "
                >
                    <span
                        style="
                            background:${reminderColour};
                            padding:0 2px;
                        "
                    >
                        ${escapeHtml(reminder)}
                    </span>
                </div>
            `;
        }

        html += `
                <div
                    style="
                        width:100%;
                        height:300px;
                        border:1px solid #000;
                        box-sizing:border-box;
                    "
                ></div>
            </div>
        `;
    });

    html += '</div>';

    container.innerHTML = html;
}

async function downloadOPGReport() {
    const button = document.getElementById('exampleDownloadBtn');

    try {
        showExampleStatus('Generating OPG Report...', 'info');
        if (button) button.disabled = true;

        /*
         * Import the browser-compatible DOCX module directly.
         * This avoids relying on window.docx or the currently
         * missing PizZip/Docxtemplater libraries.
         */
        const {
            Document,
            Packer,
            Paragraph,
            TextRun,
            Table,
            TableRow,
            TableCell,
            PageBreak,
            WidthType,
            TableLayoutType,
            BorderStyle,
            HeightRule,
            VerticalAlign
        } = await import(
            'https://cdn.jsdelivr.net/npm/docx@8.2.2/+esm'
        );

        if (
            !Array.isArray(EXAMPLE_DATA) ||
            EXAMPLE_DATA.length <= 1
        ) {
            throw new Error('No patient data is available.');
        }

        const headers = EXAMPLE_DATA[0] || [];

        const rows = EXAMPLE_DATA
            .slice(1)
            .filter(row =>
                row &&
                row.some(value =>
                    value !== undefined &&
                    value !== null &&
                    String(value).trim() !== ''
                )
            );

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

        const doctorIdx = headers.findIndex(header =>
            String(header)
                .toLowerCase()
                .includes('doctor')
        );

        const reminderIdx = headers.findIndex(header =>
            String(header)
                .toLowerCase()
                .includes('personal reminders')
        );

        const dateIdx = headers.findIndex(header =>
            String(header)
                .toLowerCase()
                .includes('visit date')
        );

        if (
            ptNoIdx === -1 ||
            nameIdx === -1 ||
            doctorIdx === -1 ||
            reminderIdx === -1
        ) {
            throw new Error(
                'Required patient columns could not be found.'
            );
        }

        const FONT = 'Arial';
        const FONT_SIZE = 20;

        const border = {
            style: BorderStyle.SINGLE,
            size: 4,
            color: '000000'
        };

        const tableBorders = {
            top: border,
            bottom: border,
            left: border,
            right: border,
            insideHorizontal: border,
            insideVertical: border
        };

        function getValue(row, index) {
            return index >= 0 &&
                row[index] !== undefined &&
                row[index] !== null
                    ? String(row[index]).trim()
                    : '';
        }

        function makeRun(
            text,
            highlight = null
        ) {
            const options = {
                text: String(text || ''),
                font: FONT,
                size: FONT_SIZE,
                bold: true,
                color: '000000'
            };

            if (highlight) {
                options.highlight = highlight;
            }

            return new TextRun(options);
        }

        function makeInformationCell(
            label,
            value,
            width
        ) {
            return new TableCell({
                width: {
                    size: width,
                    type: WidthType.PERCENTAGE
                },

                verticalAlign:
                    VerticalAlign.CENTER,

                margins: {
                    top: 35,
                    bottom: 35,
                    left: 75,
                    right: 75
                },

                children: [
                    new Paragraph({
                        spacing: {
                            before: 0,
                            after: 0,
                            line: 240
                        },

                        children: [
                            makeRun(label),
                            makeRun(value)
                        ]
                    })
                ]
            });
        }

        function createPatientTable(
            fileNumber,
            patientName,
            doctor
        ) {
            return new Table({
                width: {
                    size: 100,
                    type: WidthType.PERCENTAGE
                },

                layout:
                    TableLayoutType.FIXED,

                borders:
                    tableBorders,

                rows: [
                    new TableRow({
                        cantSplit: true,

                        children: [
                            makeInformationCell(
                                'File #  ',
                                fileNumber,
                                22
                            ),

                            makeInformationCell(
                                'Pt. Name -  ',
                                patientName,
                                55
                            ),

                            makeInformationCell(
                                'Dr. ',
                                doctor,
                                23
                            )
                        ]
                    })
                ]
            });
        }

        function createReminderParagraph(
            reminder
        ) {
            const reminderUpper =
                reminder.toUpperCase();

            let highlight = null;

            if (
                reminderUpper.startsWith(
                    'NEW PATIENT'
                ) ||
                reminderUpper.startsWith(
                    'NEW VISIT'
                )
            ) {
                highlight = 'green';
            } else if (
                reminderUpper.startsWith(
                    'LAST VISIT'
                )
            ) {
                highlight = 'yellow';
            }

            return new Paragraph({
                keepNext: true,

                spacing: {
                    before: 0,
                    after: 0,
                    line: 240
                },

                children: reminder
                    ? [
                        makeRun(
                            reminder,
                            highlight
                        )
                    ]
                    : []
            });
        }

        function createEmptyOPGTable() {
            return new Table({
                width: {
                    size: 100,
                    type: WidthType.PERCENTAGE
                },

                layout:
                    TableLayoutType.FIXED,

                borders:
                    tableBorders,

                rows: [
                    new TableRow({
                        cantSplit: true,

                        height: {
                            value: 3600,
                            rule: HeightRule.EXACT
                        },

                        children: [
                            new TableCell({
                                width: {
                                    size: 100,
                                    type:
                                        WidthType.PERCENTAGE
                                },

                                children: [
                                    new Paragraph({
                                        spacing: {
                                            before: 0,
                                            after: 0
                                        },

                                        children: []
                                    })
                                ]
                            })
                        ]
                    })
                ]
            });
        }

        const documentChildren = [];

        rows.forEach((row, index) => {
            /*
             * Start a new page after every
             * two patients.
             */
            if (
                index > 0 &&
                index % 2 === 0
            ) {
                documentChildren.push(
                    new Paragraph({
                        children: [
                            new PageBreak()
                        ]
                    })
                );
            }

            const fileNumber =
                getValue(
                    row,
                    ptNoIdx
                );

            const patientName =
                getValue(
                    row,
                    nameIdx
                );

            const doctor =
                stripDrPrefix(
                    getValue(
                        row,
                        doctorIdx
                    )
                );

            const reminder =
                getValue(
                    row,
                    reminderIdx
                );

            /*
             * 1. Patient-information row.
             *
             * The reminder is deliberately
             * not included in patientName.
             */
            documentChildren.push(
                createPatientTable(
                    fileNumber,
                    patientName,
                    doctor
                )
            );

            /*
             * 2. Separate reminder/status line.
             *
             * NEW PATIENT = green
             * LAST VISIT = yellow
             */
            documentChildren.push(
                createReminderParagraph(
                    reminder
                )
            );

            /*
             * 3. Large empty OPG area.
             */
            documentChildren.push(
                createEmptyOPGTable()
            );

            /*
             * Add the large gap between the
             * first and second patient shown
             * on each page.
             */
            if (
                index % 2 === 0 &&
                index < rows.length - 1
            ) {
                documentChildren.push(
                    new Paragraph({
                        spacing: {
                            before: 0,
                            after: 2800
                        },

                        children: []
                    })
                );
            }
        });

        const document = new Document({
            sections: [
                {
                    properties: {
                        page: {
                            /*
                             * A4 portrait.
                             */
                            size: {
                                width: 11906,
                                height: 16838
                            },

                            margin: {
                                top: 900,
                                right: 1134,
                                bottom: 900,
                                left: 1134,
                                header: 0,
                                footer: 0,
                                gutter: 0
                            }
                        }
                    },

                    children:
                        documentChildren
                }
            ]
        });

        const blob =
            await Packer.toBlob(
                document
            );

        const firstDate =
            dateIdx >= 0 &&
            rows[0]
                ? rows[0][dateIdx]
                : null;

        const formattedDate =
            firstDate !== undefined &&
            firstDate !== null &&
            firstDate !== ''
                ? formatDate(
                    firstDate
                ).toUpperCase()
                : 'OPG';

        saveAs(
            blob,
            `REPORT FOR OPG - ${formattedDate}.docx`
        );

        showExampleStatus(
            'OPG Report downloaded.',
            'success'
        );

    } catch (error) {
        console.error(
            'OPG report error:',
            error
        );

        showExampleStatus(
            'Error generating OPG Report: ' +
            error.message,
            'error'
        );

    } finally {
        if (button) {
            button.disabled = false;
        }
    }
}

function showExampleStatus(message, type) {
    const statusDiv = document.getElementById('exampleStatus');
    statusDiv.textContent = message;
    statusDiv.className = 'status-message ' + type;
}
