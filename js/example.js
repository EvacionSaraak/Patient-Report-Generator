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
    try {
        showExampleStatus('Generating OPG Report...', 'info');

        let lib =
            (typeof docxLib !== 'undefined' && docxLib) ||
            window.docx;

        if (!lib && typeof docx !== 'undefined') {
            lib = docx;
        }

        if (!lib) {
            throw new Error(
                'docx library not loaded. Please refresh the page and try again.'
            );
        }

        if (!EXAMPLE_DATA || EXAMPLE_DATA.length <= 1) {
            throw new Error('No patient data is available.');
        }

        const headers = EXAMPLE_DATA[0] || [];
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
                'One or more required input columns could not be found.'
            );
        }

        /*
         * Measurements copied from the supplied example document.
         */
        const INFO_TABLE_WIDTH = 9913;

        const INFO_COLUMN_WIDTHS = [
            2117,
            5386,
            2410
        ];

        const OPG_TABLE_WIDTH = 9891;

        const OPG_CELL_HEIGHT = 4891;

        const FONT = 'Arial';

        /*
         * Word font sizes use half-points.
         * 18 = 9 pt.
         */
        const FONT_SIZE = 18;

        const widthType = lib.WidthType
            ? lib.WidthType.DXA
            : 'dxa';

        const fixedLayout = lib.TableLayoutType
            ? lib.TableLayoutType.FIXED
            : 'fixed';

        const exactHeight = lib.HeightRule
            ? lib.HeightRule.EXACT
            : 'exact';

        const singleBorder = {
            style: lib.BorderStyle
                ? lib.BorderStyle.SINGLE
                : 'single',

            size: 12,
            color: '000000'
        };

        const tableBorders = {
            top: singleBorder,
            bottom: singleBorder,
            left: singleBorder,
            right: singleBorder,
            insideHorizontal: singleBorder,
            insideVertical: singleBorder
        };

        function makeTextRun(
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

            return new lib.TextRun(options);
        }

        function makeInformationCell(
            text,
            width
        ) {
            return new lib.TableCell({
                width: {
                    size: width,
                    type: widthType
                },

                children: [
                    new lib.Paragraph({
                        children: [
                            makeTextRun(text)
                        ],

                        spacing: {
                            before: 0,
                            after: 0,
                            line: 240
                        }
                    })
                ]
            });
        }

        function createInformationTable(
            fileNumber,
            patientName,
            doctor
        ) {
            return new lib.Table({
                width: {
                    size: INFO_TABLE_WIDTH,
                    type: widthType
                },

                columnWidths: INFO_COLUMN_WIDTHS,

                layout: fixedLayout,

                borders: tableBorders,

                rows: [
                    new lib.TableRow({
                        cantSplit: true,

                        height: {
                            value: 92,
                            rule: exactHeight
                        },

                        children: [
                            makeInformationCell(
                                `File #  ${fileNumber}`,
                                INFO_COLUMN_WIDTHS[0]
                            ),

                            makeInformationCell(
                                `Pt. Name - ${patientName}`,
                                INFO_COLUMN_WIDTHS[1]
                            ),

                            makeInformationCell(
                                doctor,
                                INFO_COLUMN_WIDTHS[2]
                            )
                        ]
                    })
                ]
            });
        }

        function createReminderParagraph(
            reminder
        ) {
            const reminderUpper = String(
                reminder || ''
            )
                .trim()
                .toUpperCase();

            let highlight = null;

            /*
             * NEW PATIENT / NEW VISIT = green
             * LAST VISIT = yellow
             */
            if (
                reminderUpper.startsWith('NEW PATIENT') ||
                reminderUpper.startsWith('NEW VISIT')
            ) {
                highlight = 'green';
            } else if (
                reminderUpper.startsWith('LAST VISIT')
            ) {
                highlight = 'yellow';
            }

            return new lib.Paragraph({
                children: reminder
                    ? [
                        makeTextRun(
                            reminder,
                            highlight
                        )
                    ]
                    : [],

                spacing: {
                    before: 0,
                    after: 0,
                    line: 240
                },

                keepNext: true
            });
        }

        function createEmptyOPGTable() {
            return new lib.Table({
                width: {
                    size: OPG_TABLE_WIDTH,
                    type: widthType
                },

                columnWidths: [
                    OPG_TABLE_WIDTH
                ],

                layout: fixedLayout,

                borders: tableBorders,

                rows: [
                    new lib.TableRow({
                        cantSplit: true,

                        height: {
                            value: OPG_CELL_HEIGHT,
                            rule: exactHeight
                        },

                        children: [
                            new lib.TableCell({
                                width: {
                                    size: OPG_TABLE_WIDTH,
                                    type: widthType
                                },

                                children: [
                                    new lib.Paragraph({
                                        children: [],

                                        spacing: {
                                            before: 0,
                                            after: 0
                                        }
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
             * The supplied example uses two patients per page.
             */
            if (
                index > 0 &&
                index % 2 === 0
            ) {
                documentChildren.push(
                    new lib.Paragraph({
                        pageBreakBefore: true,
                        children: []
                    })
                );
            }

            const fileNumber = String(
                row[ptNoIdx] !== undefined
                    ? row[ptNoIdx]
                    : ''
            ).trim();

            const patientName = String(
                row[nameIdx] !== undefined
                    ? row[nameIdx]
                    : ''
            ).trim();

            const doctor = String(
                row[doctorIdx] !== undefined
                    ? row[doctorIdx]
                    : ''
            ).trim();

            const reminder = String(
                row[reminderIdx] !== undefined
                    ? row[reminderIdx]
                    : ''
            ).trim();

            /*
             * Patient information remains entirely
             * inside the three-column table.
             */
            documentChildren.push(
                createInformationTable(
                    fileNumber,
                    patientName,
                    doctor
                )
            );

            /*
             * The reminder is a separate paragraph
             * below the information table.
             *
             * It is not included in Pt. Name.
             */
            documentChildren.push(
                createReminderParagraph(
                    reminder
                )
            );

            /*
             * Large empty OPG area below the reminder.
             */
            documentChildren.push(
                createEmptyOPGTable()
            );
        });

        const document = new lib.Document({
            sections: [
                {
                    properties: {
                        page: {
                            size: {
                                width: 12240,
                                height: 15840
                            },

                            margin: {
                                top: 1440,
                                right: 1440,
                                bottom: 1440,
                                left: 1440,
                                header: 708,
                                footer: 708,
                                gutter: 0
                            }
                        }
                    },

                    children: documentChildren
                }
            ]
        });

        const blob = await lib.Packer.toBlob(
            document
        );

        const dateValue =
            dateIdx !== -1 &&
            rows[0] &&
            rows[0][dateIdx] !== undefined
                ? rows[0][dateIdx]
                : null;

        const dateString =
            dateValue !== null
                ? formatDate(
                    dateValue
                ).toUpperCase()
                : 'OPG';

        saveAs(
            blob,
            `REPORT FOR OPG - ${dateString}.docx`
        );

        showExampleStatus(
            'OPG Report downloaded.',
            'success'
        );

    } catch (error) {
        console.error(error);

        showExampleStatus(
            'Error generating OPG Report: ' +
            error.message,
            'error'
        );
    }
}

function showExampleStatus(message, type) {
    const statusDiv = document.getElementById('exampleStatus');
    statusDiv.textContent = message;
    statusDiv.className = 'status-message ' + type;
}
