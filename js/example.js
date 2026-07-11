// OPG Report page logic

let opgData = null;
let opgDocxModulePromise = null;

const OPG_HEADERS = [
    'PT NO.',
    'Patient Name',
    'Visit Date',
    'Doctor',
    'Personal Reminders'
];

function stripDrPrefix(name) {
    return String(name || '')
        .replace(/^dr\.?\s+/i, '')
        .trim();
}

function normalizeWordText(value) {
    return String(value || '')
        .replace(/\u00a0/g, ' ')
        .replace(/[\t\r\n]+/g, ' ')
        .replace(/\s{2,}/g, ' ')
        .trim();
}

function initExample() {
    const fileInput =
        document.getElementById(
            'opgFileInput'
        );

    const downloadBtn =
        document.getElementById(
            'exampleDownloadBtn'
        );

    if (
        !fileInput ||
        !downloadBtn
    ) {
        console.error(
            'OPG Report controls were not found.'
        );

        return;
    }

    fileInput.addEventListener(
        'change',
        handleOPGFileSelect
    );

    downloadBtn.addEventListener(
        'click',
        downloadOPGReport
    );

    downloadBtn.disabled = true;

    document
        .getElementById(
            'exampleDataPreview'
        )
        .innerHTML = `
            <p class="text-muted mb-0">
                Upload a Patient Report DOCX file
                to view the extracted patient data.
            </p>
        `;

    document
        .getElementById(
            'exampleOutputPreview'
        )
        .innerHTML = `
            <p class="text-muted mb-0">
                The OPG report preview will appear
                here after a file is loaded.
            </p>
        `;
}

async function handleOPGFileSelect(
    event
) {
    const file =
        event.target.files[0];

    const fileName =
        document.getElementById(
            'opgFileName'
        );

    const downloadBtn =
        document.getElementById(
            'exampleDownloadBtn'
        );

    if (!file) {
        return;
    }

    opgData = null;

    downloadBtn.disabled = true;

    if (
        !/\.docx$/i.test(
            file.name
        )
    ) {
        fileName.textContent = '';

        showExampleStatus(
            'Please select a valid DOCX file.',
            'error'
        );

        event.target.value = '';

        return;
    }

    if (
        typeof mammoth ===
        'undefined'
    ) {
        showExampleStatus(
            'The Word-file reader did not load. ' +
            'Refresh the page and try again.',
            'error'
        );

        return;
    }

    fileName.textContent =
        `Selected: ${file.name}`;

    showExampleStatus(
        'Reading Patient Report...',
        'info'
    );

    try {
        const arrayBuffer =
            await file.arrayBuffer();

        const result =
            await mammoth.convertToHtml({
                arrayBuffer
            });

        const records =
            parsePatientReportHtml(
                result.value
            );

        if (
            !records.length
        ) {
            throw new Error(
                'No patient records were found. ' +
                'The file must use the ' +
                'Patient Report table format.'
            );
        }

        opgData = [
            OPG_HEADERS,

            ...records.map(
                record => [
                    record.fileNumber,
                    record.patientName,
                    record.visitDate,
                    record.doctor,
                    record.reminder
                ]
            )
        ];

        displayExampleDataPreview(
            opgData
        );

        renderOPGReportPreview(
            opgData
        );

        downloadBtn.disabled = false;

        if (
            result.messages &&
            result.messages.length
        ) {
            console.warn(
                'Mammoth conversion messages:',
                result.messages
            );
        }

        showExampleStatus(
            `${records.length} patient record` +
            `${records.length === 1 ? '' : 's'} ` +
            'loaded successfully.',
            'success'
        );

    } catch (error) {
        console.error(
            'OPG input error:',
            error
        );

        opgData = null;

        downloadBtn.disabled = true;

        document
            .getElementById(
                'exampleDataPreview'
            )
            .innerHTML = `
                <p class="text-muted mb-0">
                    No patient data loaded.
                </p>
            `;

        document
            .getElementById(
                'exampleOutputPreview'
            )
            .innerHTML = `
                <p class="text-muted mb-0">
                    No OPG preview available.
                </p>
            `;

        showExampleStatus(
            'Error reading Patient Report: ' +
            error.message,
            'error'
        );
    }
}

function getWordCellText(
    cell
) {
    const paragraphs =
        Array.from(
            cell.querySelectorAll(
                'p'
            )
        )
        .map(
            paragraph =>
                normalizeWordText(
                    paragraph.textContent
                )
        )
        .filter(
            Boolean
        );

    return normalizeWordText(
        paragraphs.length
            ? paragraphs.join(' ')
            : cell.textContent
    );
}

function parsePatientReportHtml(
    html
) {
    const parsedDocument =
        new DOMParser()
            .parseFromString(
                html,
                'text/html'
            );

    const records = [];

    parsedDocument
        .querySelectorAll(
            'table'
        )
        .forEach(
            table => {

                const record = {
                    reminder: '',
                    visitDate: '',
                    fileNumber: '',
                    patientName: '',
                    doctor: ''
                };

                table
                    .querySelectorAll(
                        'tr'
                    )
                    .forEach(
                        row => {

                            const cells =
                                Array.from(
                                    row.cells || []
                                )
                                .map(
                                    getWordCellText
                                );

                            if (
                                !cells.length
                            ) {
                                return;
                            }

                            const label =
                                normalizeWordText(
                                    cells[0]
                                )
                                .replace(
                                    /:$/,
                                    ''
                                )
                                .toLowerCase();

                            const value =
                                normalizeWordText(
                                    cells
                                        .slice(1)
                                        .join(' ')
                                );

                            if (
                                !label
                            ) {
                                if (
                                    /^(last visit|new patient|new visit)\b/i
                                        .test(
                                            value
                                        )
                                ) {
                                    record.reminder =
                                        value;
                                }

                                return;
                            }

                            if (
                                label ===
                                'date'
                            ) {
                                record.visitDate =
                                    value;

                            } else if (
                                label ===
                                    'file number' ||
                                label ===
                                    'file no' ||
                                label ===
                                    'file #'
                            ) {
                                record.fileNumber =
                                    value;

                            } else if (
                                label ===
                                    'patient name' ||
                                label ===
                                    'pt name'
                            ) {
                                record.patientName =
                                    value;

                            } else if (
                                label ===
                                    'doctor name' ||
                                label ===
                                    'doctor' ||
                                label ===
                                    'dr'
                            ) {
                                record.doctor =
                                    value;
                            }
                        }
                    );

                if (
                    record.fileNumber ||
                    record.patientName
                ) {
                    records.push(
                        record
                    );
                }
            }
        );

    return records;
}

function displayExampleDataPreview(
    data
) {
    const container =
        document.getElementById(
            'exampleDataPreview'
        );

    if (
        !data ||
        data.length <= 1
    ) {
        container.innerHTML = `
            <p class="text-muted mb-0">
                No patient data loaded.
            </p>
        `;

        return;
    }

    const headers =
        data[0];

    let html =
        '<table>' +
        '<thead>' +
        '<tr>';

    headers.forEach(
        header => {

            html += `
                <th>
                    ${escapeHtml(
                        String(
                            header || ''
                        )
                    )}
                </th>
            `;
        }
    );

    html +=
        '</tr>' +
        '</thead>' +
        '<tbody>';

    data
        .slice(1)
        .forEach(
            row => {

                html += '<tr>';

                headers.forEach(
                    (
                        _,
                        index
                    ) => {

                        html += `
                            <td>
                                ${escapeHtml(
                                    String(
                                        row[index] ??
                                        ''
                                    )
                                )}
                            </td>
                        `;
                    }
                );

                html += '</tr>';
            }
        );

    html +=
        '</tbody>' +
        '</table>';

    container.innerHTML =
        html;
}

function renderOPGReportPreview(
    data
) {
    const container =
        document.getElementById(
            'exampleOutputPreview'
        );

    if (
        !data ||
        data.length <= 1
    ) {
        container.innerHTML = `
            <p class="text-muted mb-0">
                No OPG preview available.
            </p>
        `;

        return;
    }

    const headers =
        data[0];

    const rows =
        data.slice(1);

    const ptNoIdx =
        headers.findIndex(
            header =>
                String(
                    header
                )
                .toLowerCase()
                .includes(
                    'pt no'
                )
        );

    const nameIdx =
        headers.findIndex(
            header =>
                String(
                    header
                )
                .toLowerCase()
                .includes(
                    'patient name'
                )
        );

    const doctorIdx =
        headers.findIndex(
            header =>
                String(
                    header
                )
                .toLowerCase()
                .includes(
                    'doctor'
                )
        );

    const reminderIdx =
        headers.findIndex(
            header =>
                String(
                    header
                )
                .toLowerCase()
                .includes(
                    'personal reminders'
                )
        );

    let html =
        '<div class="opg-document-preview">';

    rows.forEach(
        row => {

            const fileNumber =
                normalizeWordText(
                    row[ptNoIdx]
                );

            const patientName =
                normalizeWordText(
                    row[nameIdx]
                );

            const doctor =
                stripDrPrefix(
                    row[doctorIdx]
                );

            const reminder =
                normalizeWordText(
                    row[reminderIdx]
                );

            const reminderUpper =
                reminder.toUpperCase();

            const reminderClass =
                reminderUpper
                    .startsWith(
                        'LAST VISIT'
                    )
                    ? 'opg-reminder-last'

                    : /^(NEW PATIENT|NEW VISIT)/
                        .test(
                            reminderUpper
                        )
                        ? 'opg-reminder-new'

                        : '';

            html += `
                <section
                    class="opg-preview-record"
                >

                    <table
                        class="opg-info-table"
                    >

                        <colgroup>

                            <col
                                style="width:22%"
                            >

                            <col
                                style="width:55%"
                            >

                            <col
                                style="width:23%"
                            >

                        </colgroup>

                        <tr>

                            <td>
                                File #&nbsp;&nbsp;
                                ${escapeHtml(
                                    fileNumber
                                )}
                            </td>

                            <td>
                                Pt. Name -&nbsp;&nbsp;
                                ${escapeHtml(
                                    patientName
                                )}
                            </td>

                            <td>
                                Dr.
                                ${escapeHtml(
                                    doctor
                                )}
                            </td>

                        </tr>

                    </table>

                    <div
                        class="opg-reminder-line"
                    >

                        <span
                            class="${reminderClass}"
                        >

                            ${escapeHtml(
                                reminder
                            )}

                        </span>

                    </div>

                    <div
                        class="opg-empty-image-area"
                    ></div>

                </section>
            `;
        }
    );

    html += '</div>';

    container.innerHTML =
        html;
}

function loadOPGDocxModule() {
    if (
        !opgDocxModulePromise
    ) {
        opgDocxModulePromise =
            import(
                'https://cdn.jsdelivr.net/npm/docx@8.2.2/+esm'
            );
    }

    return opgDocxModulePromise;
}

async function downloadOPGReport() {
    const button =
        document.getElementById(
            'exampleDownloadBtn'
        );

    try {
        if (
            !opgData ||
            opgData.length <= 1
        ) {
            throw new Error(
                'Upload a Patient Report ' +
                'DOCX file first.'
            );
        }

        showExampleStatus(
            'Generating OPG Report...',
            'info'
        );

        button.disabled = true;

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
        } =
            await loadOPGDocxModule();

        const headers =
            opgData[0];

        const rows =
            opgData.slice(1);

        const ptNoIdx =
            headers.findIndex(
                header =>
                    String(
                        header
                    )
                    .toLowerCase()
                    .includes(
                        'pt no'
                    )
            );

        const nameIdx =
            headers.findIndex(
                header =>
                    String(
                        header
                    )
                    .toLowerCase()
                    .includes(
                        'patient name'
                    )
            );

        const doctorIdx =
            headers.findIndex(
                header =>
                    String(
                        header
                    )
                    .toLowerCase()
                    .includes(
                        'doctor'
                    )
            );

        const reminderIdx =
            headers.findIndex(
                header =>
                    String(
                        header
                    )
                    .toLowerCase()
                    .includes(
                        'personal reminders'
                    )
            );

        const dateIdx =
            headers.findIndex(
                header =>
                    String(
                        header
                    )
                    .toLowerCase()
                    .includes(
                        'visit date'
                    )
            );

        const border = {
            style:
                BorderStyle.SINGLE,

            size: 4,

            color:
                '000000'
        };

        const borders = {
            top:
                border,

            bottom:
                border,

            left:
                border,

            right:
                border,

            insideHorizontal:
                border,

            insideVertical:
                border
        };

        const makeRun = (
            text,
            highlight = null
        ) =>
            new TextRun({

                text:
                    String(
                        text || ''
                    ),

                font:
                    'Arial',

                size:
                    20,

                bold:
                    true,

                color:
                    '000000',

                ...(
                    highlight
                        ? {
                            highlight
                        }
                        : {}
                )
            });

        const makeInfoCell = (
            label,
            value,
            width
        ) =>
            new TableCell({

                width: {
                    size:
                        width,

                    type:
                        WidthType
                            .PERCENTAGE
                },

                verticalAlign:
                    VerticalAlign
                        .CENTER,

                margins: {
                    top:
                        35,

                    bottom:
                        35,

                    left:
                        75,

                    right:
                        75
                },

                children: [

                    new Paragraph({

                        spacing: {
                            before:
                                0,

                            after:
                                0,

                            line:
                                240
                        },

                        children: [

                            makeRun(
                                label
                            ),

                            makeRun(
                                value
                            )
                        ]
                    })
                ]
            });

        const createPatientTable = (
            fileNumber,
            patientName,
            doctor
        ) =>
            new Table({

                width: {
                    size:
                        100,

                    type:
                        WidthType
                            .PERCENTAGE
                },

                layout:
                    TableLayoutType
                        .FIXED,

                borders,

                rows: [

                    new TableRow({

                        cantSplit:
                            true,

                        children: [

                            makeInfoCell(
                                'File #  ',
                                fileNumber,
                                22
                            ),

                            makeInfoCell(
                                'Pt. Name -  ',
                                patientName,
                                55
                            ),

                            makeInfoCell(
                                'Dr. ',
                                doctor,
                                23
                            )
                        ]
                    })
                ]
            });

        const createReminderParagraph = (
            reminder
        ) => {

            const upper =
                reminder
                    .toUpperCase();

            const highlight =
                upper
                    .startsWith(
                        'LAST VISIT'
                    )
                    ? 'yellow'

                    : /^(NEW PATIENT|NEW VISIT)/
                        .test(
                            upper
                        )
                        ? 'green'

                        : null;

            return new Paragraph({

                keepNext:
                    true,

                spacing: {
                    before:
                        0,

                    after:
                        0,

                    line:
                        240
                },

                children:
                    reminder

                        ? [
                            makeRun(
                                reminder,
                                highlight
                            )
                        ]

                        : []
            });
        };

        const createEmptyOPGTable =
            () =>
                new Table({

                    width: {
                        size:
                            100,

                        type:
                            WidthType
                                .PERCENTAGE
                    },

                    layout:
                        TableLayoutType
                            .FIXED,

                    borders,

                    rows: [

                        new TableRow({

                            cantSplit:
                                true,

                            height: {
                                value:
                                    6000,

                                rule:
                                    HeightRule
                                        .EXACT
                            },

                            children: [

                                new TableCell({

                                    width: {
                                        size:
                                            100,

                                        type:
                                            WidthType
                                                .PERCENTAGE
                                    },

                                    children: [

                                        new Paragraph({
                                            children: []
                                        })
                                    ]
                                })
                            ]
                        })
                    ]
                });

        const children = [];

        rows.forEach(
            (
                row,
                index
            ) => {

                if (
                    index > 0 &&
                    index % 2 === 0
                ) {
                    children.push(

                        new Paragraph({

                            children: [

                                new PageBreak()
                            ]
                        })
                    );
                }

                const fileNumber =
                    normalizeWordText(
                        row[ptNoIdx]
                    );

                const patientName =
                    normalizeWordText(
                        row[nameIdx]
                    );

                const doctor =
                    stripDrPrefix(
                        row[doctorIdx]
                    );

                const reminder =
                    normalizeWordText(
                        row[reminderIdx]
                    );

                children.push(

                    createPatientTable(
                        fileNumber,
                        patientName,
                        doctor
                    )
                );

                children.push(

                    createReminderParagraph(
                        reminder
                    )
                );

                children.push(

                    createEmptyOPGTable()
                );

                if (
                    index % 2 === 0 &&
                    index <
                        rows.length - 1
                ) {
                    children.push(

                        new Paragraph({

                            spacing: {
                                before:
                                    0,

                                after:
                                    480
                            },

                            children:
                                []
                        })
                    );
                }
            }
        );

        const document =
            new Document({

                sections: [

                    {
                        properties: {

                            page: {

                                size: {
                                    width:
                                        11906,

                                    height:
                                        16838
                                },

                                margin: {
                                    top:
                                        650,

                                    right:
                                        270,

                                    bottom:
                                        650,

                                    left:
                                        270,

                                    header:
                                        0,

                                    footer:
                                        0,

                                    gutter:
                                        0
                                }
                            }
                        },

                        children
                    }
                ]
            });

        const blob =
            await Packer.toBlob(
                document
            );

        const dateValue =
            dateIdx >= 0 &&
            rows[0]

                ? normalizeWordText(
                    rows[0][dateIdx]
                )

                : '';

        const dateText =
            dateValue ||
            'OPG';

        saveAs(
            blob,

            'REPORT FOR OPG - ' +
            dateText
                .toUpperCase() +
            '.docx'
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

        if (
            button
        ) {
            button.disabled =
                !opgData ||
                opgData.length <= 1;
        }
    }
}

function showExampleStatus(
    message,
    type
) {
    const statusDiv =
        document.getElementById(
            'exampleStatus'
        );

    statusDiv.textContent =
        message;

    statusDiv.className =
        'status-message mt-3 ' +
        type;
}
