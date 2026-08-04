// Normal report helpers copied/adapted from origin/main app.js formatting logic.

function getNormalRemarks(personalReminders) {
    if (!personalReminders) {
        return '';
    }
    const remindersStr = String(personalReminders).toUpperCase();
    if (remindersStr.includes('OPG')) {
        return 'Patient with new OPG';
    }
    return '';
}

function getNormalReportHeaderText(parsed) {
    const dateRange = getDateRange(parsed.records || []);
    return dateRange.min && dateRange.max
        ? `PATIENT REPORT | ${dateRange.min} - ${dateRange.max}`
        : 'PATIENT REPORT';
}

function generateNormalWordPreviewHtml(parsed) {
    const headerText = getNormalReportHeaderText(parsed);
    const records = parsed.records || [];

    let html = '<div class="document-preview">';
    html += `<h3 class="mb-3">${escapeHtml(headerText)}</h3>`;

    records.forEach((record, index) => {
        if (index > 0) {
            html += '<hr class="my-4">';
        }

        const ptNo = record.fileNumber !== undefined ? String(record.fileNumber) : '';
        const patientName = record.patientName !== undefined ? String(record.patientName) : '';
        const visitDate = record.visitDate !== undefined ? formatDate(record.visitDate) : '';
        const doctor = record.doctor !== undefined ? String(record.doctor) : '';
        const remarks = getNormalRemarks(record.personalReminders);

        html += '<div class="patient-record mb-3">';
        html += `<p class="mb-1"><strong>Date:</strong> ${escapeHtml(visitDate)}</p>`;
        html += `<p class="mb-1 ms-2"><strong>File Number:</strong> ${escapeHtml(ptNo)}</p>`;
        html += `<p class="mb-1 ms-2"><strong>Patient Name:</strong> ${escapeHtml(patientName)}</p>`;
        html += `<p class="mb-1 ms-2"><strong>Doctor Name:</strong> ${escapeHtml(doctor)}</p>`;

        if (remarks && remarks.trim()) {
            if (remarks === 'Patient with new OPG') {
                html += `<p class="mb-1 ms-2"><strong>Remarks:</strong> <span style="background-color: yellow;">${escapeHtml(remarks)}</span></p>`;
            } else {
                html += `<p class="mb-1 ms-2"><strong>Remarks:</strong> ${escapeHtml(remarks)}</p>`;
            }
        }

        html += '</div>';
    });

    html += '</div>';
    return html;
}

function generateNormalTextReport(parsed) {
    const records = parsed.records || [];
    const headerText = getNormalReportHeaderText(parsed);
    const lines = [headerText, ''];

    records.forEach((record, index) => {
        const visitDate = record.visitDate !== undefined ? formatDate(record.visitDate) : '';
        const remarks = getNormalRemarks(record.personalReminders);

        lines.push(`Date: ${visitDate}`);
        lines.push(` File Number: ${record.fileNumber || ''}`);
        lines.push(` Patient Name: ${record.patientName || ''}`);
        lines.push(` Doctor Name: ${record.doctor || ''}`);
        if (remarks && remarks.trim()) {
            lines.push(` Remarks: ${remarks}`);
        }

        if (index < records.length - 1) {
            lines.push('', '----------------------------------------', '');
        }
    });

    return lines.join('\n');
}

function createNormalDocumentContent(parsed, lib) {
    const docxLib = lib || window.docx || docx;
    const children = [];
    const headerText = getNormalReportHeaderText(parsed);
    const records = parsed.records || [];

    children.push(
        new docxLib.Paragraph({
            children: [
                new docxLib.TextRun({
                    text: headerText,
                    font: 'Calibri',
                    size: 32,
                    bold: true,
                    color: '000000'
                })
            ],
            spacing: {
                after: 400
            }
        })
    );

    if (records.length > 0) {
        records.forEach((record, index) => {
            if (index > 0) {
                children.push(
                    new docxLib.Paragraph({
                        text: '',
                        border: {
                            top: {
                                color: '000000',
                                space: 1,
                                style: 'single',
                                size: 6
                            }
                        },
                        spacing: {
                            before: 200,
                            after: 200
                        }
                    })
                );
            }

            const ptNo = record.fileNumber !== undefined ? String(record.fileNumber) : '';
            const patientName = record.patientName !== undefined ? String(record.patientName) : '';
            const visitDate = record.visitDate !== undefined ? formatDate(record.visitDate) : '';
            const doctor = record.doctor !== undefined ? String(record.doctor) : '';
            const remarks = getNormalRemarks(record.personalReminders);

            children.push(
                new docxLib.Paragraph({
                    children: [
                        new docxLib.TextRun({
                            text: 'Date',
                            bold: true,
                            font: 'Calibri',
                            size: 24,
                            color: '000000'
                        }),
                        new docxLib.TextRun({
                            text: `: ${visitDate}`,
                            font: 'Calibri',
                            size: 24,
                            color: '000000'
                        })
                    ],
                    spacing: { after: 100 }
                })
            );

            children.push(
                new docxLib.Paragraph({
                    children: [
                        new docxLib.TextRun({
                            text: ' File Number',
                            bold: true,
                            font: 'Calibri',
                            size: 24,
                            color: '000000'
                        }),
                        new docxLib.TextRun({
                            text: `: ${ptNo}`,
                            font: 'Calibri',
                            size: 24,
                            color: '000000'
                        })
                    ],
                    spacing: { after: 100 }
                })
            );

            children.push(
                new docxLib.Paragraph({
                    children: [
                        new docxLib.TextRun({
                            text: ' Patient Name',
                            bold: true,
                            font: 'Calibri',
                            size: 24,
                            color: '000000'
                        }),
                        new docxLib.TextRun({
                            text: `: ${patientName}`,
                            font: 'Calibri',
                            size: 24,
                            color: '000000'
                        })
                    ],
                    spacing: { after: 100 }
                })
            );

            children.push(
                new docxLib.Paragraph({
                    children: [
                        new docxLib.TextRun({
                            text: ' Doctor Name',
                            bold: true,
                            font: 'Calibri',
                            size: 24,
                            color: '000000'
                        }),
                        new docxLib.TextRun({
                            text: `: ${doctor}`,
                            font: 'Calibri',
                            size: 24,
                            color: '000000'
                        })
                    ],
                    spacing: { after: 100 }
                })
            );

            if (remarks && remarks.trim()) {
                children.push(
                    new docxLib.Paragraph({
                        children: [
                            new docxLib.TextRun({
                                text: ' Remarks',
                                bold: true,
                                font: 'Calibri',
                                size: 24,
                                color: '000000'
                            }),
                            new docxLib.TextRun({
                                text: ': ',
                                font: 'Calibri',
                                size: 24,
                                color: '000000'
                            }),
                            new docxLib.TextRun({
                                text: remarks,
                                font: 'Calibri',
                                size: 24,
                                color: '000000',
                                highlight: remarks === 'Patient with new OPG' ? 'yellow' : undefined
                            })
                        ],
                        spacing: { after: 100 }
                    })
                );
            }

            if ((index + 1) % 5 === 0 && index + 1 < records.length) {
                children.push(
                    new docxLib.Paragraph({
                        text: '',
                        pageBreakBefore: true
                    })
                );
            }
        });
    } else {
        children.push(
            new docxLib.Paragraph({
                text: 'No data available.',
                font: 'Calibri',
                spacing: {
                    before: 200
                }
            })
        );
    }

    return children;
}
