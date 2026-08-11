// Render the parsed patient records as a summary table (left panel)
function renderOPGDataPreview(records) {
    const container = document.getElementById('exampleDataPreview');

    if (!records || !records.length) {
        container.innerHTML = '<p class="text-muted mb-0">No records were accepted.</p>';
        return;
    }

    container.innerHTML =
        `<table>
            <thead>
                <tr>
                    <th>File Number</th>
                    <th>Patient Name</th>
                    <th>Date</th>
                    <th>Doctor</th>
                    <th>Personal Reminders</th>
                </tr>
            </thead>
            <tbody>
                ${records.map(rec =>
                    `<tr>
                        <td>${opgEscapeHtml(rec.fileNumber)}</td>
                        <td>${opgEscapeHtml(rec.patientName)}</td>
                        <td>${opgEscapeHtml(rec.date)}</td>
                        <td>${opgEscapeHtml(rec.doctorName)}</td>
                        <td>${opgEscapeHtml(rec.personalReminders)}</td>
                    </tr>`
                ).join('')}
            </tbody>
        </table>`;
}

// Strip a leading "Dr." or "Dr " prefix (case-insensitive) so we never output "Dr. Dr. …"
function opgNormaliseDoctorName(name) {
    return String(name || '').trim().replace(/^dr\.?\s*/i, '');
}

// Render the OPG report card preview (right panel) – shows the exact OPG structure
function renderOPGReportPreview(records, reportDate) {
    const container = document.getElementById('exampleOutputPreview');

    if (!records || !records.length) {
        container.innerHTML = '<p class="text-muted mb-0">No OPG preview is available.</p>';
        return;
    }

    const dateLabel = reportDate ? `Date: ${opgEscapeHtml(reportDate)}` : '';

    container.innerHTML =
        `<div class="opg-document-preview">
            ${dateLabel ? `<p class="fw-bold mb-3">${dateLabel}</p>` : ''}
            ${records.map(rec => {
                const reminder = (rec.personalReminders || '').trim();
                const reminderHtml = reminder
                    ? `<p class="opg-reminder-para">${opgEscapeHtml(reminder)}</p>`
                    : `<p class="opg-reminder-para opg-reminder-empty">&nbsp;</p>`;

                return `<section class="opg-preview-record">
                    <table class="opg-info-table">
                        <colgroup>
                            <col style="width:${(2117/9913*100).toFixed(2)}%">
                            <col style="width:${(5386/9913*100).toFixed(2)}%">
                            <col style="width:${(2410/9913*100).toFixed(2)}%">
                        </colgroup>
                        <tr>
                            <td>File&nbsp;#&nbsp;&nbsp;${opgEscapeHtml(rec.fileNumber)}</td>
                            <td>Pt.&nbsp;Name&nbsp;&#8211;&nbsp;${opgEscapeHtml(rec.patientName)}</td>
                            <td>Dr.&nbsp;${opgEscapeHtml(opgNormaliseDoctorName(rec.doctorName))}</td>
                        </tr>
                    </table>
                    ${reminderHtml}
                    <div class="opg-empty-image-area"></div>
                </section>`;
            }).join('')}
        </div>`;
}
