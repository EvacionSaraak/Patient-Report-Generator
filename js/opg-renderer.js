// Render warning alerts for rows skipped due to missing data
function renderOPGWarnings(rows) {
    const container = document.getElementById('opgWarning');

    if (!rows.length) {
        container.innerHTML = '';
        return;
    }

    const items = rows.map(row =>
        `<li>
            <strong>${opgEscapeHtml(row.claimId || `Row ${row.rowNumber} (Claim ID missing)`)}</strong>
            — Missing: ${opgEscapeHtml(row.missing.join(', '))}
        </li>`
    ).join('');

    container.innerHTML =
        `<div class="alert alert-warning mb-0">
            <strong>${rows.length} row${rows.length === 1 ? '' : 's'} ignored due to missing data:</strong>
            <ul class="mb-0 mt-2">${items}</ul>
        </div>`;
}

// Render the accepted claim rows as a summary table
function renderOPGDataPreview(rows) {
    const container = document.getElementById('exampleDataPreview');

    if (!rows.length) {
        container.innerHTML = '<p class="text-muted mb-0">No complete rows were accepted.</p>';
        return;
    }

    container.innerHTML =
        `<table>
            <thead>
                <tr>
                    <th>Claim ID</th>
                    <th>File #</th>
                    <th>Patient</th>
                    <th>Performing Clinician</th>
                    <th>Date</th>
                    <th>Last Modified By</th>
                </tr>
            </thead>
            <tbody>
                ${rows.map(row =>
                    `<tr>
                        <td>${opgEscapeHtml(row.claimId)}</td>
                        <td>${opgEscapeHtml(row.fileNumber)}</td>
                        <td>${opgEscapeHtml(row.patientName)}</td>
                        <td>${opgEscapeHtml(row.doctor)}</td>
                        <td>${opgEscapeHtml(row.date)}</td>
                        <td>${opgEscapeHtml(row.lastModifiedBy)}</td>
                    </tr>`
                ).join('')}
            </tbody>
        </table>`;
}

// Render the OPG report card preview for each accepted row
function renderOPGReportPreview(rows) {
    const container = document.getElementById('exampleOutputPreview');

    if (!rows.length) {
        container.innerHTML = '<p class="text-muted mb-0">No OPG preview is available.</p>';
        return;
    }

    container.innerHTML =
        `<div class="opg-document-preview">
            ${rows.map(row =>
                `<section class="opg-preview-record">
                    <table class="opg-info-table">
                        <colgroup>
                            <col style="width:22%">
                            <col style="width:55%">
                            <col style="width:23%">
                        </colgroup>
                        <tr>
                            <td>File #&nbsp;&nbsp;${opgEscapeHtml(row.fileNumber)}</td>
                            <td>Pt. Name -&nbsp;&nbsp;${opgEscapeHtml(row.patientName)}</td>
                            <td>Dr. ${opgEscapeHtml(row.doctor)}</td>
                        </tr>
                    </table>
                    <div class="opg-reminder-line">&nbsp;</div>
                    <div class="opg-empty-image-area"></div>
                </section>`
            ).join('')}
        </div>`;
}
