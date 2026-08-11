// Wire up event listeners and set the default download format
fileInput.addEventListener('change', handleFileSelect);
downloadBtn.addEventListener('click', generateReport);
refreshPreviewBtn.addEventListener('click', refreshWordPreview);
wordTabBtn.addEventListener('click', () => setDownloadFormat('docx'));
textTabBtn.addEventListener('click', () => setDownloadFormat('txt'));
setDownloadFormat('docx');
