// Global state
let parsedData = null;
let selectedDownloadFormat = 'docx';

// DOM elements
const fileInput = document.getElementById('patient-report-file-input');
const fileName = document.getElementById('fileName');
const downloadBtn = document.getElementById('downloadBtn');
const downloadBtnText = document.getElementById('downloadBtnText');
const statusDiv = document.getElementById('status');
const previewSection = document.getElementById('previewSection');
const dataPreview = document.getElementById('dataPreview');
const reportPreviewSection = document.getElementById('reportPreviewSection');
const wordPreview = document.getElementById('wordPreview');
const textPreview = document.getElementById('textPreview');
const wordPreviewPanel = document.getElementById('wordPreviewPanel');
const textPreviewPanel = document.getElementById('textPreviewPanel');
const wordTabBtn = document.getElementById('wordTabBtn');
const textTabBtn = document.getElementById('textTabBtn');
const refreshPreviewBtn = document.getElementById('refreshPreviewBtn');
