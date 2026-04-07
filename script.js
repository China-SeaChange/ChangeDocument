// script.js
let filesQueue = [];
let processing = false;

const dropZone = document.getElementById('dropZone');
const fileInput = document.getElementById('fileInput');
const fileListSection = document.getElementById('fileListSection');
const fileTableBody = document.getElementById('fileTableBody');
const clearAllFilesBtn = document.getElementById('clearAllFilesBtn');
const batchAuthor = document.getElementById('batchAuthor');
const batchTitle = document.getElementById('batchTitle');
const batchSubject = document.getElementById('batchSubject');
const batchKeywords = document.getElementById('batchKeywords');
const namePrefix = document.getElementById('namePrefix');
const nameSuffix = document.getElementById('nameSuffix');
const statusMsg = document.getElementById('statusMsg');
const batchProcessBtn = document.getElementById('batchProcessBtn');
const resetBtn = document.getElementById('resetBtn');
const csvNote = document.getElementById('csvNote');
const macroNote = document.getElementById('macroNote');

function setStatus(message, isError = false) {
    statusMsg.innerHTML = isError ? `❌ ${message}` : `✨ ${message}`;
    statusMsg.style.background = isError ? '#ffe8e6' : '#eef3fc';
    statusMsg.style.borderLeftColor = isError ? '#e67e22' : '#2b8c6e';
}

function renderFileList() {
    if (filesQueue.length === 0) {
        fileListSection.style.display = 'none';
        batchProcessBtn.disabled = true;
        return;
    }
    fileListSection.style.display = 'block';
    batchProcessBtn.disabled = false;
    const tbody = fileTableBody;
    tbody.innerHTML = '';
    for (let i = 0; i < filesQueue.length; i++) {
        const item = filesQueue[i];
        const row = tbody.insertRow();
        row.insertCell(0).textContent = item.originalName;
        let typeShow = item.subType.toUpperCase();
        if (item.typeCategory === 'excel' && item.subType === 'csv') typeShow = 'CSV';
        row.insertCell(1).textContent = typeShow;
        const statusCell = row.insertCell(2);
        statusCell.textContent = item.status || 'Pending';
        statusCell.style.color = '#e67e22';
        const actionCell = row.insertCell(3);
        const removeBtn = document.createElement('button');
        removeBtn.textContent = '✖';
        removeBtn.className = 'remove-file';
        removeBtn.onclick = () => {
            filesQueue.splice(i, 1);
            renderFileList();
            checkNotesVisibility();
        };
        actionCell.appendChild(removeBtn);
    }
    checkNotesVisibility();
}

function checkNotesVisibility() {
    let hasCsv = filesQueue.some(f => f.subType === 'csv');
    let hasMacro = filesQueue.some(f => ['xlsm','pptm','docm'].includes(f.subType));
    csvNote.style.display = hasCsv ? 'block' : 'none';
    macroNote.style.display = hasMacro ? 'block' : 'none';
}

async function processNewFiles(files) {
    for (const file of files) {
        const ext = file.name.split('.').pop().toLowerCase();
        let category = null;
        let subType = ext;
        if (['xlsx','xlsm','csv'].includes(ext)) category = 'excel';
        else if (['pptx','pptm'].includes(ext)) category = 'ppt';
        else if (['docx','docm'].includes(ext)) category = 'word';
        else if (ext === 'pdf') category = 'pdf';
        else {
            setStatus(`Skipped unsupported file: ${file.name}`, true);
            continue;
        }
        try {
            const buffer = await file.arrayBuffer();
            filesQueue.push({
                file: file,
                buffer: buffer,
                typeCategory: category,
                subType: subType,
                originalName: file.name,
                status: 'Pending'
            });
        } catch(e) {
            setStatus(`Failed to read: ${file.name}`, true);
        }
    }
    renderFileList();
    setStatus(`Added ${files.length} file(s). Total: ${filesQueue.length} pending.`, false);
}

async function modifyOpenXmlProperties(buffer, newProps, originalSubType) {
    const zip = await JSZip.loadAsync(buffer);
    let coreFile = zip.file("docProps/core.xml");
    let coreContent = '';
    if (!coreFile) {
        coreContent = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"></cp:coreProperties>`;
    } else {
        coreContent = await coreFile.async("string");
    }
    const parser = new DOMParser();
    const coreDoc = parser.parseFromString(coreContent, "application/xml");
    const coreRoot = coreDoc.documentElement;
    const setNS = (ns, tag, val) => {
        let elem = coreDoc.getElementsByTagNameNS(ns, tag)[0];
        if (!elem) {
            elem = coreDoc.createElementNS(ns, tag);
            coreRoot.appendChild(elem);
        }
        elem.textContent = val || '';
    };
    if (newProps.author !== undefined) setNS('http://purl.org/dc/elements/1.1/', 'creator', newProps.author);
    if (newProps.title !== undefined) setNS('http://purl.org/dc/elements/1.1/', 'title', newProps.title);
    if (newProps.subject !== undefined) setNS('http://purl.org/dc/elements/1.1/', 'subject', newProps.subject);
    if (newProps.keywords !== undefined) setNS('http://schemas.openxmlformats.org/package/2006/metadata/core-properties', 'keywords', newProps.keywords);
    if (newProps.author !== undefined) setNS('http://schemas.openxmlformats.org/package/2006/metadata/core-properties', 'lastModifiedBy', newProps.author);
    const serializer = new XMLSerializer();
    let newCore = serializer.serializeToString(coreDoc);
    if (!newCore.startsWith('<?xml')) newCore = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + newCore;
    zip.file("docProps/core.xml", newCore);

    let appFile = zip.file("docProps/app.xml");
    let appContent = '';
    if (!appFile) {
        appContent = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"></Properties>`;
    } else {
        appContent = await appFile.async("string");
    }
    const appDoc = parser.parseFromString(appContent, "application/xml");
    const appRoot = appDoc.documentElement;
    const setAppTag = (tag, val) => {
        let elem = appDoc.getElementsByTagName(tag)[0];
        if (!elem) {
            elem = appDoc.createElement(tag);
            appRoot.appendChild(elem);
        }
        elem.textContent = val || '';
    };
    if (newProps.author !== undefined) setAppTag('Author', newProps.author);
    if (newProps.title !== undefined) setAppTag('Title', newProps.title);
    if (newProps.subject !== undefined) setAppTag('Subject', newProps.subject);
    if (newProps.keywords !== undefined) setAppTag('Keywords', newProps.keywords);
    if (newProps.author !== undefined) setAppTag('LastAuthor', newProps.author);
    let newApp = serializer.serializeToString(appDoc);
    if (!newApp.startsWith('<?xml')) newApp = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + newApp;
    zip.file("docProps/app.xml", newApp);

    let mime = 'application/octet-stream';
    if (originalSubType === 'xlsx') mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    else if (originalSubType === 'xlsm') mime = 'application/vnd.ms-excel.sheet.macroEnabled.12';
    else if (originalSubType === 'pptx') mime = 'application/vnd.openxmlformats-officedocument.presentationml.presentation';
    else if (originalSubType === 'pptm') mime = 'application/vnd.ms-powerpoint.presentation.macroEnabled.12';
    else if (originalSubType === 'docx') mime = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
    else if (originalSubType === 'docm') mime = 'application/vnd.ms-word.document.macroEnabled.12';
    const blob = await zip.generateAsync({ type: "blob", mimeType: mime });
    return blob;
}

async function convertCsvToXlsxWithProps(csvBuffer, newProps) {
    const data = new Uint8Array(csvBuffer);
    const workbook = XLSX.read(data, { type: 'array' });
    if (!workbook.Props) workbook.Props = {};
    workbook.Props.Author = newProps.author || '';
    workbook.Props.Title = newProps.title || '';
    workbook.Props.Subject = newProps.subject || '';
    workbook.Props.Keywords = newProps.keywords || '';
    const wbout = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
    return new Blob([wbout], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}

function parseKeywordsToArray(keywordsStr) {
    if (!keywordsStr || typeof keywordsStr !== 'string') return [];
    return keywordsStr.split(/[,，\s]+/).filter(k => k.trim().length > 0);
}

async function modifyPdfProperties(buffer, newProps) {
    if (!window.PDFLib) {
        throw new Error('PDF library not loaded, please refresh.');
    }
    const { PDFDocument } = window.PDFLib;
    const pdfDoc = await PDFDocument.load(buffer);
    pdfDoc.setTitle(newProps.title || '');
    pdfDoc.setAuthor(newProps.author || '');
    pdfDoc.setSubject(newProps.subject || '');
    const keywordsArray = parseKeywordsToArray(newProps.keywords);
    pdfDoc.setKeywords(keywordsArray);
    pdfDoc.setCreator('Batch Document Property Editor');
    const modifiedBytes = await pdfDoc.save();
    return new Blob([modifiedBytes], { type: 'application/pdf' });
}

async function processSingleFile(item, commonProps, prefix, suffix) {
    const newProps = {
        author: commonProps.author,
        title: commonProps.title,
        subject: commonProps.subject,
        keywords: commonProps.keywords
    };
    let outputBlob = null;
    let finalExt = '';
    const originalName = item.originalName;
    const dotIndex = originalName.lastIndexOf('.');
    const baseName = dotIndex !== -1 ? originalName.substring(0, dotIndex) : originalName;
    
    if (item.typeCategory === 'excel') {
        if (item.subType === 'csv') {
            outputBlob = await convertCsvToXlsxWithProps(item.buffer, newProps);
            finalExt = 'xlsx';
        } else {
            outputBlob = await modifyOpenXmlProperties(item.buffer, newProps, item.subType);
            finalExt = item.subType === 'xlsm' ? 'xlsm' : 'xlsx';
        }
    } 
    else if (item.typeCategory === 'ppt') {
        outputBlob = await modifyOpenXmlProperties(item.buffer, newProps, item.subType);
        finalExt = item.subType === 'pptm' ? 'pptm' : 'pptx';
    }
    else if (item.typeCategory === 'word') {
        outputBlob = await modifyOpenXmlProperties(item.buffer, newProps, item.subType);
        finalExt = item.subType === 'docm' ? 'docm' : 'docx';
    }
    else if (item.typeCategory === 'pdf') {
        outputBlob = await modifyPdfProperties(item.buffer, newProps);
        finalExt = 'pdf';
    }
    else {
        throw new Error(`Unsupported type: ${item.typeCategory}`);
    }
    let newBase = (prefix || '') + baseName + (suffix || '');
    let newFileName = `${newBase}.${finalExt}`;
    return { blob: outputBlob, newFileName };
}

async function batchProcess() {
    if (filesQueue.length === 0) {
        setStatus('No files to process. Please upload first.', true);
        return;
    }
    if (processing) {
        setStatus('Processing in progress, please wait...', false);
        return;
    }
    const commonProps = {
        author: batchAuthor.value.trim(),
        title: batchTitle.value.trim(),
        subject: batchSubject.value.trim(),
        keywords: batchKeywords.value.trim()
    };
    const prefix = namePrefix.value.trim();
    const suffix = nameSuffix.value.trim();
    
    processing = true;
    batchProcessBtn.disabled = true;
    setStatus(`⏳ Processing ${filesQueue.length} file(s)...`, false);
    
    for (let i=0; i<filesQueue.length; i++) {
        filesQueue[i].status = 'Processing';
    }
    renderFileList();
    
    const zip = new JSZip();
    let successCount = 0;
    let errorCount = 0;
    
    for (let i = 0; i < filesQueue.length; i++) {
        const item = filesQueue[i];
        try {
            const { blob, newFileName } = await processSingleFile(item, commonProps, prefix, suffix);
            zip.file(newFileName, blob);
            item.status = '✅ Success';
            successCount++;
        } catch (err) {
            console.error(`Failed ${item.originalName}:`, err);
            item.status = `❌ Failed: ${err.message}`;
            errorCount++;
        }
        renderFileList();
    }
    
    if (successCount > 0) {
        try {
            const zipBlob = await zip.generateAsync({ type: "blob" });
            const timestamp = new Date().toISOString().slice(0,19).replace(/:/g, '-');
            saveAs(zipBlob, `Batch_Modified_Docs_${timestamp}.zip`);
            setStatus(`✅ Batch completed! Success: ${successCount}, Failed: ${errorCount}. ZIP downloaded.`, false);
        } catch (err) {
            setStatus(`ZIP generation failed: ${err.message}`, true);
        }
    } else {
        setStatus(`Processing failed, no files were generated.`, true);
    }
    
    processing = false;
    batchProcessBtn.disabled = false;
}

function resetAll() {
    filesQueue = [];
    renderFileList();
    batchAuthor.value = '';
    batchTitle.value = '';
    batchSubject.value = '';
    batchKeywords.value = '';
    namePrefix.value = '';
    nameSuffix.value = '';
    setStatus('Reset all files and form.', false);
    processing = false;
    batchProcessBtn.disabled = true;
    fileInput.value = '';
}

function setupDragAndDrop() {
    dropZone.addEventListener('click', () => fileInput.click());
    dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropZone.classList.add('drag-over');
    });
    dropZone.addEventListener('dragleave', () => {
        dropZone.classList.remove('drag-over');
    });
    dropZone.addEventListener('drop', async (e) => {
        e.preventDefault();
        dropZone.classList.remove('drag-over');
        const files = Array.from(e.dataTransfer.files);
        if (files.length) await processNewFiles(files);
    });
    fileInput.addEventListener('change', async (e) => {
        if (e.target.files.length) {
            await processNewFiles(Array.from(e.target.files));
            fileInput.value = '';
        }
    });
}

clearAllFilesBtn.addEventListener('click', () => {
    filesQueue = [];
    renderFileList();
    setStatus('Cleared file list.', false);
});
batchProcessBtn.addEventListener('click', batchProcess);
resetBtn.addEventListener('click', resetAll);

setupDragAndDrop();
renderFileList();
setStatus('✨ Ready. Batch upload and modify metadata (PDF keywords fixed).', false);