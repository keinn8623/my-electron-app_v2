const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
    selectExcelFile: () => ipcRenderer.invoke('select-excel-file'),
    parseExcelFile: (filePath) => ipcRenderer.invoke('parse-excel-file', filePath),
    generatePDF: (studentData) => ipcRenderer.invoke('generate-pdf', studentData)
});
