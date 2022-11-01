// See the Electron documentation for details on how to use preload scripts:
// https://www.electronjs.org/docs/latest/tutorial/process-model#preload-scripts
const { contextBridge } = require('electron');
const jspreadsheet = require('jspreadsheet-ce');
 const XLSX = require('xlsx');
 const os = require('os');
 const path = require('path');
 const xlsx = require('node-xlsx').default;
 const fs = require('fs');


contextBridge.exposeInMainWorld('versions', {
  node: () => process.versions.node,
  chrome: () => process.versions.chrome,
  electron: () => process.versions.electron,
  slice: (n) => process.argv.slice(n),
  // we can also expose variables, not just functions
})

contextBridge.exposeInMainWorld('jspreadsheet', {
    build : (spreadsheet, ...args) => jspreadsheet(spreadsheet, ...args),
    fromSpreadsheet: (xlxx, func) => jspreadsheet.fromSpreadsheet(xlxx, func),
    destroy: (table, t) => jspreadsheet.destroy(table, t),
})

contextBridge.exposeInMainWorld('os', {
    homedir: () => os.homedir(),
    userInfo: () => os.userInfo(),
    
  });

 contextBridge.exposeInMainWorld('XLSX', {
    readFile : (filePath) => XLSX.readFile(filePath),
})
contextBridge.exposeInMainWorld('path', {
    join: (...args) => path.join(...args),
    resolve: (...args) => path.resolve(...args),
  });

contextBridge.exposeInMainWorld('xlsx', {
    parse: (...args) => xlsx.parse(...args),
    build: (...args) => xlsx.build(...args)
  });
 contextBridge.exposeInMainWorld('fs', {
    writeFileSync : (...args) => fs.writeFileSync(...args),
    existsSync: (dest) => fs.existsSync(dest),
    mkdirSync: (dest) => fs.mkdirSync(dest), 
})
  