
const spreadSheet = document.getElementById('spreadsheet');
const openBtn = document.getElementById('openBtn');
const savecloseBtn = document.getElementById('savecloseBtn');
const closeBtn = document.getElementById('closeBtn');
var sheet;
function loadExcel() {
    
    /* if(!isFileIExcel(file)) {
        alertError('Please select an image');
        return;    
    } */
    //Get original dimnsions

    showSheet(document.getElementById("openBtn").files[0].path)

}




function showSheet(file){
    /* const workbook = XLSX.readFile(file);
    const worksheet = workbook.Sheets[workbook.SheetNames[0]]; */
    const workSheetsFromFile = xlsx.parse(file);
    const posts = [];
    var data = workSheetsFromFile[0].data
    let i = 1;
    data[0].push("Sum");
    while (i< data.length) {
        let x = i + 1;
        data[i][6] = '=IFERROR(AVERAGE(C'+x+':F'+x+'),"" )';
        data[i].push('=IFERROR(SUM(C'+x+':F'+x+'),"" )');
        i = i + 1;
    }
    console.log(data)
    sheet = jspreadsheet.build(document.getElementById('spreadsheet'), {
        data: data
    })
        }


function saveAndClose() {
    data = sheet.getData();
    console.log(data);
    var buffer = xlsx.build([{name: 'Sheet', data: data}]);
    var downloadDir = path.resolve(os.homedir(), '../', os.userInfo().username, 'Documents', 'KalkuDocs');
    
    var dir = downloadDir+"\\" +'exported.xlsx'

    if (!fs.existsSync(downloadDir)) {
        fs.mkdirSync(downloadDir);
    }
    fs.writeFileSync(dir, buffer, 'binary');
    clearInputFile(openBtn);
    sheet.destroy(spreadSheet, true)
    return dir
}

function clearInputFile(f){
    if(f.value){
        try{
            f.value = ''; //for IE11, latest Chrome/Firefox/Opera...
        }catch(err){ }
        if(f.value){ //for IE5 ~ IE10
            var form = document.createElement('form'),
                parentNode = f.parentNode, ref = f.nextSibling;
            form.appendChild(f);
            form.reset();
            parentNode.insertBefore(f,ref);
        }
    }
}

openBtn.addEventListener('change', loadExcel);
savecloseBtn.onclick = saveAndClose;