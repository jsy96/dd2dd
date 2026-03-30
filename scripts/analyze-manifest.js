const XLSX = require('xlsx');
const path = require('path');

async function analyzeManifest() {
  console.log('=== 分析舱单文件格式 ===');
  try {
    const excelPath = path.join(process.cwd(), 'templates', '舱单的格式.xls');
    const workbook = XLSX.readFile(excelPath);
    
    console.log('工作表列表:', workbook.SheetNames);
    
    workbook.SheetNames.forEach(sheetName => {
      console.log(`\n=== 工作表: ${sheetName} ===`);
      const worksheet = workbook.Sheets[sheetName];
      const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
      
      jsonData.forEach((row, index) => {
        if (row && row.length > 0) {
          console.log(`行 ${index + 1}:`, JSON.stringify(row));
        }
      });
    });
  } catch (error) {
    console.error('读取舱单文件失败:', error.message);
  }
}

analyzeManifest().catch(console.error);
