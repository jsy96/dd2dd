const mammoth = require('mammoth');
const XLSX = require('xlsx');
const path = require('path');

async function analyzeTemplates() {
  // 分析 Word 模板
  console.log('=== 分析提单确认件模板 ===');
  try {
    const wordPath = path.join(process.cwd(), 'templates', '提单确认件的格式.docx');
    const result = await mammoth.extractRawText({ path: wordPath });
    console.log('Word 模板内容:');
    console.log(result.value);
    console.log('\n');
  } catch (error) {
    console.error('读取 Word 模板失败:', error.message);
  }

  // 分析 Excel 模板
  console.log('=== 分析装箱单发票模板 ===');
  try {
    const excelPath = path.join(process.cwd(), 'templates', '装箱单发票的格式.xls');
    const workbook = XLSX.readFile(excelPath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    
    console.log('Excel 模板内容:');
    jsonData.forEach((row, index) => {
      console.log(`行 ${index + 1}:`, row);
    });
  } catch (error) {
    console.error('读取 Excel 模板失败:', error.message);
  }
}

analyzeTemplates().catch(console.error);
