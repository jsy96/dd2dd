const ExcelJS = require('exceljs');
const path = require('path');

async function verifyGeneratedExcel() {
  console.log('=== 验证生成的装箱单发票 ===\n');
  try {
    // 读取最新生成的文件
    const filePath = '/tmp/装箱单发票_1774879452181.xls';
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    
    const worksheet = workbook.worksheets[0];
    
    // 检查日期
    console.log('--- 发票日期 ---');
    const dateRow = worksheet.getRow(6);
    console.log('行6列7 (G6):', dateRow.getCell(7).value);
    console.log('行6列8 (H6):', dateRow.getCell(8).value);
    
    // 检查商品列表
    console.log('\n--- 商品列表 (前15个) ---');
    for (let i = 0; i < 15; i++) {
      const rowNum = 12 + i;
      const row = worksheet.getRow(rowNum);
      const cell = row.getCell(5);
      console.log(`行${rowNum}列E (E${rowNum}):`, cell.value);
    }
    
  } catch (error) {
    console.error('读取文件失败:', error.message);
  }
}

verifyGeneratedExcel().catch(console.error);
