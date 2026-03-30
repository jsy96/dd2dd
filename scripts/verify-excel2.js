const ExcelJS = require('exceljs');

async function verifyExcel() {
  console.log('=== 验证生成的装箱单发票 ===\n');
  try {
    const filePath = '/tmp/装箱单发票_1774879674371.xls';
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    
    const worksheet = workbook.worksheets[0];
    
    // 检查标题（保持不变）
    console.log('--- 标题（应保持不变）---');
    console.log('行1:', worksheet.getRow(1).getCell(1).value);
    console.log('行2:', worksheet.getRow(2).getCell(1).value?.toString().substring(0, 50));
    
    // 检查发票编号（保持不变）
    console.log('\n--- 发票编号（应保持不变）---');
    console.log('行4列7:', worksheet.getRow(4).getCell(7).value);
    
    // 检查日期（已替换）
    console.log('\n--- 发票日期（已替换）---');
    const dateRow = worksheet.getRow(6);
    console.log('行6列7:', dateRow.getCell(7).value);
    
    // 检查表头（保持不变）
    console.log('\n--- 表头（应保持不变）---');
    const headerRow = worksheet.getRow(10);
    console.log('行10列1:', headerRow.getCell(1).value?.richText?.map(r => r.text).join('') || headerRow.getCell(1).value);
    
    // 检查商品列表
    console.log('\n--- 商品列表 ---');
    for (let i = 0; i < 15; i++) {
      const rowNum = 12 + i;
      const row = worksheet.getRow(rowNum);
      const cell = row.getCell(5);
      let value = '';
      if (cell.value?.richText) {
        value = cell.value.richText.map(r => r.text).join('');
      } else {
        value = cell.value || '(空)';
      }
      console.log(`行${rowNum}列E:`, value);
    }
    
    // 检查数量和单价列是否保持不变
    console.log('\n--- 其他列（应保持不变）---');
    const row12 = worksheet.getRow(12);
    console.log('行12列C (数量):', row12.getCell(3).value);
    console.log('行12列D (单位):', row12.getCell(4).value);
    console.log('行12列F (单价):', row12.getCell(6).value);
    
  } catch (error) {
    console.error('读取失败:', error.message);
  }
}

verifyExcel().catch(console.error);
