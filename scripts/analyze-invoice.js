const ExcelJS = require('exceljs');
const path = require('path');

async function analyzeInvoiceTemplate() {
  console.log('=== 分析装箱单发票模板 ===\n');
  try {
    const templatePath = path.join(process.cwd(), 'templates', '装箱单发票的格式.xlsx');
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(templatePath);
    
    const worksheet = workbook.worksheets[0];
    console.log('工作表名称:', worksheet.name);
    console.log('\n--- 逐行分析 ---\n');
    
    worksheet.eachRow((row, rowNumber) => {
      var values = [];
      row.eachCell((cell, colNumber) => {
        let value = '';
        if (cell.value !== null && cell.value !== undefined) {
          if (typeof cell.value === 'object') {
            value = JSON.stringify(cell.value);
          } else {
            value = String(cell.value);
          }
        }
        if (value) {
          values.push(`[${colNumber}]${value}`);
        }
      });
      if (values.length > 0) {
        console.log(`行${rowNumber}: ${values.join(' | ')}`);
      }
    });
    
    console.log('\n--- 查找发票日期位置 ---');
    let foundDate = false;
    worksheet.eachRow((row, rowNumber) => {
      row.eachCell((cell, colNumber) => {
        if (cell.value && typeof cell.value === 'string' && cell.value.includes('发票日期')) {
          console.log(`找到"发票日期"在 行${rowNumber} 列${colNumber}: ${cell.value}`);
          // 查看下一列的值
          const nextCell = row.getCell(colNumber + 1);
          console.log(`  下一列(列${colNumber + 1})值: ${nextCell.value}`);
          foundDate = true;
        }
        if (cell.value && typeof cell.value === 'string' && cell.value.includes('{发票日期}')) {
          console.log(`找到"{发票日期}"在 行${rowNumber} 列${colNumber}`);
          foundDate = true;
        }
      });
    });
    
    console.log('\n--- 查找商品位置 ---');
    for (let i = 1; i <= 35; i++) {
      const row = worksheet.getRow(i);
      row.eachCell((cell, colNumber) => {
        if (cell.value && typeof cell.value === 'string' && cell.value.includes('商品')) {
          console.log(`找到"商品"在 行${i} 列${colNumber}: ${cell.value}`);
        }
      });
    }
    
  } catch (error) {
    console.error('读取模板失败:', error.message);
  }
}

analyzeInvoiceTemplate().catch(console.error);
