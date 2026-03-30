const ExcelJS = require('exceljs');
const path = require('path');

async function createSampleManifest() {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('舱单');

  // 设置列宽
  worksheet.columns = [
    { header: '项目', key: 'item', width: 20 },
    { header: '内容', key: 'content', width: 40 },
    { header: '备注', key: 'note', width: 30 },
  ];

  // 添加基本信息
  worksheet.addRow(['船名', 'COSCO SHIPPING', '']);
  worksheet.addRow(['航次', 'V.2024E001', '']);
  worksheet.addRow(['目的港', 'LOS ANGELES, USA', '']);
  worksheet.addRow(['提单号', 'COSU1234567890', '']);
  worksheet.addRow(['箱号', 'CAXU1234567', '']);
  worksheet.addRow(['封号', 'SEAL123456', '']);
  worksheet.addRow(['箱型', '40GP', '']);
  worksheet.addRow(['发货人', 'SHANGHAI TRADING CO., LTD.\nNO.123, NANJING ROAD\nSHANGHAI, CHINA', '']);
  worksheet.addRow(['收货人', 'AMERICAN IMPORT CO.\n456 MAIN STREET\nLOS ANGELES, CA 90001\nUSA', '']);
  worksheet.addRow(['通知人', 'SAME AS CONSIGNEE', '']);
  worksheet.addRow(['日期', 'FEB. 04. 2026', '']);
  worksheet.addRow(['', '', '']);
  
  // 添加商品列表标题
  worksheet.addRow(['商品列表', '', '']);
  worksheet.addRow(['商品名称', '件数', '毛重(KG)', '体积(CBM)']);
  
  // 添加商品数据
  worksheet.addRow(['电子产品', '100', '1500', '25.5']);
  worksheet.addRow(['服装', '200', '800', '18.2']);
  worksheet.addRow(['家居用品', '150', '1200', '32.8']);
  
  // 添加合计
  worksheet.addRow(['合计', '450', '3500', '76.5']);

  // 保存文件
  const filePath = path.join(process.cwd(), 'templates', '舱单的格式.xls');
  await workbook.xlsx.writeFile(filePath);
  console.log('示例舱单文件已创建:', filePath);
}

createSampleManifest().catch(console.error);
