import { NextRequest, NextResponse } from 'next/server';
import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import PizZip from 'pizzip';
import Docxtemplater from 'docxtemplater';
import { promises as fs } from 'fs';
import path from 'path';

interface CargoData {
  船名: string;
  航次: string;
  目的港: string;
  提单号: string;
  箱号: string;
  封号: string;
  箱型: string;
  发货人: string;
  发货人名称: string;
  发货人地址: string;
  发货人电话: string;
  收货人: string;
  收货人名称: string;
  收货人地址: string;
  收货人电话: string;
  收货人联系人: string;
  通知人: string;
  通知人名称: string;
  通知人地址: string;
  通知人电话: string;
  英文品名: string;
  件数: string;
  毛重: string;
  体积: string;
  唛头: string;
  包装单位: string;
}

// 解析舱单 Excel 文件
function parseManifestExcel(buffer: Buffer): CargoData {
  const workbook = XLSX.read(buffer, { type: 'buffer' });
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];
  
  // 转换为 JSON 格式（保留空行）
  const jsonData = XLSX.utils.sheet_to_json(worksheet, { 
    header: 1,
    defval: ''
  }) as any[][];

  // 辅助函数：安全获取单元格值
  const getCellValue = (row: number, col: number): string => {
    if (row < 0 || row >= jsonData.length) return '';
    const rowData = jsonData[row];
    if (!rowData || col < 0 || col >= rowData.length) return '';
    return String(rowData[col] || '').trim();
  };

  // 解析数据（根据舱单格式）
  // 行索引从0开始，Excel行号从1开始
  const data: CargoData = {
    // 行4: 船名,航次,目的港
    船名: getCellValue(3, 1),      // B4
    航次: getCellValue(3, 4),      // E4
    目的港: getCellValue(3, 7),    // H4
    
    // 行5: 总提单号
    提单号: getCellValue(4, 1),    // B5
    
    // 行13: 箱号,封号,箱型
    箱号: getCellValue(12, 0),     // A13
    封号: getCellValue(12, 1),     // B13
    箱型: getCellValue(12, 2),     // C13
    
    // 行21: 英文品名,件数,毛重,体积
    英文品名: getCellValue(20, 4), // E21
    件数: getCellValue(20, 6),     // G21
    包装单位: getCellValue(20, 7), // H21
    毛重: getCellValue(20, 8),     // I21
    体积: getCellValue(20, 9),     // J21
    唛头: getCellValue(20, 10),    // K21
    
    // 发货人信息 (行28-31)
    发货人: '',
    发货人名称: getCellValue(27, 2),  // C28 - 名称值在第三列
    发货人地址: getCellValue(28, 2),  // C29 - 地址值在第三列
    发货人电话: getCellValue(30, 2),  // C31 - 电话值在第三列
    
    // 收货人信息 (行35-40)
    收货人: '',
    收货人名称: getCellValue(34, 2),  // C35
    收货人地址: getCellValue(35, 2),  // C36
    收货人电话: getCellValue(37, 2),  // C38
    收货人联系人: getCellValue(39, 2), // C40
    
    // 通知人信息 (行44-47)
    通知人: '',
    通知人名称: getCellValue(43, 2),  // C44
    通知人地址: getCellValue(44, 2),  // C45
    通知人电话: getCellValue(46, 2),  // C47
  };

  // 组合发货人完整信息
  data.发货人 = [
    data.发货人名称,
    data.发货人地址,
    `TEL: ${data.发货人电话}`
  ].filter(Boolean).join('\n');

  // 组合收货人完整信息
  data.收货人 = [
    data.收货人名称,
    data.收货人地址,
    `TEL: ${data.收货人电话}`,
    data.收货人联系人 ? `Contact: ${data.收货人联系人}` : ''
  ].filter(Boolean).join('\n');

  // 组合通知人完整信息
  data.通知人 = [
    data.通知人名称,
    data.通知人地址,
    `TEL: ${data.通知人电话}`
  ].filter(Boolean).join('\n');

  return data;
}

// 生成 Word 文档（使用固定模板）
async function generateWordDocument(data: CargoData): Promise<ArrayBuffer> {
  // 读取转换后的 .docx 模板
  const templatePath = path.join(process.cwd(), 'templates', '提单确认件的格式.docx');
  const templateBuffer = await fs.readFile(templatePath);
  
  const zip = new PizZip(templateBuffer);
  const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
  });

  // 将英文品名拆分为商品列表
  const goodsList = data.英文品名.split(',').map(s => s.trim()).filter(Boolean);
  const goodsData: { [key: string]: string } = {};
  // 模板中有商品1到商品22
  for (let i = 1; i <= 22; i++) {
    goodsData[`商品${i}`] = goodsList[i - 1] || '';
  }

  // 填充数据
  doc.setData({
    船名: data.船名,
    航次: data.航次,
    目的港: data.目的港,
    提单号: data.提单号,
    箱号: data.箱号,
    封号: data.封号,
    箱型: data.箱型,
    发货人: data.发货人,
    收货人: data.收货人,
    通知人: data.通知人,
    件数: data.件数,
    毛重: data.毛重,
    体积: data.体积,
    公司名: data.发货人名称,
    公司地址: data.发货人地址,
    电话: data.发货人电话,
    传真: '',
    电子邮箱: '',
    许可证号: '',
    收货地址: data.收货人地址,
    邮编: '',
    手机号: '',
    电话号码: data.收货人电话,
    // 通知人字段（新模板使用简短字段名）
    姓名: data.通知人名称,
    地址: data.通知人地址,
    // 手机号和电话号码已在上面的收货人字段中定义，这里会覆盖
    ...goodsData,
  });

  doc.render();
  const buffer = doc.getZip().generate({ type: 'nodebuffer' });
  return new Uint8Array(buffer).buffer as ArrayBuffer;
}

// 生成 Excel 文档（使用固定模板）
async function generateExcelDocument(data: CargoData): Promise<ArrayBuffer> {
  // 读取模板（使用转换后的 .xlsx 格式）
  const templatePath = path.join(process.cwd(), 'templates', '装箱单发票的格式.xlsx');
  const templateBuffer = await fs.readFile(templatePath);
  
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(templateBuffer as any);
  const worksheet = workbook.worksheets[0];

  if (!worksheet) {
    throw new Error('无法加载 Excel 模板');
  }

  // 生成当天日期，格式如 "FEB. 04. 2026"
  const today = new Date();
  const months = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'];
  const formattedDate = `${months[today.getMonth()]}. ${String(today.getDate()).padStart(2, '0')}. ${today.getFullYear()}`;

  // 辅助函数：获取单元格完整文本（支持富文本）
  const getCellText = (cell: any): string => {
    if (!cell.value) return '';
    if (typeof cell.value === 'string') return cell.value;
    if (cell.value.richText) {
      return cell.value.richText.map((rt: any) => rt.text || '').join('');
    }
    return '';
  };

  // 辅助函数：替换单元格中的占位符
  const replacePlaceholder = (cell: any, placeholder: string, replacement: string): boolean => {
    const text = getCellText(cell);
    if (text.includes(placeholder)) {
      // 获取原有格式（取第一个富文本片段的字体样式）
      let font: any = {};
      if (cell.value?.richText && cell.value.richText.length > 0) {
        font = cell.value.richText[0].font || {};
      }
      
      // 创建新的富文本，保持原有字体样式
      cell.value = {
        richText: [{
          font: font,
          text: replacement
        }]
      };
      return true;
    }
    return false;
  };

  // 填充发票日期 - 替换 {发票日期} 占位符
  worksheet.eachRow((row, rowNumber) => {
    row.eachCell((cell) => {
      replacePlaceholder(cell, '{发票日期}', formattedDate);
    });
  });

  // 将英文品名按逗号分隔成商品列表
  const goodsList = data.英文品名.split(',').map(s => s.trim()).filter(Boolean);

  // 填充商品列表 - 替换 {商品1} 到 {商品22} 占位符
  for (let i = 0; i < 22; i++) {
    const rowNum = 12 + i;
    const row = worksheet.getRow(rowNum);
    const cell = row.getCell(5); // E列
    const placeholder = `{商品${i + 1}}`;
    
    if (i < goodsList.length) {
      replacePlaceholder(cell, placeholder, goodsList[i]);
    } else {
      // 商品数量不足，清空该占位符
      replacePlaceholder(cell, placeholder, '');
    }
  }

  const buffer = await workbook.xlsx.writeBuffer();
  return buffer as ArrayBuffer;
}

// 解析舱单接口
export async function POST(request: NextRequest) {
  try {
    const formData = await request.formData();
    const manifestFile = formData.get('manifest') as File;

    if (!manifestFile) {
      return NextResponse.json(
        { success: false, message: '请上传舱单文件' },
        { status: 400 }
      );
    }

    // 读取舱单文件内容
    const manifestBuffer = Buffer.from(await manifestFile.arrayBuffer());

    // 解析舱单数据
    const cargoData = parseManifestExcel(manifestBuffer);

    // 生成文件
    const wordBuffer = await generateWordDocument(cargoData);
    const excelBuffer = await generateExcelDocument(cargoData);

    // 保存文件到临时目录
    const timestamp = Date.now();
    const wordFileName = `提单确认件_${timestamp}.doc`;
    const excelFileName = `装箱单发票_${timestamp}.xls`;
    
    const tempDir = '/tmp';
    const wordFilePath = path.join(tempDir, wordFileName);
    const excelFilePath = path.join(tempDir, excelFileName);

    await fs.writeFile(wordFilePath, Buffer.from(wordBuffer));
    await fs.writeFile(excelFilePath, Buffer.from(excelBuffer));

    // 返回结果
    return NextResponse.json({
      success: true,
      message: '文件处理成功',
      data: cargoData,
      wordFileUrl: `/api/download?file=${wordFileName}`,
      excelFileUrl: `/api/download?file=${excelFileName}`,
    });
  } catch (error) {
    console.error('处理文件失败:', error);
    return NextResponse.json(
      { success: false, message: '处理文件失败，请检查文件格式' },
      { status: 500 }
    );
  }
}
