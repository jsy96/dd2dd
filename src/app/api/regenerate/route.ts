import { NextRequest, NextResponse } from 'next/server';
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
      replacePlaceholder(cell, placeholder, '');
    }
  }

  const buffer = await workbook.xlsx.writeBuffer();
  return buffer as ArrayBuffer;
}

// 重新生成文件接口
export async function POST(request: NextRequest) {
  try {
    const body = await request.json();
    const cargoData = body.data as CargoData;

    if (!cargoData) {
      return NextResponse.json(
        { success: false, message: '缺少数据' },
        { status: 400 }
      );
    }

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
      message: '文件重新生成成功',
      wordFileUrl: `/api/download?file=${wordFileName}`,
      excelFileUrl: `/api/download?file=${excelFileName}`,
    });
  } catch (error) {
    console.error('重新生成文件失败:', error);
    return NextResponse.json(
      { success: false, message: '重新生成文件失败' },
      { status: 500 }
    );
  }
}
