/**
 * Excel生成サービス
 * 要件定義をExcelファイルとして出力
 */

import ExcelJS from 'exceljs';
import { AIExtractedRequirements } from './aiExtractor.js';

export interface ExcelGeneratorOptions {
  includeFormulas?: boolean;
  includeTables?: boolean;
  includeNotes?: boolean;
}

/**
 * スタイル定義
 */
const STYLES = {
  header: {
    font: { bold: true, size: 14, color: { argb: 'FFFFFFFF' } },
    fill: { type: 'pattern' as const, pattern: 'solid' as const, fgColor: { argb: 'FF4472C4' } },
    alignment: { horizontal: 'center' as const, vertical: 'middle' as const },
    border: {
      top: { style: 'thin' as const },
      bottom: { style: 'thin' as const },
      left: { style: 'thin' as const },
      right: { style: 'thin' as const },
    },
  },
  subHeader: {
    font: { bold: true, size: 12 },
    fill: { type: 'pattern' as const, pattern: 'solid' as const, fgColor: { argb: 'FFD9E2F3' } },
    alignment: { horizontal: 'left' as const, vertical: 'middle' as const },
    border: {
      top: { style: 'thin' as const },
      bottom: { style: 'thin' as const },
      left: { style: 'thin' as const },
      right: { style: 'thin' as const },
    },
  },
  cell: {
    alignment: { horizontal: 'left' as const, vertical: 'top' as const, wrapText: true },
    border: {
      top: { style: 'thin' as const },
      bottom: { style: 'thin' as const },
      left: { style: 'thin' as const },
      right: { style: 'thin' as const },
    },
  },
  categoryHeader: {
    font: { bold: true, size: 11 },
    fill: { type: 'pattern' as const, pattern: 'solid' as const, fgColor: { argb: 'FFE2EFDA' } },
    border: {
      top: { style: 'thin' as const },
      bottom: { style: 'thin' as const },
      left: { style: 'thin' as const },
      right: { style: 'thin' as const },
    },
  },
};

/**
 * 要件定義をExcelに変換
 */
export async function generateRequirementsExcel(
  requirements: AIExtractedRequirements,
  options: ExcelGeneratorOptions = {}
): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'RD-System';
  workbook.created = new Date();

  // 1. 概要シート
  createSummarySheet(workbook, requirements);

  // 2. 入力項目シート
  createInputItemsSheet(workbook, requirements);

  // 3. 算定方法シート
  if (options.includeFormulas !== false) {
    createCalculationSheet(workbook, requirements);
  }

  // 4. 費用項目シート
  createFeeStructureSheet(workbook, requirements);

  // 5. 別表シート
  if (options.includeTables !== false && requirements.tables.length > 0) {
    createTablesSheet(workbook, requirements);
  }

  // 6. 補足事項シート
  if (options.includeNotes !== false && requirements.additionalNotes.length > 0) {
    createNotesSheet(workbook, requirements);
  }

  // Excelファイルをバッファとして出力
  const buffer = await workbook.xlsx.writeBuffer();
  return Buffer.from(buffer);
}

/**
 * 概要シートを作成
 */
function createSummarySheet(
  workbook: ExcelJS.Workbook,
  requirements: AIExtractedRequirements
): void {
  const sheet = workbook.addWorksheet('概要', {
    properties: { tabColor: { argb: 'FF4472C4' } },
  });

  // 列幅設定
  sheet.columns = [
    { width: 20 },
    { width: 80 },
  ];

  // タイトル
  const titleRow = sheet.addRow(['要件定義書']);
  titleRow.font = { bold: true, size: 18 };
  sheet.mergeCells('A1:B1');
  titleRow.height = 30;

  sheet.addRow([]);

  // 基本情報
  const infoData = [
    ['ドキュメント名', requirements.documentTitle],
    ['種別', requirements.documentType],
    ['作成日', new Date().toLocaleDateString('ja-JP')],
    ['概要', requirements.summary],
  ];

  infoData.forEach((row) => {
    const r = sheet.addRow(row);
    r.getCell(1).font = { bold: true };
    r.getCell(1).fill = STYLES.subHeader.fill;
    r.getCell(2).alignment = { wrapText: true };
    r.height = row[0] === '概要' ? 60 : 25;
  });

  // 統計情報
  sheet.addRow([]);
  const statsHeader = sheet.addRow(['統計情報']);
  statsHeader.font = { bold: true, size: 14 };

  const statsData = [
    ['入力項目数', requirements.inputItems.reduce((sum, cat) => sum + cat.items.length, 0).toString()],
    ['計算ルール数', requirements.calculationRules.length.toString()],
    ['費用項目数', requirements.feeStructure.reduce((sum, cat) => sum + cat.items.length, 0).toString()],
    ['別表数', requirements.tables.length.toString()],
  ];

  statsData.forEach((row) => {
    const r = sheet.addRow(row);
    r.getCell(1).font = { bold: true };
  });
}

/**
 * 入力項目シートを作成
 */
function createInputItemsSheet(
  workbook: ExcelJS.Workbook,
  requirements: AIExtractedRequirements
): void {
  const sheet = workbook.addWorksheet('入力項目', {
    properties: { tabColor: { argb: 'FF70AD47' } },
  });

  // 列幅設定
  sheet.columns = [
    { width: 5 },   // No
    { width: 15 },  // カテゴリ
    { width: 25 },  // 項目名
    { width: 40 },  // 説明
    { width: 12 },  // データ型
    { width: 8 },   // 必須
    { width: 30 },  // バリデーション
  ];

  // ヘッダー
  const headers = ['No', 'カテゴリ', '項目名', '説明', 'データ型', '必須', 'バリデーション'];
  const headerRow = sheet.addRow(headers);
  headerRow.eachCell((cell) => {
    Object.assign(cell, STYLES.header);
  });
  headerRow.height = 25;

  // データ
  let no = 1;
  requirements.inputItems.forEach((category) => {
    category.items.forEach((item) => {
      const row = sheet.addRow([
        no++,
        category.category,
        item.name,
        item.description,
        item.dataType,
        item.required ? '○' : '',
        item.validationRules?.join('\n') || '',
      ]);
      row.eachCell((cell) => {
        Object.assign(cell, STYLES.cell);
      });
      row.height = Math.max(20, (item.validationRules?.length || 1) * 15);
    });
  });

  // フィルター設定
  sheet.autoFilter = {
    from: 'A1',
    to: `G${sheet.rowCount}`,
  };
}

/**
 * 算定方法シートを作成
 */
function createCalculationSheet(
  workbook: ExcelJS.Workbook,
  requirements: AIExtractedRequirements
): void {
  const sheet = workbook.addWorksheet('算定方法', {
    properties: { tabColor: { argb: 'FFFFC000' } },
  });

  // 列幅設定
  sheet.columns = [
    { width: 5 },   // No
    { width: 25 },  // 計算名
    { width: 40 },  // 説明
    { width: 35 },  // 計算式
    { width: 30 },  // 条件
    { width: 30 },  // 例
  ];

  // ヘッダー
  const headers = ['No', '計算名', '説明', '計算式', '適用条件', '計算例'];
  const headerRow = sheet.addRow(headers);
  headerRow.eachCell((cell) => {
    Object.assign(cell, STYLES.header);
  });
  headerRow.height = 25;

  // データ
  requirements.calculationRules.forEach((rule, index) => {
    const row = sheet.addRow([
      index + 1,
      rule.name,
      rule.description,
      rule.formula,
      rule.conditions?.join('\n') || '',
      rule.examples?.join('\n') || '',
    ]);
    row.eachCell((cell) => {
      Object.assign(cell, STYLES.cell);
    });
    const maxLines = Math.max(
      rule.conditions?.length || 1,
      rule.examples?.length || 1
    );
    row.height = Math.max(25, maxLines * 15);
  });
}

/**
 * 費用項目シートを作成
 */
function createFeeStructureSheet(
  workbook: ExcelJS.Workbook,
  requirements: AIExtractedRequirements
): void {
  const sheet = workbook.addWorksheet('費用項目', {
    properties: { tabColor: { argb: 'FF5B9BD5' } },
  });

  // 列幅設定
  sheet.columns = [
    { width: 5 },   // No
    { width: 15 },  // カテゴリ
    { width: 25 },  // 費用名
    { width: 40 },  // 説明
    { width: 15 },  // 金額
    { width: 10 },  // 単位
    { width: 30 },  // 条件
  ];

  // ヘッダー
  const headers = ['No', 'カテゴリ', '費用名', '説明', '金額', '単位', '適用条件'];
  const headerRow = sheet.addRow(headers);
  headerRow.eachCell((cell) => {
    Object.assign(cell, STYLES.header);
  });
  headerRow.height = 25;

  // データ
  let no = 1;
  requirements.feeStructure.forEach((category) => {
    category.items.forEach((item) => {
      const row = sheet.addRow([
        no++,
        category.category,
        item.name,
        item.description,
        item.amount || '',
        item.unit || '',
        item.conditions?.join('\n') || '',
      ]);
      row.eachCell((cell) => {
        Object.assign(cell, STYLES.cell);
      });
      row.height = Math.max(20, (item.conditions?.length || 1) * 15);
    });
  });

  // フィルター設定
  sheet.autoFilter = {
    from: 'A1',
    to: `G${sheet.rowCount}`,
  };
}

/**
 * 別表シートを作成
 */
function createTablesSheet(
  workbook: ExcelJS.Workbook,
  requirements: AIExtractedRequirements
): void {
  requirements.tables.forEach((table, index) => {
    const sheetName = `別表${index + 1}`;
    const sheet = workbook.addWorksheet(sheetName, {
      properties: { tabColor: { argb: 'FF7030A0' } },
    });

    // タイトル
    const titleRow = sheet.addRow([table.title]);
    titleRow.font = { bold: true, size: 14 };
    sheet.mergeCells(`A1:${String.fromCharCode(65 + table.headers.length - 1)}1`);

    // 説明
    if (table.description) {
      const descRow = sheet.addRow([table.description]);
      descRow.font = { italic: true };
      sheet.mergeCells(`A2:${String.fromCharCode(65 + table.headers.length - 1)}2`);
    }

    sheet.addRow([]);

    // ヘッダー
    if (table.headers.length > 0) {
      const headerRow = sheet.addRow(table.headers);
      headerRow.eachCell((cell) => {
        Object.assign(cell, STYLES.header);
      });
    }

    // データ行
    table.rows.forEach((rowData) => {
      const row = sheet.addRow(rowData);
      row.eachCell((cell) => {
        Object.assign(cell, STYLES.cell);
      });
    });

    // 列幅の自動調整
    sheet.columns.forEach((column, i) => {
      let maxLength = table.headers[i]?.length || 10;
      table.rows.forEach((row) => {
        const cellLength = row[i]?.toString().length || 0;
        maxLength = Math.max(maxLength, cellLength);
      });
      column.width = Math.min(Math.max(maxLength + 2, 10), 50);
    });
  });
}

/**
 * 補足事項シートを作成
 */
function createNotesSheet(
  workbook: ExcelJS.Workbook,
  requirements: AIExtractedRequirements
): void {
  const sheet = workbook.addWorksheet('補足事項', {
    properties: { tabColor: { argb: 'FF808080' } },
  });

  // 列幅設定
  sheet.columns = [
    { width: 5 },
    { width: 100 },
  ];

  // ヘッダー
  const headerRow = sheet.addRow(['No', '補足事項']);
  headerRow.eachCell((cell) => {
    Object.assign(cell, STYLES.header);
  });

  // データ
  requirements.additionalNotes.forEach((note, index) => {
    const row = sheet.addRow([index + 1, note]);
    row.eachCell((cell) => {
      Object.assign(cell, STYLES.cell);
    });
    row.height = Math.max(25, Math.ceil(note.length / 50) * 15);
  });
}
