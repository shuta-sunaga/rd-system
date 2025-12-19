/**
 * PDF解析サービス
 * PDFファイルからテキストを抽出し、要件定義に必要な情報を解析する
 */

// pdf-parse v1.x uses default export
import pdf from 'pdf-parse';
import { readFile } from 'fs/promises';

export interface ParsedPDFContent {
  rawText: string;
  pages: string[];
  metadata: {
    title?: string;
    author?: string;
    pageCount: number;
  };
}

export interface RequirementItem {
  category: string;
  name: string;
  description: string;
  inputType?: string;
  required?: boolean;
}

export interface CalculationFormula {
  name: string;
  description: string;
  formula: string;
  variables: string[];
  conditions?: string[];
}

export interface FeeItem {
  name: string;
  description: string;
  amount?: string;
  conditions?: string[];
}

export interface ExtractedRequirements {
  documentTitle: string;
  inputItems: RequirementItem[];
  calculationItems: RequirementItem[];
  formulas: CalculationFormula[];
  fees: FeeItem[];
  otherRequirements: string[];
  tables: TableData[];
}

export interface TableData {
  title: string;
  headers: string[];
  rows: string[][];
}

/**
 * PDFファイルを解析してテキストを抽出
 */
export async function parsePDF(filePath: string): Promise<ParsedPDFContent> {
  const dataBuffer = await readFile(filePath);
  const data = await pdf(dataBuffer);

  // ページごとにテキストを分割
  const pages = data.text.split(/\n\s*-\s*\d+\s*-\s*\n/).filter((p: string) => p.trim());

  return {
    rawText: data.text,
    pages,
    metadata: {
      title: data.info?.Title,
      author: data.info?.Author,
      pageCount: data.numpages,
    },
  };
}

/**
 * PDFバッファから解析
 */
export async function parsePDFBuffer(buffer: Buffer): Promise<ParsedPDFContent> {
  const data = await pdf(buffer);

  // ページごとにテキストを分割
  const pages = data.text.split(/\n\s*-\s*\d+\s*-\s*\n/).filter((p: string) => p.trim());

  return {
    rawText: data.text,
    pages,
    metadata: {
      title: data.info?.Title,
      author: data.info?.Author,
      pageCount: data.numpages,
    },
  };
}

/**
 * テキストから入力項目を抽出
 */
export function extractInputItems(text: string): RequirementItem[] {
  const items: RequirementItem[] = [];

  // 定義セクションを探す
  const definitionPatterns = [
    /第\d+条[（(]定義[）)]\s*([\s\S]*?)(?=第\d+条|$)/g,
    /（\d+）([^（）\n]+)/g,
  ];

  // 入力項目を示すキーワード
  const inputKeywords = [
    '入力', '記入', '申請', '届出', '提出', '選定', '選択',
  ];

  // 「〜とは」で始まる定義を抽出
  const definitionMatches = text.matchAll(/[「『]?([^「『」』\n]+)[」』]?とは[、,]([^。]+)/g);
  for (const match of definitionMatches) {
    items.push({
      category: '定義',
      name: match[1].trim(),
      description: match[2].trim(),
      inputType: 'text',
    });
  }

  return items;
}

/**
 * テキストから計算式を抽出
 */
export function extractFormulas(text: string): CalculationFormula[] {
  const formulas: CalculationFormula[] = [];

  // 計算に関するパターン
  const calculationPatterns = [
    /([^。\n]+)を乗じた金額/g,
    /([^。\n]+)の([0-9０-９]+[割分厘%％])/g,
    /([^。\n]+)を上限として/g,
  ];

  // 「〜は、〜とする」形式の計算式
  const formulaMatches = text.matchAll(/([^。\n]{5,50})は[、,]([^。]+)とする/g);
  for (const match of formulaMatches) {
    const name = match[1].trim();
    const desc = match[2].trim();

    // 数値や計算が含まれているか確認
    if (/[0-9０-９]|割|分|倍|円|率|額/.test(desc)) {
      formulas.push({
        name,
        description: desc,
        formula: convertToFormula(desc),
        variables: extractVariables(desc),
      });
    }
  }

  return formulas;
}

/**
 * 日本語の説明を計算式に変換
 */
function convertToFormula(description: string): string {
  let formula = description;

  // 割合の変換
  formula = formula.replace(/([0-9０-９]+)割/g, (_, n) => `${toHalfWidth(n)} * 0.1`);
  formula = formula.replace(/([0-9０-９]+)分/g, (_, n) => `${toHalfWidth(n)} * 0.01`);
  formula = formula.replace(/([0-9０-９]+)[%％]/g, (_, n) => `${toHalfWidth(n)} * 0.01`);
  formula = formula.replace(/([0-9０-９]+)倍/g, (_, n) => `* ${toHalfWidth(n)}`);

  // 金額の変換
  formula = formula.replace(/([0-9０-９,，]+)円/g, (_, n) => toHalfWidth(n.replace(/[,，]/g, '')));

  return formula;
}

/**
 * 全角数字を半角に変換
 */
function toHalfWidth(str: string): string {
  return str.replace(/[０-９]/g, (s) =>
    String.fromCharCode(s.charCodeAt(0) - 0xFEE0)
  );
}

/**
 * 説明文から変数を抽出
 */
function extractVariables(description: string): string[] {
  const variables: string[] = [];

  // 「〜の」で終わる名詞を変数として抽出
  const varPatterns = [
    /([^、。\n]+(?:料|額|費|金|率))/g,
  ];

  for (const pattern of varPatterns) {
    const matches = description.matchAll(pattern);
    for (const match of matches) {
      if (!variables.includes(match[1])) {
        variables.push(match[1]);
      }
    }
  }

  return variables;
}

/**
 * テーブルデータを抽出
 */
export function extractTables(text: string): TableData[] {
  const tables: TableData[] = [];

  // 「別表」パターンを検索
  const tablePatterns = [
    /[「『]?別表\s*(\d+)[」』]?\s*([^\n]+)\n([\s\S]*?)(?=[「『]?別表|付則|以\s*上|$)/g,
  ];

  for (const pattern of tablePatterns) {
    const matches = text.matchAll(pattern);
    for (const match of matches) {
      const title = `別表${match[1]} ${match[2].trim()}`;
      const content = match[3];

      // テーブル内容を解析
      const lines = content.split('\n').filter(l => l.trim());
      if (lines.length > 0) {
        tables.push({
          title,
          headers: [],
          rows: lines.map(l => l.split(/\s{2,}|\t/).filter(c => c.trim())),
        });
      }
    }
  }

  return tables;
}

/**
 * 費用項目を抽出
 */
export function extractFees(text: string): FeeItem[] {
  const fees: FeeItem[] = [];

  // 費用に関するパターン
  const feePatterns = [
    /([^。\n]+(?:費|料|金|手当))[はを][、,]([^。]+)/g,
  ];

  for (const pattern of feePatterns) {
    const matches = text.matchAll(pattern);
    for (const match of matches) {
      const name = match[1].trim();
      const desc = match[2].trim();

      // 金額を抽出
      const amountMatch = desc.match(/([0-9０-９,，]+)\s*円/);

      fees.push({
        name,
        description: desc,
        amount: amountMatch ? toHalfWidth(amountMatch[1].replace(/[,，]/g, '')) : undefined,
      });
    }
  }

  return fees;
}

/**
 * PDFから要件定義を抽出（メイン関数）
 */
export async function extractRequirementsFromPDF(
  buffer: Buffer
): Promise<ExtractedRequirements> {
  const parsed = await parsePDFBuffer(buffer);
  const text = parsed.rawText;

  // ドキュメントタイトルを抽出
  const titleMatch = text.match(/^([^\n]+)/);
  const documentTitle = titleMatch ? titleMatch[1].trim() : '要件定義書';

  return {
    documentTitle,
    inputItems: extractInputItems(text),
    calculationItems: [],
    formulas: extractFormulas(text),
    fees: extractFees(text),
    otherRequirements: [],
    tables: extractTables(text),
  };
}
