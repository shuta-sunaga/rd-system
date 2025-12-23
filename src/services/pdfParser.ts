/**
 * PDF解析サービス
 * PDFファイルからテキストを抽出し、要件定義に必要な情報を解析する
 *
 * ハイブリッド方式:
 * 1. まずpdf-parseでテキスト抽出を試行
 * 2. 日本語文字が少ない場合はOCRにフォールバック
 */

import pdf from 'pdf-parse';
import { readFile } from 'fs/promises';

const JAPANESE_CHAR_THRESHOLD = 50;

export interface ParsedPDFContent {
  rawText: string;
  pages: string[];
  metadata: {
    title?: string;
    author?: string;
    pageCount: number;
  };
  extractionMethod?: 'pdf-parse' | 'gemini-vision';
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

function countJapaneseChars(text: string): number {
  const matches = text.match(/[\u3040-\u309f\u30a0-\u30ff\u4e00-\u9faf]/g);
  return matches ? matches.length : 0;
}

/**
 * Gemini APIエラーを詳細なメッセージに変換
 */
function handleGeminiError(error: unknown): Error {
  const err = error as Error & { status?: number; code?: string };
  const message = err.message || String(error);
  const status = err.status;
  const code = err.code;

  // ネットワークエラー
  if (
    message.includes('fetch failed') ||
    message.includes('ECONNREFUSED') ||
    message.includes('ENOTFOUND') ||
    message.includes('getaddrinfo')
  ) {
    return new Error(
      `[Gemini Vision 接続エラー] サーバーに接続できません。\n` +
      `原因: ネットワーク制限またはファイアウォールでAPIアクセスがブロックされている可能性があります。\n` +
      `対象ドメイン: generativelanguage.googleapis.com\n` +
      `詳細: ${message}`
    );
  }

  // タイムアウト
  if (
    message.includes('timeout') ||
    message.includes('ETIMEDOUT') ||
    message.includes('AbortError')
  ) {
    return new Error(
      `[Gemini Vision タイムアウト] APIリクエストがタイムアウトしました。\n` +
      `原因: ネットワーク遅延またはプロキシ設定の問題の可能性があります。\n` +
      `詳細: ${message}`
    );
  }

  // SSL/TLS エラー
  if (
    message.includes('certificate') ||
    message.includes('SSL') ||
    message.includes('TLS') ||
    message.includes('CERT_')
  ) {
    return new Error(
      `[Gemini Vision SSL/TLSエラー] セキュリティ証明書の問題が発生しました。\n` +
      `原因: 企業プロキシによるSSLインスペクションの可能性があります。\n` +
      `詳細: ${message}`
    );
  }

  // 認証エラー
  if (status === 401 || status === 403 || message.includes('API key')) {
    return new Error(
      `[Gemini Vision 認証エラー] APIキーが無効または権限がありません。\n` +
      `ステータス: ${status || 'N/A'}\n` +
      `詳細: ${message}`
    );
  }

  // レート制限
  if (status === 429 || message.includes('quota') || message.includes('rate')) {
    return new Error(
      `[Gemini Vision レート制限] APIの使用制限に達しました。\n` +
      `しばらく待ってから再試行してください。\n` +
      `詳細: ${message}`
    );
  }

  // サーバーエラー
  if (status && status >= 500) {
    return new Error(
      `[Gemini Vision サーバーエラー] Googleサーバー側でエラーが発生しました。\n` +
      `ステータス: ${status}\n` +
      `詳細: ${message}`
    );
  }

  // その他のエラー
  return new Error(
    `[Gemini Vision エラー] PDF解析中にエラーが発生しました。\n` +
    `コード: ${code || 'N/A'}, ステータス: ${status || 'N/A'}\n` +
    `詳細: ${message}`
  );
}

async function extractTextWithGeminiVision(buffer: Buffer, pageCount: number): Promise<{ text: string; pageCount: number }> {
  console.log('Gemini VisionでPDF直接解析を開始...');

  const { GoogleGenerativeAI } = await import('@google/generative-ai');

  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) {
    throw new Error('GEMINI_API_KEY is not set');
  }

  const genAI = new GoogleGenerativeAI(apiKey);
  const model = genAI.getGenerativeModel({ model: 'gemini-2.0-flash-exp' });

  const base64Pdf = buffer.toString('base64');

  let result;
  try {
    result = await model.generateContent([
      {
        inlineData: {
          mimeType: 'application/pdf',
          data: base64Pdf,
        },
      },
      'このPDFに含まれるテキストを全て抽出してください。レイアウトを可能な限り維持し、表がある場合は表形式で出力してください。ページ区切りは「--- ページ区切り ---」で示してください。説明や解説は不要です。テキストのみを出力してください。',
    ]);
  } catch (error: unknown) {
    throw handleGeminiError(error);
  }

  const text = result.response.text();
  console.log('Gemini Vision完了');

  return {
    text,
    pageCount,
  };
}

export async function parsePDF(filePath: string): Promise<ParsedPDFContent> {
  const dataBuffer = await readFile(filePath);
  return parsePDFBuffer(dataBuffer);
}

export async function parsePDFBuffer(buffer: Buffer): Promise<ParsedPDFContent> {
  console.log('pdf-parseでテキスト抽出を試行...');
  const data = await pdf(buffer);
  const japaneseCount = countJapaneseChars(data.text);
  console.log('日本語文字数: ' + japaneseCount);

  if (japaneseCount >= JAPANESE_CHAR_THRESHOLD) {
    console.log('pdf-parseでの抽出成功');
    const pages = data.text.split(/\n\s*-\s*\d+\s*-\s*\n/).filter((p: string) => p.trim());
    return {
      rawText: data.text,
      pages,
      metadata: {
        title: data.info?.Title,
        author: data.info?.Author,
        pageCount: data.numpages,
      },
      extractionMethod: 'pdf-parse',
    };
  }

  console.log('日本語文字が少ないためGemini Visionにフォールバック');
  const ocrResult = await extractTextWithGeminiVision(buffer, data.numpages);
  const pages = ocrResult.text.split(/--- ページ区切り ---/).filter((p: string) => p.trim());

  return {
    rawText: ocrResult.text,
    pages,
    metadata: {
      title: data.info?.Title,
      author: data.info?.Author,
      pageCount: ocrResult.pageCount,
    },
    extractionMethod: 'gemini-vision',
  };
}

export function extractInputItems(text: string): RequirementItem[] {
  const items: RequirementItem[] = [];
  const matches = text.matchAll(/[「『]?([^「『」』\n]+)[」』]?とは[、,]([^。]+)/g);
  for (const match of matches) {
    items.push({
      category: '定義',
      name: match[1].trim(),
      description: match[2].trim(),
      inputType: 'text',
    });
  }
  return items;
}

export function extractFormulas(text: string): CalculationFormula[] {
  const formulas: CalculationFormula[] = [];
  const matches = text.matchAll(/([^。\n]{5,50})は[、,]([^。]+)とする/g);
  for (const match of matches) {
    const name = match[1].trim();
    const desc = match[2].trim();
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

function convertToFormula(description: string): string {
  let formula = description;
  formula = formula.replace(/([0-9０-９]+)割/g, (_, n) => toHalfWidth(n) + ' * 0.1');
  formula = formula.replace(/([0-9０-９]+)分/g, (_, n) => toHalfWidth(n) + ' * 0.01');
  formula = formula.replace(/([0-9０-９]+)[%％]/g, (_, n) => toHalfWidth(n) + ' * 0.01');
  formula = formula.replace(/([0-9０-９]+)倍/g, (_, n) => '* ' + toHalfWidth(n));
  formula = formula.replace(/([0-9０-９,，]+)円/g, (_, n) => toHalfWidth(n.replace(/[,，]/g, '')));
  return formula;
}

function toHalfWidth(str: string): string {
  return str.replace(/[０-９]/g, (s) => String.fromCharCode(s.charCodeAt(0) - 0xFEE0));
}

function extractVariables(description: string): string[] {
  const variables: string[] = [];
  const matches = description.matchAll(/([^、。\n]+(?:料|額|費|金|率))/g);
  for (const match of matches) {
    if (!variables.includes(match[1])) {
      variables.push(match[1]);
    }
  }
  return variables;
}

export function extractTables(text: string): TableData[] {
  const tables: TableData[] = [];
  const pattern = /[「『]?別表\s*(\d+)[」』]?\s*([^\n]+)\n([\s\S]*?)(?=[「『]?別表|付則|以\s*上|$)/g;
  const matches = text.matchAll(pattern);
  for (const match of matches) {
    const title = '別表' + match[1] + ' ' + match[2].trim();
    const content = match[3];
    const lines = content.split('\n').filter((l) => l.trim());
    if (lines.length > 0) {
      tables.push({
        title,
        headers: [],
        rows: lines.map((l) => l.split(/\s{2,}|\t/).filter((c) => c.trim())),
      });
    }
  }
  return tables;
}

export function extractFees(text: string): FeeItem[] {
  const fees: FeeItem[] = [];
  const pattern = /([^。\n]+(?:費|料|金|手当))[はを][、,]([^。]+)/g;
  const matches = text.matchAll(pattern);
  for (const match of matches) {
    const name = match[1].trim();
    const desc = match[2].trim();
    const amountMatch = desc.match(/([0-9０-９,，]+)\s*円/);
    fees.push({
      name,
      description: desc,
      amount: amountMatch ? toHalfWidth(amountMatch[1].replace(/[,，]/g, '')) : undefined,
    });
  }
  return fees;
}

export async function extractRequirementsFromPDF(buffer: Buffer): Promise<ExtractedRequirements> {
  const parsed = await parsePDFBuffer(buffer);
  const text = parsed.rawText;
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
