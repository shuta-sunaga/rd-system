/**
 * AI要件抽出サービス
 * Google Gemini APIを使用してPDFから要件定義を抽出
 */

import { GoogleGenerativeAI } from '@google/generative-ai';

export interface AIExtractedRequirements {
  documentTitle: string;
  documentType: string;
  summary: string;
  inputItems: {
    category: string;
    items: {
      name: string;
      description: string;
      dataType: string;
      required: boolean;
      validationRules?: string[];
    }[];
  }[];
  calculationRules: {
    name: string;
    description: string;
    formula: string;
    conditions?: string[];
    examples?: string[];
  }[];
  feeStructure: {
    category: string;
    items: {
      name: string;
      description: string;
      amount?: string;
      unit?: string;
      conditions?: string[];
    }[];
  }[];
  tables: {
    title: string;
    description: string;
    headers: string[];
    rows: string[][];
  }[];
  additionalNotes: string[];
}

const EXTRACTION_PROMPT = `あなたは要件定義書を作成する専門家です。
以下のPDFテキストから、システム開発のための要件定義を抽出してください。

特に以下の点に注目して抽出してください：

1. **入力が必要な項目**
   - 使用料算定に必要な入力項目
   - 申請・届出に必要な項目
   - データ型（数値、テキスト、日付、選択肢など）

2. **算定に必要な項目**
   - 計算に使用される変数
   - 参照テーブル
   - 条件分岐のパラメータ

3. **算定方法・計算式**
   - 日本語で記載された計算ルールを数式に変換
   - 条件付きの計算ロジック
   - 上限・下限のルール

4. **費用項目**
   - 各種費用の名称と説明
   - 金額や計算方法
   - 適用条件

5. **その他重要な要件**
   - 期限・期間に関するルール
   - 例外処理
   - 参照すべき別表・付表

回答は以下のJSON形式のみで出力してください（説明文は不要）：

{
  "documentTitle": "ドキュメントタイトル",
  "documentType": "規程種別（例：社宅規程、転勤取扱基準）",
  "summary": "ドキュメントの概要（100文字程度）",
  "inputItems": [
    {
      "category": "カテゴリ名",
      "items": [
        {
          "name": "項目名",
          "description": "説明",
          "dataType": "データ型",
          "required": true,
          "validationRules": ["バリデーションルール"]
        }
      ]
    }
  ],
  "calculationRules": [
    {
      "name": "計算名",
      "description": "説明",
      "formula": "計算式（例：家賃 × 本人負担率）",
      "conditions": ["適用条件"],
      "examples": ["計算例"]
    }
  ],
  "feeStructure": [
    {
      "category": "費用カテゴリ",
      "items": [
        {
          "name": "費用名",
          "description": "説明",
          "amount": "金額（あれば）",
          "unit": "単位",
          "conditions": ["適用条件"]
        }
      ]
    }
  ],
  "tables": [
    {
      "title": "テーブル名",
      "description": "説明",
      "headers": ["列1", "列2"],
      "rows": [["値1", "値2"]]
    }
  ],
  "additionalNotes": ["その他の重要事項"]
}`;

export class AIExtractor {
  private genAI: GoogleGenerativeAI;
  private model: any;

  constructor(apiKey?: string) {
    const key = apiKey || process.env.GEMINI_API_KEY || process.env.GOOGLE_API_KEY;
    if (!key) {
      throw new Error('GEMINI_API_KEY または GOOGLE_API_KEY 環境変数を設定してください');
    }
    this.genAI = new GoogleGenerativeAI(key);
    // Gemini 2.0 Flash - 最新で高速
    this.model = this.genAI.getGenerativeModel({ model: 'gemini-2.0-flash' });
  }

  /**
   * PDFテキストから要件を抽出
   */
  async extractRequirements(pdfText: string): Promise<AIExtractedRequirements> {
    const prompt = `${EXTRACTION_PROMPT}\n\n---\n\nPDFテキスト:\n${pdfText}`;

    const result = await this.model.generateContent(prompt);
    const response = await result.response;
    const text = response.text();

    // JSONを抽出
    const jsonMatch = text.match(/\{[\s\S]*\}/);
    if (!jsonMatch) {
      throw new Error('AIからの応答をJSONとして解析できませんでした');
    }

    try {
      return JSON.parse(jsonMatch[0]);
    } catch (e) {
      throw new Error(`JSON解析エラー: ${e}`);
    }
  }

  /**
   * 複数のPDFを統合して要件を抽出
   */
  async extractFromMultiplePDFs(
    pdfTexts: { filename: string; text: string }[]
  ): Promise<AIExtractedRequirements> {
    const combinedText = pdfTexts
      .map((pdf) => `\n\n=== ${pdf.filename} ===\n\n${pdf.text}`)
      .join('\n');

    return this.extractRequirements(combinedText);
  }
}

// シングルトンインスタンスは使用時に作成
export const createAIExtractor = (apiKey?: string) => new AIExtractor(apiKey);
