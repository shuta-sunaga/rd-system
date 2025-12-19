/**
 * 要件定義作成システム - Webサーバー
 * PDFアップロード → 要件抽出 → Excel生成
 */

import 'dotenv/config';
import express from 'express';
import multer from 'multer';
import cors from 'cors';
import path from 'path';
import { fileURLToPath } from 'url';
import { parsePDFBuffer } from './services/pdfParser.js';
import { AIExtractor, AIExtractedRequirements } from './services/aiExtractor.js';
import { generateRequirementsExcel } from './services/excelGenerator.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const PORT = process.env.PORT || 3003;

// ミドルウェア設定
app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, '../public')));

// ファイルアップロード設定
const storage = multer.memoryStorage();
const upload = multer({
  storage,
  limits: {
    fileSize: 10 * 1024 * 1024, // 10MB
  },
  fileFilter: (req, file, cb) => {
    if (file.mimetype === 'application/pdf') {
      cb(null, true);
    } else {
      cb(new Error('PDFファイルのみアップロード可能です'));
    }
  },
});

// AIエクストラクター（遅延初期化）
let aiExtractor: AIExtractor | null = null;

function getAIExtractor(): AIExtractor {
  if (!aiExtractor) {
    aiExtractor = new AIExtractor();
  }
  return aiExtractor;
}

/**
 * ヘルスチェック
 */
app.get('/api/health', (req, res) => {
  res.json({ status: 'ok', timestamp: new Date().toISOString() });
});

/**
 * PDF解析・要件定義生成API
 */
app.post('/api/generate', upload.array('files', 5), async (req, res) => {
  try {
    const files = req.files as Express.Multer.File[];

    if (!files || files.length === 0) {
      return res.status(400).json({
        error: 'PDFファイルがアップロードされていません',
      });
    }

    console.log(`Processing ${files.length} PDF file(s)...`);

    // 複数PDFのテキストを抽出
    const pdfTexts: { filename: string; text: string }[] = [];

    for (const file of files) {
      console.log(`Parsing: ${file.originalname}`);
      const parsed = await parsePDFBuffer(file.buffer);
      pdfTexts.push({
        filename: file.originalname,
        text: parsed.rawText,
      });
    }

    // AIで要件を抽出
    console.log('Extracting requirements with AI...');
    let requirements: AIExtractedRequirements;

    if (pdfTexts.length === 1) {
      requirements = await getAIExtractor().extractRequirements(pdfTexts[0].text);
    } else {
      requirements = await getAIExtractor().extractFromMultiplePDFs(pdfTexts);
    }

    // Excel生成
    console.log('Generating Excel file...');
    const excelBuffer = await generateRequirementsExcel(requirements);

    // レスポンス
    const filename = `要件定義書_${new Date().toISOString().slice(0, 10)}.xlsx`;

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${encodeURIComponent(filename)}`);
    res.send(excelBuffer);

    console.log('Excel file generated successfully');
  } catch (error) {
    console.error('Error processing request:', error);
    res.status(500).json({
      error: error instanceof Error ? error.message : '処理中にエラーが発生しました',
    });
  }
});

/**
 * PDF解析のみ（プレビュー用）
 */
app.post('/api/preview', upload.array('files', 5), async (req, res) => {
  try {
    const files = req.files as Express.Multer.File[];

    if (!files || files.length === 0) {
      return res.status(400).json({
        error: 'PDFファイルがアップロードされていません',
      });
    }

    // 複数PDFのテキストを抽出
    const pdfTexts: { filename: string; text: string }[] = [];

    for (const file of files) {
      const parsed = await parsePDFBuffer(file.buffer);
      pdfTexts.push({
        filename: file.originalname,
        text: parsed.rawText,
      });
    }

    // AIで要件を抽出
    let requirements: AIExtractedRequirements;

    if (pdfTexts.length === 1) {
      requirements = await getAIExtractor().extractRequirements(pdfTexts[0].text);
    } else {
      requirements = await getAIExtractor().extractFromMultiplePDFs(pdfTexts);
    }

    res.json(requirements);
  } catch (error) {
    console.error('Error processing preview:', error);
    res.status(500).json({
      error: error instanceof Error ? error.message : '処理中にエラーが発生しました',
    });
  }
});

/**
 * フロントエンドHTML
 */
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, '../public/index.html'));
});

// サーバー起動
app.listen(PORT, () => {
  console.log(`
╔════════════════════════════════════════════════════════════╗
║                                                            ║
║   🌸 要件定義作成システム                                  ║
║                                                            ║
║   Server running at: http://localhost:${PORT}               ║
║                                                            ║
║   Endpoints:                                               ║
║   - POST /api/generate  : PDF → Excel変換                  ║
║   - POST /api/preview   : PDF解析プレビュー                ║
║   - GET  /api/health    : ヘルスチェック                   ║
║                                                            ║
╚════════════════════════════════════════════════════════════╝
  `);
});

export default app;
