require('dotenv').config();

const express = require('express');
const { v4: uuidv4 } = require('uuid');
const { processYamato } = require('./carriers/yamato');
const { processSagawa } = require('./carriers/sagawa');
const { uploadToDrive, validateDriveConfig } = require('./drive-service');

const app = express();
const PORT = process.env.PORT || 3000;

// インメモリ ジョブストア（Cloud Run は1インスタンスなのでこれで十分）
const jobStore = new Map();

app.use(express.json({ limit: '10mb' }));

app.use((req, _res, next) => {
  console.log(`${req.method} ${req.path}`, { ip: req.ip });
  next();
});

// 認証
function authenticate(req, res, next) {
  const apiKey = req.body?.apiKey || req.headers['x-api-key'];
  const expectedKey = process.env.WORKER_API_KEY;
  if (!expectedKey) {
    console.warn('WORKER_API_KEY が未設定です。認証をスキップします。');
    return next();
  }
  if (apiKey !== expectedKey) {
    console.warn('認証失敗: 不正なAPIキー', { ip: req.ip });
    return res.status(401).json({ success: false, message: '認証エラー: APIキーが無効です' });
  }
  next();
}

// ヘルスチェック
app.get('/health', (_req, res) => {
  const driveStatus = validateDriveConfig();
  res.json({
    status: 'ok',
    timestamp: new Date().toISOString(),
    uptime: process.uptime(),
    driveConfigured: driveStatus.configured,
    driveMessage: driveStatus.message,
    activeJobs: [...jobStore.values()].filter(j => j.status === 'processing').length,
  });
});

// 伝票発行（非同期 fire-and-forget）
app.post('/api/print-slip', authenticate, (req, res) => {
  const { carrier, csvContent, shippingDate } = req.body;

  if (!carrier || !csvContent) {
    return res.status(400).json({
      success: false,
      message: '必須パラメータが不足しています (carrier, csvContent)',
    });
  }
  if (!['yamato', 'sagawa'].includes(carrier)) {
    return res.status(400).json({
      success: false,
      message: `不正な配送業者: ${carrier}（yamato / sagawa のみ対応）`,
    });
  }

  const jobId = uuidv4();
  jobStore.set(jobId, {
    status: 'processing',
    carrier,
    shippingDate,
    createdAt: new Date().toISOString(),
  });

  // 即座にレスポンスを返す
  res.json({ success: true, jobId });

  // バックグラウンドで処理を実行
  processSlipJob(jobId, carrier, csvContent, shippingDate);
});

// ジョブステータス確認（デバッグ用）
app.get('/api/job/:jobId', (req, res) => {
  const job = jobStore.get(req.params.jobId);
  if (!job) {
    return res.status(404).json({ success: false, message: 'ジョブが見つかりません' });
  }
  res.json({ success: true, ...job });
});

// バックグラウンド処理
async function processSlipJob(jobId, carrier, csvContent, shippingDate) {
  console.log('バックグラウンド処理開始', { jobId, carrier, shippingDate });

  try {
    // Step 1: Playwright でPDF取得
    let pdfPath;
    if (carrier === 'yamato') {
      pdfPath = await processYamato(csvContent, shippingDate);
    } else {
      pdfPath = await processSagawa(csvContent, shippingDate);
    }
    console.log('PDF ダウンロード完了', { jobId, pdfPath });

    // Step 2: Google Drive にアップロード
    const driveStatus = validateDriveConfig();
    if (!driveStatus.configured) {
      throw new Error('Google Drive設定が無効です: ' + driveStatus.message);
    }

    const driveResult = await uploadToDrive(pdfPath, carrier, shippingDate, jobId);
    console.log('Drive アップロード完了', { jobId, fileId: driveResult.fileId });

    jobStore.set(jobId, {
      ...jobStore.get(jobId),
      status: 'completed',
      driveFileId: driveResult.fileId,
      driveWebViewLink: driveResult.webViewLink,
      completedAt: new Date().toISOString(),
    });
  } catch (error) {
    console.error('バックグラウンド処理エラー', { jobId, error: error.message, stack: error.stack });
    jobStore.set(jobId, {
      ...jobStore.get(jobId),
      status: 'error',
      error: error.message,
      completedAt: new Date().toISOString(),
    });
  }
}

// エラーハンドリング
app.use((err, _req, res, _next) => {
  console.error('未処理エラー', { error: err.message, stack: err.stack });
  res.status(500).json({ success: false, message: 'サーバー内部エラー' });
});

app.listen(PORT, () => {
  console.log(`配送伝票ワーカー起動: port=${PORT}, env=${process.env.NODE_ENV || 'development'}`);
  const driveStatus = validateDriveConfig();
  console.log(`Drive設定: ${driveStatus.configured ? '有効' : '無効'} (${driveStatus.message})`);
});

module.exports = app;
