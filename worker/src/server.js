require('dotenv').config();

const crypto = require('crypto');
const express = require('express');
const { v4: uuidv4 } = require('uuid');
const fs = require('fs');
const { processYamato } = require('./carriers/yamato');
const { processSagawa } = require('./carriers/sagawa');
const { uploadToDrive, validateDriveConfig } = require('./drive-service');

const app = express();
const PORT = process.env.PORT || 3000;

// ジョブストア（TTL付き）。エントリは { job, createdAt }
const JOB_TTL_MS = Number(process.env.JOB_TTL_MS) || 24 * 60 * 60 * 1000; // デフォルト24時間
const JOB_CLEANUP_INTERVAL_MS = Number(process.env.JOB_CLEANUP_INTERVAL_MS) || 60 * 60 * 1000; // 1時間
const jobStore = new Map();

function cleanExpiredJobs() {
  const now = Date.now();
  for (const [id, entry] of jobStore.entries()) {
    if (now - entry.createdAt > JOB_TTL_MS) jobStore.delete(id);
  }
}
setInterval(cleanExpiredJobs, JOB_CLEANUP_INTERVAL_MS);

app.use(express.json({ limit: '10mb' }));

app.use((req, _res, next) => {
  console.log(`${req.method} ${req.path}`, { ip: req.ip });
  next();
});

// 認証
function authenticate(req, res, next) {
  const apiKey = req.headers['x-api-key'];
  const expectedKey = process.env.WORKER_API_KEY;
  if (!expectedKey) {
    if (process.env.NODE_ENV === 'production') {
      console.error('本番環境で WORKER_API_KEY が未設定です。');
      return res.status(500).send('サーバー設定エラー: APIキーが設定されていません');
    }
    console.warn('WORKER_API_KEY が未設定です。認証をスキップします。');
    return next();
  }
  try {
    const bufApiKey = Buffer.from(apiKey != null ? String(apiKey) : '', 'utf8');
    const bufExpected = Buffer.from(expectedKey, 'utf8');
    const len = bufExpected.length;
    if (bufApiKey.length !== len) {
      const dummy = Buffer.alloc(len, 0);
      crypto.timingSafeEqual(dummy, bufExpected);
      return res.status(401).json({ success: false, message: '認証エラー: APIキーが無効です' });
    }
    if (!crypto.timingSafeEqual(bufApiKey, bufExpected)) {
      console.warn('認証失敗: 不正なAPIキー', { ip: req.ip });
      return res.status(401).json({ success: false, message: '認証エラー: APIキーが無効です' });
    }
  } catch {
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
    activeJobs: [...jobStore.values()].filter(entry => entry.job?.status === 'processing').length,
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
  const now = Date.now();
  jobStore.set(jobId, {
    job: { status: 'processing', carrier, shippingDate, createdAt: new Date().toISOString() },
    createdAt: now,
  });
  cleanExpiredJobs();

  // 即座にレスポンスを返す
  res.json({ success: true, jobId });

  // バックグラウンドで処理を実行
  processSlipJob(jobId, carrier, csvContent, shippingDate);
});

// ジョブステータス確認（デバッグ用）
app.get('/api/job/:jobId', (req, res) => {
  const entry = jobStore.get(req.params.jobId);
  if (!entry?.job) {
    return res.status(404).json({ success: false, message: 'ジョブが見つかりません' });
  }
  res.json({ success: true, ...entry.job });
});

// バックグラウンド処理
async function processSlipJob(jobId, carrier, csvContent, shippingDate) {
  console.log('バックグラウンド処理開始', { jobId, carrier, shippingDate });

  let tmpDir;
  try {
    // Step 1: Playwright でPDF取得
    let pdfPath;
    if (carrier === 'yamato') {
      const result = await processYamato(csvContent, shippingDate);
      pdfPath = result.pdfPath;
      tmpDir = result.tmpDir;
    } else {
      const result = await processSagawa(csvContent, shippingDate);
      pdfPath = result.pdfPath;
      tmpDir = result.tmpDir;
    }
    console.log('PDF ダウンロード完了', { jobId, pdfPath });

    // Step 2: Google Drive にアップロード
    const driveStatus = validateDriveConfig();
    if (!driveStatus.configured) {
      throw new Error('Google Drive設定が無効です: ' + driveStatus.message);
    }

    const driveResult = await uploadToDrive(pdfPath, carrier, shippingDate, jobId);
    console.log('Drive アップロード完了', { jobId, fileId: driveResult.fileId });

    try {
      await fs.promises.unlink(pdfPath);
      console.log('ローカルPDF削除完了', { jobId, fileId: driveResult.fileId });
    } catch (unlinkErr) {
      console.warn('ローカルPDF削除に失敗（アップロードは成功）', { jobId, pdfPath, error: unlinkErr?.message });
    }
    if (tmpDir) {
      fs.promises.rm(tmpDir, { recursive: true, force: true }).catch(() => {});
    }

    const entry = jobStore.get(jobId);
    jobStore.set(jobId, {
      ...entry,
      job: {
        ...entry?.job,
        status: 'completed',
        driveFileId: driveResult.fileId,
        driveWebViewLink: driveResult.webViewLink,
        completedAt: new Date().toISOString(),
      },
      createdAt: entry?.createdAt ?? Date.now(),
    });

    // GASへ完了コールバック
    await notifyGas(jobId, 'completed', { driveFileId: driveResult.fileId, driveWebViewLink: driveResult.webViewLink });
  } catch (error) {
    console.error('バックグラウンド処理エラー', { jobId, error: error.message, stack: error.stack });
    if (tmpDir) {
      try {
        fs.promises.rm(tmpDir, { recursive: true, force: true }).catch(() => {});
      } catch (_) {}
    }
    const entry = jobStore.get(jobId);
    jobStore.set(jobId, {
      ...entry,
      job: {
        ...entry?.job,
        status: 'error',
        error: error.message,
        completedAt: new Date().toISOString(),
      },
      createdAt: entry?.createdAt ?? Date.now(),
    });

    // GASへエラーコールバック
    await notifyGas(jobId, 'error', { error: error.message });
  }
}

/**
 * GAS Web App へステータスコールバックを送信
 */
async function notifyGas(jobId, status, extra = {}) {
  const callbackUrl = process.env.GAS_CALLBACK_URL;
  if (!callbackUrl) {
    console.log('GAS_CALLBACK_URL 未設定、コールバックスキップ');
    return;
  }

  const payload = {
    jobId,
    status,
    apiKey: process.env.WORKER_API_KEY || '',
    ...extra,
  };

  try {
    const res = await fetch(callbackUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload),
      signal: AbortSignal.timeout(10000),
    });
    console.log('GASコールバック送信完了', { jobId, status, httpStatus: res.status });
  } catch (err) {
    // コールバック失敗はログのみ（メイン処理には影響させない）
    console.warn('GASコールバック送信失敗', { jobId, error: err.message });
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
