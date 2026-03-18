/**
 * 伝票キャリア共通定数
 * ヤマト・佐川の PDF 検証などで共有する。
 */

/** 有効なPDFとみなす最小バイト数（空やHTMLエラーページを除外） */
const MIN_PDF_SIZE = 1000;

module.exports = { MIN_PDF_SIZE };
