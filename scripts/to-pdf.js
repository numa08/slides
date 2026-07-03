#!/usr/bin/env node
// スライドを PDF として出力するランナー（SpeakerDeck アップロード用）。
//
// pptx を生成したうえで、閲覧環境と同じ Windows の PowerPoint 本体で PDF に
// 変換する。Linux 側のレンダラ(LibreOffice / Chromium 等)を使うと Windows フォント
// (Yu Gothic / Consolas)が置換され、手計算したレイアウトが崩れる。フォント忠実性を
// 優先し、閲覧時と同一のレンダラである PowerPoint COM を使う（WSL2 + Windows PowerPoint 前提）。
//
// 使い方: pnpm slides:pdf <path-to-md>
//   例)   pnpm slides:pdf 202607/FT8の話.md  ->  202607/FT8の話.pdf
const path = require("path");
const fs = require("fs");
const { execFileSync } = require("child_process");

const PP_SAVE_AS_PDF = 32; // PowerPoint の SaveAs フォーマット定数 ppSaveAsPDF

const mdPath = process.argv[2];
if (!mdPath) {
  console.error("使い方: pnpm slides:pdf <path-to-md>");
  process.exit(1);
}
const absMd = path.resolve(mdPath);
if (!fs.existsSync(absMd)) {
  console.error(`Markdown が見つかりません: ${mdPath}`);
  process.exit(1);
}

const dir = path.dirname(absMd);
const stem = path.basename(mdPath).replace(/\.md$/i, "");
const pptxPath = path.join(dir, stem + ".pptx");
const pdfPath = path.join(dir, stem + ".pdf");

// 1) pptx を生成（既存ランナーを子プロセスで実行し、非同期の書き出し完了まで待つ）
execFileSync(process.execPath, [path.join(__dirname, "gen-slides.js"), mdPath], { stdio: "inherit" });
if (!fs.existsSync(pptxPath)) {
  console.error(`pptx が生成されていません: ${pptxPath}`);
  process.exit(1);
}

// 2) Windows の PowerPoint で pptx -> pdf。
//    PowerPoint COM は UNC パス(\\wsl.localhost\...)や日本語 argv で不安定になりうるため、
//    Windows ローカルの一時フォルダに ASCII 名でコピーしてから変換し、PDF を WSL 側へ戻す。
const winToWsl = (p) => execFileSync("wslpath", ["-u", p]).toString().trim();

let winTmpDir;
try {
  // GetTempPath は末尾に "\\" を含む Windows パスを返す
  winTmpDir = execFileSync("powershell.exe", ["-NoProfile", "-Command", "[IO.Path]::GetTempPath()"]).toString().replace(/[\r\n]/g, "");
} catch (e) {
  console.error("PDF 変換には Windows の powershell.exe が必要です（WSL2 + Windows PowerPoint 前提）。");
  console.error(e.message);
  process.exit(1);
}
const wslTmpDir = winToWsl(winTmpDir);
const tag = `_slidegen_${process.pid}`; // 並行実行の衝突回避
const tmpPptxWsl = path.join(wslTmpDir, tag + ".pptx");
const tmpPdfWsl = path.join(wslTmpDir, tag + ".pdf");
const tmpPptxWin = winTmpDir + tag + ".pptx";
const tmpPdfWin = winTmpDir + tag + ".pdf";

const cleanup = () => {
  for (const f of [tmpPptxWsl, tmpPdfWsl]) {
    try { fs.unlinkSync(f); } catch { /* 無ければ無視 */ }
  }
};

try {
  fs.copyFileSync(pptxPath, tmpPptxWsl);
  const ps = [
    "$ErrorActionPreference='Stop'",
    "$pp=New-Object -ComObject PowerPoint.Application",
    // Open(FileName, ReadOnly, Untitled, WithWindow)
    `$d=$pp.Presentations.Open('${tmpPptxWin}',$true,$false,$false)`,
    `$d.SaveAs('${tmpPdfWin}',${PP_SAVE_AS_PDF})`,
    "$d.Close()",
    "$pp.Quit()",
  ].join("; ");
  execFileSync("powershell.exe", ["-NoProfile", "-Command", ps], { stdio: "inherit" });
  if (!fs.existsSync(tmpPdfWsl)) throw new Error("PowerPoint が PDF を出力しませんでした");
  fs.copyFileSync(tmpPdfWsl, pdfPath);
} finally {
  cleanup();
}
console.log("pdf:", pdfPath);
