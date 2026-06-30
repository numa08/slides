#!/usr/bin/env node
// スライド生成の汎用ランナー。
// 渡された Markdown と同じディレクトリにある build_slides.js を実行し、
// スライド(.pptx)を同ディレクトリに生成する。
//
// 使い方: pnpm slides <path-to-md>
//   例)   pnpm slides 202607/FT8の話.md  ->  202607/FT8の話.pptx
//
// 規約: 各イベントディレクトリに build_slides.js を置く（その中で内容を定義する）。
const path = require("path");
const fs = require("fs");

const mdPath = process.argv[2];
if (!mdPath) {
  console.error("使い方: pnpm slides <path-to-md>");
  process.exit(1);
}
const absMd = path.resolve(mdPath);
if (!fs.existsSync(absMd)) {
  console.error(`Markdown が見つかりません: ${mdPath}`);
  process.exit(1);
}
const builder = path.join(path.dirname(absMd), "build_slides.js");
if (!fs.existsSync(builder)) {
  console.error(`build_slides.js が見つかりません: ${builder}`);
  process.exit(1);
}

// build_slides.js は process.argv[2]（= md パス）を読んで出力先を決める
require(builder);
