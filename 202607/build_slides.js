// FT8 LT スライド生成スクリプト（裏テック無尽 2026/07/03 @numa08）
// 原稿: FT8の話.md に基づく。ダークテーマ＝FT8ウォーターフォール/夜空モチーフ。
const pptxgen = require("pptxgenjs");
const path = require("path");

// 生成対象の Markdown パスを受け取り、同じディレクトリに同名の .pptx を出力する。
// 使い方: pnpm slides <path-to-md>  /  node build_slides.js <path-to-md>
// 注: スライドの内容はこのスクリプト内に定義（FT8専用デザイン）。md パスは出力先の決定に使う。
const mdPath = process.argv[2];
if (!mdPath) {
  console.error("使い方: pnpm slides <path-to-md>");
  process.exit(1);
}
const outPath = path.join(
  path.dirname(path.resolve(mdPath)),
  path.basename(mdPath).replace(/\.md$/i, "") + ".pptx"
);

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5 inch
pres.author = "numa08";
pres.title = "FT8 の話";

const W = 13.33, H = 7.5, M = 0.7;

const C = {
  bg: "0A0F2C", panel: "141D40", panel2: "0E1530",
  ink: "EAF2FF", mute: "93A6CC", mute2: "5E6E96",
  cyan: "2EE6D6", gold: "FFC94D", red: "FF6B6B",
  line: "27325E", green: "8DE08D",
};
const F = { jp: "Yu Gothic", mono: "Consolas" };

const shadow = () => ({ type: "outer", color: "000000", blur: 10, offset: 3, angle: 90, opacity: 0.35 });

// 決定論的擬似乱数（再現性のため）
let _seed = 20260704;
function rnd() { _seed = (_seed * 1103515245 + 12345) & 0x7fffffff; return _seed / 0x7fffffff; }

// ウォーターフォール風の装飾バンド
function waterfall(slide, x, y, w, h, n = 56) {
  const cw = w / n;
  for (let i = 0; i < n; i++) {
    const v = rnd();
    const col = v > 0.86 ? C.gold : v > 0.62 ? C.cyan : C.line;
    const tr = v > 0.62 ? 100 - Math.round(v * 55) : 78;
    const bh = h * (0.35 + v * 0.65);
    slide.addShape(pres.shapes.RECTANGLE, {
      x: x + i * cw, y: y + (h - bh), w: cw * 0.7, h: bh,
      fill: { color: col, transparency: tr }, line: { type: "none" },
    });
  }
}

// 共通の下地（背景・キッカー・ノンブル）
function base(slide, kicker, page) {
  slide.background = { color: C.bg };
  if (kicker) {
    slide.addText([
      { text: "FT8", options: { color: C.cyan, bold: true } },
      { text: "  ▸  " + kicker, options: { color: C.mute2 } },
    ], { x: M, y: 0.34, w: 10, h: 0.32, fontSize: 11, fontFace: F.mono, margin: 0, charSpacing: 1 });
  }
  if (page) {
    slide.addText(String(page).padStart(2, "0") + " / 12", {
      x: W - 1.9, y: 6.95, w: 1.2, h: 0.3, align: "right",
      fontSize: 10, color: C.mute2, fontFace: F.mono, margin: 0,
    });
  }
}

// セクション見出し（タイトル左にアクセントバー）
function heading(slide, text, opt = {}) {
  const y = opt.y ?? 0.95;
  slide.addShape(pres.shapes.RECTANGLE, { x: M, y: y + 0.06, w: 0.09, h: 0.62, fill: { color: opt.bar || C.cyan }, line: { type: "none" } });
  slide.addText(text, {
    x: M + 0.28, y, w: opt.w || W - 2 * M - 0.3, h: 0.74,
    fontSize: opt.size || 30, bold: true, color: C.ink, fontFace: F.jp, valign: "middle", margin: 0,
  });
}

// カード
function card(slide, x, y, w, h, fill) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x, y, w, h, rectRadius: 0.08,
    fill: { color: fill || C.panel }, line: { color: C.line, width: 1 }, shadow: shadow(),
  });
}

/* ---------------- 1. TITLE ---------------- */
{
  const s = pres.addSlide();
  s.background = { color: C.bg };
  waterfall(s, 0, 0, W, 1.5, 80);
  waterfall(s, 0, H - 1.2, W, 1.2, 80);
  s.addText("FT8 の話", {
    x: M, y: 2.35, w: W - 2 * M, h: 1.4, fontSize: 70, bold: true, color: C.ink, fontFace: F.jp, margin: 0,
  });
  s.addText([
    { text: "77bit", options: { color: C.gold, bold: true } },
    { text: " で地球の裏側と交信する、", options: { color: C.ink } },
    { text: "執念", options: { color: C.cyan, bold: true } },
    { text: "のプロトコル", options: { color: C.ink } },
  ], { x: M, y: 3.75, w: W - 2 * M, h: 0.6, fontSize: 23, fontFace: F.jp, margin: 0 });
  s.addText([
    { text: "裏テック無尽  2026/07/03", options: { color: C.mute } },
    { text: "      @numa08", options: { color: C.mute2 } },
  ], { x: M, y: 4.7, w: 8, h: 0.4, fontSize: 15, fontFace: F.mono, margin: 0 });
  s.addText("🍺 ゆるく聞いてください / 気になったら後で捕まえて", {
    x: M, y: 5.2, w: 9, h: 0.4, fontSize: 13, italic: true, color: C.mute2, fontFace: F.jp, margin: 0,
  });
  s.addNotes("アマチュア無線で使うFT8の話。Web/モバイル/クラウドの人には低レイヤー。ビール片手に、データが必ず欠ける世界の執念を見てください。");
}

/* ---------------- 2. DEMO LINK + QR ---------------- */
{
  const s = pres.addSlide();
  base(s, "live demo", 2);
  heading(s, "デモ");

  // QRコード（白カードで quiet zone とコントラストを確保）
  const qw = 3.9, qx = M + 0.2, qy = 2.1, qh = 3.9;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: qx, y: qy, w: qw, h: qh, rectRadius: 0.08,
    fill: { color: "FFFFFF" }, line: { color: C.line, width: 1 }, shadow: shadow(),
  });
  const pad = 0.5, img = qw - 2 * pad;
  s.addImage({ path: path.join(__dirname, "qr-demo.png"), x: qx + pad, y: qy + pad, w: img, h: img });

  // URL（QRの右に、垂直中央で大きく）
  const ux = qx + qw + 0.7;
  s.addText("https://ft8.numa08.dev", {
    x: ux, y: qy, w: W - M - ux, h: qh, fontSize: 34, bold: true,
    color: C.cyan, fontFace: F.mono, valign: "middle", margin: 0,
  });

  s.addNotes("デモページ(https://ft8.numa08.dev)。使い方や『開いたまま待っていてほしい』お願いは口頭で伝える。");
}

/* ---------------- 3. WHAT IS FT8 ---------------- */
{
  const s = pres.addSlide();
  base(s, "what is it", 3);
  heading(s, "FT8 とは？");
  s.addText([
    { text: "微弱電波", options: { color: C.cyan, bold: true } },
    { text: "で離れた無線機同士がデジタル信号をやり取りする ", options: { color: C.ink } },
    { text: "オープンなプロトコル", options: { color: C.gold, bold: true } },
  ], { x: M, y: 1.95, w: W - 2 * M, h: 0.5, fontSize: 19, fontFace: F.jp, margin: 0 });

  const stats = [
    ["77 bit", "1回に運べる情報量\nUTF-8なら9文字に満たない"],
    ["数ワット", "送信電力\nLED電球ほどの小さな電力で"],
    ["地球の裏側〜宇宙", "到達距離\nノイズと減衰でデータは欠ける"],
  ];
  const cw = (W - 2 * M - 2 * 0.4) / 3;
  stats.forEach(([big, sub], i) => {
    const x = M + i * (cw + 0.4);
    card(s, x, 2.95, cw, 2.5);
    s.addShape(pres.shapes.RECTANGLE, { x, y: 2.95, w: cw, h: 0.09, fill: { color: i === 0 ? C.gold : C.cyan }, line: { type: "none" } });
    s.addText(big, { x: x + 0.1, y: 3.35, w: cw - 0.2, h: 1.0, fontSize: big.length > 6 ? 26 : 40, bold: true, color: i === 0 ? C.gold : C.ink, fontFace: F.jp, align: "center", valign: "middle", margin: 0 });
    s.addText(sub, { x: x + 0.25, y: 4.45, w: cw - 0.5, h: 0.9, fontSize: 13, color: C.mute, fontFace: F.jp, align: "center", valign: "top", margin: 0, lineSpacingMultiple: 1.1 });
  });
  s.addText("「データは欠損する」前提に真っ向から立ち向かう、執念の仕様", {
    x: M, y: 5.75, w: W - 2 * M, h: 0.4, fontSize: 14, italic: true, color: C.mute2, fontFace: F.jp, margin: 0,
  });
  s.addNotes("数ワットの電波が地球の裏側へ。ノイズと減衰でデータは欠ける。それでも77bitでどうにか交信を成立させる。");
}

/* ---------------- 4. PROBLEM ---------------- */
{
  const s = pres.addSlide();
  base(s, "the problem", 4);
  heading(s, "普段のネット と 微弱電波");
  const cw = (W - 2 * M - 0.5) / 2;
  // 左: 普段
  card(s, M, 2.0, cw, 3.7, C.panel2);
  s.addText("普段の通信", { x: M + 0.3, y: 2.2, w: cw - 0.6, h: 0.5, fontSize: 20, bold: true, color: C.green, fontFace: F.jp, margin: 0 });
  s.addText([
    { text: "強靭な有線（光・LAN）で常時接続", options: { bullet: true, breakLine: true } },
    { text: "数KB〜数GBを難なく処理", options: { bullet: true, breakLine: true } },
    { text: "5Gも基地局から先は有線の世界", options: { bullet: true, breakLine: true } },
    { text: "データはまず欠けない前提", options: { bullet: true } },
  ], { x: M + 0.35, y: 2.85, w: cw - 0.7, h: 2.6, fontSize: 16, color: C.ink, fontFace: F.jp, paraSpaceAfter: 10, margin: 0 });
  // 右: 微弱電波
  const rx = M + cw + 0.5;
  card(s, rx, 2.0, cw, 3.7, C.panel2);
  s.addShape(pres.shapes.RECTANGLE, { x: rx, y: 2.0, w: cw, h: 0.09, fill: { color: C.red }, line: { type: "none" } });
  s.addText("微弱電波の世界", { x: rx + 0.3, y: 2.2, w: cw - 0.6, h: 0.5, fontSize: 20, bold: true, color: C.red, fontFace: F.jp, margin: 0 });
  s.addText([
    { text: "ケーブル無し。直接ぶつける", options: { bullet: true, breakLine: true } },
    { text: "距離で 減衰 → 信号が弱る", options: { bullet: true, breakLine: true } },
    { text: "他の電波の影響で かき消される", options: { bullet: true, breakLine: true } },
    { text: "データは 必ず欠ける", options: { bullet: true } },
  ], { x: rx + 0.35, y: 2.85, w: cw - 0.7, h: 2.6, fontSize: 16, color: C.ink, fontFace: F.jp, paraSpaceAfter: 10, margin: 0 });
  s.addText("→ この「欠ける前提」とどう戦うかが FT8 の全て", {
    x: M, y: 5.95, w: W - 2 * M, h: 0.4, fontSize: 15, italic: true, color: C.cyan, fontFace: F.jp, align: "center", margin: 0,
  });
  s.addNotes("山梨から地球の裏側へ直接。ケーブルは非現実的だから電波。でも電波は減衰と妨害。データは必ず欠ける、が出発点。");
}

/* ---------------- 5. COMPRESSION ---------------- */
{
  const s = pres.addSlide();
  base(s, "source encoding", 5);
  heading(s, "無線に最適化した圧縮術");
  s.addText("定型文（コールサイン・位置・信号レポート）だから、徹底的に圧縮できる", {
    x: M, y: 1.9, w: W - 2 * M, h: 0.4, fontSize: 16, color: C.mute, fontFace: F.jp, margin: 0,
  });
  // 表
  const rows = [
    [{ text: "情報", options: { bold: true, color: C.bg, fill: { color: C.cyan }, align: "left" } },
     { text: "中身", options: { bold: true, color: C.bg, fill: { color: C.cyan } } },
     { text: "ビット数", options: { bold: true, color: C.bg, fill: { color: C.cyan }, align: "center" } }],
    [{ text: "コールサイン", options: { color: C.ink } },
     { text: "標準は約2.7億通り (2²⁸)", options: { color: C.mute } },
     { text: "28 bit", options: { color: C.gold, bold: true, align: "center" } }],
    [{ text: "位置 (グリッド)", options: { color: C.ink } },
     { text: "4文字＝32,400通り (<2¹⁵)", options: { color: C.mute } },
     { text: "15 bit", options: { color: C.gold, bold: true, align: "center" } }],
  ];
  s.addTable(rows, {
    x: M, y: 2.5, w: 7.3, colW: [2.0, 3.7, 1.6], rowH: [0.5, 0.62, 0.62],
    fontFace: F.jp, fontSize: 15, valign: "middle", border: { type: "solid", pt: 1, color: C.line },
    fill: { color: C.panel }, align: "left",
  });
  // 例カード（右）
  const ex = M + 7.7;
  card(s, ex, 2.5, W - M - ex, 3.05);
  s.addText("例：この1メッセージ", { x: ex + 0.3, y: 2.7, w: 4, h: 0.4, fontSize: 14, color: C.mute, fontFace: F.jp, margin: 0 });
  s.addText("CQ JK1TUT PM95", { x: ex + 0.3, y: 3.1, w: W - M - ex - 0.6, h: 0.6, fontSize: 24, bold: true, color: C.cyan, fontFace: F.mono, margin: 0 });
  s.addText([
    { text: "UTF-8 (12文字)", options: { color: C.mute, breakLine: true } },
    { text: "96 bit", options: { color: C.red, bold: true, fontSize: 22, breakLine: true } },
    { text: "圧縮後 (実質)", options: { color: C.mute, breakLine: true } },
    { text: "43 bit ", options: { color: C.gold, bold: true, fontSize: 22 } },
    { text: "→ 固定77bit枠に余裕", options: { color: C.mute, fontSize: 13 } },
  ], { x: ex + 0.3, y: 3.85, w: W - M - ex - 0.6, h: 1.6, fontSize: 14, fontFace: F.jp, margin: 0, lineSpacingMultiple: 1.05 });
  s.addText("JSONで文字列を投げる日常からすると、ちょっと感動する", {
    x: M, y: 5.85, w: W - 2 * M, h: 0.4, fontSize: 14, italic: true, color: C.mute2, fontFace: F.jp, margin: 0,
  });
  s.addNotes("アマチュア無線の文脈。定型文だから圧縮できる。コールサイン28bit、位置15bit。96bit相当が43bitに。");
}

/* ---------------- 6. SECTION DIVIDER (LDPC) ---------------- */
{
  const s = pres.addSlide();
  s.background = { color: C.panel2 };
  waterfall(s, 0, H - 1.15, W, 1.15, 90);
  s.addText("ここが本題", { x: M, y: 1.3, w: 8, h: 0.45, fontSize: 18, color: C.cyan, fontFace: F.jp, bold: true, charSpacing: 2, margin: 0 });
  s.addText("データは必ず欠ける、\nそれでも直す。", {
    x: M, y: 1.92, w: W - 2 * M, h: 1.55, fontSize: 44, bold: true, color: C.ink, fontFace: F.jp, margin: 0, lineSpacingMultiple: 1.08,
  });
  s.addText("LDPC ── 低密度パリティ検査符号", {
    x: M, y: 3.7, w: W - 2 * M, h: 0.6, fontSize: 26, bold: true, color: C.gold, fontFace: F.jp, margin: 0,
  });
  s.addText("今日いちばん感動したところ", { x: M, y: 4.45, w: 8, h: 0.4, fontSize: 14, italic: true, color: C.mute2, fontFace: F.jp, margin: 0 });
  s.addNotes("ここからが山場。誤り訂正のLDPC。数学的なのに、解く実態は確率論的。そこに痺れた。");
}

/* ---------------- 7. TX SIDE ---------------- */
{
  const s = pres.addSlide();
  base(s, "LDPC · 送信側", 7);
  heading(s, "送信側：数学で「保険」をかける", { bar: C.gold });
  // ビットレイアウトバー
  const bx = M, by = 2.1, bw = W - 2 * M, bh = 0.95, total = 174;
  const segs = [["データ", 77, C.cyan], ["CRC", 14, C.gold], ["パリティ", 83, C.mute2]];
  let cx = bx;
  segs.forEach(([lab, n, col]) => {
    const sw = bw * n / total;
    s.addShape(pres.shapes.RECTANGLE, { x: cx, y: by, w: sw, h: bh, fill: { color: col }, line: { color: C.bg, width: 2 } });
    s.addText([{ text: lab + "\n", options: { fontSize: 14, bold: true } }, { text: n + " bit", options: { fontSize: 12 } }],
      { x: cx, y: by, w: sw, h: bh, align: "center", valign: "middle", color: C.bg, fontFace: F.jp, margin: 0 });
    cx += sw;
  });
  s.addText("合計 174 bit ＝ LDPC (174, 91) 符号の1ワード", { x: bx, y: by + bh + 0.12, w: bw, h: 0.35, align: "right", fontSize: 13, color: C.mute, fontFace: F.jp, margin: 0 });

  // 手順
  s.addText([
    { text: "① ", options: { color: C.gold, bold: true } },
    { text: "データ77bit に多項式 ", options: { color: C.ink } },
    { text: "0x6757", options: { color: C.cyan, fontFace: F.mono } },
    { text: " を当て 14bit CRC（誤り検査用チェックサム）を計算", options: { color: C.ink } },
  ], { x: M, y: 3.55, w: W - 2 * M, h: 0.4, fontSize: 16, fontFace: F.jp, margin: 0 });
  s.addText([
    { text: "② ", options: { color: C.gold, bold: true } },
    { text: "生成行列 G の 83 行それぞれを 91bit と mod2 内積 → 83bit パリティ（誤り訂正）", options: { color: C.ink } },
  ], { x: M, y: 4.05, w: W - 2 * M, h: 0.4, fontSize: 16, fontFace: F.jp, margin: 0 });

  // 数式カード
  card(s, M, 4.7, W - 2 * M, 1.55, C.panel);
  s.addText([
    { text: "parity(83) = G(83×91) · m(91)   (mod 2)", options: { color: C.cyan, breakLine: true, fontSize: 19, bold: true } },
    { text: "m = データ77 + CRC14 = 91 bit", options: { color: C.gold, fontSize: 18, bold: true } },
  ], { x: M + 0.4, y: 4.85, w: W - 2 * M - 0.8, h: 1.25, fontFace: F.mono, valign: "middle", margin: 0, lineSpacingMultiple: 1.3 });
  s.addText("全部 mod 2（XOR）の世界。数学的に隙のない「保険」をかけてから電波に乗せる", {
    x: M, y: 6.35, w: W - 2 * M, h: 0.35, fontSize: 13, italic: true, color: C.mute2, fontFace: F.jp, margin: 0,
  });
  s.addNotes("77にCRC14足して91。生成行列で83パリティ。合計174bit。肝は『正しい符号語はH·c=0』という厳密な制約。これが受信側の武器になる。");
}

/* ---------------- 8. RX SIDE ---------------- */
{
  const s = pres.addSlide();
  base(s, "LDPC · 受信側", 8);
  heading(s, "受信側：怪しいビットを「特定」する流れ", { bar: C.gold });
  s.addText([
    { text: "届くのは曖昧な値（軟判定）。", options: { color: C.ink } },
    { text: "検査式が 0 にならない＝矛盾", options: { color: C.cyan, bold: true } },
    { text: " → その矛盾から怪しいビットを絞り込む", options: { color: C.ink } },
  ], { x: M, y: 1.78, w: W - 2 * M, h: 0.4, fontSize: 15, fontFace: F.jp, margin: 0 });
  s.addText([
    { text: "検査式 ＝ 事前定義のスパースな ", options: { color: C.mute } },
    { text: "検査行列 H", options: { color: C.cyan, bold: true } },
    { text: " の各行。", options: { color: C.mute } },
    { text: "H × 受信ビット列 の各行がすべて 0 なら正常", options: { color: C.gold, bold: true } },
    { text: "（0 でない行に誤りの疑い）", options: { color: C.mute } },
  ], { x: M, y: 2.22, w: W - 2 * M, h: 0.35, fontSize: 12.5, fontFace: F.jp, margin: 0 });

  const cardY = 2.72, cardH = 3.25, arrowGap = 0.5;
  const cardW = (W - 2 * M - 2 * arrowGap) / 3;
  const xs = [M, M + cardW + arrowGap, M + 2 * (cardW + arrowGap)];
  const accents = [C.cyan, C.gold, C.cyan];
  const titles = ["① 軟判定ビット", "② 検査式で矛盾を探す", "③ 容疑者を特定"];
  xs.forEach((x, i) => {
    card(s, x, cardY, cardW, cardH);
    s.addShape(pres.shapes.RECTANGLE, { x, y: cardY, w: cardW, h: 0.08, fill: { color: accents[i] }, line: { type: "none" } });
    s.addText(titles[i], { x: x + 0.25, y: cardY + 0.18, w: cardW - 0.5, h: 0.4, fontSize: 15, bold: true, color: accents[i], fontFace: F.jp, margin: 0 });
  });
  // カード間の矢印
  [M + cardW + 0.04, M + 2 * cardW + arrowGap + 0.04].forEach((x) => {
    s.addText("→", { x, y: cardY + cardH / 2 - 0.35, w: arrowGap - 0.08, h: 0.7, align: "center", valign: "middle", fontSize: 28, bold: true, color: C.mute2, margin: 0 });
  });

  // カード1: 軟判定ビット
  const x1 = xs[0];
  const bits = [["1", C.cyan], ["0", C.cyan], ["?", C.gold], ["1", C.cyan], ["?", C.red], ["0", C.cyan]];
  const idx = ["c1", "c2", "c3", "c4", "c5", "c6"];
  const bw = 0.44, bg = 0.07, btot = bits.length * bw + (bits.length - 1) * bg;
  const bsx = x1 + (cardW - btot) / 2, bsy = cardY + 0.95;
  bits.forEach(([v, col], i) => {
    const x = bsx + i * (bw + bg);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: bsy, w: bw, h: 0.5, rectRadius: 0.05, fill: { color: C.panel }, line: { color: col, width: v === "?" ? 1.75 : 1 } });
    s.addText(v, { x, y: bsy, w: bw, h: 0.5, align: "center", valign: "middle", fontSize: 17, bold: true, color: col, fontFace: F.mono, margin: 0 });
    s.addText(idx[i], { x, y: bsy + 0.52, w: bw, h: 0.24, align: "center", fontSize: 9, color: C.mute2, fontFace: F.mono, margin: 0 });
  });
  s.addText([
    { text: "ノイズで 0/1 を断定できないビット ＝ ", options: { color: C.mute } },
    { text: "?", options: { color: C.gold, bold: true } },
    { text: "\n（この例では c3, c5）", options: { color: C.mute } },
  ], { x: x1 + 0.3, y: cardY + 2.1, w: cardW - 0.6, h: 1.0, fontSize: 12.5, fontFace: F.jp, margin: 0, lineSpacingMultiple: 1.2 });

  // カード2: 検査式で矛盾を探す
  const x2 = xs[1];
  const eqs = [
    ["A", "c1+c2+c4 → 0", "成立", C.green, false],
    ["B", "c2+c3+c5 → 1", "矛盾", C.red, true],
    ["C", "c3+c5+c6 → 1", "矛盾", C.red, true],
  ];
  eqs.forEach(([lab, eq, tag, col, bad], i) => {
    const y = cardY + 0.95 + i * 0.5;
    s.addText([
      { text: lab + "  ", options: { fontFace: F.mono, color: C.mute2, bold: true } },
      { text: eq + "   ", options: { fontFace: F.mono, color: col, bold: bad } },
      { text: tag, options: { fontFace: F.jp, color: col, bold: true, fontSize: 12 } },
    ], { x: x2 + 0.3, y, w: cardW - 0.5, h: 0.42, fontSize: 14, valign: "middle", margin: 0 });
  });
  s.addText("各行が 0 ＝ 正常 ／ 0 でない ＝ 誤りの疑い", { x: x2 + 0.3, y: cardY + 2.65, w: cardW - 0.6, h: 0.4, fontSize: 12, color: C.mute, fontFace: F.jp, margin: 0 });

  // カード3: 容疑者を特定
  const x3 = xs[2];
  s.addText([
    { text: "矛盾した ", options: { color: C.ink } },
    { text: "B・C に共通", options: { color: C.gold, bold: true } },
    { text: " するビット", options: { color: C.ink } },
  ], { x: x3 + 0.3, y: cardY + 0.95, w: cardW - 0.55, h: 0.4, fontSize: 13.5, fontFace: F.jp, margin: 0 });
  s.addText("→  c3   c5  が容疑者", { x: x3 + 0.3, y: cardY + 1.42, w: cardW - 0.55, h: 0.45, fontSize: 17, bold: true, color: C.cyan, fontFace: F.jp, margin: 0 });
  s.addText([
    { text: "自信のない方を反転して再計算", options: { bullet: true, breakLine: true } },
    { text: "全式が 0 になれば確定", options: { bullet: true } },
  ], { x: x3 + 0.35, y: cardY + 2.1, w: cardW - 0.6, h: 1.0, fontSize: 12.5, color: C.ink, fontFace: F.jp, paraSpaceAfter: 7, margin: 0 });

  // 下部: BP / OSD への接続
  s.addText([
    { text: "この「確からしさの更新 → 再計算」を高速に反復するのが ", options: { color: C.mute } },
    { text: "Belief Propagation", options: { color: C.cyan, bold: true } },
    { text: "（多くは数回で収束 / 弱すぎる時は重い OSD にフォールバック）", options: { color: C.mute } },
  ], { x: M, y: 6.15, w: W - 2 * M, h: 0.35, fontSize: 13, fontFace: F.jp, align: "center", margin: 0 });
  s.addText("※ ビット番号・式は説明用の簡略例（実際は 174bit・83式）", { x: M, y: 6.55, w: W - 2 * M, h: 0.3, fontSize: 9.5, color: C.mute2, fontFace: F.jp, align: "center", margin: 0 });

  s.addNotes("流れ：①受信ビットは軟判定で一部が怪しい→②全検査式を計算、0にならない式が矛盾→③矛盾する式に共通するビットが容疑者。自信のない方を反転して全式0なら確定。これを確率で高速反復するのがBP、ダメならOSD。数字は説明用の例。");
}

/* ---------------- 9. PUNCHLINE ---------------- */
{
  const s = pres.addSlide();
  s.background = { color: C.bg };
  base(s, "the punchline", 9);
  s.addShape(pres.shapes.RECTANGLE, { x: M, y: 2.2, w: 0.12, h: 2.6, fill: { color: C.cyan }, line: { type: "none" } });
  s.addText([
    { text: "仕組みは ", options: { color: C.ink } },
    { text: "厳密な線形代数", options: { color: C.gold, bold: true } },
    { text: "。\n解く過程は ", options: { color: C.ink } },
    { text: "確率論的な交渉", options: { color: C.cyan, bold: true } },
    { text: "。", options: { color: C.ink } },
  ], { x: M + 0.45, y: 2.2, w: W - 2 * M - 0.45, h: 1.8, fontSize: 40, bold: true, fontFace: F.jp, valign: "middle", margin: 0, lineSpacingMultiple: 1.15 });
  s.addText("数学で受け止めて、確率で押し返す。", {
    x: M + 0.45, y: 4.15, w: W - 2 * M, h: 0.6, fontSize: 22, color: C.ink, fontFace: F.jp, margin: 0,
  });
  s.addText([
    { text: "最後の番人：", options: { color: C.mute } },
    { text: "復元した77bitから CRC を再計算 → 一致すれば確信して表示、ズレたら捨てる", options: { color: C.mute } },
  ], { x: M + 0.45, y: 5.15, w: W - 2 * M - 0.45, h: 0.5, fontSize: 15, fontFace: F.jp, margin: 0 });
  s.addNotes("ここが感動ポイント。厳密な数学を確率で解く。確率で殴って、最後はCRCで答え合わせする二段構え。");
}

/* ---------------- 10. 15s CYCLE ---------------- */
{
  const s = pres.addSlide();
  base(s, "timing", 10);
  heading(s, "15秒に詰め込む「絶対に通信する」意思");
  // タイムライン 0-60s, 4窓
  const tx = M, ty = 2.5, tw = W - 2 * M, th = 0.9;
  const labels = ["送信", "受信", "送信", "受信"];
  for (let i = 0; i < 4; i++) {
    const x = tx + i * tw / 4;
    s.addShape(pres.shapes.RECTANGLE, { x, y: ty, w: tw / 4 - 0.06, h: th, fill: { color: i % 2 === 0 ? C.cyan : C.panel }, line: { color: C.line, width: 1 } });
    s.addText(labels[i], { x, y: ty, w: tw / 4 - 0.06, h: th, align: "center", valign: "middle", fontSize: 15, bold: true, color: i % 2 === 0 ? C.bg : C.mute, fontFace: F.jp, margin: 0 });
  }
  ["0", "15", "30", "45", "60s"].forEach((t, i) => {
    s.addText(t, { x: tx + i * tw / 4 - 0.3, y: ty + th + 0.05, w: 0.6, h: 0.3, align: "center", fontSize: 11, color: C.mute2, fontFace: F.mono, margin: 0 });
  });

  // 内訳カード
  const cw = (W - 2 * M - 0.4) / 2;
  card(s, M, 4.0, cw, 1.9);
  s.addText([
    { text: "実送信 ≈ 12.6 秒", options: { color: C.gold, bold: true, fontSize: 24, breakLine: true } },
    { text: "0.16s × 79シンボル", options: { color: C.mute, fontFace: F.mono, fontSize: 14, breakLine: true } },
    { text: "残り ≈ 2.4秒 は時刻同期・デコードの余白（NTP前提）", options: { color: C.ink, fontSize: 13 } },
  ], { x: M + 0.35, y: 4.25, w: cw - 0.7, h: 1.5, fontFace: F.jp, margin: 0, lineSpacingMultiple: 1.1 });

  card(s, M + cw + 0.4, 4.0, cw, 1.9);
  s.addText([
    { text: "1回ダメでも、15秒ごとに再挑戦", options: { color: C.cyan, bold: true, fontSize: 19, breakLine: true } },
    { text: "確率的な復元を、繰り返しで底上げする作戦。", options: { color: C.ink, fontSize: 15 } },
  ], { x: M + cw + 0.75, y: 4.45, w: cw - 0.7, h: 1.2, fontFace: F.jp, margin: 0, lineSpacingMultiple: 1.15 });
  s.addNotes("GMTの15秒サイクルで送受切替。実送信は12.6秒、残りは同期マージン。だから全員NTP同期。1回失敗しても繰り返して確率を上げる。");
}

/* ---------------- 11. STATE MACHINE ---------------- */
{
  const s = pres.addSlide();
  base(s, "state machine", 11);
  heading(s, "会話の文脈を、機械が状態で持つ");
  // 6状態フロー
  const states = ["Calling", "Replying", "Report", "Roger Rpt", "Rogers", "Signoff"];
  const n = states.length, gap = 0.32;
  const bw = (W - 2 * M - (n - 1) * gap) / n, by = 2.15, bh = 0.85;
  states.forEach((st, i) => {
    const x = M + i * (bw + gap);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: by, w: bw, h: bh, rectRadius: 0.06, fill: { color: i === 0 ? C.cyan : C.panel }, line: { color: C.line, width: 1 } });
    s.addText(st, { x, y: by, w: bw, h: bh, align: "center", valign: "middle", fontSize: 12.5, bold: true, color: i === 0 ? C.bg : C.ink, fontFace: F.mono, margin: 0 });
    if (i < n - 1) s.addText("›", { x: x + bw, y: by, w: gap, h: bh, align: "center", valign: "middle", fontSize: 18, color: C.mute2, margin: 0 });
  });
  s.addText("CQ → 応答 → 信号レポート → 受領 → ラジャー → 73(さよなら) を WSJT-X が管理（仕様書にもUML）", {
    x: M, y: by + bh + 0.15, w: W - 2 * M, h: 0.35, fontSize: 13, color: C.mute, fontFace: F.jp, align: "center", margin: 0,
  });
  // a priori 削減
  card(s, M, 4.05, W - 2 * M, 1.65, C.panel);
  s.addText("状態が分かる → 次のメッセージ型が読める（a priori）→ 探索すべきビットが激減", {
    x: M + 0.4, y: 4.2, w: W - 2 * M - 0.8, h: 0.4, fontSize: 14, color: C.mute, fontFace: F.jp, margin: 0,
  });
  s.addText([
    { text: "未知 77 bit", options: { color: C.red, bold: true, fontSize: 26 } },
    { text: "   →   ", options: { color: C.mute2, fontSize: 22 } },
    { text: "残り 15 bit", options: { color: C.cyan, bold: true, fontSize: 26 } },
    { text: "    （探索 約80%減）", options: { color: C.gold, bold: true, fontSize: 20 } },
  ], { x: M + 0.4, y: 4.7, w: W - 2 * M - 0.8, h: 0.8, fontFace: F.mono, valign: "middle", align: "center", margin: 0 });
  s.addText("ちなみに「完全ロボット運用」は禁止 ── 交信の起点は必ず人間", {
    x: M, y: 5.95, w: W - 2 * M, h: 0.4, fontSize: 13, italic: true, color: C.mute2, fontFace: F.jp, margin: 0,
  });
  s.addNotes("交信は決まった型＝ステートマシン。次の型が読めるからa prioriが効く。77bit総当たりが15bitに。これがさっきの80%削減の正体。なお起点は必ず人間。");
}

/* ---------------- 12. SUMMARY + CLOSING ---------------- */
{
  const s = pres.addSlide();
  s.background = { color: C.panel2 };
  base(s, "wrap up", 12);
  heading(s, "「欠ける前提」への、多層的な最適化");
  const items = [
    ["① 数学を確率で解く", "LDPC × Belief Propagation", C.gold],
    ["② 繰り返しで取り返す", "15秒サイクルで確率を底上げ", C.cyan],
    ["③ 文脈を状態で持つ", "ステートマシンで探索空間を削る", C.green],
  ];
  const cw = (W - 2 * M - 2 * 0.4) / 3;
  items.forEach(([t, sub, col], i) => {
    const x = M + i * (cw + 0.4);
    card(s, x, 2.0, cw, 1.7);
    s.addShape(pres.shapes.RECTANGLE, { x, y: 2.0, w: cw, h: 0.09, fill: { color: col }, line: { type: "none" } });
    s.addText(t, { x: x + 0.25, y: 2.3, w: cw - 0.5, h: 0.6, fontSize: 17, bold: true, color: col, fontFace: F.jp, margin: 0 });
    s.addText(sub, { x: x + 0.25, y: 2.95, w: cw - 0.5, h: 0.6, fontSize: 13, color: C.mute, fontFace: F.jp, margin: 0 });
  });
  s.addText("レイヤーは違うのに、全部が「微弱電波でどうにか通信する」一点に向いている。", {
    x: M, y: 3.95, w: W - 2 * M, h: 0.4, fontSize: 16, italic: true, color: C.ink, fontFace: F.jp, align: "center", margin: 0,
  });
  // 余談カード
  card(s, M, 4.6, W - 2 * M, 1.35, C.panel);
  s.addText([
    { text: "余談：", options: { color: C.gold, bold: true } },
    { text: "開発者 Joe Taylor (K1JT) は ", options: { color: C.ink } },
    { text: "宇宙の信号から微弱電波を検出するスペシャリスト", options: { color: C.cyan, bold: true } },
    { text: "（ノーベル物理学賞）。", options: { color: C.ink } },
    { text: "そして LDPC は今や 5G・WiFi・SSD でも当たり前に使われている。", options: { color: C.mute } },
  ], { x: M + 0.35, y: 4.75, w: W - 2 * M - 0.7, h: 1.05, fontSize: 14.5, fontFace: F.jp, valign: "middle", margin: 0, lineSpacingMultiple: 1.15 });
  s.addText("🍺 興味が湧いたら後で話しかけてください ── @numa08", {
    x: M, y: 6.15, w: W - 2 * M, h: 0.45, fontSize: 16, bold: true, color: C.ink, fontFace: F.jp, align: "center", margin: 0,
  });
  s.addNotes("3つの最適化が全部『微弱電波でどうにか』に収束。開発者K1JTは宇宙の微弱信号のスペシャリストでノーベル賞。LDPCは5G/WiFi/SSDにも。続きは後で、ビール片手に。");
}

pres.writeFile({ fileName: outPath }).then((f) => console.log("written:", f));
