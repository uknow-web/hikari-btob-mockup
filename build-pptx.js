// Generate HIKARI BtoB EC proposal PPTX
const pptxgen = require("pptxgenjs");

const COLOR = {
  navy: "0E2A4A",
  navyDark: "091F37",
  navyLight: "1A4678",
  orange: "C9731C",
  orangeDark: "9E5510",
  white: "FFFFFF",
  bg: "F6F7F9",
  border: "D9DEE5",
  text: "1A1A1A",
  muted: "666666",
  mutedLight: "8B95A1",
  accent: "7FD1A9",
  warn: "C83C3C",
};

const FONT_BODY = "Meiryo";
const FONT_HEADER = "Meiryo";

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.3 x 7.5
pres.author = "HIKARI / uknow-web";
pres.title = "HIKARI BtoB EC Proposal";
pres.subject = "makeshopを用いたBtoB EC構築提案書";

const W = 13.3;
const H = 7.5;

// ====== SLIDE MASTER: CONTENT ======
pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: COLOR.white },
  objects: [
    // Top navy bar
    { rect: { x: 0, y: 0, w: W, h: 0.12, fill: { color: COLOR.navy } } },
    // Bottom footer
    { rect: { x: 0, y: H - 0.35, w: W, h: 0.35, fill: { color: COLOR.bg } } },
    { text: {
        text: "HIKARI BtoB EC 構築提案書 ｜ makeshop ベース",
        options: { x: 0.5, y: H - 0.32, w: 8, h: 0.3, fontSize: 9, color: COLOR.muted, fontFace: FONT_BODY, valign: "middle" }
    }},
    { text: {
        text: "uknow-web",
        options: { x: W - 2.5, y: H - 0.32, w: 2, h: 0.3, fontSize: 9, color: COLOR.muted, fontFace: FONT_BODY, align: "right", valign: "middle" }
    }},
  ],
  slideNumber: { x: W - 0.6, y: H - 0.32, w: 0.4, h: 0.3, fontSize: 9, color: COLOR.muted, align: "right", fontFace: FONT_BODY }
});

// ====== HELPERS ======
function addTitle(slide, num, title, sub) {
  // Orange accent block
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 0.5, w: 0.1, h: 0.7, fill: { color: COLOR.orange }, line: { type: "none" }
  });
  // Number
  slide.addText(String(num).padStart(2, "0"), {
    x: 0.75, y: 0.42, w: 1.2, h: 0.5, fontSize: 32, bold: true, color: COLOR.orange, fontFace: FONT_HEADER, margin: 0
  });
  // Title
  slide.addText(title, {
    x: 1.8, y: 0.45, w: 10.5, h: 0.6, fontSize: 26, bold: true, color: COLOR.navy, fontFace: FONT_HEADER, valign: "middle", margin: 0
  });
  // Subtitle
  if (sub) {
    slide.addText(sub, {
      x: 1.8, y: 1.02, w: 10.5, h: 0.3, fontSize: 11, color: COLOR.muted, fontFace: FONT_BODY, margin: 0
    });
  }
  // Divider
  slide.addShape(pres.shapes.LINE, {
    x: 0.5, y: 1.45, w: W - 1.0, h: 0, line: { color: COLOR.border, width: 0.75 }
  });
}

function bullets(slide, items, opt = {}) {
  const arr = items.map((t, i) => ({
    text: t,
    options: { bullet: { code: "25A0" }, breakLine: i < items.length - 1, paraSpaceAfter: 6 }
  }));
  slide.addText(arr, {
    x: opt.x || 0.7, y: opt.y || 1.7, w: opt.w || 12, h: opt.h || 5,
    fontSize: opt.fontSize || 13, color: COLOR.text, fontFace: FONT_BODY, valign: "top"
  });
}

function section(slide, x, y, w, h, title, body, opts = {}) {
  slide.addShape(pres.shapes.RECTANGLE, {
    x, y, w, h, fill: { color: opts.bg || COLOR.bg }, line: { color: COLOR.border, width: 0.5 }
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x, y, w: 0.08, h, fill: { color: opts.accent || COLOR.orange }, line: { type: "none" }
  });
  slide.addText(title, {
    x: x + 0.25, y: y + 0.12, w: w - 0.3, h: 0.35, fontSize: 13, bold: true, color: COLOR.navy, fontFace: FONT_HEADER, margin: 0
  });
  if (Array.isArray(body)) {
    slide.addText(body.map((t, i) => ({
      text: t, options: { bullet: { code: "25A0" }, breakLine: i < body.length - 1, paraSpaceAfter: 4 }
    })), {
      x: x + 0.25, y: y + 0.55, w: w - 0.4, h: h - 0.7,
      fontSize: 11, color: COLOR.text, fontFace: FONT_BODY, valign: "top", margin: 0
    });
  } else {
    slide.addText(body, {
      x: x + 0.25, y: y + 0.5, w: w - 0.4, h: h - 0.6,
      fontSize: 11, color: COLOR.text, fontFace: FONT_BODY, valign: "top", margin: 0
    });
  }
}

// ====== 01. TITLE SLIDE ======
{
  const s = pres.addSlide();
  s.background = { color: COLOR.navy };

  // Orange diagonal accent
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.3, h: H, fill: { color: COLOR.orange }, line: { type: "none" } });

  // HIKARI mark (simulated with text since logo access in pptxgenjs requires file)
  s.addImage({ path: "images/logo-hikari.png", x: 1, y: 1, w: 1.0, h: 1.0, transparency: 0 });

  s.addText("BtoB EC サイト構築提案書", {
    x: 1, y: 2.8, w: 11, h: 0.8, fontSize: 40, bold: true, color: COLOR.white, fontFace: FONT_HEADER, margin: 0
  });
  s.addText("製造業ブランド「HIKARI」／ makeshop を用いた現実的な立ち上げ案", {
    x: 1, y: 3.7, w: 11, h: 0.5, fontSize: 18, color: "CADCFC", fontFace: FONT_BODY, margin: 0
  });

  // Badges
  s.addShape(pres.shapes.RECTANGLE, { x: 1, y: 4.8, w: 3.2, h: 0.5, fill: { color: COLOR.orange }, line: { type: "none" } });
  s.addText("初期公開 10 SKU", { x: 1, y: 4.8, w: 3.2, h: 0.5, fontSize: 14, bold: true, color: COLOR.white, align: "center", valign: "middle", fontFace: FONT_HEADER, margin: 0 });

  s.addShape(pres.shapes.RECTANGLE, { x: 4.3, y: 4.8, w: 3.2, h: 0.5, fill: { color: COLOR.navyLight }, line: { color: "3A6BA1", width: 0.5 } });
  s.addText("makeshop 構築", { x: 4.3, y: 4.8, w: 3.2, h: 0.5, fontSize: 14, bold: true, color: COLOR.white, align: "center", valign: "middle", fontFace: FONT_HEADER, margin: 0 });

  s.addShape(pres.shapes.RECTANGLE, { x: 7.6, y: 4.8, w: 3.2, h: 0.5, fill: { color: COLOR.navyLight }, line: { color: "3A6BA1", width: 0.5 } });
  s.addText("モック動作検証済", { x: 7.6, y: 4.8, w: 3.2, h: 0.5, fontSize: 14, bold: true, color: COLOR.white, align: "center", valign: "middle", fontFace: FONT_HEADER, margin: 0 });

  s.addText("モックアップ： https://hikari-btob-mockup.vercel.app/", {
    x: 1, y: 6.2, w: 11, h: 0.3, fontSize: 12, color: "CADCFC", fontFace: FONT_BODY, margin: 0
  });
  s.addText("提案：uknow-web ｜ 2026年4月", {
    x: 1, y: 6.5, w: 11, h: 0.3, fontSize: 12, color: "8EA4BD", fontFace: FONT_BODY, margin: 0
  });
}

// ====== 02. 目次 ======
{
  const s = pres.addSlide({ masterName: "CONTENT" });
  addTitle(s, "", "目次 ｜ Contents", "本提案書の構成");
  // Remove "00" number for TOC
  s.addText("", { x: 0.75, y: 0.42, w: 1.2, h: 0.5 }); // not reliable - skip

  const items = [
    ["01", "プロジェクト概要"], ["02", "背景と目的"],
    ["03", "現状サイトの課題"], ["04", "参考サイトの学び"],
    ["05", "makeshop 採用理由"], ["06", "想定ターゲット"],
    ["07", "サイトコンセプト"], ["08", "サイトマップ"],
    ["09", "初期公開 10 SKU"], ["10", "必要機能一覧"],
    ["11", "BtoB 導線設計"], ["12", "デザイン方針"],
    ["13", "コンテンツ方針"], ["14", "商品登録方針"],
    ["15", "運用フロー"], ["16", "フェーズ設計"],
    ["17", "制作スケジュール"], ["18", "リスクと注意点"],
    ["19", "まとめ / 次のアクション"], ["—", "（参考）モック画面"]
  ];
  const colW = 5.8;
  const rowH = 0.42;
  items.forEach((it, i) => {
    const col = Math.floor(i / 10);
    const row = i % 10;
    const x = 0.8 + col * (colW + 0.5);
    const y = 1.8 + row * rowH;
    s.addText(it[0], { x, y, w: 0.7, h: rowH, fontSize: 12, bold: true, color: COLOR.orange, fontFace: FONT_HEADER, valign: "middle", margin: 0 });
    s.addText(it[1], { x: x + 0.7, y, w: colW - 0.7, h: rowH, fontSize: 13, color: COLOR.text, fontFace: FONT_BODY, valign: "middle", margin: 0 });
  });
}

// ====== 03. プロジェクト概要 ======
{
  const s = pres.addSlide({ masterName: "CONTENT" });
  addTitle(s, 1, "プロジェクト概要", "Project Overview");

  // Left: key points
  section(s, 0.5, 1.7, 6.1, 5.2, "本プロジェクトで実現すること", [
    "HIKARIブランドの業務用什器を、法人向けに直販するBtoB EC",
    "makeshop の標準機能を中心にスピーディに立ち上げ",
    "規格品は即時受注、特注品は見積導線で確実に接続",
    "会員価格・再注文・注文履歴で継続取引を支援",
    "製造業らしい信頼感のあるデザイン・UX",
    "担当1〜2名で回る運用体制を前提"
  ]);

  // Right: numeric cards
  const cards = [
    { label: "初期公開 SKU", value: "10", suffix: "型番" },
    { label: "構築期間", value: "3-4", suffix: "ヶ月" },
    { label: "運用体制", value: "1-2", suffix: "名" },
    { label: "BtoB導線", value: "5", suffix: "系統" },
  ];
  cards.forEach((c, i) => {
    const row = Math.floor(i / 2);
    const col = i % 2;
    const x = 6.9 + col * 3.1;
    const y = 1.7 + row * 2.6;
    s.addShape(pres.shapes.RECTANGLE, { x, y, w: 2.95, h: 2.4, fill: { color: COLOR.navy }, line: { type: "none" } });
    s.addShape(pres.shapes.RECTANGLE, { x, y, w: 2.95, h: 0.08, fill: { color: COLOR.orange }, line: { type: "none" } });
    s.addText(c.label, { x: x + 0.2, y: y + 0.3, w: 2.55, h: 0.3, fontSize: 11, color: "CADCFC", fontFace: FONT_BODY, margin: 0 });
    s.addText(c.value, { x: x + 0.2, y: y + 0.65, w: 2.55, h: 1.3, fontSize: 60, bold: true, color: COLOR.white, fontFace: FONT_HEADER, margin: 0 });
    s.addText(c.suffix, { x: x + 0.2, y: y + 1.85, w: 2.55, h: 0.35, fontSize: 13, color: "CADCFC", fontFace: FONT_BODY, margin: 0 });
  });
}

// ====== 04. 背景と目的 ======
{
  const s = pres.addSlide({ masterName: "CONTENT" });
  addTitle(s, 2, "背景と目的", "Background & Objectives");

  section(s, 0.5, 1.7, 6.1, 5.2, "現状の背景", [
    "現行ブランドサイトは製品紹介どまりで、見積・発注導線がない",
    "法人顧客から「型番で注文」「見積書が欲しい」「再注文を楽に」というニーズ",
    "受注対応がメール／電話に依存、属人化が発生",
    "商品追加・価格改定の運用フローが未整理"
  ]);
  section(s, 6.8, 1.7, 6.0, 5.2, "本プロジェクトの目的", [
    "規格品をオンラインで即時受注できる状態をつくる",
    "特注・大口は見積／問い合わせ導線へ確実に接続",
    "会員価格・再注文・注文履歴の仕組みを整備",
    "製造業らしい信頼感のあるブランド体験",
    "担当1〜2名で回る運用体制の構築"
  ], { accent: COLOR.navyLight });
}

// ====== 05. 現状サイトの課題 ======
{
  const s = pres.addSlide({ masterName: "CONTENT" });
  addTitle(s, 3, "現状サイトの課題整理", "Current Site Issues");

  const headers = ["観点", "現状", "課題"];
  const rows = [
    ["情報設計", "1枚構成の製品紹介", "商品比較・型番選定ができない"],
    ["購入導線", "なし（紹介のみ）", "問い合わせのみで商機を取りこぼし"],
    ["法人対応", "未整備", "会員価格／請求書払い未対応"],
    ["商品管理", "固定HTML", "SKU追加のたびに更新破綻の恐れ"],
    ["SEO集客", "製品単位ページが弱い", "型番・用途キーワードで流入しない"],
    ["運用", "属人的", "価格改定・商品追加のフローなし"]
  ];
  const tableData = [
    headers.map(h => ({ text: h, options: { fill: { color: COLOR.navy }, color: COLOR.white, bold: true, fontFace: FONT_HEADER, fontSize: 12, align: "center", valign: "middle" } })),
    ...rows.map(r => r.map((c, i) => ({
      text: c,
      options: {
        fill: { color: i === 0 ? COLOR.bg : COLOR.white },
        color: COLOR.text, fontFace: FONT_BODY, fontSize: 12,
        bold: i === 0, valign: "middle"
      }
    })))
  ];
  s.addTable(tableData, {
    x: 0.7, y: 1.85, w: 11.9, colW: [2.1, 4.4, 5.4],
    border: { pt: 0.5, color: COLOR.border },
    rowH: 0.55
  });
}

// ====== 06. 参考サイトの学び ======
{
  const s = pres.addSlide({ masterName: "CONTENT" });
  addTitle(s, 4, "参考サイトから学ぶべきポイント", "steel-labo-shopping.com を参考に");

  const items = [
    { n: "01", t: "多層的なカテゴリ", d: "素材 × 用途 × 形状の3軸で絞り込み可能" },
    { n: "02", t: "商品ページの仕様表", d: "サイズ・耐荷重・素材・規格を明快に表示" },
    { n: "03", t: "3つの相談導線の常時表示", d: "お見積り・お問い合わせ・カタログ請求" },
    { n: "04", t: "型番検索に強い構成", d: "BtoBはカテゴリ回遊より型番直行が多い" }
  ];
  items.forEach((it, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = 0.5 + col * 6.2;
    const y = 1.75 + row * 2.65;
    s.addShape(pres.shapes.RECTANGLE, { x, y, w: 6.0, h: 2.4, fill: { color: COLOR.white }, line: { color: COLOR.border, width: 0.5 } });
    s.addShape(pres.shapes.RECTANGLE, { x, y, w: 0.08, h: 2.4, fill: { color: COLOR.orange }, line: { type: "none" } });
    s.addText(it.n, { x: x + 0.3, y: y + 0.25, w: 1.0, h: 0.6, fontSize: 32, bold: true, color: COLOR.orange, fontFace: FONT_HEADER, margin: 0 });
    s.addText(it.t, { x: x + 0.3, y: y + 0.95, w: 5.5, h: 0.5, fontSize: 17, bold: true, color: COLOR.navy, fontFace: FONT_HEADER, margin: 0 });
    s.addText(it.d, { x: x + 0.3, y: y + 1.5, w: 5.5, h: 0.8, fontSize: 12, color: COLOR.text, fontFace: FONT_BODY, margin: 0 });
  });
  s.addText("→ 本提案モックは、これら4要素をmakeshop標準機能＋デザイン調整で再現済みです。", {
    x: 0.5, y: 6.75, w: 12, h: 0.35, fontSize: 12, italic: true, color: COLOR.navyLight, fontFace: FONT_BODY, margin: 0
  });
}

// ====== 07. makeshop 採用理由 ======
{
  const s = pres.addSlide({ masterName: "CONTENT" });
  addTitle(s, 5, "makeshop 採用理由", "Why makeshop?");

  section(s, 0.5, 1.7, 6.1, 2.5, "採用根拠", [
    "BtoB機能が標準装備（会員ランク別価格／見積／請求書払い）",
    "初期・月額コストがフルスクラッチの1/5〜1/10",
    "CSV一括登録でSKU拡張に強い",
    "管理画面の学習コストが低く、クライアント自走可能"
  ], { accent: COLOR.orange });

  section(s, 6.8, 1.7, 6.0, 2.5, "できること（モック実装済み）", [
    "会員ランク別価格／非会員は会員価格を伏せる",
    "見積依頼フォーム（複数商品まとめて一括）",
    "注文履歴からのワンクリック再注文",
    "商品CSV一括更新"
  ], { accent: COLOR.navyLight });

  // Lower: what needs workaround
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.4, w: 12.3, h: 2.6, fill: { color: "FFF4E5" }, line: { color: "E8CBA2", width: 0.5 } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.4, w: 0.08, h: 2.6, fill: { color: COLOR.orange }, line: { type: "none" } });
  s.addText("工夫が必要なこと", {
    x: 0.75, y: 4.55, w: 11, h: 0.35, fontSize: 13, bold: true, color: COLOR.orangeDark, fontFace: FONT_HEADER, margin: 0
  });
  const tbl = [
    [{ text: "項目", options: { bold: true, fill: { color: COLOR.white }, color: COLOR.navy, fontSize: 11, fontFace: FONT_HEADER, valign: "middle" } },
     { text: "対応方針", options: { bold: true, fill: { color: COLOR.white }, color: COLOR.navy, fontSize: 11, fontFace: FONT_HEADER, valign: "middle" } }],
    [{ text: "複雑な承認ワークフロー（部長承認→発注）", options: { fontSize: 11, fontFace: FONT_BODY, valign: "middle" } },
     { text: "初期はメール承認で代替", options: { fontSize: 11, fontFace: FONT_BODY, valign: "middle" } }],
    [{ text: "取引先ごとの完全個別単価", options: { fontSize: 11, fontFace: FONT_BODY, valign: "middle" } },
     { text: "会員ランクを活用、細分化は見積へ", options: { fontSize: 11, fontFace: FONT_BODY, valign: "middle" } }],
    [{ text: "特注品の自動見積", options: { fontSize: 11, fontFace: FONT_BODY, valign: "middle" } },
     { text: "自動化せず、問い合わせフォームに寄せる", options: { fontSize: 11, fontFace: FONT_BODY, valign: "middle" } }],
    [{ text: "基幹システム連携", options: { fontSize: 11, fontFace: FONT_BODY, valign: "middle" } },
     { text: "初期は手動CSV、拡張フェーズでAPI検討", options: { fontSize: 11, fontFace: FONT_BODY, valign: "middle" } }]
  ];
  s.addTable(tbl, { x: 0.75, y: 4.95, w: 11.85, colW: [5.3, 6.55], border: { pt: 0.5, color: COLOR.border }, rowH: 0.38 });
}

// ====== 08. 想定ターゲット ======
{
  const s = pres.addSlide({ masterName: "CONTENT" });
  addTitle(s, 6, "想定ターゲット（法人顧客像）", "Target Customers");

  const targets = [
    { icon: "🏭", t: "食品工場／セントラルキッチン", d: "ステンレスワゴン、衛生什器" },
    { icon: "🏥", t: "医療・クリニック・介護", d: "医療用ステンレスワゴン" },
    { icon: "📦", t: "物流・倉庫", d: "スチール棚（軽量／中量／重量）" },
    { icon: "🍴", t: "飲食・ホテル", d: "バックヤード什器、厨房ワゴン" },
    { icon: "🧪", t: "研究機関・実験室", d: "小型ワゴン、作業台" }
  ];
  targets.forEach((tg, i) => {
    const x = 0.5 + (i % 5) * 2.5;
    const y = 1.75;
    s.addShape(pres.shapes.RECTANGLE, { x, y, w: 2.35, h: 2.6, fill: { color: COLOR.bg }, line: { color: COLOR.border, width: 0.5 } });
    s.addText(tg.icon, { x, y: y + 0.15, w: 2.35, h: 0.8, fontSize: 40, align: "center", fontFace: FONT_BODY, margin: 0 });
    s.addText(tg.t, { x: x + 0.1, y: y + 1.05, w: 2.15, h: 0.8, fontSize: 12, bold: true, color: COLOR.navy, align: "center", fontFace: FONT_HEADER, margin: 0 });
    s.addText(tg.d, { x: x + 0.1, y: y + 1.85, w: 2.15, h: 0.6, fontSize: 10, color: COLOR.muted, align: "center", fontFace: FONT_BODY, margin: 0 });
  });

  // Persona
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.6, w: 12.3, h: 2.4, fill: { color: COLOR.navy }, line: { type: "none" } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.6, w: 0.08, h: 2.4, fill: { color: COLOR.orange }, line: { type: "none" } });
  s.addText("購買担当者のペルソナ", { x: 0.75, y: 4.75, w: 11, h: 0.35, fontSize: 14, bold: true, color: COLOR.white, fontFace: FONT_HEADER, margin: 0 });
  s.addText([
    { text: "総務・購買・現場責任者（30〜50代）", options: { bullet: { code: "25A0" }, breakLine: true, paraSpaceAfter: 4, color: "CADCFC" } },
    { text: "「型番で探す」「納期が読める」「見積が取れる」を重視", options: { bullet: { code: "25A0" }, breakLine: true, paraSpaceAfter: 4, color: "CADCFC" } },
    { text: "PC業務中に検索、相見積り前提で複数サイトを回遊", options: { bullet: { code: "25A0" }, paraSpaceAfter: 4, color: "CADCFC" } }
  ], { x: 0.75, y: 5.2, w: 11.8, h: 1.6, fontSize: 12, fontFace: FONT_BODY, valign: "top", margin: 0 });
}

// ====== 09. サイトコンセプト ======
{
  const s = pres.addSlide({ masterName: "CONTENT" });
  addTitle(s, 7, "サイトコンセプト", "Site Concept");

  // Big concept
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.7, w: 12.3, h: 2.0, fill: { color: COLOR.navy }, line: { type: "none" } });
  s.addText("現場で使う什器を、型番から最短で発注できる。", {
    x: 0.5, y: 1.9, w: 12.3, h: 0.7, fontSize: 28, bold: true, color: COLOR.white, align: "center", fontFace: FONT_HEADER, margin: 0
  });
  s.addText("相談もすぐできる。", {
    x: 0.5, y: 2.6, w: 12.3, h: 0.7, fontSize: 28, bold: true, color: COLOR.orange, align: "center", fontFace: FONT_HEADER, margin: 0
  });
  s.addText("— 3 Design Principles —", {
    x: 0.5, y: 3.3, w: 12.3, h: 0.3, fontSize: 11, color: "CADCFC", align: "center", italic: true, fontFace: FONT_BODY, margin: 0
  });

  const principles = [
    { n: "01", t: "規格品はすぐ買える", d: "カート決済でワンストップ完結" },
    { n: "02", t: "迷ったら相談できる", d: "見積・問い合わせ・特注相談の3導線" },
    { n: "03", t: "一度使ったらリピート", d: "会員価格と再注文でLTVを高める" }
  ];
  principles.forEach((p, i) => {
    const x = 0.5 + i * 4.1;
    const y = 4.1;
    s.addShape(pres.shapes.RECTANGLE, { x, y, w: 4.0, h: 2.8, fill: { color: COLOR.white }, line: { color: COLOR.border, width: 0.5 } });
    s.addShape(pres.shapes.OVAL, { x: x + 0.3, y: y + 0.3, w: 0.9, h: 0.9, fill: { color: COLOR.orange }, line: { type: "none" } });
    s.addText(p.n, { x: x + 0.3, y: y + 0.3, w: 0.9, h: 0.9, fontSize: 18, bold: true, color: COLOR.white, align: "center", valign: "middle", fontFace: FONT_HEADER, margin: 0 });
    s.addText(p.t, { x: x + 0.3, y: y + 1.3, w: 3.5, h: 0.5, fontSize: 17, bold: true, color: COLOR.navy, fontFace: FONT_HEADER, margin: 0 });
    s.addText(p.d, { x: x + 0.3, y: y + 1.85, w: 3.5, h: 0.85, fontSize: 12, color: COLOR.text, fontFace: FONT_BODY, margin: 0 });
  });
}

// ====== 10. サイトマップ ======
{
  const s = pres.addSlide({ masterName: "CONTENT" });
  addTitle(s, 8, "サイトマップ", "Sitemap");

  const cols = [
    { t: "商品まわり", items: ["トップ", "商品カテゴリ（4区分）", "用途から探す（5用途）", "全商品（型番検索・絞り込み）", "商品詳細（仕様表＋3導線）"] },
    { t: "受注・相談", items: ["お見積り依頼", "特注・オーダー相談", "カタログ請求", "お問い合わせ", "ショッピングカート / 見積カート"] },
    { t: "会員・マイページ", items: ["会員登録（法人審査）", "ログイン", "マイページ", "注文履歴 / 再注文", "見積履歴"] },
    { t: "企業・情報", items: ["会社情報", "工場・製造背景", "品質・保証", "ご利用ガイド", "よくあるご質問"] }
  ];
  cols.forEach((c, i) => {
    const x = 0.5 + i * 3.15;
    const y = 1.75;
    s.addShape(pres.shapes.RECTANGLE, { x, y, w: 3.0, h: 5.2, fill: { color: COLOR.bg }, line: { color: COLOR.border, width: 0.5 } });
    s.addShape(pres.shapes.RECTANGLE, { x, y, w: 3.0, h: 0.5, fill: { color: COLOR.navy }, line: { type: "none" } });
    s.addText(c.t, { x, y, w: 3.0, h: 0.5, fontSize: 13, bold: true, color: COLOR.white, align: "center", valign: "middle", fontFace: FONT_HEADER, margin: 0 });
    s.addText(c.items.map((it, j) => ({
      text: it, options: { bullet: { code: "25A0" }, breakLine: j < c.items.length - 1, paraSpaceAfter: 6, color: COLOR.text }
    })), { x: x + 0.2, y: y + 0.7, w: 2.7, h: 4.4, fontSize: 11, fontFace: FONT_BODY, valign: "top", margin: 0 });
  });
}

// ====== 11. 初期公開 10 SKU ======
{
  const s = pres.addSlide({ masterName: "CONTENT" });
  addTitle(s, 9, "初期公開 10 SKU", "Initial Lineup");

  const hdr = ["No", "型番", "商品名", "カテゴリ", "販売区分"];
  const rows = [
    ["01", "SS-900-L", "軽量スチール棚 W900", "スチール棚", "在庫品"],
    ["02", "SS-1200-M", "中量スチール棚 W1200", "スチール棚", "在庫品"],
    ["03", "SS-1500-H", "重量スチール棚 W1500", "スチール棚", "受注生産"],
    ["04", "WG-600-3", "汎用ワゴン W600", "ワゴン", "在庫品"],
    ["05", "WG-750-S", "静音ワゴン W750", "ワゴン", "在庫品"],
    ["06", "SW-600-MED", "ステンレス医療用ワゴン", "ステンレスワゴン", "在庫品"],
    ["07", "SW-750-KIT", "ステンレス厨房ワゴン", "ステンレスワゴン", "在庫品"],
    ["08", "SF-WT-1200", "ステンレス作業台", "ステンレス什器", "受注生産"],
    ["09", "SF-SK-900", "ステンレス一槽シンク", "ステンレス什器", "受注生産"],
    ["10", "SF-CUSTOM", "ステンレス特注什器（オーダー）", "ステンレス什器", "要見積"]
  ];
  const kindColor = { "在庫品": "206848", "受注生産": "C9731C", "要見積": "666666" };
  const tbl = [
    hdr.map(h => ({ text: h, options: { fill: { color: COLOR.navy }, color: COLOR.white, bold: true, fontFace: FONT_HEADER, fontSize: 11, align: "center", valign: "middle" } })),
    ...rows.map((r) => r.map((c, i) => {
      const opts = { fontSize: 11, fontFace: FONT_BODY, valign: "middle", color: COLOR.text };
      if (i === 0) opts.align = "center";
      if (i === 1) { opts.fontFace = "Consolas"; opts.color = COLOR.navyLight; }
      if (i === 4) {
        opts.color = kindColor[c] || COLOR.text;
        opts.bold = true;
        opts.align = "center";
      }
      return { text: c, options: opts };
    }))
  ];
  s.addTable(tbl, {
    x: 0.5, y: 1.7, w: 12.3, colW: [0.8, 1.8, 4.5, 2.8, 2.4],
    border: { pt: 0.5, color: COLOR.border },
    rowH: 0.42
  });

  s.addText("狙い：4カテゴリをバランスよくカバーし、「在庫品／受注生産／要見積」の3区分の見え方を初期段階で検証する。", {
    x: 0.5, y: 6.45, w: 12.3, h: 0.35, fontSize: 11, italic: true, color: COLOR.navyLight, fontFace: FONT_BODY, margin: 0
  });
}

// ====== 12. 必要機能一覧 ======
{
  const s = pres.addSlide({ masterName: "CONTENT" });
  addTitle(s, 10, "必要機能一覧", "Functional Scope");

  section(s, 0.5, 1.7, 4.0, 5.2, "標準機能で対応", [
    "法人会員登録・会員ランク",
    "会員ランク別価格表示",
    "カート・決済（振込/クレカ/請求書）",
    "商品検索（型番・キーワード・カテゴリ）",
    "見積依頼／カタログ請求／問合せフォーム",
    "注文履歴・再注文",
    "商品CSV一括登録／更新",
    "メルマガ／クーポン"
  ], { accent: "206848" });

  section(s, 4.65, 1.7, 4.0, 5.2, "カスタマイズ・デザイン対応", [
    "商品詳細の仕様表テンプレート",
    "販売区分バッジ（在庫/受注/要見積/静音）",
    "見積カートと買い物カートの並立",
    "製造業トーンのデザイン統一",
    "PC最適化レイアウト",
    "ヒーロー・カテゴリビジュアル",
    "法人会員限定価格の表示制御"
  ], { accent: COLOR.orange });

  section(s, 8.8, 1.7, 4.0, 5.2, "初期スコープ外（拡張）", [
    "基幹／在庫システム連携",
    "複雑な承認ワークフロー",
    "特注のセミオート見積",
    "多言語対応",
    "3D・AR表示",
    "定期購入"
  ], { accent: COLOR.mutedLight });
}

// ====== 13. BtoB 導線設計 ======
{
  const s = pres.addSlide({ masterName: "CONTENT" });
  addTitle(s, 11, "BtoB 向け導線設計", "5 Conversion Paths");

  const hdr = ["#", "導線", "実装", "モックでの動作"];
  const rows = [
    ["①", "会員価格", "法人会員ランク", "ログイン切替で会員価格が表示／非表示"],
    ["②", "見積依頼", "見積カート", "複数商品まとめて一括見積依頼が可能"],
    ["③", "案件相談（特注）", "特注相談フォーム", "SF-CUSTOM＋全商品の「問い合わせ」ボタン"],
    ["④", "大口注文", "数量別価格・バナー制御", "数量連動バナー（拡張で実装）"],
    ["⑤", "再注文", "マイページ注文履歴", "ワンクリック再注文（makeshop標準）"]
  ];
  const tbl = [
    hdr.map(h => ({ text: h, options: { fill: { color: COLOR.navy }, color: COLOR.white, bold: true, fontFace: FONT_HEADER, fontSize: 12, align: "center", valign: "middle" } })),
    ...rows.map(r => r.map((c, i) => ({
      text: c,
      options: {
        fontSize: 12, fontFace: FONT_BODY, valign: "middle",
        color: i === 0 ? COLOR.orange : COLOR.text,
        bold: i === 0 || i === 1,
        align: i === 0 ? "center" : "left"
      }
    })))
  ];
  s.addTable(tbl, {
    x: 0.5, y: 1.7, w: 12.3, colW: [0.8, 2.8, 3.2, 5.5],
    border: { pt: 0.5, color: COLOR.border },
    rowH: 0.7
  });
  s.addText("→ 「買う／相談する／続ける」の3サイクルを標準機能で実現。特注は割り切って人手運用で品質担保。", {
    x: 0.5, y: 6.8, w: 12.3, h: 0.35, fontSize: 11, italic: true, color: COLOR.navyLight, fontFace: FONT_BODY, margin: 0
  });
}

// ====== 14. デザイン方針 ======
{
  const s = pres.addSlide({ masterName: "CONTENT" });
  addTitle(s, 12, "デザイン方針", "Design Direction");

  // Color palette
  section(s, 0.5, 1.7, 6.1, 3.0, "カラーパレット", []);
  const palette = [
    { c: COLOR.navy, n: "Navy", v: "#0E2A4A" },
    { c: COLOR.orange, n: "Orange", v: "#C9731C" },
    { c: COLOR.white, n: "White", v: "#FFFFFF" },
    { c: COLOR.bg, n: "Light BG", v: "#F6F7F9" }
  ];
  palette.forEach((p, i) => {
    const x = 0.8 + i * 1.45;
    const y = 2.35;
    s.addShape(pres.shapes.RECTANGLE, { x, y, w: 1.3, h: 1.3, fill: { color: p.c }, line: { color: COLOR.border, width: 0.5 } });
    s.addText(p.n, { x, y: y + 1.35, w: 1.3, h: 0.3, fontSize: 10, bold: true, color: COLOR.navy, align: "center", fontFace: FONT_HEADER, margin: 0 });
    s.addText(p.v, { x, y: y + 1.65, w: 1.3, h: 0.25, fontSize: 9, color: COLOR.muted, align: "center", fontFace: "Consolas", margin: 0 });
  });

  // UI Policy
  section(s, 6.8, 1.7, 6.0, 3.0, "UI 方針", [
    "PC最適ファースト（スマホは破綻しないレベル）",
    "情報密度高め（BtoBユーザーは情報量を歓迎）",
    "商品詳細は仕様表 → 価格 → 3導線の固定順",
    "装飾的イラストを排し、写真・数値・表で語る"
  ], { accent: COLOR.navyLight });

  // Visual Tone
  section(s, 0.5, 4.9, 12.3, 2.1, "ビジュアルトーン", [
    "白 × ダークネイビー × アクセントオレンジのミニマル構成",
    "写真は白背景の商品単体＋現場背景ヒーローの2種構成",
    "フォントはゴシック系で読みやすさ優先",
    "アイコンは線画ベースで業務系の堅実さを演出"
  ]);
}

// ====== 15. コンテンツ方針 ======
{
  const s = pres.addSlide({ masterName: "CONTENT" });
  addTitle(s, 13, "コンテンツ方針", "Content Strategy");

  const hdr = ["コンテンツ", "役割", "優先度"];
  const rows = [
    ["商品詳細（仕様表・写真・3導線）", "受注の本丸", "★★★"],
    ["ご利用ガイド（送料／支払／納期）", "問い合わせ削減", "★★★"],
    ["よくある質問", "同上", "★★★"],
    ["製造背景・工場紹介", "信頼感 / 既存サイト流用", "★★"],
    ["用途別特集（SEO）", "回遊・流入強化", "★★（拡張）"],
    ["導入事例", "比較検討支援", "★★（拡張）"],
    ["お知らせ／ブログ", "運用で徐々に", "★"]
  ];
  const tbl = [
    hdr.map(h => ({ text: h, options: { fill: { color: COLOR.navy }, color: COLOR.white, bold: true, fontFace: FONT_HEADER, fontSize: 12, align: "center", valign: "middle" } })),
    ...rows.map(r => r.map((c, i) => ({
      text: c,
      options: {
        fontSize: 12, fontFace: FONT_BODY, valign: "middle",
        color: COLOR.text,
        align: i === 2 ? "center" : "left",
        bold: i === 2
      }
    })))
  ];
  s.addTable(tbl, {
    x: 0.5, y: 1.7, w: 12.3, colW: [5.5, 4.5, 2.3],
    border: { pt: 0.5, color: COLOR.border },
    rowH: 0.55
  });
}

// ====== 16. 商品登録方針 ======
{
  const s = pres.addSlide({ masterName: "CONTENT" });
  addTitle(s, 14, "商品登録方針", "Product Data Policy");

  section(s, 0.5, 1.7, 6.1, 5.2, "型番・データ設計", [
    "CSVテンプレート（制作側準備）で登録",
    "型番体系は4プレフィックス：SS / WG / SW / SF",
    "カテゴリは3階層まで、それ以上は絞り込みタグ",
    "共通仕様項目を固定（外寸／段数／耐荷重／素材／区分）",
    "特注は SF-CUSTOM 1枠に集約し管理簡素化"
  ]);
  section(s, 6.8, 1.7, 6.0, 5.2, "画像ガイドライン", [
    "白背景の商品単体写真を基本（1200px四方以上推奨）",
    "使用シーン写真を1〜2枚追加（現場想起のため）",
    "寸法図を差し込み、BtoB検索の信頼感を担保",
    "画像の権利は公開前に棚卸し、必要に応じ再撮影",
    "ファイル命名は型番ベースで統一（SS-1200-M_01.jpg 等）"
  ], { accent: COLOR.navyLight });
}

// ====== 17. 運用フロー ======
{
  const s = pres.addSlide({ masterName: "CONTENT" });
  addTitle(s, 15, "運用フロー", "Operations");

  const flows = [
    { t: "日次", items: ["注文確認 → 出荷 → 発送通知", "見積／問い合わせの一次対応（24時間以内返信をKPI化）"] },
    { t: "週次", items: ["新規法人会員の承認", "在庫・価格の差分確認"] },
    { t: "月次", items: ["商品追加／改定（CSV一括）", "アクセス解析レビュー", "見積→受注転換率のモニタリング"] }
  ];
  flows.forEach((f, i) => {
    const x = 0.5 + i * 4.1;
    const y = 1.75;
    s.addShape(pres.shapes.RECTANGLE, { x, y, w: 4.0, h: 3.5, fill: { color: COLOR.white }, line: { color: COLOR.border, width: 0.5 } });
    s.addShape(pres.shapes.RECTANGLE, { x, y, w: 4.0, h: 0.55, fill: { color: COLOR.navy }, line: { type: "none" } });
    s.addText(f.t, { x, y, w: 4.0, h: 0.55, fontSize: 16, bold: true, color: COLOR.white, align: "center", valign: "middle", fontFace: FONT_HEADER, margin: 0 });
    s.addText(f.items.map((it, j) => ({
      text: it, options: { bullet: { code: "25A0" }, breakLine: j < f.items.length - 1, paraSpaceAfter: 6, color: COLOR.text }
    })), { x: x + 0.25, y: y + 0.8, w: 3.6, h: 2.5, fontSize: 12, fontFace: FONT_BODY, valign: "top", margin: 0 });
  });

  // Team
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 5.5, w: 12.3, h: 1.5, fill: { color: COLOR.bg }, line: { color: COLOR.border, width: 0.5 } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 5.5, w: 0.08, h: 1.5, fill: { color: COLOR.orange }, line: { type: "none" } });
  s.addText("想定体制", { x: 0.75, y: 5.65, w: 11, h: 0.3, fontSize: 13, bold: true, color: COLOR.navy, fontFace: FONT_HEADER, margin: 0 });
  s.addText([
    { text: "EC運用担当 1名（受注・問い合わせ一次対応）", options: { bullet: { code: "25A0" }, breakLine: true, paraSpaceAfter: 2 } },
    { text: "営業／見積担当 1〜2名（特注・大口対応）", options: { bullet: { code: "25A0" }, breakLine: true, paraSpaceAfter: 2 } },
    { text: "制作会社（月次保守・軽微な改修）", options: { bullet: { code: "25A0" } } }
  ], { x: 0.75, y: 6.0, w: 12, h: 1.0, fontSize: 11, color: COLOR.text, fontFace: FONT_BODY, valign: "top", margin: 0 });
}

// ====== 18. フェーズ設計 ======
{
  const s = pres.addSlide({ masterName: "CONTENT" });
  addTitle(s, 16, "フェーズ設計", "Phasing Plan");

  // Initial phase
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.75, w: 6.1, h: 5.2, fill: { color: COLOR.navy }, line: { type: "none" } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.75, w: 6.1, h: 0.08, fill: { color: COLOR.orange }, line: { type: "none" } });
  s.addText("PHASE 1", { x: 0.75, y: 1.95, w: 5, h: 0.4, fontSize: 12, bold: true, color: COLOR.orange, fontFace: FONT_HEADER, margin: 0 });
  s.addText("初期フェーズ（公開時）", { x: 0.75, y: 2.35, w: 5.8, h: 0.5, fontSize: 20, bold: true, color: COLOR.white, fontFace: FONT_HEADER, margin: 0 });
  s.addText([
    { text: "10 SKU（規格品9＋特注受付1）", options: { bullet: { code: "25A0" }, breakLine: true, paraSpaceAfter: 5, color: "CADCFC" } },
    { text: "会員登録・見積・問い合わせ・カタログ請求の4導線", options: { bullet: { code: "25A0" }, breakLine: true, paraSpaceAfter: 5, color: "CADCFC" } },
    { text: "法人会員ランク 2段階（標準／優良）", options: { bullet: { code: "25A0" }, breakLine: true, paraSpaceAfter: 5, color: "CADCFC" } },
    { text: "決済：銀行振込＋クレジット＋請求書払い", options: { bullet: { code: "25A0" }, breakLine: true, paraSpaceAfter: 5, color: "CADCFC" } },
    { text: "運用体制 1〜2名", options: { bullet: { code: "25A0" }, color: "CADCFC" } }
  ], { x: 0.75, y: 3.0, w: 5.7, h: 3.8, fontSize: 12, fontFace: FONT_BODY, valign: "top", margin: 0 });

  // Expansion
  s.addShape(pres.shapes.RECTANGLE, { x: 6.75, y: 1.75, w: 6.05, h: 5.2, fill: { color: COLOR.white }, line: { color: COLOR.border, width: 0.5 } });
  s.addShape(pres.shapes.RECTANGLE, { x: 6.75, y: 1.75, w: 6.05, h: 0.08, fill: { color: COLOR.navyLight }, line: { type: "none" } });
  s.addText("PHASE 2 (3-12m)", { x: 7.0, y: 1.95, w: 5, h: 0.4, fontSize: 12, bold: true, color: COLOR.navyLight, fontFace: FONT_HEADER, margin: 0 });
  s.addText("拡張フェーズ", { x: 7.0, y: 2.35, w: 5.8, h: 0.5, fontSize: 20, bold: true, color: COLOR.navy, fontFace: FONT_HEADER, margin: 0 });
  s.addText([
    { text: "SKU拡張（30〜100型番）", options: { bullet: { code: "25A0" }, breakLine: true, paraSpaceAfter: 5, color: COLOR.text } },
    { text: "特注のセミオート見積（選択式フォーム）", options: { bullet: { code: "25A0" }, breakLine: true, paraSpaceAfter: 5, color: COLOR.text } },
    { text: "基幹／在庫システム連携", options: { bullet: { code: "25A0" }, breakLine: true, paraSpaceAfter: 5, color: COLOR.text } },
    { text: "用途別特集ページ・導入事例拡充", options: { bullet: { code: "25A0" }, breakLine: true, paraSpaceAfter: 5, color: COLOR.text } },
    { text: "広告・SEO強化", options: { bullet: { code: "25A0" }, breakLine: true, paraSpaceAfter: 5, color: COLOR.text } },
    { text: "定期購入（消耗品向け）", options: { bullet: { code: "25A0" }, color: COLOR.text } }
  ], { x: 7.0, y: 3.0, w: 5.7, h: 3.8, fontSize: 12, fontFace: FONT_BODY, valign: "top", margin: 0 });
}

// ====== 19. 制作スケジュール ======
{
  const s = pres.addSlide({ masterName: "CONTENT" });
  addTitle(s, 17, "制作スケジュール", "Schedule (3-4 months)");

  const phases = [
    { n: "①", t: "要件定義", w: 2, span: 2, color: COLOR.orange },
    { n: "②", t: "デザイン", w: 3, span: 3, color: COLOR.navyLight },
    { n: "③", t: "構築", w: 4, span: 4, color: COLOR.navy },
    { n: "④", t: "商品登録", w: 2, span: 2, color: "206848" },
    { n: "⑤", t: "テスト", w: 2, span: 2, color: COLOR.navyLight },
    { n: "⑥", t: "研修", w: 1, span: 1, color: COLOR.mutedLight },
    { n: "⑦", t: "公開", w: 0.5, span: 1, color: COLOR.orange }
  ];
  // Gantt-like
  const startX = 1.8;
  const totalWeeks = 14;
  const weekW = (11 / totalWeeks);
  // Week scale
  for (let w = 0; w <= totalWeeks; w += 2) {
    s.addText(`W${w}`, { x: startX + w * weekW - 0.3, y: 1.7, w: 0.6, h: 0.3, fontSize: 9, color: COLOR.muted, align: "center", fontFace: FONT_BODY, margin: 0 });
  }
  let cursor = 0;
  phases.forEach((p, i) => {
    const y = 2.1 + i * 0.55;
    s.addText(`${p.n} ${p.t}`, { x: 0.5, y, w: 1.3, h: 0.4, fontSize: 12, bold: true, color: COLOR.navy, fontFace: FONT_HEADER, valign: "middle", margin: 0 });
    s.addShape(pres.shapes.RECTANGLE, {
      x: startX + cursor * weekW, y: y + 0.06, w: p.span * weekW - 0.05, h: 0.35,
      fill: { color: p.color }, line: { type: "none" }
    });
    s.addText(`${p.span}週`, {
      x: startX + cursor * weekW, y: y + 0.06, w: p.span * weekW - 0.05, h: 0.35,
      fontSize: 10, color: COLOR.white, align: "center", valign: "middle", bold: true, fontFace: FONT_HEADER, margin: 0
    });
    cursor += p.span;
  });

  s.addText("※ 10SKUへの絞り込みにより、前案（50-150SKU想定）より1〜2ヶ月短縮。", {
    x: 0.5, y: 6.5, w: 12.3, h: 0.35, fontSize: 11, italic: true, color: COLOR.navyLight, fontFace: FONT_BODY, margin: 0
  });
}

// ====== 20. リスクと注意点 ======
{
  const s = pres.addSlide({ masterName: "CONTENT" });
  addTitle(s, 18, "想定リスクと注意点", "Risks & Mitigations");

  const hdr = ["リスク", "内容", "対策"];
  const rows = [
    ["画像権利", "既存サイト画像の権利が未確認", "公開前に棚卸し・必要に応じ再撮影"],
    ["情報粒度のバラつき", "型番・仕様の表記が不統一", "マスタ整備＋テンプレート固定"],
    ["特注の過剰自動化", "自動見積の作り込みは破綻する", "初期は人手運用で割り切る"],
    ["会員審査", "厳しすぎると離脱", "最小項目＋後追い審査"],
    ["運用属人化", "担当1名に集中", "マニュアル化／CSV運用で標準化"],
    ["10SKUの見栄え", "カテゴリが埋まらない印象", "「取扱を絞った精鋭ラインナップ」として訴求"]
  ];
  const tbl = [
    hdr.map(h => ({ text: h, options: { fill: { color: COLOR.navy }, color: COLOR.white, bold: true, fontFace: FONT_HEADER, fontSize: 12, align: "center", valign: "middle" } })),
    ...rows.map(r => r.map((c, i) => ({
      text: c,
      options: {
        fontSize: 11, fontFace: FONT_BODY, valign: "middle",
        color: i === 0 ? COLOR.warn : COLOR.text,
        bold: i === 0,
        align: "left"
      }
    })))
  ];
  s.addTable(tbl, {
    x: 0.5, y: 1.7, w: 12.3, colW: [2.8, 4.8, 4.7],
    border: { pt: 0.5, color: COLOR.border },
    rowH: 0.65
  });
}

// ====== 21. まとめ / 次のアクション ======
{
  const s = pres.addSlide({ masterName: "CONTENT" });
  addTitle(s, 19, "まとめ / 次のアクション", "Summary & Next Steps");

  section(s, 0.5, 1.7, 6.1, 5.2, "本提案のポイント", [
    "makeshopで十分現実的にBtoB ECは立ち上げ可能",
    "参考サイトの勝ち筋を標準機能＋デザインで再現",
    "10SKUで3〜4ヶ月公開、運用は1〜2名で回る",
    "「規格品は即買／特注は見積」の割り切りで運用破綻を防止",
    "拡張は SKU増・連携・コンテンツ拡充へ段階的に"
  ]);

  // Next actions
  s.addShape(pres.shapes.RECTANGLE, { x: 6.8, y: 1.7, w: 6.0, h: 5.2, fill: { color: COLOR.navy }, line: { type: "none" } });
  s.addShape(pres.shapes.RECTANGLE, { x: 6.8, y: 1.7, w: 0.08, h: 5.2, fill: { color: COLOR.orange }, line: { type: "none" } });
  s.addText("次のアクション", { x: 7.05, y: 1.9, w: 5.6, h: 0.45, fontSize: 16, bold: true, color: COLOR.white, fontFace: FONT_HEADER, margin: 0 });

  const actions = [
    { n: "01", t: "10SKUの型番・仕様・価格マスタ確定" },
    { n: "02", t: "画像の権利確認（特に特注什器の写真）" },
    { n: "03", t: "法人会員ランク（2段階）・価格ポリシー決定" },
    { n: "04", t: "本提案＋モックURLで要件定義キックオフ" }
  ];
  actions.forEach((a, i) => {
    const y = 2.55 + i * 1.0;
    s.addShape(pres.shapes.OVAL, { x: 7.05, y: y, w: 0.65, h: 0.65, fill: { color: COLOR.orange }, line: { type: "none" } });
    s.addText(a.n, { x: 7.05, y: y, w: 0.65, h: 0.65, fontSize: 13, bold: true, color: COLOR.white, align: "center", valign: "middle", fontFace: FONT_HEADER, margin: 0 });
    s.addText(a.t, { x: 7.85, y: y, w: 4.9, h: 0.65, fontSize: 13, color: COLOR.white, valign: "middle", fontFace: FONT_BODY, margin: 0 });
  });
}

// ====== 22. A案 vs B案 デザイン比較 ======
{
  const s = pres.addSlide({ masterName: "CONTENT" });
  addTitle(s, "A/B", "デザイン比較 ｜ A案（makeshop）vs B案（BASE）", "Mockup Design Comparison");

  // ---- A案 (makeshop) ----
  const ax = 0.5, ay = 1.7, aw = 6.1, ah = 5.2;
  s.addShape(pres.shapes.RECTANGLE, { x: ax, y: ay, w: aw, h: ah, fill: { color: COLOR.white }, line: { color: COLOR.border, width: 0.5 } });
  s.addShape(pres.shapes.RECTANGLE, { x: ax, y: ay, w: aw, h: 0.55, fill: { color: COLOR.navy }, line: { type: "none" } });
  s.addText("A 案  ｜  makeshop", { x: ax + 0.2, y: ay, w: aw - 0.4, h: 0.55, fontSize: 14, bold: true, color: COLOR.white, fontFace: FONT_HEADER, valign: "middle", margin: 0 });
  s.addText("BtoB EC 推奨案", { x: ax + aw - 2.0, y: ay, w: 1.7, h: 0.55, fontSize: 10, color: COLOR.orange, align: "right", valign: "middle", fontFace: FONT_HEADER, bold: true, margin: 0 });

  // Mini UI preview (A)
  const apx = ax + 0.25, apy = ay + 0.75;
  s.addShape(pres.shapes.RECTANGLE, { x: apx, y: apy, w: aw - 0.5, h: 1.3, fill: { color: COLOR.navy }, line: { type: "none" } });
  s.addText("HIKARI  BtoB STORE", { x: apx + 0.15, y: apy + 0.1, w: 3, h: 0.3, fontSize: 11, bold: true, color: COLOR.white, fontFace: FONT_HEADER, margin: 0 });
  s.addShape(pres.shapes.RECTANGLE, { x: apx + 3.4, y: apy + 0.1, w: 1.9, h: 0.28, fill: { color: COLOR.orange }, line: { type: "none" } });
  s.addText("法人会員ログイン", { x: apx + 3.4, y: apy + 0.1, w: 1.9, h: 0.28, fontSize: 8, color: COLOR.white, align: "center", valign: "middle", fontFace: FONT_HEADER, margin: 0 });
  s.addText("🗄  商品カテゴリ／用途／全商品／特注・オーダー相談", { x: apx + 0.15, y: apy + 0.5, w: aw - 0.8, h: 0.3, fontSize: 9, color: "CADCFC", fontFace: FONT_BODY, margin: 0 });
  s.addShape(pres.shapes.RECTANGLE, { x: apx + 0.15, y: apy + 0.9, w: 1.2, h: 0.28, fill: { color: COLOR.orange }, line: { type: "none" } });
  s.addText("見積カート(2)", { x: apx + 0.15, y: apy + 0.9, w: 1.2, h: 0.28, fontSize: 8, color: COLOR.white, align: "center", valign: "middle", fontFace: FONT_HEADER, margin: 0 });
  s.addShape(pres.shapes.RECTANGLE, { x: apx + 1.45, y: apy + 0.9, w: 1.2, h: 0.28, fill: { color: COLOR.navyLight }, line: { type: "none" } });
  s.addText("カート(1)", { x: apx + 1.45, y: apy + 0.9, w: 1.2, h: 0.28, fontSize: 8, color: COLOR.white, align: "center", valign: "middle", fontFace: FONT_HEADER, margin: 0 });

  // Feature list A
  const aFeatures = [
    "カラー：ネイビー × オレンジ（BtoB堅実）",
    "ナビゲーション：8項目（カテゴリ/用途/特注/見積 等）",
    "絞り込み：カテゴリ × 素材 × 販売区分の3軸",
    "価格表示：会員ログイン時のみ会員価格を表示",
    "カート：ショッピングカート + 見積カートの2系統",
    "特注：独立した「特注・オーダー相談」専用ページ",
    "請求書払い・会員ランク別価格：標準対応"
  ];
  s.addText(aFeatures.map((t, i) => ({
    text: t, options: { bullet: { code: "25A0" }, breakLine: i < aFeatures.length - 1, paraSpaceAfter: 3, color: COLOR.text }
  })), { x: apx, y: apy + 1.5, w: aw - 0.5, h: 2.8, fontSize: 11, fontFace: FONT_BODY, valign: "top", margin: 0 });

  // ---- B案 (BASE) ----
  const bx = 6.8, by = 1.7, bw = 6.0, bh = 5.2;
  s.addShape(pres.shapes.RECTANGLE, { x: bx, y: by, w: bw, h: bh, fill: { color: COLOR.white }, line: { color: COLOR.border, width: 0.5 } });
  s.addShape(pres.shapes.RECTANGLE, { x: bx, y: by, w: bw, h: 0.55, fill: { color: "3A3A3A" }, line: { type: "none" } });
  s.addText("B 案  ｜  BASE", { x: bx + 0.2, y: by, w: bw - 0.4, h: 0.55, fontSize: 14, bold: true, color: COLOR.white, fontFace: FONT_HEADER, valign: "middle", margin: 0 });
  s.addText("BtoC テンプレート寄り", { x: bx + bw - 2.3, y: by, w: 2.0, h: 0.55, fontSize: 10, color: "E8C18B", align: "right", valign: "middle", fontFace: FONT_HEADER, bold: true, margin: 0 });

  // Mini UI preview (B)
  const bpx = bx + 0.25, bpy = by + 0.75;
  s.addShape(pres.shapes.RECTANGLE, { x: bpx, y: bpy, w: bw - 0.5, h: 1.3, fill: { color: "FAFAF8" }, line: { color: "EEE9E2", width: 0.5 } });
  s.addText("HIKARI  STORE", { x: bpx + 0.15, y: bpy + 0.1, w: 3, h: 0.3, fontSize: 11, bold: true, color: "222222", fontFace: FONT_HEADER, margin: 0 });
  s.addText("🔍  👤  🛒", { x: bpx + bw - 1.5, y: bpy + 0.1, w: 1.2, h: 0.3, fontSize: 11, color: "222222", align: "right", fontFace: FONT_BODY, margin: 0 });
  s.addText("ITEMS   ALL   ABOUT   CUSTOM   CONTACT", { x: bpx + 0.15, y: bpy + 0.55, w: bw - 0.8, h: 0.3, fontSize: 8, color: "222222", fontFace: FONT_HEADER, charSpacing: 2, margin: 0 });
  s.addShape(pres.shapes.RECTANGLE, { x: bpx + 0.15, y: bpy + 0.95, w: 1.6, h: 0.25, fill: { color: "222222" }, line: { type: "none" } });
  s.addText("ADD TO CART", { x: bpx + 0.15, y: bpy + 0.95, w: 1.6, h: 0.25, fontSize: 8, color: COLOR.white, align: "center", valign: "middle", fontFace: FONT_HEADER, margin: 0 });

  // Feature list B
  const bFeatures = [
    "カラー：ウォームホワイト × ブラック × 金茶",
    "ナビゲーション：5項目（ITEMS/ALL/ABOUT/CUSTOM/CONTACT）",
    "絞り込み：カテゴリタブ1軸のみ",
    "価格表示：常時表示（会員価格の動的切替なし）",
    "カート：ショッピングカート1系統のみ",
    "特注：共通のお問い合わせモーダル",
    "請求書払い・会員ランク別価格：標準では不可"
  ];
  s.addText(bFeatures.map((t, i) => ({
    text: t, options: { bullet: { code: "25A0" }, breakLine: i < bFeatures.length - 1, paraSpaceAfter: 3, color: COLOR.text }
  })), { x: bpx, y: bpy + 1.5, w: bw - 0.5, h: 2.8, fontSize: 11, fontFace: FONT_BODY, valign: "top", margin: 0 });

  // Bottom annotation
  s.addText([
    { text: "A案 ", options: { bold: true, color: COLOR.navy } },
    { text: "： 法人取引の導線（会員価格・見積・請求書払い）を標準機能で再現 ", options: { color: COLOR.text } },
    { text: "｜ ", options: { color: COLOR.mutedLight } },
    { text: "B案 ", options: { bold: true, color: "3A3A3A" } },
    { text: "： BtoC運用に近く、BtoB機能は回避策／外部ツール／手運用で補完が必要", options: { color: COLOR.text } }
  ], { x: 0.5, y: 7.0, w: 12.3, h: 0.3, fontSize: 10, fontFace: FONT_BODY, italic: true, margin: 0 });
}

// ====== 23. BASE vs makeshop ランニングコスト比較 ======
{
  const s = pres.addSlide({ masterName: "CONTENT" });
  addTitle(s, "COST", "BASE vs makeshop ランニングコスト比較", "5-Year Running Cost / BtoB機能対応込みの実質コスト");

  // ---- 上段：基本プラン費用 ----
  const plan1 = [
    [{ text: "項目", options: { bold: true, fill: { color: COLOR.navy }, color: COLOR.white, fontSize: 10, fontFace: FONT_HEADER, align: "center", valign: "middle" } },
     { text: "BASE（スタンダード）", options: { bold: true, fill: { color: COLOR.navy }, color: COLOR.white, fontSize: 10, fontFace: FONT_HEADER, align: "center", valign: "middle" } },
     { text: "makeshop（プレミアム＋B2B）", options: { bold: true, fill: { color: COLOR.orange }, color: COLOR.white, fontSize: 10, fontFace: FONT_HEADER, align: "center", valign: "middle" } }],
    [{ text: "初期費用", options: { fill: { color: COLOR.bg }, bold: true, fontSize: 10, fontFace: FONT_BODY, valign: "middle" } },
     { text: "0円", options: { fontSize: 10, fontFace: FONT_BODY, align: "center", valign: "middle" } },
     { text: "11,000円", options: { fontSize: 10, fontFace: FONT_BODY, align: "center", valign: "middle" } }],
    [{ text: "月額（年払い）", options: { fill: { color: COLOR.bg }, bold: true, fontSize: 10, fontFace: FONT_BODY, valign: "middle" } },
     { text: "16,580円／月", options: { fontSize: 10, fontFace: FONT_BODY, align: "center", valign: "middle" } },
     { text: "22,000円／月（B2B込）", options: { fontSize: 10, fontFace: FONT_BODY, align: "center", valign: "middle" } }],
    [{ text: "決済手数料", options: { fill: { color: COLOR.bg }, bold: true, fontSize: 10, fontFace: FONT_BODY, valign: "middle" } },
     { text: "2.9%", options: { fontSize: 10, fontFace: FONT_BODY, align: "center", valign: "middle" } },
     { text: "3.14% + 40円/件", options: { fontSize: 10, fontFace: FONT_BODY, align: "center", valign: "middle" } }],
  ];
  s.addTable(plan1, {
    x: 0.5, y: 1.7, w: 6.0, colW: [1.7, 2.1, 2.2],
    border: { pt: 0.5, color: COLOR.border }, rowH: 0.38
  });

  // ---- 右上：BtoB機能対応可否 ----
  const feat = [
    [{ text: "BtoB機能", options: { bold: true, fill: { color: COLOR.navy }, color: COLOR.white, fontSize: 10, fontFace: FONT_HEADER, align: "center", valign: "middle" } },
     { text: "BASE", options: { bold: true, fill: { color: COLOR.navy }, color: COLOR.white, fontSize: 10, fontFace: FONT_HEADER, align: "center", valign: "middle" } },
     { text: "makeshop", options: { bold: true, fill: { color: COLOR.orange }, color: COLOR.white, fontSize: 10, fontFace: FONT_HEADER, align: "center", valign: "middle" } }],
    [{ text: "会員ランク別価格", options: { fill: { color: COLOR.bg }, bold: true, fontSize: 10, fontFace: FONT_BODY, valign: "middle" } },
     { text: "×", options: { fontSize: 14, bold: true, color: COLOR.warn, fontFace: FONT_HEADER, align: "center", valign: "middle" } },
     { text: "◎", options: { fontSize: 14, bold: true, color: "206848", fontFace: FONT_HEADER, align: "center", valign: "middle" } }],
    [{ text: "請求書払い（掛売）", options: { fill: { color: COLOR.bg }, bold: true, fontSize: 10, fontFace: FONT_BODY, valign: "middle" } },
     { text: "△ 外部連携", options: { fontSize: 10, color: COLOR.orangeDark, fontFace: FONT_BODY, align: "center", valign: "middle" } },
     { text: "◎", options: { fontSize: 14, bold: true, color: "206848", fontFace: FONT_HEADER, align: "center", valign: "middle" } }],
    [{ text: "見積書PDF発行", options: { fill: { color: COLOR.bg }, bold: true, fontSize: 10, fontFace: FONT_BODY, valign: "middle" } },
     { text: "×", options: { fontSize: 14, bold: true, color: COLOR.warn, fontFace: FONT_HEADER, align: "center", valign: "middle" } },
     { text: "◎", options: { fontSize: 14, bold: true, color: "206848", fontFace: FONT_HEADER, align: "center", valign: "middle" } }],
  ];
  s.addTable(feat, {
    x: 6.8, y: 1.7, w: 6.0, colW: [3.0, 1.5, 1.5],
    border: { pt: 0.5, color: COLOR.border }, rowH: 0.38
  });

  // ---- 中段：試算前提 ----
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.55, w: 12.3, h: 0.5, fill: { color: "FFF6E6" }, line: { color: "E8C18B", width: 0.5 } });
  s.addText("前提条件： 年商3,000万円 ／ 年間1,000件受注 ／ 掛売比率30% ／ BtoB運用による追加コストを加味", {
    x: 0.65, y: 3.55, w: 12.0, h: 0.5, fontSize: 11, color: "8A6A30", bold: true, valign: "middle", fontFace: FONT_BODY, margin: 0
  });

  // ---- 下段：5年累計コスト大比較 ----
  // BASE side
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.25, w: 6.0, h: 2.7, fill: { color: "3A3A3A" }, line: { type: "none" } });
  s.addText("BASE（5年累計 / 実質）", { x: 0.7, y: 4.4, w: 5.5, h: 0.35, fontSize: 12, bold: true, color: "E8C18B", fontFace: FONT_HEADER, margin: 0 });
  s.addText("約 926 万円", { x: 0.7, y: 4.8, w: 5.5, h: 1.0, fontSize: 42, bold: true, color: COLOR.white, fontFace: FONT_HEADER, margin: 0 });
  s.addText([
    { text: "内訳： ", options: { color: "AAAAAA" } },
    { text: "プラン費 ", options: { color: "CCCCCC" } },
    { text: "+ 決済 ", options: { color: "CCCCCC" } },
    { text: "+ 掛売手数料（NP等） ", options: { color: "CCCCCC" } },
    { text: "+ 手作業工数 ", options: { color: "CCCCCC" } },
    { text: "+ 外部ツール ", options: { color: "CCCCCC" } },
    { text: "+ 追加開発費", options: { color: "CCCCCC" } }
  ], { x: 0.7, y: 5.95, w: 5.6, h: 0.85, fontSize: 9, fontFace: FONT_BODY, valign: "top", margin: 0 });

  // makeshop side
  s.addShape(pres.shapes.RECTANGLE, { x: 6.8, y: 4.25, w: 6.0, h: 2.7, fill: { color: COLOR.navy }, line: { type: "none" } });
  s.addShape(pres.shapes.RECTANGLE, { x: 6.8, y: 4.25, w: 6.0, h: 0.1, fill: { color: COLOR.orange }, line: { type: "none" } });
  s.addText("makeshop（5年累計）", { x: 7.0, y: 4.4, w: 5.5, h: 0.35, fontSize: 12, bold: true, color: COLOR.orange, fontFace: FONT_HEADER, margin: 0 });
  s.addText("約 604 万円", { x: 7.0, y: 4.8, w: 5.5, h: 1.0, fontSize: 42, bold: true, color: COLOR.white, fontFace: FONT_HEADER, margin: 0 });
  s.addText([
    { text: "内訳： ", options: { color: "AAAAAA" } },
    { text: "プラン費（B2B含） ", options: { color: "CADCFC" } },
    { text: "+ 決済 ", options: { color: "CADCFC" } },
    { text: "※ BtoB機能は全て標準対応のため、追加運用コストは発生しません。", options: { color: COLOR.orange, italic: true, breakLine: true } }
  ], { x: 7.0, y: 5.95, w: 5.6, h: 0.85, fontSize: 9, fontFace: FONT_BODY, valign: "top", margin: 0 });

  // Diff banner
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 7.0, w: 12.3, h: 0.3, fill: { color: COLOR.orange }, line: { type: "none" } });
  s.addText("→ 5年累計で makeshop が 約 322 万円 お得。さらにBtoB運用の標準化による属人化リスク低減を加味すると優位性はさらに拡大。", {
    x: 0.5, y: 7.0, w: 12.3, h: 0.3, fontSize: 11, bold: true, color: COLOR.white, align: "center", valign: "middle", fontFace: FONT_HEADER, margin: 0
  });
}

// ====== 24. (参考) モック画面 ======
{
  const s = pres.addSlide({ masterName: "CONTENT" });
  addTitle(s, "", "（参考）モックアップ画面", "https://hikari-btob-mockup.vercel.app/");

  // URL card
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.75, w: 12.3, h: 0.7, fill: { color: COLOR.navy }, line: { type: "none" } });
  s.addText("🌐  https://hikari-btob-mockup.vercel.app/", {
    x: 0.5, y: 1.75, w: 12.3, h: 0.7, fontSize: 16, bold: true, color: COLOR.white, align: "center", valign: "middle", fontFace: "Consolas", margin: 0
  });

  const features = [
    { t: "ヘッダー・ヒーロー", d: "HIKARIロゴ / 法人会員ログイン切替 / 検索バー / 工場写真ヒーロー" },
    { t: "商品一覧・絞り込み", d: "カテゴリ／素材／販売区分の3軸フィルタ・バッジ表示・10SKUを一望" },
    { t: "商品詳細モーダル", d: "仕様表／会員価格表示／カート・見積・問合せの3導線" },
    { t: "見積カート", d: "買い物カートと別系統・複数商品まとめて一括見積依頼" },
    { t: "法人会員ログイン", d: "ログインON/OFFで会員価格の表示／非表示を切替（デモ）" },
    { t: "特注品導線", d: "SF-CUSTOM は価格非表示＋「要見積」バッジ＋専用フォーム" }
  ];
  features.forEach((f, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = 0.5 + col * 6.2;
    const y = 2.7 + row * 1.4;
    s.addShape(pres.shapes.RECTANGLE, { x, y, w: 6.0, h: 1.3, fill: { color: COLOR.bg }, line: { color: COLOR.border, width: 0.5 } });
    s.addShape(pres.shapes.RECTANGLE, { x, y, w: 0.08, h: 1.3, fill: { color: COLOR.orange }, line: { type: "none" } });
    s.addText(f.t, { x: x + 0.25, y: y + 0.15, w: 5.5, h: 0.35, fontSize: 13, bold: true, color: COLOR.navy, fontFace: FONT_HEADER, margin: 0 });
    s.addText(f.d, { x: x + 0.25, y: y + 0.52, w: 5.5, h: 0.75, fontSize: 11, color: COLOR.text, fontFace: FONT_BODY, margin: 0 });
  });
}

pres.writeFile({ fileName: "hikari-btob-proposal.pptx" })
  .then(() => console.log("✅ generated: hikari-btob-proposal.pptx"))
  .catch(err => { console.error("❌", err); process.exit(1); });
