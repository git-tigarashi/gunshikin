// PowerPoint スライド生成スクリプト（統合版 / 株式会社Vision Creator）
// 集客見込み客・既存受講者 共通の勉強会用

const PptxGenJS = require('C:/Users/haman/AppData/Roaming/npm/node_modules/pptxgenjs');
const path = require('path');

const pptx = new PptxGenJS();
pptx.layout = 'LAYOUT_WIDE'; // 16:9

// ========== カラーパレット（集客用ベース） ==========
const C = {
  navy:    '1A2744',
  gold:    'D4A843',
  orange:  'E8602C',
  white:   'FFFFFF',
  light:   'F4F6FA',
  gray:    '7A8A9A',
  darkgray:'333333',
  green:   '2E7D32',
  red:     'C62828',
  accent:  '1565C0',
  navyMid: '243060',
  navyBg:  '1E2A44',
};

const FONT_JA = 'メイリオ';
const FONT_EN = 'Calibri';
const COMPANY = '株式会社 Vision Creator';

// ========== ユーティリティ ==========
function addBg(slide, color) {
  slide.addShape(pptx.ShapeType.rect, {
    x: 0, y: 0, w: '100%', h: '100%', fill: { color }
  });
}

// フッター（全スライド共通）
function addFooter(slide) {
  // 左：会社名
  slide.addText(COMPANY, {
    x: 0.3, y: 6.88, w: 6.0, h: 0.28,
    fontSize: 9, color: C.gray, fontFace: FONT_JA, align: 'left', italic: true
  });
  // 右：講義タイトル
  slide.addText('投資の軍資金を作る8つの方法', {
    x: 6.5, y: 6.88, w: 6.8, h: 0.28,
    fontSize: 9, color: C.gray, fontFace: FONT_JA, align: 'right'
  });
}

// 区切り線付きセクションタイトル
function sectionTitle(slide, text, sub, opts = {}) {
  const y    = opts.y  || 0.45;
  const fs   = opts.fs || 28;
  const th   = opts.th || 0.72;
  const barC = opts.barC || C.gold;
  slide.addShape(pptx.ShapeType.rect, { x: 0.4, y, w: 0.09, h: th, fill: { color: barC } });
  slide.addText(text, {
    x: 0.65, y, w: 12.0, h: th,
    fontSize: fs, bold: true, color: C.navy, fontFace: FONT_JA
  });
  if (sub) {
    slide.addText(sub, {
      x: 0.65, y: y + th + 0.06, w: 12.0, h: 0.38,
      fontSize: 13, color: C.gray, fontFace: FONT_JA
    });
  }
}

// ダークヘッダー付き手法スライドのヘッダー
function methodHeader(slide, no, title, badge) {
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: '100%', h: 1.1, fill: { color: C.navy } });
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 1.08, w: '100%', h: 0.05, fill: { color: C.gold } });
  // 番号バッジ
  slide.addShape(pptx.ShapeType.rect, { x: 0.4, y: 0.15, w: 0.82, h: 0.82, fill: { color: C.gold } });
  slide.addText(no, {
    x: 0.4, y: 0.15, w: 0.82, h: 0.82,
    fontSize: 30, bold: true, color: C.navy, fontFace: FONT_EN, align: 'center'
  });
  slide.addText(title, {
    x: 1.38, y: 0.18, w: 8.6, h: 0.78,
    fontSize: 26, bold: true, color: C.white, fontFace: FONT_JA
  });
  // 月額バッジ
  slide.addShape(pptx.ShapeType.rect, { x: 10.15, y: 0.18, w: 2.55, h: 0.72, fill: { color: C.orange } });
  slide.addText(badge, {
    x: 10.15, y: 0.18, w: 2.55, h: 0.72,
    fontSize: 16, bold: true, color: C.white, fontFace: FONT_JA, align: 'center'
  });
}

// チェックポイント行
function checkRow(slide, y, text, color) {
  color = color || C.accent;
  slide.addShape(pptx.ShapeType.rect, { x: 0.42, y: y + 0.08, w: 0.48, h: 0.48, fill: { color } });
  slide.addText('✓', {
    x: 0.42, y: y + 0.08, w: 0.48, h: 0.48,
    fontSize: 16, bold: true, color: C.white, fontFace: FONT_EN, align: 'center'
  });
  slide.addText(text, {
    x: 1.04, y, w: 11.65, h: 0.66,
    fontSize: 15.5, color: C.darkgray, fontFace: FONT_JA
  });
}

// アクションボックス（今日の一歩）
function actionBox(slide, y, text) {
  slide.addShape(pptx.ShapeType.rect, {
    x: 0.42, y, w: 12.28, h: 0.72,
    fill: { color: 'FFF8E1' }, line: { color: C.gold, pt: 2 }
  });
  slide.addShape(pptx.ShapeType.rect, { x: 0.42, y, w: 1.55, h: 0.72, fill: { color: C.gold } });
  slide.addText('今日の一歩', {
    x: 0.42, y: y + 0.12, w: 1.55, h: 0.48,
    fontSize: 13, bold: true, color: C.navy, fontFace: FONT_JA, align: 'center'
  });
  slide.addText(text, {
    x: 2.08, y: y + 0.12, w: 10.5, h: 0.48,
    fontSize: 14, color: C.darkgray, fontFace: FONT_JA
  });
}

// ポイントボックス（淡い背景）
function noteBox(slide, y, text, h) {
  h = h || 0.62;
  slide.addShape(pptx.ShapeType.rect, {
    x: 0.42, y, w: 12.28, h,
    fill: { color: 'EEF3FB' }, line: { color: C.accent, pt: 1 }
  });
  slide.addText(text, {
    x: 0.58, y: y + 0.06, w: 12.0, h: h - 0.12,
    fontSize: 13.5, color: C.accent, fontFace: FONT_JA, bold: true
  });
}

// ============================================================
//  スライド 01: 表紙
// ============================================================
(function() {
  const s = pptx.addSlide();
  addBg(s, C.navy);

  // 上下ライン
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0,    w: '100%', h: 0.07, fill: { color: C.gold } });
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 6.3,  w: '100%', h: 0.07, fill: { color: C.gold } });
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 6.37, w: '100%', h: 0.04, fill: { color: C.orange } });

  // メインタイトル
  s.addText('投資の軍資金を作る', {
    x: 0.8, y: 0.9, w: 11.7, h: 1.0,
    fontSize: 46, bold: true, color: C.white, fontFace: FONT_JA, align: 'center'
  });
  s.addText('8つの方法', {
    x: 0.8, y: 1.88, w: 11.7, h: 1.3,
    fontSize: 64, bold: true, color: C.gold, fontFace: FONT_JA, align: 'center'
  });

  // サブタイトル
  s.addShape(pptx.ShapeType.rect, { x: 3.2, y: 3.3, w: 6.9, h: 0.06, fill: { color: C.orange } });
  s.addText('〜 将来不安をゼロにする、資金作りの全技術 〜', {
    x: 0.8, y: 3.42, w: 11.7, h: 0.52,
    fontSize: 17, color: 'B0C4DE', fontFace: FONT_JA, align: 'center'
  });

  // 会社名ロゴエリア
  s.addShape(pptx.ShapeType.rect, { x: 3.8, y: 4.2, w: 5.7, h: 0.85, fill: { color: C.navyMid } });
  s.addText(COMPANY, {
    x: 3.8, y: 4.2, w: 5.7, h: 0.85,
    fontSize: 22, bold: true, color: C.gold, fontFace: FONT_JA, align: 'center'
  });

  // フッター会社名
  s.addText(COMPANY, {
    x: 0, y: 6.55, w: '100%', h: 0.35,
    fontSize: 10, color: 'AABBCC', fontFace: FONT_JA, align: 'center', italic: true
  });
})();

// ============================================================
//  スライド 02: 本日のゴール
// ============================================================
(function() {
  const s = pptx.addSlide();
  addBg(s, C.light);
  sectionTitle(s, '本日のゴール');

  const items = [
    { no: '①', text: '軍資金（投資元本）を作る8つの具体的手法を理解する', color: C.navy },
    { no: '②', text: '月3〜10万円の軍資金を継続的に確保できるイメージを持つ', color: C.accent },
    { no: '③', text: '貯めた軍資金を「どう増やすか」の選択肢を広げる', color: C.orange },
  ];

  items.forEach((item, i) => {
    const y = 1.55 + i * 1.45;
    s.addShape(pptx.ShapeType.rect, {
      x: 0.42, y, w: 12.28, h: 1.22,
      fill: { color: C.white }, line: { color: 'D8DEE6', pt: 1 }
    });
    s.addShape(pptx.ShapeType.ellipse, { x: 0.62, y: y + 0.23, w: 0.76, h: 0.76, fill: { color: item.color } });
    s.addText(item.no, {
      x: 0.62, y: y + 0.23, w: 0.76, h: 0.76,
      fontSize: 18, bold: true, color: C.white, fontFace: FONT_JA, align: 'center'
    });
    s.addText(item.text, {
      x: 1.55, y: y + 0.23, w: 10.95, h: 0.76,
      fontSize: 19, color: C.darkgray, fontFace: FONT_JA
    });
  });

  addFooter(s);
})();

// ============================================================
//  スライド 03: 将来不安の現実
// ============================================================
(function() {
  const s = pptx.addSlide();
  addBg(s, C.navy);
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: '100%', h: 0.07, fill: { color: C.gold } });

  s.addText('あなたの老後は、今のままで大丈夫ですか？', {
    x: 0.5, y: 0.18, w: 12.3, h: 0.8,
    fontSize: 30, bold: true, color: C.white, fontFace: FONT_JA, align: 'center'
  });

  const facts = [
    { num: '¥148,000', label: '厚生年金の平均月額', sub: '（2024年度）', color: C.orange },
    { num: '0.1%',     label: '銀行普通預金の金利', sub: '100万円預けて年わずか1,000円', color: C.red },
    { num: '+30%',     label: '物価上昇（2020年比）', sub: '実質的な購買力は年々低下', color: C.orange },
    { num: '2,000万円', label: '老後の資金不足試算額', sub: '（金融庁レポートより）', color: C.red },
  ];

  facts.forEach((f, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = col === 0 ? 0.42 : 6.78;
    const y = 1.2 + row * 2.35;
    s.addShape(pptx.ShapeType.rect, {
      x, y, w: 6.0, h: 2.1,
      fill: { color: C.navyMid }, line: { color: f.color, pt: 2 }
    });
    s.addText(f.num, {
      x, y: y + 0.1, w: 6.0, h: 1.0,
      fontSize: 40, bold: true, color: f.color, fontFace: FONT_EN, align: 'center'
    });
    s.addText(f.label, {
      x, y: y + 1.05, w: 6.0, h: 0.52,
      fontSize: 16, bold: true, color: C.white, fontFace: FONT_JA, align: 'center'
    });
    s.addText(f.sub, {
      x, y: y + 1.55, w: 6.0, h: 0.4,
      fontSize: 12, color: 'AABBCC', fontFace: FONT_JA, align: 'center'
    });
  });

  addFooter(s);
})();

// ============================================================
//  スライド 04: 「軍資金」という発想の転換
// ============================================================
(function() {
  const s = pptx.addSlide();
  addBg(s, C.light);
  sectionTitle(s, '「軍資金」という発想の転換');

  // Before → After フロー
  const bgBox = (x, y, w, h, col) =>
    s.addShape(pptx.ShapeType.rect, { x, y, w, h, fill: { color: col } });

  bgBox(0.42, 1.45, 12.28, 0.72, C.navy);
  s.addText('給料 ＝ 「生活費」だけではなく「生活費」＋「軍資金」に分けて考える', {
    x: 0.5, y: 1.45, w: 12.1, h: 0.72,
    fontSize: 18, bold: true, color: C.gold, fontFace: FONT_JA, align: 'center'
  });

  // フロー矢印図
  const boxes = [
    { x: 0.4,  label: '給与\n収入',     color: C.accent },
    { x: 3.85, label: '生活費\n（支出）', color: C.gray },
    { x: 7.3,  label: '軍資金\n（元本）', color: C.green },
    { x: 10.75,label: '資産形成\n（未来）', color: C.orange },
  ];

  boxes.forEach((b, i) => {
    bgBox(b.x, 2.45, 3.1, 1.5, b.color);
    s.addText(b.label, {
      x: b.x, y: 2.45, w: 3.1, h: 1.5,
      fontSize: 19, bold: true, color: C.white, fontFace: FONT_JA, align: 'center'
    });
    if (i < 3) {
      bgBox(b.x + 3.14, 2.9, 0.42, 0.6, C.gold);
      s.addText('▶', {
        x: b.x + 3.14, y: 2.9, w: 0.42, h: 0.6,
        fontSize: 18, bold: true, color: C.navy, fontFace: FONT_EN, align: 'center'
      });
    }
  });

  s.addText('「軍資金思考」のポイント：先に確保して、残りで生活する', {
    x: 0.42, y: 4.15, w: 12.28, h: 0.52,
    fontSize: 16, color: C.gray, fontFace: FONT_JA
  });
  s.addText('今日の講義では、この軍資金を作る8つの方法をすべてお伝えします', {
    x: 0.42, y: 4.72, w: 12.28, h: 0.52,
    fontSize: 17, bold: true, color: C.orange, fontFace: FONT_JA
  });

  addFooter(s);
})();

// ============================================================
//  スライド 05: 8手法 全体マップ
// ============================================================
(function() {
  const s = pptx.addSlide();
  addBg(s, C.navy);
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: '100%', h: 0.07, fill: { color: C.gold } });

  s.addText('8つの軍資金作成法　全体マップ', {
    x: 0.5, y: 0.15, w: 12.3, h: 0.72,
    fontSize: 26, bold: true, color: C.white, fontFace: FONT_JA
  });

  const methods = [
    { no: '01', title: '先取り貯蓄の自動化',      amt: '月 1〜3万円',    cat: '支出設計',   catC: C.accent },
    { no: '02', title: '固定費の徹底見直し',      amt: '月 5千〜3万円',  cat: '支出設計',   catC: C.accent },
    { no: '03', title: '変動費の最適化',          amt: '月 3千〜1.5万円', cat: '支出設計',  catC: C.accent },
    { no: '04', title: 'キャッシュレス還元の集約', amt: '月 2千〜1万円',  cat: '還元活用',   catC: C.green },
    { no: '05', title: 'ポイントサイト活用',       amt: '月 3千〜3万円',  cat: '収入獲得',   catC: C.orange },
    { no: '06', title: '不用品売却・断捨離',       amt: '初回〜5万円',    cat: '収入獲得',   catC: C.orange },
    { no: '07', title: '副業・スキル収益化',       amt: '月 1〜10万円',   cat: '収入獲得',   catC: C.orange },
    { no: '08', title: '税制優遇制度の活用',       amt: '年 5〜20万円',   cat: '節税',       catC: C.gold },
  ];

  methods.forEach((m, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = col === 0 ? 0.42 : 6.72;
    const y = 1.05 + row * 1.38;

    s.addShape(pptx.ShapeType.rect, { x, y, w: 5.95, h: 1.2, fill: { color: C.navyMid }, line: { color: '3A4A80', pt: 1 } });
    // 番号
    s.addShape(pptx.ShapeType.rect, { x, y, w: 0.58, h: 1.2, fill: { color: C.gold } });
    s.addText(m.no, { x, y: y + 0.3, w: 0.58, h: 0.6, fontSize: 18, bold: true, color: C.navy, fontFace: FONT_EN, align: 'center' });
    // タイトル
    s.addText(m.title, { x: x + 0.65, y: y + 0.1, w: 3.45, h: 0.62, fontSize: 15, bold: true, color: C.white, fontFace: FONT_JA });
    // カテゴリタグ
    s.addShape(pptx.ShapeType.rect, { x: x + 0.65, y: y + 0.74, w: 1.0, h: 0.3, fill: { color: m.catC } });
    s.addText(m.cat, { x: x + 0.65, y: y + 0.74, w: 1.0, h: 0.3, fontSize: 10, color: C.white, fontFace: FONT_JA, align: 'center' });
    // 金額
    s.addText(m.amt, { x: x + 4.15, y: y + 0.32, w: 1.7, h: 0.56, fontSize: 13, bold: true, color: C.gold, fontFace: FONT_JA, align: 'right' });
  });

  addFooter(s);
})();

// ============================================================
//  手法スライド 共通関数（統合版）
//  - チェック行：概念説明 + 実践ポイント
//  - actionBox：今日の一歩（全員向け汎用アクション）
//  - noteBox：補足情報
// ============================================================
function buildMethodSlide(no, title, sub, badge, checks, actionText, noteText) {
  const s = pptx.addSlide();
  addBg(s, C.light);
  methodHeader(s, no, title, badge);

  // サブタイトル
  s.addText(sub, {
    x: 0.5, y: 1.16, w: 12.2, h: 0.42,
    fontSize: 14, color: C.gray, fontFace: FONT_JA, italic: true
  });

  // チェック行
  checks.forEach((c, i) => {
    checkRow(s, 1.68 + i * 0.86, c);
  });

  // 今日の一歩
  const actionY = 1.68 + checks.length * 0.86 + 0.18;
  actionBox(s, actionY, actionText);

  // 補足ノート
  if (noteText) {
    noteBox(s, actionY + 0.88, noteText);
  }

  addFooter(s);
}

// ============================================================
//  スライド 06: 手法①
// ============================================================
buildMethodSlide(
  '01', '先取り貯蓄の自動化', '〜「残ったら貯める」から「先に確保する」への転換〜', '月 1〜3万円',
  [
    '給与振込日の翌日に専用口座・証券口座へ「自動転送」を設定する',
    '手取りの10%を目標に。月収25万円なら2.5万円を"存在しないお金"として扱う',
    '住信SBIネット銀行「目的別口座」や楽天銀行「自動振替」が実用的',
  ],
  '今月中に証券口座または専用貯蓄口座への自動積立をオンに設定する',
  '米国発「Pay Yourself First」── 資産家に共通する、収入管理の世界的鉄則'
);

// ============================================================
//  スライド 07: 手法②
// ============================================================
buildMethodSlide(
  '02', '固定費の徹底見直し', '〜一度やれば毎月ずっと効く"最高コスパ"の節約〜', '月 5千〜3万円',
  [
    'スマートフォン：大手キャリア（〜8,000円）→ 格安SIM（〜1,500円）で月6,500円削減',
    '生命保険：保障が過大になりやすい。FP無料相談で月5,000〜15,000円の削減余地を確認',
    'サブスク棚卸し：使っていないサービスをリストアップ → 月2,000〜5,000円削減',
    '電気・ガス：自由化を活用して切り替えるだけで月1,000〜3,000円の削減も',
  ],
  '今週中にスマホ料金と保険の月額を確認し、見直しを検討する',
  '「削減＝我慢」ではなく「同等のサービスを安く買う」という発想で取り組む'
);

// ============================================================
//  スライド 08: 手法③
// ============================================================
buildMethodSlide(
  '03', '変動費の最適化', '〜ゼロにしなくていい。まず"漏れ"をふさぐだけ〜', '月 3千〜1.5万円',
  [
    '食費：まとめ買い＋冷凍活用。業務スーパー・コストコの活用で月5,000〜10,000円削減',
    '外食費：QRコード決済のポイント還元デー（PayPay等）を活用し、実質割引で外食する',
    '趣味・娯楽：図書館・動画配信の共有プランなどで月2,000〜5,000円の代替',
  ],
  '今月の食費・外食費・サブスクを1週間分だけ記録し「漏れ」を可視化する',
  '米国発「ノースペンドデー」：週に1日を消費ゼロの日にするだけで意識が大きく変わる'
);

// ============================================================
//  スライド 09: 手法④
// ============================================================
buildMethodSlide(
  '04', 'キャッシュレス還元の集約', '〜すでに使っているお金を"ポイント"に変換〜', '月 2千〜1万円',
  [
    '年会費無料×高還元カード1〜2枚に支出を集中（楽天カード最大3%・PayPayカード最大5%）',
    '公共料金・保険料・通販・ガソリンもすべてカード払いに統一する',
    '月30万円支出 × 還元率2% ＝ 月6,000円 → 年72,000円の軍資金が自動的に積み上がる',
    '楽天ポイント投資・PayPay資産運用で「ポイントをそのまま投資に回す」設定も活用可能',
  ],
  '今月中に高還元クレカに支出を集約し、ポイントの投資振替を設定する',
  'ポイントは「貯めるもの」ではなく「現金と同価値の資産」として即活用するのが鉄則'
);

// ============================================================
//  スライド 10: 手法⑤
// ============================================================
buildMethodSlide(
  '05', 'ポイントサイト・アフィリエイト活用', '〜日常のサービス申し込みを"収入"に変える〜', '月 3千〜3万円',
  [
    'A8.net：国内最大のASP。ブログ不要で証券口座・クレカ開設の自己アフィリが可能',
    'ハピタス・モッピー：クレカ発行・口座開設等で1案件あたり数千〜数万ポイント獲得',
    'ポイントタウン：アンケート・ゲーム・ショッピングで毎日コツコツと換金可能ポイントを獲得',
  ],
  '今週中にハピタスまたはモッピーに登録し、証券口座開設の案件を1件経由する',
  '「ポイ活」自体を目的にしない。既定の支出・行動をポイントサイト経由に変えるだけでOK'
);

// ============================================================
//  スライド 11: 手法⑥
// ============================================================
buildMethodSlide(
  '06', '不用品売却・断捨離収益化', '〜眠っている「資産」を現金化する〜', '初回 〜5万円',
  [
    'メルカリ：衣類・雑貨・本。スマホ1枚で写真を撮って即出品。初回断捨離で数万円も',
    'ヤフオク：ブランド品・趣味グッズ・レア品。入札形式で高値がつきやすい',
    'ハードオフ・買取専門店：持ち込み即日現金化。家電・楽器・カメラ類に有効',
  ],
  '今日：家の中を15分で見回し「半年使っていないもの」をスマホでリストアップする',
  '「一度きり」だからこそ種銭作りに最適。副産物として"衝動買い防止"の意識も生まれる'
);

// ============================================================
//  スライド 12: 手法⑦
// ============================================================
buildMethodSlide(
  '07', '副業・スキル収益化', '〜本業以外の「第二の収入源」を作る〜', '月 1〜10万円',
  [
    'クラウドワークス・ランサーズ：ライティング・データ入力・翻訳。スキル不要案件も豊富',
    'ストアカ・ ココナラ：職歴・趣味・専門知識を「教える商品」に変換。月1〜10万円',
    'せどり・転売：仕入れ→販売のサイクルを回す。月3〜20万円の実績者も多い',
    'まず月1〜2万円の達成を目標に。「全額を投資口座に回す」ルールを最初から決める',
  ],
  '今月中：自分の職歴・趣味から「売れるスキル」を1つ書き出してみる',
  '長年のキャリア・人脈・専門知識は「副業の武器」。経験があるほど高単価になりやすい'
);

// ============================================================
//  スライド 13: 手法⑧
// ============================================================
buildMethodSlide(
  '08', '税制優遇制度の活用', '〜払わなくていい税金を取り戻し、投資に回す〜', '年 5〜20万円',
  [
    'iDeCo：掛金が全額所得控除。年収600万円で月2万円拠出すれば年間約4〜6万円の節税',
    'ふるさと納税：実質2,000円の自己負担で返礼品受取＋住民税控除。食費削減にも直結',
    'NISA（新NISA）：運用益・配当が非課税。長期資産形成と組み合わせてフル活用',
    '医療費控除・副業の青色申告：経費計上で課税所得を大幅に圧縮できる',
  ],
  '今週中：ふるさと納税のシミュレーションサイトで今年の上限額を確認し1件注文する',
  '節税の恩恵は制度で上限が決まっているが、使い切るだけで年5〜20万円の軍資金になる'
);

// ============================================================
//  スライド 14: 8手法まとめ（合計）
// ============================================================
(function() {
  const s = pptx.addSlide();
  addBg(s, C.navy);
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: '100%', h: 0.07, fill: { color: C.gold } });

  s.addText('8手法を組み合わせると…', {
    x: 0.5, y: 0.18, w: 12.3, h: 0.72,
    fontSize: 28, bold: true, color: C.white, fontFace: FONT_JA
  });

  const groups = [
    { range: '手法①〜④', label: '支出設計・還元活用（4手法）', amt: '月 1.5〜7万円',  barC: C.accent },
    { range: '手法⑤〜⑦', label: '収入獲得（3手法）',           amt: '月 1.3〜13万円', barC: C.orange },
    { range: '手法⑧',     label: '節税',                       amt: '年 5〜20万円',   barC: C.gold },
  ];

  groups.forEach((g, i) => {
    const y = 1.1 + i * 1.28;
    s.addShape(pptx.ShapeType.rect, { x: 0.42, y, w: 12.28, h: 1.08, fill: { color: C.navyMid }, line: { color: g.barC, pt: 2 } });
    s.addShape(pptx.ShapeType.rect, { x: 0.42, y, w: 0.95, h: 1.08, fill: { color: g.barC } });
    s.addText(g.range, { x: 0.42, y: y + 0.2, w: 0.95, h: 0.68, fontSize: 12, bold: true, color: C.white, fontFace: FONT_JA, align: 'center' });
    s.addText(g.label, { x: 1.5, y: y + 0.22, w: 7.5, h: 0.64, fontSize: 19, color: C.white, fontFace: FONT_JA });
    s.addText(g.amt,   { x: 9.5, y: y + 0.2,  w: 3.1, h: 0.68, fontSize: 22, bold: true, color: C.gold, fontFace: FONT_JA, align: 'right' });
  });

  // 合計ボックス
  s.addShape(pptx.ShapeType.rect, { x: 0.42, y: 4.95, w: 12.28, h: 1.38, fill: { color: C.gold } });
  s.addText('8手法をすべて取り組めば…', { x: 0.62, y: 5.02, w: 5.0, h: 0.5, fontSize: 16, color: C.navy, fontFace: FONT_JA });
  s.addText('月 3〜20万円 の軍資金を作ることが可能', {
    x: 0.62, y: 5.45, w: 12.0, h: 0.72,
    fontSize: 28, bold: true, color: C.navy, fontFace: FONT_JA
  });

  addFooter(s);
})();

// ============================================================
//  スライド 15: ブリッジ「その軍資金、どこに置く？」
// ============================================================
(function() {
  const s = pptx.addSlide();
  addBg(s, C.light);
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: '100%', h: 1.15, fill: { color: C.navy } });
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 1.13, w: '100%', h: 0.05, fill: { color: C.orange } });
  s.addText('さあ、次の問いです', { x: 0.5, y: 0.08, w: 12.3, h: 0.42, fontSize: 16, color: 'B0C4DE', fontFace: FONT_JA });
  s.addText('貯めた軍資金を、どこに置きますか？', {
    x: 0.5, y: 0.48, w: 12.3, h: 0.68,
    fontSize: 30, bold: true, color: C.gold, fontFace: FONT_JA
  });

  const opts = [
    { label: '銀行預金', detail: '金利0.1%\n100万円で年1,000円', verdict: '論外', vc: C.red },
    { label: '長期積立\nインデックス', detail: '優良な選択肢\nただし成果まで20〜30年かかる', verdict: '時間がかかる', vc: C.orange },
    { label: '高配当株\nREIT', detail: '年4〜5%の配当\n月15万円には3,000〜4,500万円が必要', verdict: '元本が足りない', vc: C.orange },
  ];

  opts.forEach((o, i) => {
    const x = 0.42 + i * 4.3;
    s.addShape(pptx.ShapeType.rect, { x, y: 1.38, w: 4.0, h: 3.55, fill: { color: C.white }, line: { color: 'D8DEE6', pt: 1 } });
    s.addShape(pptx.ShapeType.rect, { x, y: 1.38, w: 4.0, h: 0.7, fill: { color: C.accent } });
    s.addText(o.label, { x, y: 1.4, w: 4.0, h: 0.66, fontSize: 17, bold: true, color: C.white, fontFace: FONT_JA, align: 'center' });
    s.addText(o.detail, { x: x + 0.18, y: 2.2, w: 3.64, h: 1.8, fontSize: 14.5, color: C.darkgray, fontFace: FONT_JA, align: 'center' });
    s.addShape(pptx.ShapeType.rect, { x: x + 0.3, y: 4.2, w: 3.4, h: 0.56, fill: { color: o.vc } });
    s.addText('→ ' + o.verdict, { x: x + 0.3, y: 4.2, w: 3.4, h: 0.56, fontSize: 14, bold: true, color: C.white, fontFace: FONT_JA, align: 'center' });
  });

  s.addShape(pptx.ShapeType.rect, { x: 0.42, y: 5.18, w: 12.28, h: 0.82, fill: { color: C.navy } });
  s.addText('「もし、もっと早く・少ない元手で、月数万〜15万円を得られる方法があったとしたら？」', {
    x: 0.5, y: 5.22, w: 12.1, h: 0.74,
    fontSize: 18, bold: true, color: C.gold, fontFace: FONT_JA, align: 'center'
  });

  addFooter(s);
})();

// ============================================================
//  スライド 16: 米国株オプション取引とは
// ============================================================
(function() {
  const s = pptx.addSlide();
  addBg(s, C.navy);
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: '100%', h: 0.07, fill: { color: C.gold } });

  s.addText('解決策：米国株オプション取引', {
    x: 0.5, y: 0.18, w: 12.3, h: 0.72,
    fontSize: 28, bold: true, color: C.white, fontFace: FONT_JA
  });
  s.addText('〜「買う・売る」の2択から、「権利を売買する」世界へ〜', {
    x: 0.5, y: 0.9, w: 12.3, h: 0.42,
    fontSize: 15, color: 'AABBCC', fontFace: FONT_JA
  });

  const points = [
    {
      no: '01', title: '毎月収益が狙える仕組み',
      detail: 'オプションの時間的価値を「売る」ことで、株価が大きく動かなくても毎月プレミアムを受け取ることができる'
    },
    {
      no: '02', title: '長期投資と並行して活用できる',
      detail: '保有株を活用したオプション戦略（カバードコール等）は、既存の資産と組み合わせて効率的に運用できる'
    },
    {
      no: '03', title: '月数万〜15万円の収益実績',
      detail: '元本100〜300万円から始め、月5〜15万円の収益を上げているメンバーが多数。実際の事例はセミナーでご紹介'
    },
  ];

  points.forEach((p, i) => {
    const y = 1.5 + i * 1.58;
    s.addShape(pptx.ShapeType.rect, { x: 0.42, y, w: 12.28, h: 1.38, fill: { color: C.navyMid }, line: { color: C.gold, pt: 1 } });
    s.addShape(pptx.ShapeType.rect, { x: 0.42, y, w: 0.82, h: 1.38, fill: { color: C.gold } });
    s.addText(p.no, { x: 0.42, y: y + 0.35, w: 0.82, h: 0.68, fontSize: 24, bold: true, color: C.navy, fontFace: FONT_EN, align: 'center' });
    s.addText(p.title,  { x: 1.38, y: y + 0.1,  w: 11.2, h: 0.58, fontSize: 20, bold: true, color: C.gold,  fontFace: FONT_JA });
    s.addText(p.detail, { x: 1.38, y: y + 0.7,  w: 11.2, h: 0.58, fontSize: 14, color: 'CCDDEE', fontFace: FONT_JA });
  });

  addFooter(s);
})();

// ============================================================
//  スライド 17: セミナー案内（CTA）
// ============================================================
(function() {
  const s = pptx.addSlide();
  addBg(s, C.navy);
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 6.25, w: '100%', h: 0.07, fill: { color: C.gold } });
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 6.32, w: '100%', h: 0.04, fill: { color: C.orange } });

  s.addText('次のステップ', { x: 0.5, y: 0.18, w: 12.3, h: 0.52, fontSize: 20, color: 'B0C4DE', fontFace: FONT_JA, align: 'center' });

  s.addShape(pptx.ShapeType.rect, { x: 1.3, y: 0.75, w: 10.7, h: 1.2, fill: { color: C.gold } });
  s.addText('高速資産形成セミナー', {
    x: 1.3, y: 0.75, w: 10.7, h: 1.2,
    fontSize: 40, bold: true, color: C.navy, fontFace: FONT_JA, align: 'center'
  });

  s.addText('米国株オプション取引の詳細を無料でお伝えします', {
    x: 0.5, y: 2.05, w: 12.3, h: 0.6,
    fontSize: 20, color: C.white, fontFace: FONT_JA, align: 'center'
  });

  const items = [
    '✅  軍資金を作りながら「同時に増やす仕組み」の全体像を解説',
    '✅  月5〜15万円を実現しているメンバーの実例を紹介',
    '✅  オプション取引の基礎〜実践戦略まで丁寧にお伝え',
    '✅  個別相談・質疑応答あり（先着順）',
  ];
  items.forEach((item, i) => {
    s.addText(item, {
      x: 1.8, y: 2.82 + i * 0.58, w: 9.7, h: 0.52,
      fontSize: 16.5, color: C.white, fontFace: FONT_JA
    });
  });

  s.addShape(pptx.ShapeType.rect, { x: 2.2, y: 5.25, w: 8.9, h: 0.82, fill: { color: C.orange } });
  s.addText('▶  まずは無料セミナーにご参加ください', {
    x: 2.2, y: 5.25, w: 8.9, h: 0.82,
    fontSize: 21, bold: true, color: C.white, fontFace: FONT_JA, align: 'center'
  });

  // 会社名
  s.addText(COMPANY, {
    x: 0.5, y: 6.1, w: 12.3, h: 0.3,
    fontSize: 11, color: 'AABBCC', fontFace: FONT_JA, align: 'center', italic: true
  });
})();

// ============================================================
//  スライド 18: まとめ・3ステップロードマップ
// ============================================================
(function() {
  const s = pptx.addSlide();
  addBg(s, C.light);
  sectionTitle(s, 'まとめ：今日から始める3ステップ');

  const steps = [
    {
      step: 'STEP 1', sub: '今すぐ（〜1ヶ月）', color: C.accent,
      items: ['先取り貯蓄の自動化を設定', '固定費の見直し（スマホ・保険）', 'キャッシュレスを1〜2枚に集約']
    },
    {
      step: 'STEP 2', sub: '仕組みを整える（1〜3ヶ月）', color: C.green,
      items: ['ふるさと納税 / iDeCo を申込む', 'ポイントサイト登録・自己アフィ活用', '断捨離・不用品を売却し種銭を作る']
    },
    {
      step: 'STEP 3', sub: '加速させる（3ヶ月〜）', color: C.orange,
      items: ['副業をスタート（小さく始める）', '全手法の収益を投資口座に自動移動', '「増やす仕組み」を手に入れる']
    },
  ];

  steps.forEach((st, i) => {
    const x = 0.38 + i * 4.3;
    s.addShape(pptx.ShapeType.rect, { x, y: 1.35, w: 4.0, h: 5.15, fill: { color: C.white }, line: { color: 'D8DEE6', pt: 1 } });
    s.addShape(pptx.ShapeType.rect, { x, y: 1.35, w: 4.0, h: 1.05, fill: { color: st.color } });
    s.addText(st.step, { x, y: 1.38, w: 4.0, h: 0.58, fontSize: 21, bold: true, color: C.white, fontFace: FONT_EN, align: 'center' });
    s.addText(st.sub,  { x, y: 1.9,  w: 4.0, h: 0.42, fontSize: 11, color: C.white, fontFace: FONT_JA, align: 'center' });

    st.items.forEach((item, j) => {
      const iy = 2.6 + j * 0.96;
      s.addShape(pptx.ShapeType.rect, { x: x + 0.2, y: iy,        w: 0.48, h: 0.48, fill: { color: st.color } });
      s.addText('✓', {               x: x + 0.2, y: iy,        w: 0.48, h: 0.48, fontSize: 15, bold: true, color: C.white, fontFace: FONT_EN, align: 'center' });
      s.addText(item, {               x: x + 0.78, y: iy,        w: 3.1,  h: 0.82, fontSize: 13, color: C.darkgray, fontFace: FONT_JA });
    });
  });

  // 締めメッセージ
  s.addShape(pptx.ShapeType.rect, { x: 0.38, y: 6.62, w: 12.52, h: 0.58, fill: { color: C.navy } });
  s.addText('軍資金を作りながら、同時に増やす仕組みを手に入れる。それが「高速資産形成」です。  ─  ' + COMPANY, {
    x: 0.38, y: 6.62, w: 12.52, h: 0.58,
    fontSize: 13, bold: true, color: C.gold, fontFace: FONT_JA, align: 'center'
  });

  addFooter(s);
})();

// ============================================================
//  保存
// ============================================================
const outputPath = path.join('D:/dev/軍資金作成資料/docs', '軍資金作成講義_統合版スライド.pptx');
pptx.writeFile({ fileName: outputPath })
  .then(() => console.log('✅ PowerPoint saved:', outputPath))
  .catch(err => console.error('❌ Error:', err));
