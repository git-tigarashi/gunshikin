// ============================================================
//  PowerPoint スライド生成スクリプト v2
//  タイトル：投資で成功する人が最初にやっている「軍資金の作り方」
//  株式会社 Vision Creator
//  ※セミナー誘導なし・純講義コンテンツ版
// ============================================================
const PptxGenJS = require('C:/Users/haman/AppData/Roaming/npm/node_modules/pptxgenjs');
const path = require('path');

const pptx = new PptxGenJS();
pptx.layout = 'LAYOUT_WIDE';

// ========== カラーパレット ==========
const C = {
  navy:    '1A2744',
  gold:    'D4A843',
  orange:  'E8602C',
  white:   'FFFFFF',
  light:   'F4F6FA',
  gray:    '7A8A9A',
  darkgray:'2C2C2C',
  green:   '2E7D32',
  red:     'C62828',
  accent:  '1565C0',
  navyMid: '243060',
  silver:  'B0BEC5',
};

const FONT_JA  = 'メイリオ';
const FONT_EN  = 'Calibri';
const COMPANY  = '株式会社 Vision Creator';

// ========== ユーティリティ ==========
function bg(s, color) {
  s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:'100%', h:'100%', fill:{color} });
}

function footer(s) {
  s.addText(COMPANY, {
    x:0.35, y:6.88, w:5.8, h:0.28,
    fontSize:9, color:C.gray, fontFace:FONT_JA, italic:true
  });
  s.addText('投資で成功する人が最初にやっている「軍資金の作り方」', {
    x:6.3, y:6.88, w:7.0, h:0.28,
    fontSize:9, color:C.gray, fontFace:FONT_JA, align:'right'
  });
}

// 縦バー付きセクション見出し
function h1(s, text, sub, y, barC) {
  y    = y    || 0.42;
  barC = barC || C.gold;
  s.addShape(pptx.ShapeType.rect, { x:0.4, y, w:0.09, h:0.72, fill:{color:barC} });
  s.addText(text, {
    x:0.65, y, w:12.1, h:0.72,
    fontSize:28, bold:true, color:C.navy, fontFace:FONT_JA
  });
  if (sub) s.addText(sub, {
    x:0.65, y:y+0.76, w:12.1, h:0.38,
    fontSize:13, color:C.gray, fontFace:FONT_JA
  });
}

// ダーク背景ヘッダー（手法スライド）
function methodHeader(s, no, title, badge) {
  s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:'100%', h:1.1, fill:{color:C.navy} });
  s.addShape(pptx.ShapeType.rect, { x:0, y:1.08, w:'100%', h:0.05, fill:{color:C.gold} });
  s.addShape(pptx.ShapeType.rect, { x:0.4, y:0.15, w:0.82, h:0.82, fill:{color:C.gold} });
  s.addText(no, { x:0.4, y:0.15, w:0.82, h:0.82, fontSize:30, bold:true, color:C.navy, fontFace:FONT_EN, align:'center' });
  s.addText(title, { x:1.38, y:0.18, w:8.6, h:0.78, fontSize:26, bold:true, color:C.white, fontFace:FONT_JA });
  s.addShape(pptx.ShapeType.rect, { x:10.15, y:0.18, w:2.55, h:0.72, fill:{color:C.orange} });
  s.addText(badge, { x:10.15, y:0.18, w:2.55, h:0.72, fontSize:16, bold:true, color:C.white, fontFace:FONT_JA, align:'center' });
}

// チェックリスト行
function checkRow(s, y, text, iconC) {
  iconC = iconC || C.accent;
  s.addShape(pptx.ShapeType.rect, { x:0.42, y:y+0.09, w:0.48, h:0.48, fill:{color:iconC} });
  s.addText('✓', { x:0.42, y:y+0.09, w:0.48, h:0.48, fontSize:16, bold:true, color:C.white, fontFace:FONT_EN, align:'center' });
  s.addText(text, { x:1.04, y, w:11.65, h:0.66, fontSize:15.5, color:C.darkgray, fontFace:FONT_JA });
}

// 「今日の一歩」アクションボックス
function actionBox(s, y, text) {
  s.addShape(pptx.ShapeType.rect, { x:0.42, y, w:12.28, h:0.72, fill:{color:'FFF8E1'}, line:{color:C.gold, pt:2} });
  s.addShape(pptx.ShapeType.rect, { x:0.42, y, w:1.55, h:0.72, fill:{color:C.gold} });
  s.addText('今日の一歩', { x:0.42, y:y+0.12, w:1.55, h:0.48, fontSize:13, bold:true, color:C.navy, fontFace:FONT_JA, align:'center' });
  s.addText(text, { x:2.08, y:y+0.12, w:10.5, h:0.48, fontSize:14, color:C.darkgray, fontFace:FONT_JA });
}

// 補足ノートボックス
function noteBox(s, y, text, h) {
  h = h || 0.62;
  s.addShape(pptx.ShapeType.rect, { x:0.42, y, w:12.28, h, fill:{color:'EEF3FB'}, line:{color:C.accent, pt:1} });
  s.addText(text, { x:0.58, y:y+0.08, w:12.0, h:h-0.16, fontSize:13.5, color:C.accent, fontFace:FONT_JA, bold:true });
}

// 引用ブロック（フック・メッセージ表示用）
function quoteBox(s, y, text, w, x) {
  w = w || 12.28; x = x || 0.42;
  s.addShape(pptx.ShapeType.rect, { x, y, w, h:0.08, fill:{color:C.gold} });
  s.addShape(pptx.ShapeType.rect, { x, y:y+0.08, w:0.06, h:0.74, fill:{color:C.gold} });
  s.addText(text, { x:x+0.2, y:y+0.1, w:w-0.28, h:0.7, fontSize:17, color:C.navy, fontFace:FONT_JA, bold:true, italic:true });
}

// ============================================================
//  スライド 01: 表紙
// ============================================================
(function() {
  const s = pptx.addSlide();
  bg(s, C.navy);

  // 上下アクセントライン
  s.addShape(pptx.ShapeType.rect, { x:0, y:0,    w:'100%', h:0.08, fill:{color:C.gold} });
  s.addShape(pptx.ShapeType.rect, { x:0, y:6.3,  w:'100%', h:0.08, fill:{color:C.gold} });
  s.addShape(pptx.ShapeType.rect, { x:0, y:6.38, w:'100%', h:0.04, fill:{color:C.orange} });

  // サブテーマライン群
  const subs = [
    '40代からでも間に合う資産形成の始め方',
    '資産形成の9割は「投資前」で決まる',
    'なぜ多くの人は投資で結果が出ないのか',
  ];
  subs.forEach((t, i) => {
    s.addShape(pptx.ShapeType.rect, { x:1.6, y:0.38+i*0.42, w:9.5, h:0.32, fill:{color:'0D1830'} });
    s.addText(t, { x:1.65, y:0.42+i*0.42, w:9.4, h:0.28, fontSize:12, color:C.silver, fontFace:FONT_JA });
  });

  // メインタイトル
  s.addText('投資で成功する人が', {
    x:0.6, y:1.72, w:12.1, h:0.9,
    fontSize:44, bold:true, color:C.white, fontFace:FONT_JA, align:'center'
  });
  s.addText('最初にやっている', {
    x:0.6, y:2.58, w:12.1, h:0.9,
    fontSize:44, bold:true, color:C.white, fontFace:FONT_JA, align:'center'
  });

  // 強調ワード
  s.addShape(pptx.ShapeType.rect, { x:1.5, y:3.48, w:10.3, h:1.3, fill:{color:C.gold} });
  s.addText('「軍資金」の作り方', {
    x:1.5, y:3.48, w:10.3, h:1.3,
    fontSize:54, bold:true, color:C.navy, fontFace:FONT_JA, align:'center'
  });

  // 会社名
  s.addShape(pptx.ShapeType.rect, { x:3.8, y:4.98, w:5.7, h:0.78, fill:{color:C.navyMid} });
  s.addText(COMPANY, {
    x:3.8, y:4.98, w:5.7, h:0.78,
    fontSize:20, bold:true, color:C.gold, fontFace:FONT_JA, align:'center'
  });

  s.addText(COMPANY, {
    x:0, y:6.55, w:'100%', h:0.3,
    fontSize:10, color:C.silver, fontFace:FONT_JA, align:'center', italic:true
  });
})();

// ============================================================
//  スライド 02: 本日のテーマ・目次
// ============================================================
(function() {
  const s = pptx.addSlide();
  bg(s, C.light);
  h1(s, '本日お伝えすること', '〜 投資を始める前に知っておくべき話 〜');

  // 左カラム：問いかけ3題
  const qs = [
    { q: 'なぜ多くの人は\n投資で結果が出ないのか？', y:1.42, color:C.red },
    { q: '資産形成は銘柄ではなく\n「ここ」で決まる', y:2.88, color:C.orange },
    { q: '40代からの\n資産形成の正しい順番', y:4.34, color:C.accent },
  ];

  qs.forEach(q => {
    s.addShape(pptx.ShapeType.rect, { x:0.42, y:q.y, w:5.5, h:1.22, fill:{color:q.color}, line:{color:q.color, pt:1} });
    s.addText(q.q, { x:0.52, y:q.y+0.12, w:5.3, h:0.98, fontSize:17, bold:true, color:C.white, fontFace:FONT_JA });
  });

  // 右カラム：本日の講義内容
  s.addShape(pptx.ShapeType.rect, { x:6.2, y:1.42, w:6.5, h:5.2, fill:{color:C.white}, line:{color:'D8DEE6', pt:1} });
  s.addShape(pptx.ShapeType.rect, { x:6.2, y:1.42, w:6.5, h:0.62, fill:{color:C.navy} });
  s.addText('本日の講義内容', { x:6.2, y:1.42, w:6.5, h:0.62, fontSize:16, bold:true, color:C.gold, fontFace:FONT_JA, align:'center' });

  const items = [
    'Part 1  投資で失敗する人・成功する人の違い',
    'Part 2  「軍資金」という発想の転換',
    'Part 3  軍資金を作る8つの方法',
    '   ① 先取り貯蓄の自動化',
    '   ② 固定費の徹底見直し',
    '   ③ 変動費の最適化',
    '   ④ キャッシュレス還元の集約',
    '   ⑤ ポイントサイト活用',
    '   ⑥ 不用品売却・断捨離',
    '   ⑦ 副業・スキル収益化',
    '   ⑧ 税制優遇制度の活用',
    'Part 4  40代からの資産形成ロードマップ',
  ];

  items.forEach((item, i) => {
    const isMain = item.startsWith('Part');
    s.addText(item, {
      x:6.4, y:2.16+i*0.34, w:6.1, h:0.32,
      fontSize: isMain ? 13 : 12,
      bold: isMain,
      color: isMain ? C.navy : C.darkgray,
      fontFace:FONT_JA
    });
  });

  footer(s);
})();

// ============================================================
//  スライド 03: なぜ多くの人は投資で結果が出ないのか
// ============================================================
(function() {
  const s = pptx.addSlide();
  bg(s, C.navy);
  s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:'100%', h:0.08, fill:{color:C.red} });

  s.addText('なぜ多くの人は投資で結果が出ないのか？', {
    x:0.5, y:0.2, w:12.3, h:0.78,
    fontSize:28, bold:true, color:C.white, fontFace:FONT_JA
  });

  // 3つの理由
  const reasons = [
    {
      no: '01',
      title: '「銘柄選び」に集中しすぎる',
      detail: '「どの株を買えば儲かるか」ばかり考える。しかし利益の大半は銘柄ではなく元本の大きさで決まる。',
      icon: '📈',
    },
    {
      no: '02',
      title: '元本（軍資金）が少なすぎる',
      detail: '10万円を運用して年5%の利益 ＝ 5,000円。\n100万円なら50,000円。同じ努力でも10倍の差になる。',
      icon: '💰',
    },
    {
      no: '03',
      title: '正しい「順番」を知らずに始める',
      detail: '投資の知識を得る前に、軍資金を作る準備をする。この順番を間違えると、知識があっても動けない。',
      icon: '🔢',
    },
  ];

  reasons.forEach((r, i) => {
    const y = 1.18 + i * 1.72;
    s.addShape(pptx.ShapeType.rect, { x:0.42, y, w:12.28, h:1.52, fill:{color:C.navyMid}, line:{color:C.orange, pt:1} });
    s.addShape(pptx.ShapeType.rect, { x:0.42, y, w:0.82, h:1.52, fill:{color:C.orange} });
    s.addText(r.no, { x:0.42, y:y+0.42, w:0.82, h:0.68, fontSize:26, bold:true, color:C.white, fontFace:FONT_EN, align:'center' });
    s.addText(r.title, { x:1.38, y:y+0.1, w:11.1, h:0.58, fontSize:20, bold:true, color:C.gold, fontFace:FONT_JA });
    s.addText(r.detail, { x:1.38, y:y+0.68, w:11.1, h:0.74, fontSize:13.5, color:'CCDDEE', fontFace:FONT_JA });
  });

  quoteBox(s, 6.3, '資産形成は銘柄ではなく「ここ」で決まる ── それが今日の本題です');
  footer(s);
})();

// ============================================================
//  スライド 04: 資産形成の9割は「投資前」で決まる
// ============================================================
(function() {
  const s = pptx.addSlide();
  bg(s, C.light);
  h1(s, '資産形成の9割は「投資前」で決まる');

  // インパクト比較
  s.addShape(pptx.ShapeType.rect, { x:0.42, y:1.42, w:12.28, h:0.62, fill:{color:C.navy} });
  s.addText('同じ利回りでも「元本」の大きさで結果がまったく変わる', {
    x:0.5, y:1.44, w:12.1, h:0.58, fontSize:18, bold:true, color:C.gold, fontFace:FONT_JA, align:'center'
  });

  // 比較カード
  const cards = [
    { label: '元本 10万円', rate: '年利5%', profit: '5,000円/年', note: '毎月コーヒー代にもならない', color: C.gray },
    { label: '元本 100万円', rate: '年利5%', profit: '50,000円/年', note: '月4,000円超の追加収入', color: C.accent },
    { label: '元本 300万円', rate: '年利5%', profit: '150,000円/年', note: '月12,500円の安定収益', color: C.green },
    { label: '元本 500万円', rate: '年利5%', profit: '250,000円/年', note: '月2万円超のキャッシュフロー', color: C.orange },
  ];

  cards.forEach((c, i) => {
    const x = 0.42 + i * 3.1;
    s.addShape(pptx.ShapeType.rect, { x, y:2.22, w:2.95, h:3.1, fill:{color:C.white}, line:{color:'D8DEE6', pt:1} });
    s.addShape(pptx.ShapeType.rect, { x, y:2.22, w:2.95, h:0.62, fill:{color:c.color} });
    s.addText(c.label, { x, y:2.24, w:2.95, h:0.58, fontSize:15, bold:true, color:C.white, fontFace:FONT_JA, align:'center' });
    s.addText(c.rate,   { x, y:2.96, w:2.95, h:0.48, fontSize:13, color:C.gray, fontFace:FONT_JA, align:'center' });
    s.addText(c.profit, { x, y:3.42, w:2.95, h:0.88, fontSize:22, bold:true, color:c.color, fontFace:FONT_JA, align:'center' });
    s.addText(c.note,   { x:x+0.12, y:4.3, w:2.7, h:0.88, fontSize:12.5, color:C.gray, fontFace:FONT_JA, align:'center' });
  });

  // 結論
  s.addShape(pptx.ShapeType.rect, { x:0.42, y:5.5, w:12.28, h:0.82, fill:{color:C.navy} });
  s.addText('投資の「結果」は技術よりも「元本の大きさ」で8〜9割が決まる。だから最初に軍資金を作る。', {
    x:0.55, y:5.52, w:12.0, h:0.78, fontSize:17, bold:true, color:C.gold, fontFace:FONT_JA, align:'center'
  });

  footer(s);
})();

// ============================================================
//  スライド 05: 「軍資金」という発想の転換
// ============================================================
(function() {
  const s = pptx.addSlide();
  bg(s, C.light);
  h1(s, '「軍資金」という発想の転換', '〜 資産形成が早い人は最初にこれをやる 〜');

  // Before/After
  const baf = [
    { label:'Before\n（ほとんどの人）', desc:'給料 → 生活費をすべて使う\n→ 残ったら（ほぼ残らず）投資へ', color:C.gray },
    { label:'After\n（軍資金思考）',    desc:'給料 → 先に軍資金を確保\n→ 残りで生活 → 軍資金を投資へ', color:C.green },
  ];

  baf.forEach((b, i) => {
    const x = i === 0 ? 0.42 : 6.72;
    s.addShape(pptx.ShapeType.rect, { x, y:1.42, w:5.95, h:2.38, fill:{color:C.white}, line:{color:'D8DEE6', pt:1} });
    s.addShape(pptx.ShapeType.rect, { x, y:1.42, w:5.95, h:0.68, fill:{color:b.color} });
    s.addText(b.label, { x, y:1.44, w:5.95, h:0.64, fontSize:16, bold:true, color:C.white, fontFace:FONT_JA, align:'center' });
    s.addText(b.desc,  { x:x+0.2, y:2.2, w:5.55, h:1.46, fontSize:15.5, color:C.darkgray, fontFace:FONT_JA });
    if (i===0) {
      s.addShape(pptx.ShapeType.rect, { x:6.4, y:2.28, w:0.42, h:0.62, fill:{color:C.gold} });
      s.addText('▶', { x:6.4, y:2.28, w:0.42, h:0.62, fontSize:20, bold:true, color:C.navy, fontFace:FONT_EN, align:'center' });
    }
  });

  // 3つのポイント
  s.addShape(pptx.ShapeType.rect, { x:0.42, y:4.0, w:12.28, h:0.52, fill:{color:C.navyMid} });
  s.addText('軍資金思考の3つのポイント', { x:0.52, y:4.02, w:12.0, h:0.48, fontSize:15, bold:true, color:C.gold, fontFace:FONT_JA });

  const pts = [
    '① 「残ったら貯める」ではなく「先に確保して、残りで生活する」という順序に変える',
    '② 軍資金＝投資専用口座を別に持ち、生活費と完全に分離する',
    '③ 40代からでも間に合う。月3〜5万円の積み上げが、数年後の大きな元本になる',
  ];
  pts.forEach((p, i) => {
    s.addShape(pptx.ShapeType.rect, { x:0.42, y:4.62+i*0.58, w:0.48, h:0.42, fill:{color:C.gold} });
    s.addText(['①','②','③'][i], { x:0.42, y:4.66+i*0.58, w:0.48, h:0.38, fontSize:14, bold:true, color:C.navy, fontFace:FONT_JA, align:'center' });
    s.addText(p.slice(2), { x:1.02, y:4.62+i*0.58, w:11.68, h:0.52, fontSize:14.5, color:C.darkgray, fontFace:FONT_JA });
  });

  footer(s);
})();

// ============================================================
//  スライド 06: 8手法の全体マップ
// ============================================================
(function() {
  const s = pptx.addSlide();
  bg(s, C.navy);
  s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:'100%', h:0.08, fill:{color:C.gold} });

  s.addText('軍資金を作る8つの方法　全体マップ', {
    x:0.5, y:0.15, w:12.3, h:0.72, fontSize:26, bold:true, color:C.white, fontFace:FONT_JA
  });
  s.addText('まずは全体を俯瞰してから、順番に詳しく見ていきます', {
    x:0.5, y:0.85, w:12.3, h:0.36, fontSize:14, color:C.silver, fontFace:FONT_JA
  });

  const methods = [
    { no:'01', title:'先取り貯蓄の自動化',      amt:'月 1〜3万円',     cat:'支出設計', catC:C.accent },
    { no:'02', title:'固定費の徹底見直し',      amt:'月 5千〜3万円',   cat:'支出設計', catC:C.accent },
    { no:'03', title:'変動費の最適化',          amt:'月 3千〜1.5万円', cat:'支出設計', catC:C.accent },
    { no:'04', title:'キャッシュレス還元の集約', amt:'月 2千〜1万円',   cat:'還元活用', catC:C.green },
    { no:'05', title:'ポイントサイト活用',       amt:'月 3千〜3万円',   cat:'収入獲得', catC:C.orange },
    { no:'06', title:'不用品売却・断捨離',       amt:'初回 〜5万円',    cat:'収入獲得', catC:C.orange },
    { no:'07', title:'副業・スキル収益化',       amt:'月 1〜10万円',    cat:'収入獲得', catC:C.orange },
    { no:'08', title:'税制優遇制度の活用',       amt:'年 5〜20万円',    cat:'節税',     catC:C.gold },
  ];

  methods.forEach((m, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = col === 0 ? 0.42 : 6.72;
    const y = 1.32 + row * 1.35;
    s.addShape(pptx.ShapeType.rect, { x, y, w:5.95, h:1.18, fill:{color:C.navyMid}, line:{color:'3A4A80', pt:1} });
    s.addShape(pptx.ShapeType.rect, { x, y, w:0.58, h:1.18, fill:{color:C.gold} });
    s.addText(m.no, { x, y:y+0.29, w:0.58, h:0.6, fontSize:18, bold:true, color:C.navy, fontFace:FONT_EN, align:'center' });
    s.addText(m.title, { x:x+0.66, y:y+0.1, w:3.4, h:0.62, fontSize:15, bold:true, color:C.white, fontFace:FONT_JA });
    s.addShape(pptx.ShapeType.rect, { x:x+0.66, y:y+0.74, w:1.0, h:0.3, fill:{color:m.catC} });
    s.addText(m.cat, { x:x+0.66, y:y+0.74, w:1.0, h:0.3, fontSize:10, color:C.white, fontFace:FONT_JA, align:'center' });
    s.addText(m.amt, { x:x+4.18, y:y+0.32, w:1.68, h:0.54, fontSize:13, bold:true, color:C.gold, fontFace:FONT_JA, align:'right' });
  });

  footer(s);
})();

// ============================================================
//  手法スライド共通関数（v2）
// ============================================================
function methodSlide(no, title, sub, badge, checks, actionText, noteText) {
  const s = pptx.addSlide();
  bg(s, C.light);
  methodHeader(s, no, title, badge);
  s.addText(sub, { x:0.5, y:1.16, w:12.2, h:0.42, fontSize:14, color:C.gray, fontFace:FONT_JA, italic:true });
  checks.forEach((c, i) => checkRow(s, 1.68 + i * 0.86, c));
  const ay = 1.68 + checks.length * 0.86 + 0.18;
  actionBox(s, ay, actionText);
  if (noteText) noteBox(s, ay + 0.88, noteText);
  footer(s);
}

// ============================================================
//  スライド 07〜14: 8手法
// ============================================================
methodSlide('01','先取り貯蓄の自動化',
  '〜「残ったら貯める」から「先に確保する」への転換〜','月 1〜3万円',
  [
    '給与振込日の翌日に証券口座・投資専用口座へ「自動転送」を設定する',
    '手取りの10%を目標に。月収25万円なら2.5万円を"存在しないお金"として管理する',
    '住信SBIネット銀行「目的別口座」・楽天銀行「自動振替」が実用的で設定も簡単',
  ],
  '今月中に投資専用口座への自動積立を設定する（10分でできる）',
  '米国発「Pay Yourself First」── 資産家に共通する収入管理の世界的鉄則'
);

methodSlide('02','固定費の徹底見直し',
  '〜一度やれば毎月ずっと効く"最高コスパ"の節約〜','月 5千〜3万円',
  [
    'スマートフォン：大手（〜8,000円）→ 格安SIM（〜1,500円）で月6,500円を永続削減',
    '生命保険：40代以降は保障過大になりやすい。FP無料相談で月5,000〜15,000円の削減余地を確認',
    'サブスク棚卸し：使っていないサービスを洗い出して解約。月2,000〜5,000円の削減',
    '電気・ガス：切り替えるだけで月1,000〜3,000円。一度やれば永続効果',
  ],
  '今週中にスマホ料金・保険の月額を確認し、見直しの優先順位を決める',
  '「削減＝我慢」ではなく「同等のサービスを安く買う」という発想で取り組む'
);

methodSlide('03','変動費の最適化',
  '〜ゼロにしなくていい。まず"漏れ"をふさぐだけ〜','月 3千〜1.5万円',
  [
    '食費：まとめ買い＋冷凍活用。業務スーパー・コストコの活用で月5,000〜10,000円削減',
    '外食費：QRコード決済のポイント還元デーを活用し、実質割引で外食する習慣をつくる',
    '趣味・娯楽：図書館・動画配信の共有プランなどで月2,000〜5,000円の代替を検討する',
  ],
  '今月の食費・外食費を1週間分だけ記録し「気づかない漏れ」を可視化する',
  '米国発「ノースペンドデー」：週1日を消費ゼロの日にするだけで支出意識が変わる'
);

methodSlide('04','キャッシュレス還元の集約',
  '〜すでに使っているお金を"ポイント"に変換〜','月 2千〜1万円',
  [
    '年会費無料×高還元カード1〜2枚に支出を集中（楽天カード最大3%・PayPayカード最大5%）',
    '公共料金・保険料・通販・ガソリンもすべてカード払いに統一してポイントを積み上げる',
    '月30万円の支出 × 還元率2% ＝ 月6,000円 → 年72,000円の軍資金が自動的に積み上がる',
    '楽天ポイント投資・PayPay資産運用で「ポイントをそのまま投資に回す」設定も活用可能',
  ],
  '今月中に高還元クレカへ支出を集約し、ポイントを即換金・投資振替する設定を行う',
  'ポイントは「貯めるもの」ではなく「現金と同価値の資産」として即活用するのが鉄則'
);

methodSlide('05','ポイントサイト・アフィリエイト活用',
  '〜日常のサービス申し込みを"収入"に変える〜','月 3千〜3万円',
  [
    'A8.net：国内最大のASP。ブログ不要で証券口座・クレカ開設の自己アフィリエイトが可能',
    'ハピタス・モッピー：クレカ発行・口座開設等の高額案件で1件あたり数千〜数万ポイント獲得',
    'ポイントタウン：日常のショッピング・アンケート・ゲームでコツコツと換金可能ポイントを獲得',
  ],
  '今週中にハピタスまたはモッピーに無料登録し、証券口座開設の案件を1件経由する',
  '「ポイ活」を目的にするのではなく、必要な行動・申し込みをポイントサイト経由にするだけ'
);

methodSlide('06','不用品売却・断捨離収益化',
  '〜眠っている「資産」を現金化し、最初の種銭を作る〜','初回 〜5万円',
  [
    'メルカリ：衣類・雑貨・本。スマホ1台で写真を撮って即出品。初回断捨離で数万円になることも',
    'ヤフオク：ブランド品・趣味グッズ・レア品。入札形式で市場価格に近い高値がつきやすい',
    'ハードオフ・買取専門店：家電・楽器・カメラは持ち込み即日現金化。査定は無料',
  ],
  '今日：家の中を15分で見回し「半年使っていないもの」をスマホでリストアップする',
  '「一度きり」だからこそ最初の種銭作りに最適。副産物として"衝動買い防止"の意識も生まれる'
);

methodSlide('07','副業・スキル収益化',
  '〜本業以外の「第二の収入源」を軍資金に変える〜','月 1〜10万円',
  [
    'クラウドワークス・ランサーズ：ライティング・データ入力・翻訳。スキル不要の案件も多数あり',
    'ストアカ・ ココナラ：職歴・趣味・専門知識を「教える商品」に変換。高単価になりやすい',
    'せどり・転売：仕入れ→販売のサイクル。月3〜20万円の実績者が多い初期コスト低の副業',
    '大原則：副業収益は生活費に使わず「全額を投資口座に入金する」ルールを最初に決める',
  ],
  '今月中：自分の職歴・趣味から「売れるスキル・経験」を1つ書き出してみる',
  'キャリア・人脈・専門知識は「副業の武器」。40代以降ほど高単価になりやすい傾向がある'
);

methodSlide('08','税制優遇制度の活用',
  '〜払わなくていい税金を取り戻し、投資原資に回す〜','年 5〜20万円',
  [
    'iDeCo：掛金が全額所得控除。年収600万円で月2万円拠出なら年間約4〜6万円の節税効果',
    'ふるさと納税：実質2,000円の自己負担で返礼品受取＋住民税控除。食費削減にも直結する',
    '新NISA：運用益・配当が非課税。長期資産形成の土台として併用するのが基本形',
    '副業がある場合は青色申告：最大65万円の特別控除。経費計上で課税所得を大幅に圧縮できる',
  ],
  '今週中：ふるさと納税のシミュレーションで今年の上限額を確認し、1件注文する',
  '節税の上限は制度で決まっているが、使い切るだけで年5〜20万円の軍資金が生まれる'
);

// ============================================================
//  スライド 15: 8手法の合計まとめ
// ============================================================
(function() {
  const s = pptx.addSlide();
  bg(s, C.navy);
  s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:'100%', h:0.08, fill:{color:C.gold} });

  s.addText('8手法を組み合わせると…', {
    x:0.5, y:0.2, w:12.3, h:0.72, fontSize:28, bold:true, color:C.white, fontFace:FONT_JA
  });

  const groups = [
    { range:'手法①〜④', label:'支出設計・還元活用（4手法）', amt:'月 1.5〜7万円',  barC:C.accent },
    { range:'手法⑤〜⑦', label:'収入獲得（3手法）',           amt:'月 1.3〜13万円', barC:C.orange },
    { range:'手法⑧',     label:'節税',                       amt:'年 5〜20万円',   barC:C.gold },
  ];

  groups.forEach((g, i) => {
    const y = 1.12 + i * 1.28;
    s.addShape(pptx.ShapeType.rect, { x:0.42, y, w:12.28, h:1.08, fill:{color:C.navyMid}, line:{color:g.barC, pt:2} });
    s.addShape(pptx.ShapeType.rect, { x:0.42, y, w:0.95, h:1.08, fill:{color:g.barC} });
    s.addText(g.range, { x:0.42, y:y+0.2, w:0.95, h:0.68, fontSize:12, bold:true, color:C.white, fontFace:FONT_JA, align:'center' });
    s.addText(g.label, { x:1.5, y:y+0.22, w:7.5, h:0.64, fontSize:19, color:C.white, fontFace:FONT_JA });
    s.addText(g.amt,   { x:9.5, y:y+0.2, w:3.1, h:0.68, fontSize:22, bold:true, color:C.gold, fontFace:FONT_JA, align:'right' });
  });

  // 合計ボックス
  s.addShape(pptx.ShapeType.rect, { x:0.42, y:5.0, w:12.28, h:1.35, fill:{color:C.gold} });
  s.addText('8手法すべてに取り組むと…', { x:0.62, y:5.06, w:5.0, h:0.48, fontSize:16, color:C.navy, fontFace:FONT_JA });
  s.addText('月 3〜20万円 の軍資金を継続的に作ることが可能', {
    x:0.62, y:5.48, w:12.0, h:0.74, fontSize:26, bold:true, color:C.navy, fontFace:FONT_JA
  });

  footer(s);
})();

// ============================================================
//  スライド 16: 40代からの資産形成の正しい順番
// ============================================================
(function() {
  const s = pptx.addSlide();
  bg(s, C.light);
  h1(s, '40代からの資産形成の正しい順番',
    '〜 40代からでも間に合う。ただし「順番」を間違えないことが重要 〜');

  const steps = [
    {
      step:'STEP 1', sub:'まず土台を作る（〜3ヶ月）', color:C.accent,
      items:[
        '先取り貯蓄の自動化を設定する',
        '固定費の見直し（スマホ・保険）',
        'キャッシュレス・ポイント還元を整備',
        '税制優遇（ふるさと納税・iDeCo）を申込む',
      ]
    },
    {
      step:'STEP 2', sub:'軍資金を加速させる（3〜6ヶ月）', color:C.orange,
      items:[
        '断捨離・不用品売却で種銭を作る',
        'ポイントサイトで日常行動を収益化',
        '副業をスタート（小さく・得意なことから）',
        '毎月の軍資金を投資専用口座に蓄積',
      ]
    },
    {
      step:'STEP 3', sub:'軍資金を「増やす仕組み」へ（6ヶ月〜）', color:C.green,
      items:[
        '元本100万円を目安に運用スタート',
        '長期積立（新NISA）で時間を味方につける',
        'さらなる加速の方法を学び続ける',
        '収益の一部を再投資に回すサイクルを作る',
      ]
    },
  ];

  steps.forEach((st, i) => {
    const x = 0.38 + i * 4.3;
    s.addShape(pptx.ShapeType.rect, { x, y:1.38, w:4.05, h:5.35, fill:{color:C.white}, line:{color:'D8DEE6', pt:1} });
    s.addShape(pptx.ShapeType.rect, { x, y:1.38, w:4.05, h:1.06, fill:{color:st.color} });
    s.addText(st.step, { x, y:1.41, w:4.05, h:0.58, fontSize:21, bold:true, color:C.white, fontFace:FONT_EN, align:'center' });
    s.addText(st.sub,  { x, y:1.93, w:4.05, h:0.42, fontSize:11, color:C.white, fontFace:FONT_JA, align:'center' });
    st.items.forEach((item, j) => {
      const iy = 2.62 + j * 0.96;
      s.addShape(pptx.ShapeType.rect, { x:x+0.2, y:iy+0.04, w:0.48, h:0.48, fill:{color:st.color} });
      s.addText('✓', { x:x+0.2, y:iy+0.04, w:0.48, h:0.48, fontSize:15, bold:true, color:C.white, fontFace:FONT_EN, align:'center' });
      s.addText(item, { x:x+0.78, y:iy, w:3.12, h:0.84, fontSize:13, color:C.darkgray, fontFace:FONT_JA });
    });
  });

  footer(s);
})();

// ============================================================
//  スライド 17: まとめ
// ============================================================
(function() {
  const s = pptx.addSlide();
  bg(s, C.navy);
  s.addShape(pptx.ShapeType.rect, { x:0, y:0,   w:'100%', h:0.08, fill:{color:C.gold} });
  s.addShape(pptx.ShapeType.rect, { x:0, y:6.3, w:'100%', h:0.08, fill:{color:C.gold} });
  s.addShape(pptx.ShapeType.rect, { x:0, y:6.38,w:'100%', h:0.04, fill:{color:C.orange} });

  s.addText('本日のまとめ', {
    x:0.5, y:0.2, w:12.3, h:0.72, fontSize:28, bold:true, color:C.gold, fontFace:FONT_JA
  });

  const summaries = [
    { no:'01', text:'投資の結果は「銘柄」より「元本（軍資金）の大きさ」で8〜9割が決まる', color:C.orange },
    { no:'02', text:'資産形成の9割は「投資前」の準備で決まる。軍資金作りが最初の仕事', color:C.gold },
    { no:'03', text:'8つの手法で月3〜20万円の軍資金を作り続けることは、今日から実行できる', color:C.green },
    { no:'04', text:'40代からでも間に合う。ただし「正しい順番」で取り組むことが重要', color:C.accent },
  ];

  summaries.forEach((sm, i) => {
    const y = 1.12 + i * 1.26;
    s.addShape(pptx.ShapeType.rect, { x:0.42, y, w:12.28, h:1.08, fill:{color:C.navyMid}, line:{color:sm.color, pt:2} });
    s.addShape(pptx.ShapeType.rect, { x:0.42, y, w:0.82, h:1.08, fill:{color:sm.color} });
    s.addText(sm.no, { x:0.42, y:y+0.22, w:0.82, h:0.64, fontSize:26, bold:true, color:C.navy, fontFace:FONT_EN, align:'center' });
    s.addText(sm.text, { x:1.38, y:y+0.18, w:11.1, h:0.72, fontSize:17, color:C.white, fontFace:FONT_JA });
  });

  // 締め
  s.addShape(pptx.ShapeType.rect, { x:0.42, y:6.18, w:12.28, h:0.48, fill:{color:C.gold} });
  s.addText('軍資金を作ることが、資産形成の第一歩であり最も重要なステップです。  ─  ' + COMPANY, {
    x:0.42, y:6.2, w:12.28, h:0.44, fontSize:13, bold:true, color:C.navy, fontFace:FONT_JA, align:'center'
  });

  s.addText(COMPANY, { x:0, y:6.55, w:'100%', h:0.28, fontSize:10, color:C.silver, fontFace:FONT_JA, align:'center', italic:true });
})();

// ============================================================
//  保存
// ============================================================
const outputPath = 'D:/dev/軍資金作成資料/docs/軍資金作成講義_v2_スライド.pptx';
pptx.writeFile({ fileName: outputPath })
  .then(() => console.log('✅ Saved:', outputPath))
  .catch(err => console.error('❌', err));
