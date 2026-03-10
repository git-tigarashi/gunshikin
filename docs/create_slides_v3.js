// ============================================================
//  PowerPoint スライド生成 v3
//  投資で成功する人が最初にやっている「軍資金の作り方」
//  株式会社 Vision Creator
//  v3: 4つのポイント追加（自給思考・高時給転職・低金利借金・家族共有）
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
  navyBg:  '1E2A44',
  silver:  'B0BEC5',
  teal:    '00695C',   // 手法⑨〜⑫ のアクセント
  amber:   'F57F17',   // 注意・警告
  purple:  '6A1B9A',   // 家族テーマ
};

const FONT_JA = 'メイリオ';
const FONT_EN = 'Calibri';
const COMPANY = '株式会社 Vision Creator';

// ========== 共通ユーティリティ ==========
function bg(s, color) {
  s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:'100%', h:'100%', fill:{color} });
}
function footer(s) {
  s.addText(COMPANY, { x:0.35, y:6.88, w:5.8, h:0.28, fontSize:9, color:C.gray, fontFace:FONT_JA, italic:true });
  s.addText('投資で成功する人が最初にやっている「軍資金の作り方」', { x:6.3, y:6.88, w:7.0, h:0.28, fontSize:9, color:C.gray, fontFace:FONT_JA, align:'right' });
}
function h1(s, text, sub, y, barC) {
  y = y || 0.42; barC = barC || C.gold;
  s.addShape(pptx.ShapeType.rect, { x:0.4, y, w:0.09, h:0.72, fill:{color:barC} });
  s.addText(text, { x:0.65, y, w:12.1, h:0.72, fontSize:28, bold:true, color:C.navy, fontFace:FONT_JA });
  if (sub) s.addText(sub, { x:0.65, y:y+0.76, w:12.1, h:0.38, fontSize:13, color:C.gray, fontFace:FONT_JA });
}
// 通常手法ヘッダー（ネイビー）
function methodHeader(s, no, title, badge, headerC) {
  headerC = headerC || C.navy;
  s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:'100%', h:1.1, fill:{color:headerC} });
  s.addShape(pptx.ShapeType.rect, { x:0, y:1.08, w:'100%', h:0.05, fill:{color:C.gold} });
  s.addShape(pptx.ShapeType.rect, { x:0.4, y:0.15, w:0.82, h:0.82, fill:{color:C.gold} });
  s.addText(no, { x:0.4, y:0.15, w:0.82, h:0.82, fontSize:30, bold:true, color:headerC, fontFace:FONT_EN, align:'center' });
  s.addText(title, { x:1.38, y:0.18, w:8.6, h:0.78, fontSize:25, bold:true, color:C.white, fontFace:FONT_JA });
  s.addShape(pptx.ShapeType.rect, { x:10.15, y:0.18, w:2.55, h:0.72, fill:{color:C.orange} });
  s.addText(badge, { x:10.15, y:0.18, w:2.55, h:0.72, fontSize:15, bold:true, color:C.white, fontFace:FONT_JA, align:'center' });
}
// 続きスライドヘッダー（同テーマ）
function subHeader(s, no, title, subtitle, headerC) {
  headerC = headerC || C.navy;
  s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:'100%', h:1.1, fill:{color:headerC} });
  s.addShape(pptx.ShapeType.rect, { x:0, y:1.08, w:'100%', h:0.05, fill:{color:C.gold} });
  s.addShape(pptx.ShapeType.rect, { x:0.4, y:0.15, w:0.82, h:0.82, fill:{color:C.gold} });
  s.addText(no, { x:0.4, y:0.15, w:0.82, h:0.82, fontSize:24, bold:true, color:headerC, fontFace:FONT_EN, align:'center' });
  s.addText(title, { x:1.38, y:0.12, w:8.6, h:0.52, fontSize:20, bold:true, color:C.white, fontFace:FONT_JA });
  s.addText(subtitle, { x:1.38, y:0.62, w:8.6, h:0.42, fontSize:15, color:'CCDDEE', fontFace:FONT_JA, italic:true });
  // 続きタグ
  s.addShape(pptx.ShapeType.rect, { x:10.15, y:0.18, w:2.55, h:0.72, fill:{color:headerC === C.navy ? C.teal : headerC} });
  s.addText('続き', { x:10.15, y:0.18, w:2.55, h:0.72, fontSize:14, bold:true, color:C.white, fontFace:FONT_JA, align:'center' });
}

function checkRow(s, y, text, iconC) {
  iconC = iconC || C.accent;
  s.addShape(pptx.ShapeType.rect, { x:0.42, y:y+0.09, w:0.48, h:0.48, fill:{color:iconC} });
  s.addText('✓', { x:0.42, y:y+0.09, w:0.48, h:0.48, fontSize:16, bold:true, color:C.white, fontFace:FONT_EN, align:'center' });
  s.addText(text, { x:1.04, y, w:11.65, h:0.66, fontSize:15.5, color:C.darkgray, fontFace:FONT_JA });
}
function actionBox(s, y, text) {
  s.addShape(pptx.ShapeType.rect, { x:0.42, y, w:12.28, h:0.72, fill:{color:'FFF8E1'}, line:{color:C.gold, pt:2} });
  s.addShape(pptx.ShapeType.rect, { x:0.42, y, w:1.55, h:0.72, fill:{color:C.gold} });
  s.addText('今日の一歩', { x:0.42, y:y+0.12, w:1.55, h:0.48, fontSize:13, bold:true, color:C.navy, fontFace:FONT_JA, align:'center' });
  s.addText(text, { x:2.08, y:y+0.12, w:10.5, h:0.48, fontSize:14, color:C.darkgray, fontFace:FONT_JA });
}
function noteBox(s, y, text, h) {
  h = h || 0.62;
  s.addShape(pptx.ShapeType.rect, { x:0.42, y, w:12.28, h, fill:{color:'EEF3FB'}, line:{color:C.accent, pt:1} });
  s.addText(text, { x:0.58, y:y+0.08, w:12.0, h:h-0.16, fontSize:13.5, color:C.accent, fontFace:FONT_JA, bold:true });
}
// ディスカッション促進ボックス（オレンジ）
function discussBox(s, y, question, h) {
  h = h || 0.76;
  s.addShape(pptx.ShapeType.rect, { x:0.42, y, w:12.28, h, fill:{color:'FFF3E0'}, line:{color:C.orange, pt:2} });
  s.addShape(pptx.ShapeType.rect, { x:0.42, y, w:0.82, h, fill:{color:C.orange} });
  s.addText('Q', { x:0.42, y:y+(h/2)-0.22, w:0.82, h:0.44, fontSize:26, bold:true, color:C.white, fontFace:FONT_EN, align:'center' });
  s.addText(question, { x:1.35, y:y+(h-0.38)/2, w:11.25, h:0.38+((h-0.76)*0.6), fontSize:15, color:'8B4513', fontFace:FONT_JA, bold:true, italic:true });
}
// 警告・注意ボックス（赤）
function warnBox(s, y, text, h) {
  h = h || 0.64;
  s.addShape(pptx.ShapeType.rect, { x:0.42, y, w:12.28, h, fill:{color:'FFF8F8'}, line:{color:C.red, pt:2} });
  s.addText('⚠  ' + text, { x:0.55, y:y+0.08, w:12.0, h:h-0.16, fontSize:13.5, color:C.red, fontFace:FONT_JA, bold:true });
}
// ポイントカード（横並び）
function pointCards(s, cards, yStart) {
  const w = (12.84 - 0.42 * 2 - 0.22 * (cards.length-1)) / cards.length;
  cards.forEach((c, i) => {
    const x = 0.42 + i * (w + 0.22);
    s.addShape(pptx.ShapeType.rect, { x, y:yStart, w, h:c.h||2.2, fill:{color:C.white}, line:{color:'D8DEE6', pt:1} });
    s.addShape(pptx.ShapeType.rect, { x, y:yStart, w, h:0.52, fill:{color:c.color} });
    s.addText(c.label, { x, y:yStart+0.04, w, h:0.44, fontSize:13, bold:true, color:C.white, fontFace:FONT_JA, align:'center' });
    s.addText(c.body,  { x:x+0.12, y:yStart+0.6, w:w-0.24, h:(c.h||2.2)-0.72, fontSize:13, color:C.darkgray, fontFace:FONT_JA });
  });
}
function quoteBox(s, y, text) {
  s.addShape(pptx.ShapeType.rect, { x:0.42, y, w:12.28, h:0.08, fill:{color:C.gold} });
  s.addShape(pptx.ShapeType.rect, { x:0.42, y:y+0.08, w:0.06, h:0.74, fill:{color:C.gold} });
  s.addText(text, { x:0.62, y:y+0.1, w:12.0, h:0.7, fontSize:17, color:C.navy, fontFace:FONT_JA, bold:true, italic:true });
}

// ============================================================
//  スライド 01: 表紙
// ============================================================
(function() {
  const s = pptx.addSlide();
  bg(s, C.navy);
  s.addShape(pptx.ShapeType.rect, { x:0, y:0,    w:'100%', h:0.08, fill:{color:C.gold} });
  s.addShape(pptx.ShapeType.rect, { x:0, y:6.3,  w:'100%', h:0.08, fill:{color:C.gold} });
  s.addShape(pptx.ShapeType.rect, { x:0, y:6.38, w:'100%', h:0.04, fill:{color:C.orange} });

  const subs = ['40代からでも間に合う資産形成の始め方', '資産形成の9割は「投資前」で決まる', 'なぜ多くの人は投資で結果が出ないのか'];
  subs.forEach((t, i) => {
    s.addShape(pptx.ShapeType.rect, { x:1.6, y:0.38+i*0.42, w:9.5, h:0.32, fill:{color:'0D1830'} });
    s.addText(t, { x:1.65, y:0.42+i*0.42, w:9.4, h:0.28, fontSize:12, color:C.silver, fontFace:FONT_JA });
  });

  s.addText('投資で成功する人が', { x:0.6, y:1.72, w:12.1, h:0.9, fontSize:44, bold:true, color:C.white, fontFace:FONT_JA, align:'center' });
  s.addText('最初にやっている', { x:0.6, y:2.58, w:12.1, h:0.9, fontSize:44, bold:true, color:C.white, fontFace:FONT_JA, align:'center' });
  s.addShape(pptx.ShapeType.rect, { x:1.5, y:3.48, w:10.3, h:1.3, fill:{color:C.gold} });
  s.addText('「軍資金」の作り方', { x:1.5, y:3.48, w:10.3, h:1.3, fontSize:54, bold:true, color:C.navy, fontFace:FONT_JA, align:'center' });
  s.addShape(pptx.ShapeType.rect, { x:3.8, y:4.98, w:5.7, h:0.78, fill:{color:C.navyMid} });
  s.addText(COMPANY, { x:3.8, y:4.98, w:5.7, h:0.78, fontSize:20, bold:true, color:C.gold, fontFace:FONT_JA, align:'center' });
  s.addText(COMPANY, { x:0, y:6.55, w:'100%', h:0.3, fontSize:10, color:C.silver, fontFace:FONT_JA, align:'center', italic:true });
})();

// ============================================================
//  スライド 02: 本日のテーマ・目次（12手法版）
// ============================================================
(function() {
  const s = pptx.addSlide();
  bg(s, C.light);
  h1(s, '本日お伝えすること', '〜 投資を始める前に知っておくべき話 〜');

  const qs = [
    { q:'なぜ多くの人は\n投資で結果が出ないのか？',  y:1.42, color:C.red },
    { q:'資産形成は銘柄ではなく\n「ここ」で決まる',     y:2.88, color:C.orange },
    { q:'40代からの\n資産形成の正しい順番',         y:4.34, color:C.accent },
  ];
  qs.forEach(q => {
    s.addShape(pptx.ShapeType.rect, { x:0.42, y:q.y, w:5.5, h:1.22, fill:{color:q.color} });
    s.addText(q.q, { x:0.52, y:q.y+0.12, w:5.3, h:0.98, fontSize:17, bold:true, color:C.white, fontFace:FONT_JA });
  });

  // 右：目次
  s.addShape(pptx.ShapeType.rect, { x:6.2, y:1.42, w:6.5, h:5.2, fill:{color:C.white}, line:{color:'D8DEE6', pt:1} });
  s.addShape(pptx.ShapeType.rect, { x:6.2, y:1.42, w:6.5, h:0.62, fill:{color:C.navy} });
  s.addText('本日の講義内容', { x:6.2, y:1.42, w:6.5, h:0.62, fontSize:16, bold:true, color:C.gold, fontFace:FONT_JA, align:'center' });

  const items = [
    { t:'Part 1  投資で失敗する人・成功する人の違い', main:true },
    { t:'Part 2  「軍資金」という発想の転換', main:true },
    { t:'Part 3  軍資金を作る12の方法', main:true },
    { t:'   ①〜④  支出設計・還元活用', main:false },
    { t:'   ⑤〜⑧  収入獲得・節税', main:false },
    { t:'   ⑨  自給思考', main:false },
    { t:'   ⑩  高時給への転職', main:false },
    { t:'   ⑪  低金利の借金を積極的に活用', main:false },
    { t:'   ⑫  家族でお金を共有する', main:false },
    { t:'Part 4  40代からの資産形成ロードマップ', main:true },
  ];
  items.forEach((item, i) => {
    s.addText(item.t, { x:6.4, y:2.16+i*0.34, w:6.1, h:0.32, fontSize: item.main?13:12, bold:item.main, color: item.main?C.navy:C.darkgray, fontFace:FONT_JA });
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
  s.addText('なぜ多くの人は投資で結果が出ないのか？', { x:0.5, y:0.2, w:12.3, h:0.78, fontSize:28, bold:true, color:C.white, fontFace:FONT_JA });

  const reasons = [
    { no:'01', title:'「銘柄選び」に集中しすぎる', detail:'「どの株を買えば儲かるか」ばかり考える。しかし利益の大半は銘柄ではなく元本の大きさで決まる。' },
    { no:'02', title:'元本（軍資金）が少なすぎる', detail:'10万円を運用して年5%の利益 ＝ 5,000円。\n100万円なら50,000円。同じ努力でも10倍の差になる。' },
    { no:'03', title:'正しい「順番」を知らずに始める', detail:'軍資金を作る準備なしに投資をスタートする。この順番を間違えると、知識があっても動けない。' },
  ];
  reasons.forEach((r, i) => {
    const y = 1.18 + i * 1.72;
    s.addShape(pptx.ShapeType.rect, { x:0.42, y, w:12.28, h:1.52, fill:{color:C.navyMid}, line:{color:C.orange, pt:1} });
    s.addShape(pptx.ShapeType.rect, { x:0.42, y, w:0.82, h:1.52, fill:{color:C.orange} });
    s.addText(r.no, { x:0.42, y:y+0.42, w:0.82, h:0.68, fontSize:26, bold:true, color:C.white, fontFace:FONT_EN, align:'center' });
    s.addText(r.title,  { x:1.38, y:y+0.1,  w:11.1, h:0.58, fontSize:20, bold:true, color:C.gold, fontFace:FONT_JA });
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

  s.addShape(pptx.ShapeType.rect, { x:0.42, y:1.42, w:12.28, h:0.62, fill:{color:C.navy} });
  s.addText('同じ利回りでも「元本」の大きさで結果がまったく変わる', { x:0.5, y:1.44, w:12.1, h:0.58, fontSize:18, bold:true, color:C.gold, fontFace:FONT_JA, align:'center' });

  const cards = [
    { label:'元本 10万円',  rate:'年利5%', profit:'5,000円/年',   note:'コーヒー代にもならない', color:C.gray },
    { label:'元本 100万円', rate:'年利5%', profit:'50,000円/年',  note:'月4,000円超の収入',    color:C.accent },
    { label:'元本 300万円', rate:'年利5%', profit:'150,000円/年', note:'月12,500円の収益',     color:C.green },
    { label:'元本 500万円', rate:'年利5%', profit:'250,000円/年', note:'月2万円超のキャッシュ', color:C.orange },
  ];
  cards.forEach((c, i) => {
    const x = 0.42 + i * 3.1;
    s.addShape(pptx.ShapeType.rect, { x, y:2.22, w:2.95, h:3.1, fill:{color:C.white}, line:{color:'D8DEE6', pt:1} });
    s.addShape(pptx.ShapeType.rect, { x, y:2.22, w:2.95, h:0.62, fill:{color:c.color} });
    s.addText(c.label,  { x, y:2.24, w:2.95, h:0.58, fontSize:15, bold:true, color:C.white, fontFace:FONT_JA, align:'center' });
    s.addText(c.rate,   { x, y:2.96, w:2.95, h:0.48, fontSize:13, color:C.gray, fontFace:FONT_JA, align:'center' });
    s.addText(c.profit, { x, y:3.42, w:2.95, h:0.88, fontSize:22, bold:true, color:c.color, fontFace:FONT_JA, align:'center' });
    s.addText(c.note,   { x:x+0.12, y:4.3, w:2.7, h:0.88, fontSize:12.5, color:C.gray, fontFace:FONT_JA, align:'center' });
  });
  s.addShape(pptx.ShapeType.rect, { x:0.42, y:5.5, w:12.28, h:0.82, fill:{color:C.navy} });
  s.addText('投資の「結果」は技術よりも「元本の大きさ」で8〜9割が決まる。だから最初に軍資金を作る。', { x:0.55, y:5.52, w:12.0, h:0.78, fontSize:17, bold:true, color:C.gold, fontFace:FONT_JA, align:'center' });
  footer(s);
})();

// ============================================================
//  スライド 05: 「軍資金」という発想の転換
// ============================================================
(function() {
  const s = pptx.addSlide();
  bg(s, C.light);
  h1(s, '「軍資金」という発想の転換', '〜 資産形成が早い人は最初にこれをやる 〜');

  const baf = [
    { label:'Before（ほとんどの人）', desc:'給料 → 生活費をすべて使う\n→ 残ったら（ほぼ残らず）投資へ', color:C.gray },
    { label:'After（軍資金思考）',   desc:'給料 → 先に軍資金を確保\n→ 残りで生活 → 軍資金を投資へ', color:C.green },
  ];
  baf.forEach((b, i) => {
    const x = i === 0 ? 0.42 : 6.72;
    s.addShape(pptx.ShapeType.rect, { x, y:1.42, w:5.95, h:2.4, fill:{color:C.white}, line:{color:'D8DEE6', pt:1} });
    s.addShape(pptx.ShapeType.rect, { x, y:1.42, w:5.95, h:0.68, fill:{color:b.color} });
    s.addText(b.label, { x, y:1.44, w:5.95, h:0.64, fontSize:16, bold:true, color:C.white, fontFace:FONT_JA, align:'center' });
    s.addText(b.desc,  { x:x+0.2, y:2.2, w:5.55, h:1.48, fontSize:15.5, color:C.darkgray, fontFace:FONT_JA });
    if (i===0) { s.addShape(pptx.ShapeType.rect, { x:6.4, y:2.28, w:0.42, h:0.62, fill:{color:C.gold} }); s.addText('▶', { x:6.4, y:2.28, w:0.42, h:0.62, fontSize:20, bold:true, color:C.navy, fontFace:FONT_EN, align:'center' }); }
  });

  s.addShape(pptx.ShapeType.rect, { x:0.42, y:4.0, w:12.28, h:0.52, fill:{color:C.navyMid} });
  s.addText('軍資金思考の3つのポイント', { x:0.52, y:4.02, w:12.0, h:0.48, fontSize:15, bold:true, color:C.gold, fontFace:FONT_JA });
  const pts = ['「残ったら貯める」ではなく「先に確保して、残りで生活する」という順序に変える', '軍資金＝投資専用口座を別に持ち、生活費と完全に分離する', '40代からでも間に合う。月3〜5万円の積み上げが数年後の大きな元本になる'];
  pts.forEach((p, i) => {
    s.addShape(pptx.ShapeType.rect, { x:0.42, y:4.62+i*0.58, w:0.48, h:0.42, fill:{color:C.gold} });
    s.addText(['①','②','③'][i], { x:0.42, y:4.66+i*0.58, w:0.48, h:0.38, fontSize:14, bold:true, color:C.navy, fontFace:FONT_JA, align:'center' });
    s.addText(p, { x:1.02, y:4.62+i*0.58, w:11.68, h:0.52, fontSize:14.5, color:C.darkgray, fontFace:FONT_JA });
  });
  footer(s);
})();

// ============================================================
//  スライド 06: 12手法の全体マップ
// ============================================================
(function() {
  const s = pptx.addSlide();
  bg(s, C.navy);
  s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:'100%', h:0.08, fill:{color:C.gold} });
  s.addText('軍資金を作る12の方法　全体マップ', { x:0.5, y:0.15, w:12.3, h:0.72, fontSize:24, bold:true, color:C.white, fontFace:FONT_JA });

  const methods = [
    { no:'01', t:'先取り貯蓄の自動化',      a:'月 1〜3万円',    c:'支出設計', cC:C.accent },
    { no:'02', t:'固定費の徹底見直し',      a:'月 5千〜3万円',  c:'支出設計', cC:C.accent },
    { no:'03', t:'変動費の最適化',          a:'月 3千〜1.5万円',c:'支出設計', cC:C.accent },
    { no:'04', t:'キャッシュレス還元の集約', a:'月 2千〜1万円',  c:'還元活用', cC:C.green },
    { no:'05', t:'ポイントサイト活用',       a:'月 3千〜3万円',  c:'収入獲得', cC:C.orange },
    { no:'06', t:'不用品売却・断捨離',       a:'初回〜5万円',    c:'収入獲得', cC:C.orange },
    { no:'07', t:'副業・スキル収益化',       a:'月 1〜10万円',   c:'収入獲得', cC:C.orange },
    { no:'08', t:'税制優遇制度の活用',       a:'年 5〜20万円',   c:'節税',     cC:C.gold },
    { no:'09', t:'自給思考',               a:'月 1〜3万円',    c:'生活設計', cC:C.teal },
    { no:'10', t:'高時給への転職',          a:'月 3〜30万円↑', c:'収入増加', cC:C.teal },
    { no:'11', t:'低金利の借金を積極活用',   a:'レバレッジ',     c:'資産活用', cC:C.teal },
    { no:'12', t:'家族でお金を共有する',     a:'家族一体で増加', c:'家族戦略', cC:C.purple },
  ];

  methods.forEach((m, i) => {
    const col = i % 3;
    const row = Math.floor(i / 3);
    const x = 0.42 + col * 4.28;
    const y = 1.06 + row * 1.44;
    s.addShape(pptx.ShapeType.rect, { x, y, w:4.02, h:1.28, fill:{color:C.navyMid}, line:{color:'3A4A80', pt:1} });
    s.addShape(pptx.ShapeType.rect, { x, y, w:0.55, h:1.28, fill:{color:C.gold} });
    s.addText(m.no, { x, y:y+0.34, w:0.55, h:0.6, fontSize:16, bold:true, color:C.navy, fontFace:FONT_EN, align:'center' });
    s.addText(m.t, { x:x+0.62, y:y+0.1, w:2.52, h:0.6, fontSize:13, bold:true, color:C.white, fontFace:FONT_JA });
    s.addShape(pptx.ShapeType.rect, { x:x+0.62, y:y+0.76, w:0.98, h:0.3, fill:{color:m.cC} });
    s.addText(m.c, { x:x+0.62, y:y+0.76, w:0.98, h:0.3, fontSize:10, color:C.white, fontFace:FONT_JA, align:'center' });
    s.addText(m.a, { x:x+3.0, y:y+0.38, w:0.94, h:0.52, fontSize:11, bold:true, color:C.gold, fontFace:FONT_JA, align:'right' });
  });
  footer(s);
})();

// ============================================================
//  スライド 07〜14: 手法①〜⑧（v2と同じ内容）
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

methodSlide('01','先取り貯蓄の自動化','〜「残ったら貯める」から「先に確保する」への転換〜','月 1〜3万円',
  ['給与振込日の翌日に証券口座・投資専用口座へ「自動転送」を設定する','手取りの10%を目標に。月収25万円なら2.5万円を"存在しないお金"として管理する','住信SBIネット銀行「目的別口座」・楽天銀行「自動振替」が実用的で設定も簡単'],
  '今月中に投資専用口座への自動積立を設定する（10分でできる）','米国発「Pay Yourself First」── 資産家に共通する収入管理の世界的鉄則');

methodSlide('02','固定費の徹底見直し','〜一度やれば毎月ずっと効く"最高コスパ"の節約〜','月 5千〜3万円',
  ['スマートフォン：大手（〜8,000円）→ 格安SIM（〜1,500円）で月6,500円を永続削減','生命保険：40代以降は保障過大になりやすい。FP無料相談で月5,000〜15,000円の削減余地を確認','サブスク棚卸し：使っていないサービスを洗い出して解約。月2,000〜5,000円の削減','電気・ガス：切り替えるだけで月1,000〜3,000円。一度やれば永続効果'],
  '今週中にスマホ料金・保険の月額を確認し、見直しの優先順位を決める','「削減＝我慢」ではなく「同等のサービスを安く買う」という発想で取り組む');

methodSlide('03','変動費の最適化','〜ゼロにしなくていい。まず"漏れ"をふさぐだけ〜','月 3千〜1.5万円',
  ['食費：まとめ買い＋冷凍活用。業務スーパー・コストコの活用で月5,000〜10,000円削減','外食費：QRコード決済のポイント還元デーを活用し、実質割引で外食する習慣をつくる','趣味・娯楽：図書館・動画配信の共有プランなどで月2,000〜5,000円の代替を検討'],
  '今月の食費・外食費を1週間分だけ記録し「気づかない漏れ」を可視化する','米国発「ノースペンドデー」：週1日を消費ゼロの日にするだけで支出意識が変わる');

methodSlide('04','キャッシュレス還元の集約','〜すでに使っているお金を"ポイント"に変換〜','月 2千〜1万円',
  ['年会費無料×高還元カード1〜2枚に支出を集中（楽天カード最大3%・PayPayカード最大5%）','公共料金・保険料・通販・ガソリンもすべてカード払いに統一してポイントを積み上げる','月30万円の支出 × 還元率2% ＝ 月6,000円 → 年72,000円の軍資金が自動的に積み上がる','楽天ポイント投資・PayPay資産運用で「ポイントをそのまま投資に回す」設定も活用可能'],
  '今月中に高還元クレカへ支出を集約し、ポイントを即換金・投資振替する設定を行う','ポイントは「貯めるもの」ではなく「現金と同価値の資産」として即活用するのが鉄則');

methodSlide('05','ポイントサイト・アフィリエイト活用','〜日常のサービス申し込みを"収入"に変える〜','月 3千〜3万円',
  ['A8.net：国内最大のASP。ブログ不要で証券口座・クレカ開設の自己アフィリエイトが可能','ハピタス・モッピー：クレカ発行・口座開設等の高額案件で1件あたり数千〜数万ポイント獲得','ポイントタウン：日常のショッピング・アンケート・ゲームでコツコツと換金可能ポイントを獲得'],
  '今週中にハピタスまたはモッピーに無料登録し、証券口座開設の案件を1件経由する','「ポイ活」を目的にするのではなく、必要な行動・申し込みをポイントサイト経由にするだけ');

methodSlide('06','不用品売却・断捨離収益化','〜眠っている「資産」を現金化し、最初の種銭を作る〜','初回〜5万円',
  ['メルカリ：衣類・雑貨・本。スマホ1台で写真を撮って即出品。初回断捨離で数万円になることも','ヤフオク：ブランド品・趣味グッズ・レア品。入札形式で市場価格に近い高値がつきやすい','ハードオフ・買取専門店：家電・楽器・カメラは持ち込み即日現金化。査定は無料'],
  '今日：家の中を15分で見回し「半年使っていないもの」をスマホでリストアップする','「一度きり」だからこそ最初の種銭作りに最適。副産物として"衝動買い防止"の意識も生まれる');

methodSlide('07','副業・スキル収益化','〜本業以外の「第二の収入源」を軍資金に変える〜','月 1〜10万円',
  ['クラウドワークス・ランサーズ：ライティング・データ入力・翻訳。スキル不要の案件も多数あり','ストアカ・ ココナラ：職歴・趣味・専門知識を「教える商品」に変換。高単価になりやすい','せどり・転売：仕入れ→販売のサイクル。月3〜20万円の実績者が多い初期コスト低の副業','大原則：副業収益は生活費に使わず「全額を投資口座に入金する」ルールを最初に決める'],
  '今月中：自分の職歴・趣味から「売れるスキル・経験」を1つ書き出してみる','キャリア・人脈・専門知識は「副業の武器」。40代以降ほど高単価になりやすい傾向がある');

methodSlide('08','税制優遇制度の活用','〜払わなくていい税金を取り戻し、投資原資に回す〜','年 5〜20万円',
  ['iDeCo：掛金が全額所得控除。年収600万円で月2万円拠出なら年間約4〜6万円の節税効果','ふるさと納税：実質2,000円の自己負担で返礼品受取＋住民税控除。食費削減にも直結する','新NISA：運用益・配当が非課税。長期資産形成の土台として併用するのが基本形','副業がある場合は青色申告：最大65万円の特別控除。経費計上で課税所得を大幅に圧縮できる'],
  '今週中：ふるさと納税のシミュレーションで今年の上限額を確認し1件注文する','節税の上限は制度で決まっているが、使い切るだけで年5〜20万円の軍資金が生まれる');

// ============================================================
//  スライド 15: 手法⑨ 自給思考
// ============================================================
(function() {
  const s = pptx.addSlide();
  bg(s, C.light);
  methodHeader(s, '09', '自給思考', '月 1〜3万円', C.teal);
  s.addText('〜「買う」から「作る・育てる」への発想転換。消費を生産に変える〜', { x:0.5, y:1.16, w:12.2, h:0.42, fontSize:14, color:C.gray, fontFace:FONT_JA, italic:true });

  // 4枚カード
  const cards = [
    { label:'家庭菜園', color:C.green,  h:2.4, body:'ベランダや庭でトマト・葉野菜・ハーブを育てる。月2,000〜5,000円分の食材を自給。土・プランターの初期費用は数千円から。' },
    { label:'料理・弁当持参', color:C.teal, h:2.4, body:'外食・コンビニをやめて自炊に切り替え。弁当1個で節約500〜800円×20日＝月1〜1.6万円。料理スキル自体が資産になる。' },
    { label:'DIY・修繕', color:C.accent, h:2.4, body:'業者に頼まず自分で修理・修繕する。ペンキ塗り、棚の修理、自転車整備など。YouTubeで多くが学べる。' },
    { label:'スキルシェア', color:C.orange, h:2.4, body:'近隣との物々交換や助け合い。野菜を分け合う、得意な作業を交換する。お金を使わず豊かになる仕組み。' },
  ];
  const w = (12.84 - 0.42*2 - 0.22*3) / 4;
  cards.forEach((c, i) => {
    const x = 0.42 + i*(w+0.22);
    s.addShape(pptx.ShapeType.rect, { x, y:1.72, w, h:2.4, fill:{color:C.white}, line:{color:'D8DEE6', pt:1} });
    s.addShape(pptx.ShapeType.rect, { x, y:1.72, w, h:0.52, fill:{color:c.color} });
    s.addText(c.label, { x, y:1.74, w, h:0.46, fontSize:14, bold:true, color:C.white, fontFace:FONT_JA, align:'center' });
    s.addText(c.body,  { x:x+0.1, y:2.32, w:w-0.2, h:1.68, fontSize:12, color:C.darkgray, fontFace:FONT_JA });
  });

  actionBox(s, 4.28, '今週末：家で「作れそうなもの」を1つ決め、材料リストを作る');
  discussBox(s, 5.16, 'あなたが今「お金を払って買っている」ものの中で、自分で作れそうなものは何ですか？');
  footer(s);
})();

// ============================================================
//  スライド 16: 手法⑩ 高時給への転職（1/2）
// ============================================================
(function() {
  const s = pptx.addSlide();
  bg(s, C.light);
  methodHeader(s, '10', '高時給への転職', '収入増 = 最強の軍資金加速', C.teal);
  s.addText('〜 収入を増やす効果は、節約の限界をはるかに超える 〜', { x:0.5, y:1.16, w:12.2, h:0.42, fontSize:14, color:C.gray, fontFace:FONT_JA, italic:true });

  // 時給計算の仕組み
  s.addShape(pptx.ShapeType.rect, { x:0.42, y:1.72, w:12.28, h:0.52, fill:{color:C.navyMid} });
  s.addText('まず「自分の今の時給」を計算してみる', { x:0.5, y:1.74, w:12.1, h:0.48, fontSize:16, bold:true, color:C.gold, fontFace:FONT_JA });

  // 時給計算式
  s.addShape(pptx.ShapeType.rect, { x:0.42, y:2.36, w:7.5, h:1.1, fill:{color:C.white}, line:{color:C.teal, pt:2} });
  s.addText('時給 ＝ 月収 ÷ （1日の実働時間 × 月間労働日数）', { x:0.55, y:2.46, w:7.24, h:0.7, fontSize:17, bold:true, color:C.teal, fontFace:FONT_JA });
  s.addText('※通勤時間・残業を含めて計算するとより正確', { x:0.55, y:3.06, w:7.24, h:0.32, fontSize:12, color:C.gray, fontFace:FONT_JA });

  // 例表
  const rows = [
    ['月収25万円', '9時間労働×20日', '約1,389円/時'],
    ['月収35万円', '9時間労働×20日', '約1,944円/時'],
    ['月収50万円', '9時間労働×20日', '約2,778円/時'],
  ];
  const hdrY = 2.36;
  const colX = [8.18, 9.58, 11.08];
  const colW = [1.32, 1.42, 1.82];
  const hdrs = ['月収', '勤務', '時給'];
  colX.forEach((x, i) => {
    s.addShape(pptx.ShapeType.rect, { x, y:hdrY, w:colW[i], h:0.38, fill:{color:C.navy} });
    s.addText(hdrs[i], { x, y:hdrY, w:colW[i], h:0.38, fontSize:12, bold:true, color:C.white, fontFace:FONT_JA, align:'center' });
  });
  rows.forEach((row, ri) => {
    const ry = hdrY + 0.38 + ri * 0.44;
    const bg2 = ri % 2 === 0 ? C.white : 'F0F4FF';
    colX.forEach((x, ci) => {
      s.addShape(pptx.ShapeType.rect, { x, y:ry, w:colW[ci], h:0.44, fill:{color:bg2}, line:{color:'DDDDDD', pt:1} });
      s.addText(row[ci], { x:x+0.04, y:ry+0.06, w:colW[ci]-0.08, h:0.32, fontSize:11.5, color:ci===2?C.orange:C.darkgray, fontFace:FONT_JA, bold:ci===2, align:ci===2?'center':'left' });
    });
  });

  s.addShape(pptx.ShapeType.rect, { x:0.42, y:3.68, w:12.28, h:0.58, fill:{color:C.navyMid} });
  s.addText('固定費を月1万円削減するより、時給を500円上げる方が生涯収入への影響は圧倒的に大きい', { x:0.52, y:3.7, w:12.0, h:0.54, fontSize:14.5, bold:true, color:C.gold, fontFace:FONT_JA });

  actionBox(s, 4.4, '今日：月収 ÷（実働時間×労働日数）で「自分の時給」を計算してみる');
  discussBox(s, 5.28, '今の職場の時給は「納得できる金額」ですか？同じ時間を使うなら、より高く評価される場所はないか？');
  footer(s);
})();

// ============================================================
//  スライド 17: 手法⑩ 高時給への転職（2/2）
// ============================================================
(function() {
  const s = pptx.addSlide();
  bg(s, C.light);
  subHeader(s, '10', '高時給への転職', '〜 転職・キャリアアップの3つの戦略 〜', C.teal);

  const strategies = [
    { no:'A', title:'業界・職種を変える転職', color:C.teal,
      points:['同じスキルでも業界が変わると年収が大きく変わる', 'IT・金融・医療系は他業界より平均年収が高い傾向', '転職エージェントの無料活用で市場価値を把握する'] },
    { no:'B', title:'社内での昇給・昇進戦略', color:C.accent,
      points:['資格取得・スキルアップで評価の見える化をする', '副業で副収入を得ながら、本業での交渉材料を増やす', 'マネジメントポジションへの移行で年収ジャンプを狙う'] },
    { no:'C', title:'スキル投資で時給を上げる', color:C.orange,
      points:['プログラミング・英語・資格など、市場価値の高いスキルを取得', '資格取得・セミナー費用は「未来の軍資金を生む投資」として考える', '副業でスキルを磨きながら本業に活かす相乗効果を狙う'] },
  ];

  strategies.forEach((st, i) => {
    const y = 1.3 + i * 1.72;
    s.addShape(pptx.ShapeType.rect, { x:0.42, y, w:12.28, h:1.52, fill:{color:C.white}, line:{color:'D8DEE6', pt:1} });
    s.addShape(pptx.ShapeType.rect, { x:0.42, y, w:0.62, h:1.52, fill:{color:st.color} });
    s.addText(st.no, { x:0.42, y:y+0.42, w:0.62, h:0.68, fontSize:26, bold:true, color:C.white, fontFace:FONT_EN, align:'center' });
    s.addText(st.title, { x:1.18, y:y+0.1, w:4.5, h:0.58, fontSize:17, bold:true, color:st.color, fontFace:FONT_JA });
    st.points.forEach((p, j) => {
      s.addShape(pptx.ShapeType.rect, { x:1.18, y:y+0.72+j*0.25, w:0.22, h:0.22, fill:{color:st.color} });
      s.addText(p, { x:1.5, y:y+0.68+j*0.26, w:11.0, h:0.3, fontSize:13, color:C.darkgray, fontFace:FONT_JA });
    });
  });

  noteBox(s, 6.52, '収入増加の天井は節約よりはるかに高い。軍資金作りの中で最も大きなレバレッジが効く手法。', 0.58);
  footer(s);
})();

// ============================================================
//  スライド 18: 手法⑪ 低金利の借金（1/2）概念編
// ============================================================
(function() {
  const s = pptx.addSlide();
  bg(s, C.light);
  methodHeader(s, '11', '低金利の借金を積極的に活用する', 'レバレッジ', C.teal);
  s.addText('〜「良い借金」と「悪い借金」を区別する。お金にお金を稼いでもらう発想〜', { x:0.5, y:1.16, w:12.2, h:0.42, fontSize:13.5, color:C.gray, fontFace:FONT_JA, italic:true });

  // 良い借金・悪い借金 比較
  s.addShape(pptx.ShapeType.rect, { x:0.42, y:1.72, w:12.28, h:0.48, fill:{color:C.navyMid} });
  s.addText('良い借金 vs 悪い借金', { x:0.52, y:1.74, w:12.0, h:0.44, fontSize:16, bold:true, color:C.gold, fontFace:FONT_JA });

  // 悪い借金（左）
  s.addShape(pptx.ShapeType.rect, { x:0.42, y:2.32, w:5.9, h:3.22, fill:{color:'FFF8F8'}, line:{color:C.red, pt:2} });
  s.addShape(pptx.ShapeType.rect, { x:0.42, y:2.32, w:5.9, h:0.52, fill:{color:C.red} });
  s.addText('❌  悪い借金（消費のための借金）', { x:0.55, y:2.34, w:5.64, h:0.46, fontSize:14, bold:true, color:C.white, fontFace:FONT_JA });
  const badDebts = ['消費者金融・キャッシング　金利15〜18%', 'クレジットカードのリボ払い　金利15%前後', 'ブランド品・高級車ローン（価値が下がるもの）', '生活費のための借入（食費・娯楽等）'];
  badDebts.forEach((t, i) => { s.addText('• ' + t, { x:0.58, y:2.96+i*0.5, w:5.6, h:0.44, fontSize:13, color:C.red, fontFace:FONT_JA }); });

  // 良い借金（右）
  s.addShape(pptx.ShapeType.rect, { x:6.78, y:2.32, w:5.92, h:3.22, fill:{color:'F1F8E9'}, line:{color:C.green, pt:2} });
  s.addShape(pptx.ShapeType.rect, { x:6.78, y:2.32, w:5.92, h:0.52, fill:{color:C.green} });
  s.addText('✅  良い借金（資産を生む借金）', { x:6.92, y:2.34, w:5.66, h:0.46, fontSize:14, bold:true, color:C.white, fontFace:FONT_JA });
  const goodDebts = ['住宅ローン　金利0.5〜2%（固定・変動）', '事業資金・設備ローン　金利1〜3%', '不動産投資ローン（利回りが金利を上回る場合）', '教育ローン（収入増につながるスキル投資）'];
  goodDebts.forEach((t, i) => { s.addText('• ' + t, { x:6.92, y:2.96+i*0.5, w:5.68, h:0.44, fontSize:13, color:C.green, fontFace:FONT_JA }); });

  // 矢印・条件
  s.addShape(pptx.ShapeType.rect, { x:6.4, y:3.52, w:0.42, h:0.62, fill:{color:C.gold} });
  s.addText('→', { x:6.4, y:3.52, w:0.42, h:0.62, fontSize:22, bold:true, color:C.navy, fontFace:FONT_EN, align:'center' });

  warnBox(s, 5.68, '大原則：借金の金利 ＜ 投資の期待利回り　であることを必ず確認してから活用する', 0.62);
  discussBox(s, 6.42, '今の「借金」の金利は何%ですか？その金利は低金利ですか、高金利ですか？');
  footer(s);
})();

// ============================================================
//  スライド 19: 手法⑪ 低金利の借金（2/2）実践例
// ============================================================
(function() {
  const s = pptx.addSlide();
  bg(s, C.light);
  subHeader(s, '11', '低金利の借金を積極的に活用する', '〜 実践事例と注意点 〜', C.teal);

  const examples = [
    { title:'住宅ローン × 家賃収入', color:C.green,
      body:'変動金利0.5〜1%で自宅購入。それまで払っていた家賃が消え、余剰資金を投資に回せる。区分マンション等では家賃収入で返済を賄うことも可能。' },
    { title:'事業資金ローン × 副業加速', color:C.teal,
      body:'副業の設備・ツール・広告費に低金利の事業ローンを活用。月10万円の副業収入に対して返済が月1〜2万円なら手元に残る利益は大きい。' },
    { title:'低金利カードローン × 投資元本', color:C.accent,
      body:'一部の銀行カードローン（金利1〜4%）や証券会社の信用取引・FXレバレッジ。活用できる場合は投資元本を拡大できる。ただし高リスク。' },
  ];

  examples.forEach((e, i) => {
    const y = 1.3 + i * 1.56;
    s.addShape(pptx.ShapeType.rect, { x:0.42, y, w:12.28, h:1.36, fill:{color:C.white}, line:{color:'D8DEE6', pt:1} });
    s.addShape(pptx.ShapeType.rect, { x:0.42, y, w:3.2, h:1.36, fill:{color:e.color} });
    s.addText(e.title, { x:0.52, y:y+0.34, w:3.0, h:0.68, fontSize:15, bold:true, color:C.white, fontFace:FONT_JA });
    s.addText(e.body,  { x:3.76, y:y+0.12, w:8.82, h:1.12, fontSize:13.5, color:C.darkgray, fontFace:FONT_JA });
  });

  warnBox(s, 6.0, '投機・ギャンブルのための借金、返済計画のない借金は絶対に避ける。「低金利」でも返せない借金は悪い借金。', 0.7);
  actionBox(s, 6.82, '今月中：住宅ローンや銀行ローンの現在金利を確認し、繰上げ返済 or 借り換えの可否を検討する');
  footer(s);
})();

// ============================================================
//  スライド 20: 手法⑫ 家族でお金を共有する（1/2）家族会議
// ============================================================
(function() {
  const s = pptx.addSlide();
  bg(s, C.light);
  methodHeader(s, '12', '家族でお金を共有する', '家族一体で増加', C.purple);
  s.addText('〜 家族の財布を「見える化」するだけで、無駄が消えて軍資金が増える 〜', { x:0.5, y:1.16, w:12.2, h:0.42, fontSize:13.5, color:C.gray, fontFace:FONT_JA, italic:true });

  // なぜ家族共有か
  s.addShape(pptx.ShapeType.rect, { x:0.42, y:1.7, w:12.28, h:0.48, fill:{color:C.purple} });
  s.addText('「知らない」「話さない」が一番の損失', { x:0.52, y:1.72, w:12.0, h:0.44, fontSize:15, bold:true, color:C.white, fontFace:FONT_JA });

  // 左：よくある問題
  s.addShape(pptx.ShapeType.rect, { x:0.42, y:2.3, w:5.9, h:2.48, fill:{color:C.white}, line:{color:'D8DEE6', pt:1} });
  s.addShape(pptx.ShapeType.rect, { x:0.42, y:2.3, w:5.9, h:0.48, fill:{color:C.gray} });
  s.addText('よくある「家族の財布問題」', { x:0.52, y:2.32, w:5.7, h:0.44, fontSize:13, bold:true, color:C.white, fontFace:FONT_JA });
  const problems = ['夫婦で同じサブスクに加入していた', '片方が知らずに高い保険に入っていた', '子どもの教育費を誰も把握していない', '老後資金の認識が夫婦でずれていた'];
  problems.forEach((p, i) => { s.addText('• ' + p, { x:0.58, y:2.88+i*0.46, w:5.6, h:0.4, fontSize:13, color:C.darkgray, fontFace:FONT_JA }); });

  // 右：月次家族会議の進め方
  s.addShape(pptx.ShapeType.rect, { x:6.78, y:2.3, w:5.92, h:2.48, fill:{color:C.white}, line:{color:C.purple, pt:2} });
  s.addShape(pptx.ShapeType.rect, { x:6.78, y:2.3, w:5.92, h:0.48, fill:{color:C.purple} });
  s.addText('月次・家族財務会議の進め方', { x:6.88, y:2.32, w:5.72, h:0.44, fontSize:13, bold:true, color:C.white, fontFace:FONT_JA });
  const agenda = ['今月の収支を全員で確認する（15分）', '無駄な支出・ダブりを洗い出す（10分）', '来月の目標軍資金額を設定する（5分）', '子どもも参加させる（お金の教育にも）'];
  agenda.forEach((a, i) => { s.addText(['①','②','③','④'][i] + '  ' + a, { x:6.92, y:2.88+i*0.46, w:5.68, h:0.4, fontSize:13, color:C.darkgray, fontFace:FONT_JA }); });

  actionBox(s, 4.96, '今月中：家族（または夫婦）でお金について話す時間を30分作る。まず「全サブスクの洗い出し」から');
  discussBox(s, 5.84, '家族全員が「今の資産残高」「毎月の収支」を把握していますか？お金の話を家族でできていますか？');
  footer(s);
})();

// ============================================================
//  スライド 21: 手法⑫ 家族でお金を共有する（2/2）相続・生前贈与
// ============================================================
(function() {
  const s = pptx.addSlide();
  bg(s, C.light);
  subHeader(s, '12', '家族でお金を共有する', '〜 相続・生前贈与は「今から話す」テーマ 〜', C.purple);

  s.addShape(pptx.ShapeType.rect, { x:0.42, y:1.3, w:12.28, h:0.52, fill:{color:C.navyMid} });
  s.addText('「相続の話は縁起が悪い」は過去の話。今すぐ話すことが資産を守る。', { x:0.52, y:1.32, w:12.0, h:0.48, fontSize:15, bold:true, color:C.gold, fontFace:FONT_JA });

  // 生前贈与の主な制度
  s.addShape(pptx.ShapeType.rect, { x:0.42, y:1.96, w:12.28, h:0.44, fill:{color:C.purple} });
  s.addText('主な生前贈与の非課税制度', { x:0.52, y:1.98, w:12.0, h:0.4, fontSize:14, bold:true, color:C.white, fontFace:FONT_JA });

  const gifts = [
    { title:'暦年贈与', limit:'年間 110万円まで非課税', detail:'毎年110万円以内であれば贈与税がかからない。親から子・孫へ毎年贈与することで、数年かけて大きな金額を移転できる。', color:C.purple },
    { title:'教育資金の\n一括贈与', limit:'最大 1,500万円', detail:'30歳未満の子・孫の教育費として一括贈与する場合、1,500万円（習い事は500万円）まで非課税。金融機関でプランを作る。', color:C.accent },
    { title:'住宅取得等\n資金の贈与', limit:'最大 1,000万円', detail:'子・孫が住宅を取得する際の資金として贈与する場合、一定額まで非課税。省エネ住宅なら非課税枠が拡大。', color:C.teal },
  ];
  const w = (12.84 - 0.42*2 - 0.22*2) / 3;
  gifts.forEach((g, i) => {
    const x = 0.42 + i*(w+0.22);
    s.addShape(pptx.ShapeType.rect, { x, y:2.52, w, h:2.58, fill:{color:C.white}, line:{color:'D8DEE6', pt:1} });
    s.addShape(pptx.ShapeType.rect, { x, y:2.52, w, h:0.5, fill:{color:g.color} });
    s.addText(g.title, { x, y:2.54, w, h:0.46, fontSize:14, bold:true, color:C.white, fontFace:FONT_JA, align:'center' });
    s.addShape(pptx.ShapeType.rect, { x:x+0.1, y:3.1, w:w-0.2, h:0.38, fill:{color:'EDE7F6'} });
    s.addText(g.limit, { x:x+0.1, y:3.12, w:w-0.2, h:0.34, fontSize:13, bold:true, color:g.color, fontFace:FONT_JA, align:'center' });
    s.addText(g.detail, { x:x+0.1, y:3.58, w:w-0.2, h:1.4, fontSize:12, color:C.darkgray, fontFace:FONT_JA });
  });

  noteBox(s, 5.24, '相続税の基礎控除：3,000万円 ＋（600万円 × 法定相続人数）。これを超える場合は税理士への相談を検討する。');
  discussBox(s, 5.98, '親に資産がある方：「生前贈与の話」を今日から家族でできるよう、まず自分が学ぶところから始めよう。');
  footer(s);
})();

// ============================================================
//  スライド 22: 12手法の合計まとめ
// ============================================================
(function() {
  const s = pptx.addSlide();
  bg(s, C.navy);
  s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:'100%', h:0.08, fill:{color:C.gold} });
  s.addText('12手法を組み合わせると…', { x:0.5, y:0.2, w:12.3, h:0.68, fontSize:28, bold:true, color:C.white, fontFace:FONT_JA });

  const groups = [
    { range:'①〜④', label:'支出設計・還元活用',     amt:'月 1.5〜7万円',  barC:C.accent },
    { range:'⑤〜⑧', label:'収入獲得・節税',         amt:'月 1.3〜13万円', barC:C.orange },
    { range:'⑨',    label:'自給思考',               amt:'月 1〜3万円',    barC:C.teal },
    { range:'⑩',    label:'高時給への転職',          amt:'月 3〜30万円↑', barC:C.teal },
    { range:'⑪',    label:'低金利の借金を積極活用',   amt:'資産拡大効果',   barC:C.teal },
    { range:'⑫',    label:'家族でお金を共有する',     amt:'家族全体で最適化',barC:C.purple },
  ];

  groups.forEach((g, i) => {
    const y = 1.05 + i * 0.88;
    s.addShape(pptx.ShapeType.rect, { x:0.42, y, w:12.28, h:0.76, fill:{color:C.navyMid}, line:{color:g.barC, pt:1} });
    s.addShape(pptx.ShapeType.rect, { x:0.42, y, w:0.82, h:0.76, fill:{color:g.barC} });
    s.addText(g.range, { x:0.42, y:y+0.12, w:0.82, h:0.52, fontSize:12, bold:true, color:C.white, fontFace:FONT_JA, align:'center' });
    s.addText(g.label, { x:1.38, y:y+0.16, w:7.5, h:0.44, fontSize:16, color:C.white, fontFace:FONT_JA });
    s.addText(g.amt,   { x:9.3, y:y+0.12, w:3.3, h:0.52, fontSize:17, bold:true, color:C.gold, fontFace:FONT_JA, align:'right' });
  });

  s.addShape(pptx.ShapeType.rect, { x:0.42, y:6.42, w:12.28, h:0.82, fill:{color:C.gold} });
  s.addText('12手法の組み合わせで、軍資金の最大化を実現できる', { x:0.55, y:6.46, w:12.0, h:0.74, fontSize:22, bold:true, color:C.navy, fontFace:FONT_JA });
  footer(s);
})();

// ============================================================
//  スライド 23: 40代からの正しい順番（ロードマップ）
// ============================================================
(function() {
  const s = pptx.addSlide();
  bg(s, C.light);
  h1(s, '40代からの資産形成の正しい順番', '〜 40代からでも間に合う。ただし「順番」を間違えないことが重要 〜');

  const steps = [
    { step:'STEP 1', sub:'まず土台を作る（〜3ヶ月）', color:C.accent, items:['先取り貯蓄の自動化', '固定費の見直し（スマホ・保険）', '税制優遇（ふるさと納税・iDeCo）', '家族でお金の共有を始める'] },
    { step:'STEP 2', sub:'軍資金を加速させる（3〜6ヶ月）', color:C.orange, items:['断捨離・不用品売却で種銭を作る', 'ポイントサイトで収益化', '副業スタート＋自給思考の実践', '転職・昇給の可能性を検討'] },
    { step:'STEP 3', sub:'さらに加速（6ヶ月〜）', color:C.teal, items:['低金利借金の戦略的活用を検討', '相続・生前贈与の家族会議を開く', '元本100万円を目安に運用スタート', '収益を再投資するサイクルを作る'] },
  ];

  steps.forEach((st, i) => {
    const x = 0.38 + i * 4.3;
    s.addShape(pptx.ShapeType.rect, { x, y:1.38, w:4.05, h:5.38, fill:{color:C.white}, line:{color:'D8DEE6', pt:1} });
    s.addShape(pptx.ShapeType.rect, { x, y:1.38, w:4.05, h:1.06, fill:{color:st.color} });
    s.addText(st.step, { x, y:1.41, w:4.05, h:0.58, fontSize:21, bold:true, color:C.white, fontFace:FONT_EN, align:'center' });
    s.addText(st.sub,  { x, y:1.93, w:4.05, h:0.42, fontSize:11, color:C.white, fontFace:FONT_JA, align:'center' });
    st.items.forEach((item, j) => {
      const iy = 2.62 + j * 0.97;
      s.addShape(pptx.ShapeType.rect, { x:x+0.2, y:iy+0.04, w:0.48, h:0.48, fill:{color:st.color} });
      s.addText('✓', { x:x+0.2, y:iy+0.04, w:0.48, h:0.48, fontSize:15, bold:true, color:C.white, fontFace:FONT_EN, align:'center' });
      s.addText(item, { x:x+0.78, y:iy, w:3.12, h:0.84, fontSize:13, color:C.darkgray, fontFace:FONT_JA });
    });
  });

  footer(s);
})();

// ============================================================
//  スライド 24: まとめ
// ============================================================
(function() {
  const s = pptx.addSlide();
  bg(s, C.navy);
  s.addShape(pptx.ShapeType.rect, { x:0, y:0,    w:'100%', h:0.08, fill:{color:C.gold} });
  s.addShape(pptx.ShapeType.rect, { x:0, y:6.3,  w:'100%', h:0.08, fill:{color:C.gold} });
  s.addShape(pptx.ShapeType.rect, { x:0, y:6.38, w:'100%', h:0.04, fill:{color:C.orange} });

  s.addText('本日のまとめ', { x:0.5, y:0.2, w:12.3, h:0.68, fontSize:28, bold:true, color:C.gold, fontFace:FONT_JA });

  const summaries = [
    { no:'01', text:'投資の結果は「銘柄」より「元本（軍資金）の大きさ」で8〜9割が決まる', color:C.orange },
    { no:'02', text:'資産形成の9割は「投資前」の準備で決まる。軍資金作りが最初の仕事', color:C.gold },
    { no:'03', text:'12の手法で軍資金を最大化できる。支出削減・収入増・レバレッジ・家族連携が4本柱', color:C.teal },
    { no:'04', text:'40代からでも間に合う。大切なのは「正しい順番」で・今日から動き始めること', color:C.accent },
  ];

  summaries.forEach((sm, i) => {
    const y = 1.1 + i * 1.24;
    s.addShape(pptx.ShapeType.rect, { x:0.42, y, w:12.28, h:1.06, fill:{color:C.navyMid}, line:{color:sm.color, pt:2} });
    s.addShape(pptx.ShapeType.rect, { x:0.42, y, w:0.82, h:1.06, fill:{color:sm.color} });
    s.addText(sm.no, { x:0.42, y:y+0.2, w:0.82, h:0.66, fontSize:26, bold:true, color:C.navy, fontFace:FONT_EN, align:'center' });
    s.addText(sm.text, { x:1.38, y:y+0.18, w:11.1, h:0.7, fontSize:17, color:C.white, fontFace:FONT_JA });
  });

  s.addShape(pptx.ShapeType.rect, { x:0.42, y:6.18, w:12.28, h:0.48, fill:{color:C.gold} });
  s.addText('軍資金を作ることが、資産形成の第一歩であり最も重要なステップです。  ─  ' + COMPANY, { x:0.42, y:6.2, w:12.28, h:0.44, fontSize:13, bold:true, color:C.navy, fontFace:FONT_JA, align:'center' });

  s.addText(COMPANY, { x:0, y:6.55, w:'100%', h:0.28, fontSize:10, color:C.silver, fontFace:FONT_JA, align:'center', italic:true });
})();

// ============================================================
//  保存
// ============================================================
const outputPath = 'D:/dev/軍資金作成資料/docs/軍資金作成講義_v3_スライド.pptx';
pptx.writeFile({ fileName: outputPath })
  .then(() => console.log('✅ Saved:', outputPath))
  .catch(err => console.error('❌', err));
