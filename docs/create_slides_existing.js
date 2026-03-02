// PowerPoint スライド生成スクリプト（既存受講者向け）
const PptxGenJS = require('C:/Users/haman/AppData/Roaming/npm/node_modules/pptxgenjs');
const path = require('path');

const pptx = new PptxGenJS();
pptx.layout = 'LAYOUT_WIDE';

const C = {
  navy:    '1A2744',
  gold:    'D4A843',
  orange:  'E8602C',
  white:   'FFFFFF',
  light:   'F4F6FA',
  gray:    '666666',
  darkgray:'333333',
  green:   '2E7D32',
  red:     'C62828',
  accent:  '1565C0',
  teal:    '00695C',  // 既存受講者向けのアクセントカラー
};

const FONT_JA = 'メイリオ';
const FONT_EN = 'Calibri';

function addBg(slide, color) {
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: '100%', h: '100%', fill: { color } });
}
function addFooter(slide, text) {
  slide.addText(text, { x: 0.3, y: 6.9, w: 12.7, h: 0.3, fontSize: 9, color: C.gray, fontFace: FONT_JA, align: 'right' });
}
function titleBox(slide, title, sub, opts = {}) {
  const y = opts.y || 0.5;
  slide.addShape(pptx.ShapeType.rect, { x: 0.4, y, w: 0.08, h: opts.th || 0.7, fill: { color: C.teal } });
  slide.addText(title, { x: 0.65, y, w: 11.7, h: opts.th || 0.7, fontSize: opts.fs || 28, bold: true, color: C.navy, fontFace: FONT_JA });
  if (sub) slide.addText(sub, { x: 0.65, y: y + (opts.th || 0.7) + 0.05, w: 11.7, h: 0.4, fontSize: 14, color: C.gray, fontFace: FONT_JA });
}

// ============================================================
//  スライド 01: 表紙（既存受講者向け）
// ============================================================
(function() {
  const s = pptx.addSlide();
  addBg(s, C.navy);
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 5.8, w: '100%', h: 0.06, fill: { color: C.teal } });
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 5.86, w: '100%', h: 0.04, fill: { color: C.gold } });

  s.addShape(pptx.ShapeType.rect, { x: 1.5, y: 0.4, w: 10.3, h: 0.55, fill: { color: C.teal } });
  s.addText('高速資産形成セミナー 受講者限定', {
    x: 1.5, y: 0.4, w: 10.3, h: 0.55,
    fontSize: 18, bold: true, color: C.white, fontFace: FONT_JA, align: 'center'
  });

  s.addText('投資の軍資金を作る', { x: 0.8, y: 1.15, w: 11.7, h: 0.9, fontSize: 40, bold: true, color: C.white, fontFace: FONT_JA, align: 'center' });
  s.addText('8つの方法', { x: 0.8, y: 2.05, w: 11.7, h: 1.2, fontSize: 58, bold: true, color: C.gold, fontFace: FONT_JA, align: 'center' });
  s.addText('〜 軍資金不足を解消し、最速でトレードをスタートする実践ガイド 〜', {
    x: 0.8, y: 3.35, w: 11.7, h: 0.5,
    fontSize: 16, color: 'AABBCC', fontFace: FONT_JA, align: 'center'
  });
  s.addShape(pptx.ShapeType.rect, { x: 3.5, y: 4.05, w: 6.3, h: 0.05, fill: { color: C.teal } });
  s.addText('「学んでいるのに動けない」を今日で終わりにする', {
    x: 0.8, y: 4.2, w: 11.7, h: 0.5,
    fontSize: 18, bold: true, color: C.gold, fontFace: FONT_JA, align: 'center'
  });
})();

// ============================================================
//  スライド 02: 「学んでいるのに動けない」の正体
// ============================================================
(function() {
  const s = pptx.addSlide();
  addBg(s, C.light);
  titleBox(s, '「学んでいるのに動けない」の正体');

  s.addShape(pptx.ShapeType.rect, { x: 0.4, y: 1.4, w: 12.3, h: 0.65, fill: { color: C.navy } });
  s.addText('これはあなただけではありません。受講者の多くが通る"あるある"です。', {
    x: 0.5, y: 1.42, w: 12.1, h: 0.61, fontSize: 16, bold: true, color: C.gold, fontFace: FONT_JA, align: 'center'
  });

  const patterns = [
    { label: 'パターン①', text: '受講費（44〜66万円）を捻出したことで手元の余裕資金が一時的に減った' },
    { label: 'パターン②', text: 'そもそも投資用の余剰資金がなく、講座を受けたが動けていない' },
    { label: 'パターン③', text: '軍資金の作り方を知らなかった・意識していなかった' },
    { label: 'パターン④', text: '収入は十分だが支出管理ができておらず、なかなか貯まらない' },
  ];

  patterns.forEach((p, i) => {
    const y = 2.25 + i * 1.0;
    s.addShape(pptx.ShapeType.rect, { x: 0.4, y, w: 12.3, h: 0.85, fill: { color: C.white }, line: { color: 'DDDDDD', pt: 1 } });
    s.addShape(pptx.ShapeType.rect, { x: 0.4, y, w: 1.7, h: 0.85, fill: { color: C.teal } });
    s.addText(p.label, { x: 0.4, y: y + 0.15, w: 1.7, h: 0.55, fontSize: 13, bold: true, color: C.white, fontFace: FONT_JA, align: 'center' });
    s.addText(p.text, { x: 2.2, y: y + 0.15, w: 10.3, h: 0.55, fontSize: 15, color: C.darkgray, fontFace: FONT_JA });
  });

  s.addShape(pptx.ShapeType.rect, { x: 0.4, y: 6.45, w: 12.3, h: 0.5, fill: { color: C.teal } });
  s.addText('今日の講義で全パターンの解決策をお伝えします', { x: 0.4, y: 6.45, w: 12.3, h: 0.5, fontSize: 16, bold: true, color: C.white, fontFace: FONT_JA, align: 'center' });
  addFooter(s, '高速資産形成セミナー ／ 受講者限定講義');
})();

// ============================================================
//  スライド 03: 機会損失の見える化
// ============================================================
(function() {
  const s = pptx.addSlide();
  addBg(s, C.navy);
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: '100%', h: 0.06, fill: { color: C.teal } });

  s.addText('1ヶ月遅れるたびに「逃がしている利益」がある', {
    x: 0.5, y: 0.2, w: 12.3, h: 0.7, fontSize: 26, bold: true, color: C.white, fontFace: FONT_JA
  });

  // 累積収益テーブル
  s.addShape(pptx.ShapeType.rect, { x: 0.5, y: 1.05, w: 12.3, h: 0.55, fill: { color: C.teal } });
  ['開始時期', '3ヶ月後', '6ヶ月後', '12ヶ月後'].forEach((h, i) => {
    const x = 0.5 + i * 3.1;
    s.addText(h, { x, y: 1.05, w: 3.0, h: 0.55, fontSize: 15, bold: true, color: C.white, fontFace: FONT_JA, align: 'center' });
  });

  const rows = [
    { label: '今月スタート', vals: ['約15万円', '約30万円', '約60万円'], color: C.gold },
    { label: '3ヶ月後スタート', vals: ['—', '約15万円', '約45万円'], color: 'AABBCC' },
    { label: '6ヶ月後スタート', vals: ['—', '—', '約30万円'], color: 'AABBCC' },
  ];

  rows.forEach((r, i) => {
    const y = 1.7 + i * 0.9;
    const bg = i === 0 ? '243060' : '1E2A44';
    s.addShape(pptx.ShapeType.rect, { x: 0.5, y, w: 12.3, h: 0.8, fill: { color: bg } });
    s.addText(r.label, { x: 0.7, y: y + 0.15, w: 2.8, h: 0.5, fontSize: 15, color: r.color, fontFace: FONT_JA, bold: i === 0 });
    r.vals.forEach((v, j) => {
      s.addText(v, { x: 0.5 + (j + 1) * 3.1, y: y + 0.15, w: 3.0, h: 0.5, fontSize: 16, bold: i === 0, color: r.color, fontFace: FONT_EN, align: 'center' });
    });
  });

  s.addText('※元本100万円、月5%収益目標の試算', { x: 0.5, y: 4.5, w: 12.3, h: 0.4, fontSize: 12, color: 'AABBCC', fontFace: FONT_JA });

  // 受講費回収
  s.addShape(pptx.ShapeType.rect, { x: 0.5, y: 5.0, w: 12.3, h: 1.3, fill: { color: '243060' }, line: { color: C.gold, pt: 2 } });
  s.addText('受講費の回収シミュレーション', { x: 0.7, y: 5.1, w: 5.0, h: 0.5, fontSize: 16, bold: true, color: C.gold, fontFace: FONT_JA });
  s.addText('44万円コース：月5万円収益なら 約9ヶ月で回収', { x: 0.7, y: 5.55, w: 11.9, h: 0.4, fontSize: 15, color: C.white, fontFace: FONT_JA });
  s.addText('66万円コース：月10万円収益なら 約7ヶ月で回収', { x: 0.7, y: 5.95, w: 11.9, h: 0.4, fontSize: 15, color: C.white, fontFace: FONT_JA });
  addFooter(s, '高速資産形成セミナー ／ 受講者限定講義');
})();

// ============================================================
//  手法スライド（既存受講者向け）共通ヘルパー
// ============================================================
function methodSlideEx(no, title, sub, actions, milestone, monthAmt) {
  const s = pptx.addSlide();
  addBg(s, C.light);

  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: '100%', h: 1.1, fill: { color: C.navy } });
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 1.08, w: '100%', h: 0.05, fill: { color: C.teal } });
  s.addShape(pptx.ShapeType.rect, { x: 0.4, y: 0.15, w: 0.8, h: 0.8, fill: { color: C.teal } });
  s.addText(no, { x: 0.4, y: 0.15, w: 0.8, h: 0.8, fontSize: 28, bold: true, color: C.white, fontFace: FONT_EN, align: 'center' });
  s.addText(title, { x: 1.4, y: 0.15, w: 8.5, h: 0.8, fontSize: 26, bold: true, color: C.white, fontFace: FONT_JA });
  s.addShape(pptx.ShapeType.rect, { x: 10.1, y: 0.2, w: 2.7, h: 0.7, fill: { color: C.gold } });
  s.addText(monthAmt, { x: 10.1, y: 0.2, w: 2.7, h: 0.7, fontSize: 17, bold: true, color: C.navy, fontFace: FONT_JA, align: 'center' });

  s.addText(sub, { x: 0.5, y: 1.2, w: 12.3, h: 0.4, fontSize: 14, color: C.gray, fontFace: FONT_JA });

  // 今日のアクション
  s.addShape(pptx.ShapeType.rect, { x: 0.4, y: 1.75, w: 2.2, h: 0.4, fill: { color: C.orange } });
  s.addText('今日のアクション', { x: 0.4, y: 1.75, w: 2.2, h: 0.4, fontSize: 13, bold: true, color: C.white, fontFace: FONT_JA, align: 'center' });

  actions.forEach((a, i) => {
    const y = 2.25 + i * 0.78;
    s.addShape(pptx.ShapeType.rect, { x: 0.4, y, w: 0.5, h: 0.55, fill: { color: C.orange } });
    s.addText(String(i + 1), { x: 0.4, y, w: 0.5, h: 0.55, fontSize: 18, bold: true, color: C.white, fontFace: FONT_EN, align: 'center' });
    s.addShape(pptx.ShapeType.rect, { x: 1.0, y, w: 11.7, h: 0.55, fill: { color: C.white }, line: { color: 'DDDDDD', pt: 1 } });
    s.addText(a, { x: 1.1, y: y + 0.05, w: 11.5, h: 0.45, fontSize: 15, color: C.darkgray, fontFace: FONT_JA });
  });

  // マイルストーン
  const milestoneY = 2.25 + actions.length * 0.78 + 0.2;
  s.addShape(pptx.ShapeType.rect, { x: 0.4, y: milestoneY, w: 12.3, h: 0.75, fill: { color: 'E8F5E9' }, line: { color: C.green, pt: 2 } });
  s.addText('🎯 マイルストーン：' + milestone, { x: 0.55, y: milestoneY + 0.1, w: 12.0, h: 0.55, fontSize: 14, color: C.green, fontFace: FONT_JA, bold: true });

  addFooter(s, '高速資産形成セミナー ／ 受講者限定講義');
}

// スライド 04〜11: 8手法（既存受講者向け）
methodSlideEx('01', '先取り貯蓄の自動化', '軍資金専用口座への自動積立を今日設定する',
  ['証券口座（楽天証券・SBI証券等）に「軍資金専用」の積立設定を今日中に行う',
   '手取りの10〜20%を目標に。月収25万円なら2.5〜5万円を先に確保',
   '「オプション資金」と明示したラベルの口座を作り、他の貯蓄と分ける'],
  '翌月末までに自動積立が1回完了していること', '月 2〜5万円');

methodSlideEx('02', '固定費の徹底見直し', '一度やれば毎月ずっと削減効果が続く',
  ['今週中：スマホ料金プランを確認し、格安SIM（MVNO）への切り替えを検討する',
   '今月中：生命保険の月額を確認し、FP無料相談を1件予約する',
   '今日：有料サブスクを全てリストアップし、使っていないものを解約する'],
  '来月の固定費が今月比5,000円以上減っていること', '月 5千〜3万円');

methodSlideEx('03', '変動費の最適化', '「助走期間」の一時的な生活コンパクト化',
  ['今週から：外食を週1回減らし、その分を現金で軍資金口座に入金する',
   '今日から：コンビニ利用をスーパーに切り替えるルールを設定する',
   '今月中：家計の「ノースペンドデー（消費ゼロの日）」を週1回設ける'],
  '今月の変動費が先月比10%（5,000〜10,000円）削減できていること', '月 3千〜2万円');

methodSlideEx('04', 'キャッシュレス還元の集約', 'すでに使っているお金をポイントに変換し即投資',
  ['高還元クレカ（楽天・PayPay等）に支出を集中させる設定を今月中に完了する',
   '楽天カードなら楽天証券の「ハッピープログラム」でポイントを直接投資に設定',
   '毎月末にポイント残高を確認し、全額を軍資金口座に移動するルーティンを作る'],
  '来月から毎月のポイント収益が軍資金口座に入金されていること', '月 2千〜1万円');

methodSlideEx('05', 'ポイントサイト活用（自己アフィリ）', 'オプション取引の準備作業でお金をもらう',
  ['今日：ハピタスまたはモッピーに登録（無料・5分）',
   '今週中：オプション取引に必要な証券口座をポイントサイト経由で開設する（1〜3万pt）',
   '毎月：ネットショッピング・各種サービスの利用を全てポイントサイト経由にする'],
  '初月に1万円以上のポイント収入を軍資金に加算していること', '初月 1〜5万円');

methodSlideEx('06', '不用品売却・断捨離収益化', '眠っている資産をオプションの種銭に変換',
  ['今日：家の中を15分で見回し「半年使っていないもの」をスマホでメモする',
   '今週末：メルカリまたはヤフオクに5点以上出品する（写真撮影含め1〜2時間）',
   '来月末まで：不用品売却で3〜10万円の一時収入を軍資金に加算する'],
  '1ヶ月以内に不用品売却で3万円以上を軍資金に追加していること', '初回 3〜10万円');

methodSlideEx('07', '副業・スキル収益化', '「軍資金フェーズ限定の副業」で加速する',
  ['今月中：クラウドワークスまたはランサーズに登録し、受注可能な案件を3件リストアップ',
   '今月中：ストアカまたはココナラで「自分が教えられること」を1件出品する',
   '副業収益は生活費に使わず「全額オプション口座に入金」ルールを今日決める'],
  '翌月末までに副業で1万円以上を稼ぎ、軍資金口座に入金していること', '月 1〜10万円');

methodSlideEx('08', '税制優遇制度の活用', '節税効果で軍資金を増やしながら確定申告の準備もする',
  ['今日：ふるさと納税のシミュレーションサイトで上限額を確認し、1件注文する',
   '今月中：iDeCoの公式サイトで掛金シミュレーションを行い、申込書を取り寄せる',
   '今日から：投資関連の経費（書籍・通信費等）のレシートを保管し始める'],
  '今年度中にふるさと納税・iDeCoの両方を設定していること', '年 5〜20万円');

// ============================================================
//  スライド 12: 個人別ロードマップ（ワーク）
// ============================================================
(function() {
  const s = pptx.addSlide();
  addBg(s, C.light);
  titleBox(s, '個人別ロードマップを今日作る', '〜 ワークシート 〜');

  s.addShape(pptx.ShapeType.rect, { x: 0.4, y: 1.4, w: 5.8, h: 5.2, fill: { color: C.white }, line: { color: 'DDDDDD', pt: 1 } });
  s.addShape(pptx.ShapeType.rect, { x: 0.4, y: 1.4, w: 5.8, h: 0.6, fill: { color: C.teal } });
  s.addText('現在の状況を確認', { x: 0.4, y: 1.4, w: 5.8, h: 0.6, fontSize: 16, bold: true, color: C.white, fontFace: FONT_JA, align: 'center' });

  const fields = ['現在の軍資金：＿＿＿万円', '目標の軍資金：＿＿＿万円（最低50万円）', '不足額：＿＿＿万円', '目標スタート月：20＿＿年＿＿月'];
  fields.forEach((f, i) => {
    s.addShape(pptx.ShapeType.rect, { x: 0.5, y: 2.15 + i * 0.95, w: 5.6, h: 0.75, fill: { color: 'F9F9F9' }, line: { color: 'CCCCCC', pt: 1 } });
    s.addText(f, { x: 0.65, y: 2.25 + i * 0.95, w: 5.4, h: 0.55, fontSize: 14, color: C.darkgray, fontFace: FONT_JA });
  });

  s.addShape(pptx.ShapeType.rect, { x: 6.5, y: 1.4, w: 6.3, h: 5.2, fill: { color: C.white }, line: { color: 'DDDDDD', pt: 1 } });
  s.addShape(pptx.ShapeType.rect, { x: 6.5, y: 1.4, w: 6.3, h: 0.6, fill: { color: C.orange } });
  s.addText('活用する手法と月間効果', { x: 6.5, y: 1.4, w: 6.3, h: 0.6, fontSize: 16, bold: true, color: C.white, fontFace: FONT_JA, align: 'center' });

  const methods = ['①先取り貯蓄', '②固定費削減', '③変動費最適化', '④キャッシュレス', '⑤ポイントサイト', '⑦副業'];
  methods.forEach((m, i) => {
    const y = 2.1 + i * 0.75;
    s.addText(m + '：月＿＿＿円', { x: 6.7, y, w: 5.9, h: 0.6, fontSize: 13, color: C.darkgray, fontFace: FONT_JA });
    s.addShape(pptx.ShapeType.rect, { x: 6.5, y: y + 0.55, w: 6.3, h: 0.04, fill: { color: 'EEEEEE' } });
  });

  s.addShape(pptx.ShapeType.rect, { x: 6.5, y: 6.05, w: 6.3, h: 0.55, fill: { color: C.gold } });
  s.addText('合計：月＿＿＿円　→　あと＿＿ヶ月でスタート', { x: 6.5, y: 6.05, w: 6.3, h: 0.55, fontSize: 13, bold: true, color: C.navy, fontFace: FONT_JA, align: 'center' });
  addFooter(s, '高速資産形成セミナー ／ 受講者限定講義');
})();

// ============================================================
//  スライド 13: コミット宣言
// ============================================================
(function() {
  const s = pptx.addSlide();
  addBg(s, C.navy);
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: '100%', h: 0.06, fill: { color: C.teal } });

  s.addText('コミット宣言', { x: 0.5, y: 0.2, w: 12.3, h: 0.6, fontSize: 28, bold: true, color: C.gold, fontFace: FONT_JA, align: 'center' });

  s.addShape(pptx.ShapeType.rect, { x: 0.8, y: 1.0, w: 11.7, h: 3.2, fill: { color: '243060' }, line: { color: C.teal, pt: 2 } });
  s.addText('「 私は　　　　ヶ月後（20　　年　　月）までに、', { x: 1.0, y: 1.2, w: 11.3, h: 0.7, fontSize: 20, color: C.white, fontFace: FONT_JA });
  s.addText('　月　　　　万円の軍資金を準備して、', { x: 1.0, y: 1.9, w: 11.3, h: 0.7, fontSize: 20, color: C.white, fontFace: FONT_JA });
  s.addText('　オプション取引を開始します。', { x: 1.0, y: 2.6, w: 11.3, h: 0.7, fontSize: 20, color: C.white, fontFace: FONT_JA });
  s.addText('　そのために今月中に＿＿＿＿＿を実行します。 」', { x: 1.0, y: 3.3, w: 11.3, h: 0.7, fontSize: 18, color: C.gold, fontFace: FONT_JA });

  // サポート体制
  s.addShape(pptx.ShapeType.rect, { x: 0.5, y: 4.4, w: 12.3, h: 0.5, fill: { color: C.teal } });
  s.addText('サポート体制', { x: 0.5, y: 4.4, w: 12.3, h: 0.5, fontSize: 16, bold: true, color: C.white, fontFace: FONT_JA, align: 'center' });

  const supports = ['コミュニティ（進捗報告・情報共有）', '月次フォローアップ講義', '個別相談（家計・副業アドバイス）', '取引開始サポート（最初の注文を一緒に確認）'];
  supports.forEach((sp, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = col === 0 ? 0.5 : 6.7;
    const y = 5.05 + row * 0.7;
    s.addShape(pptx.ShapeType.rect, { x, y, w: 5.9, h: 0.55, fill: { color: '243060' } });
    s.addText('✅  ' + sp, { x: x + 0.2, y: y + 0.05, w: 5.6, h: 0.45, fontSize: 14, color: C.white, fontFace: FONT_JA });
  });
  addFooter(s, '高速資産形成セミナー ／ 受講者限定講義');
})();

// ============================================================
//  保存
// ============================================================
const outputPath = path.join('D:/dev/軍資金作成資料/docs', '軍資金作成講義_既存受講者向けスライド.pptx');
pptx.writeFile({ fileName: outputPath })
  .then(() => console.log('✅ PowerPoint saved:', outputPath))
  .catch(err => console.error('❌ Error:', err));
