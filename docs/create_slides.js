// PowerPoint スライド生成スクリプト
// 高速資産形成セミナー集客講義：投資の軍資金を作る8つの方法

const PptxGenJS = require('C:/Users/haman/AppData/Roaming/npm/node_modules/pptxgenjs');
const path = require('path');

const pptx = new PptxGenJS();

// ========== カラーパレット ==========
const C = {
  navy:    '1A2744',  // メイン背景（濃紺）
  gold:    'D4A843',  // アクセント（金）
  orange:  'E8602C',  // 強調（オレンジ）
  white:   'FFFFFF',
  light:   'F4F6FA',  // コンテンツ背景（薄グレー）
  gray:    '666666',
  darkgray:'333333',
  green:   '2E7D32',
  red:     'C62828',
  accent:  '1565C0',  // サブアクセント（青）
};

// ========== レイアウト ==========
pptx.layout = 'LAYOUT_WIDE'; // 16:9

// ========== スライドマスター的な共通設定 ==========
const FONT_JA = 'メイリオ';
const FONT_EN = 'Calibri';

// ---------- ユーティリティ ----------
function addBg(slide, color) {
  slide.addShape(pptx.ShapeType.rect, {
    x: 0, y: 0, w: '100%', h: '100%',
    fill: { color }
  });
}

function addFooter(slide, text) {
  slide.addText(text, {
    x: 0.3, y: 6.9, w: 12.7, h: 0.3,
    fontSize: 9, color: C.gray, fontFace: FONT_JA,
    align: 'right'
  });
}

function titleBox(slide, title, sub, opts = {}) {
  const y = opts.y || 0.5;
  slide.addShape(pptx.ShapeType.rect, {
    x: 0.4, y: y, w: 0.08, h: opts.th || 0.7,
    fill: { color: C.gold }
  });
  slide.addText(title, {
    x: 0.65, y: y, w: 11.7, h: opts.th || 0.7,
    fontSize: opts.fs || 28, bold: true,
    color: C.navy, fontFace: FONT_JA,
  });
  if (sub) {
    slide.addText(sub, {
      x: 0.65, y: y + (opts.th || 0.7) + 0.05, w: 11.7, h: 0.4,
      fontSize: 14, color: C.gray, fontFace: FONT_JA,
    });
  }
}

// ============================================================
//  スライド 01: 表紙
// ============================================================
(function() {
  const s = pptx.addSlide();
  addBg(s, C.navy);

  // 背景装飾ライン
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 5.8, w: '100%', h: 0.06, fill: { color: C.gold } });
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 5.9, w: '100%', h: 0.04, fill: { color: C.orange } });

  s.addText('投資の軍資金を作る', {
    x: 0.8, y: 1.0, w: 11.7, h: 1.0,
    fontSize: 44, bold: true, color: C.white, fontFace: FONT_JA, align: 'center'
  });
  s.addText('8つの方法', {
    x: 0.8, y: 2.0, w: 11.7, h: 1.2,
    fontSize: 60, bold: true, color: C.gold, fontFace: FONT_JA, align: 'center'
  });
  s.addText('〜 将来不安をゼロにする資金作りの全技術 〜', {
    x: 0.8, y: 3.3, w: 11.7, h: 0.5,
    fontSize: 18, color: 'AABBCC', fontFace: FONT_JA, align: 'center'
  });
  s.addShape(pptx.ShapeType.rect, { x: 3.5, y: 4.0, w: 6.3, h: 0.05, fill: { color: C.gold } });
  s.addText('高速資産形成セミナー  集客講義', {
    x: 0.8, y: 4.2, w: 11.7, h: 0.5,
    fontSize: 16, color: 'CCDDEE', fontFace: FONT_JA, align: 'center'
  });
})();

// ============================================================
//  スライド 02: 本日のゴール
// ============================================================
(function() {
  const s = pptx.addSlide();
  addBg(s, C.light);
  titleBox(s, '本日の講義ゴール');

  const items = [
    { icon: '①', text: '軍資金（投資元本）を作る8つの具体的手法を知る', color: C.navy },
    { icon: '②', text: '月3〜10万円の軍資金を継続的に確保できるようになる', color: C.navy },
    { icon: '③', text: '貯めた軍資金を「どう増やすか」の選択肢を広げる', color: C.orange },
  ];

  items.forEach((item, i) => {
    const y = 1.6 + i * 1.4;
    s.addShape(pptx.ShapeType.rect, { x: 0.5, y, w: 12.3, h: 1.1, fill: { color: C.white }, line: { color: 'DDDDDD', pt: 1 } });
    s.addShape(pptx.ShapeType.ellipse, { x: 0.7, y: y + 0.2, w: 0.7, h: 0.7, fill: { color: item.color } });
    s.addText(item.icon, { x: 0.7, y: y + 0.2, w: 0.7, h: 0.7, fontSize: 16, bold: true, color: C.white, fontFace: FONT_EN, align: 'center' });
    s.addText(item.text, { x: 1.6, y: y + 0.2, w: 11.0, h: 0.7, fontSize: 18, color: C.darkgray, fontFace: FONT_JA });
  });
  addFooter(s, '高速資産形成セミナー集客講義');
})();

// ============================================================
//  スライド 03: 将来不安の現実
// ============================================================
(function() {
  const s = pptx.addSlide();
  addBg(s, C.navy);
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: '100%', h: 0.06, fill: { color: C.gold } });

  s.addText('あなたの老後、今のままで大丈夫ですか？', {
    x: 0.5, y: 0.3, w: 12.3, h: 0.8,
    fontSize: 30, bold: true, color: C.white, fontFace: FONT_JA, align: 'center'
  });

  const facts = [
    { num: '¥148,000', label: '厚生年金の平均月額', sub: '（2024年度）', color: C.orange },
    { num: '0.1%', label: '銀行普通預金の金利', sub: '100万円で年1,000円', color: C.red },
    { num: '+30%', label: '物価上昇（2020年比）', sub: '実質的な購買力は低下', color: C.orange },
    { num: '2,000万円', label: '老後に必要と言われる不足額', sub: '（金融庁試算）', color: C.red },
  ];

  facts.forEach((f, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = col === 0 ? 0.5 : 6.7;
    const y = 1.3 + row * 2.3;
    s.addShape(pptx.ShapeType.rect, { x, y, w: 5.9, h: 2.0, fill: { color: '243060' }, line: { color: f.color, pt: 2 } });
    s.addText(f.num, { x, y: y + 0.1, w: 5.9, h: 1.0, fontSize: 38, bold: true, color: f.color, fontFace: FONT_EN, align: 'center' });
    s.addText(f.label, { x, y: y + 1.0, w: 5.9, h: 0.5, fontSize: 16, bold: true, color: C.white, fontFace: FONT_JA, align: 'center' });
    s.addText(f.sub, { x, y: y + 1.5, w: 5.9, h: 0.4, fontSize: 12, color: 'AABBCC', fontFace: FONT_JA, align: 'center' });
  });
  addFooter(s, '高速資産形成セミナー集客講義');
})();

// ============================================================
//  スライド 04: 軍資金という発想
// ============================================================
(function() {
  const s = pptx.addSlide();
  addBg(s, C.light);
  titleBox(s, '「軍資金」という発想の転換');

  s.addShape(pptx.ShapeType.rect, { x: 0.5, y: 1.5, w: 12.3, h: 1.0, fill: { color: C.navy } });
  s.addText('給料 = 生活費だけではなく「生活費」＋「軍資金」に分けて考える', {
    x: 0.5, y: 1.5, w: 12.3, h: 1.0,
    fontSize: 20, bold: true, color: C.gold, fontFace: FONT_JA, align: 'center'
  });

  // フロー図
  const boxes = [
    { x: 0.6, label: '給与収入', color: C.accent },
    { x: 4.1, label: '生活費\n（支出）', color: C.gray },
    { x: 7.6, label: '軍資金\n（投資元本）', color: C.green },
    { x: 11.1, label: '資産形成\n（未来の収入）', color: C.orange },
  ];

  boxes.forEach((b, i) => {
    s.addShape(pptx.ShapeType.rect, { x: b.x, y: 2.9, w: 3.0, h: 1.4, fill: { color: b.color }, line: { color: 'FFFFFF', pt: 1 } });
    s.addText(b.label, { x: b.x, y: 2.9, w: 3.0, h: 1.4, fontSize: 18, bold: true, color: C.white, fontFace: FONT_JA, align: 'center' });
    if (i < 3) {
      s.addShape(pptx.ShapeType.rect, { x: b.x + 3.05, y: 3.35, w: 0.4, h: 0.5, fill: { color: C.gold } });
      s.addText('▶', { x: b.x + 3.05, y: 3.35, w: 0.4, h: 0.5, fontSize: 18, bold: true, color: C.navy, fontFace: FONT_EN, align: 'center' });
    }
  });

  s.addText('今日の講義：軍資金を作る「8つの方法」をすべてお伝えします', {
    x: 0.5, y: 4.6, w: 12.3, h: 0.6,
    fontSize: 18, bold: true, color: C.orange, fontFace: FONT_JA, align: 'center'
  });
  addFooter(s, '高速資産形成セミナー集客講義');
})();

// ============================================================
//  スライド 05: 8手法の全体マップ
// ============================================================
(function() {
  const s = pptx.addSlide();
  addBg(s, C.navy);
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: '100%', h: 0.06, fill: { color: C.gold } });

  s.addText('8つの軍資金作成法　全体マップ', {
    x: 0.5, y: 0.2, w: 12.3, h: 0.7,
    fontSize: 26, bold: true, color: C.white, fontFace: FONT_JA
  });

  const methods = [
    { no: '01', title: '先取り貯蓄の自動化',     amt: '月1〜3万円',     cat: '支出削減' },
    { no: '02', title: '固定費の徹底見直し',     amt: '月5千〜3万円',   cat: '支出削減' },
    { no: '03', title: '変動費の最適化',         amt: '月3千〜1.5万円', cat: '支出削減' },
    { no: '04', title: 'キャッシュレス還元の集約', amt: '月2千〜1万円',   cat: '支出最適化' },
    { no: '05', title: 'ポイントサイト活用',      amt: '月3千〜3万円',   cat: '収入増加' },
    { no: '06', title: '不用品売却・断捨離',      amt: '初回〜5万円',    cat: '収入増加' },
    { no: '07', title: '副業・スキル収益化',      amt: '月1〜10万円',    cat: '収入増加' },
    { no: '08', title: '税制優遇制度の活用',      amt: '年5〜20万円',    cat: '節税' },
  ];

  methods.forEach((m, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = col === 0 ? 0.4 : 6.65;
    const y = 1.1 + row * 1.35;
    const catColor = m.cat === '支出削減' ? C.accent : m.cat === '収入増加' ? C.green : C.orange;

    s.addShape(pptx.ShapeType.rect, { x, y, w: 5.9, h: 1.15, fill: { color: '243060' }, line: { color: '3A4A80', pt: 1 } });
    s.addShape(pptx.ShapeType.rect, { x, y, w: 0.55, h: 1.15, fill: { color: C.gold } });
    s.addText(m.no, { x, y: y + 0.25, w: 0.55, h: 0.65, fontSize: 16, bold: true, color: C.navy, fontFace: FONT_EN, align: 'center' });
    s.addText(m.title, { x: x + 0.6, y: y + 0.05, w: 3.5, h: 0.65, fontSize: 15, bold: true, color: C.white, fontFace: FONT_JA });
    s.addShape(pptx.ShapeType.rect, { x: x + 0.6, y: y + 0.72, w: 0.9, h: 0.3, fill: { color: catColor } });
    s.addText(m.cat, { x: x + 0.6, y: y + 0.72, w: 0.9, h: 0.3, fontSize: 10, color: C.white, fontFace: FONT_JA, align: 'center' });
    s.addText(m.amt, { x: x + 4.15, y: y + 0.3, w: 1.7, h: 0.55, fontSize: 13, bold: true, color: C.gold, fontFace: FONT_JA, align: 'right' });
  });
  addFooter(s, '高速資産形成セミナー集客講義');
})();

// ============================================================
//  手法スライド共通ヘルパー
// ============================================================
function methodSlide(no, title, sub, details, limit, monthAmt) {
  const s = pptx.addSlide();
  addBg(s, C.light);

  // ヘッダー
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: '100%', h: 1.1, fill: { color: C.navy } });
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 1.08, w: '100%', h: 0.05, fill: { color: C.gold } });
  s.addShape(pptx.ShapeType.rect, { x: 0.4, y: 0.15, w: 0.8, h: 0.8, fill: { color: C.gold } });
  s.addText(no, { x: 0.4, y: 0.15, w: 0.8, h: 0.8, fontSize: 28, bold: true, color: C.navy, fontFace: FONT_EN, align: 'center' });
  s.addText(title, { x: 1.4, y: 0.15, w: 8.5, h: 0.8, fontSize: 28, bold: true, color: C.white, fontFace: FONT_JA });
  s.addShape(pptx.ShapeType.rect, { x: 10.1, y: 0.2, w: 2.7, h: 0.7, fill: { color: C.orange } });
  s.addText(monthAmt, { x: 10.1, y: 0.2, w: 2.7, h: 0.7, fontSize: 18, bold: true, color: C.white, fontFace: FONT_JA, align: 'center' });

  s.addText(sub, { x: 0.5, y: 1.2, w: 12.3, h: 0.5, fontSize: 16, color: C.gray, fontFace: FONT_JA });

  // 詳細
  details.forEach((d, i) => {
    const y = 1.85 + i * 0.85;
    s.addShape(pptx.ShapeType.rect, { x: 0.4, y: y + 0.1, w: 0.5, h: 0.5, fill: { color: C.accent } });
    s.addText('✓', { x: 0.4, y: y + 0.1, w: 0.5, h: 0.5, fontSize: 16, bold: true, color: C.white, fontFace: FONT_EN, align: 'center' });
    s.addText(d, { x: 1.05, y, w: 11.6, h: 0.75, fontSize: 16, color: C.darkgray, fontFace: FONT_JA });
  });

  // 限界の伏線ボックス
  s.addShape(pptx.ShapeType.rect, { x: 0.4, y: 5.5, w: 12.3, h: 0.7, fill: { color: 'FFF3E0' }, line: { color: C.orange, pt: 2 } });
  s.addText('⚠  限界：' + limit, { x: 0.5, y: 5.55, w: 12.1, h: 0.6, fontSize: 14, color: C.orange, fontFace: FONT_JA, bold: true });

  addFooter(s, '高速資産形成セミナー集客講義');
  return s;
}

// ============================================================
//  スライド 06: 手法①
// ============================================================
methodSlide('01', '先取り貯蓄の自動化', '〜「意志力」に頼らない仕組みを作る〜',
  [
    '給与振込日の翌日に証券口座・専用口座へ「自動転送」を設定する',
    '手取りの10%からスタート。月20万円なら2万円をそのまま"存在しないお金"にする',
    '住信SBIネット銀行「目的別口座」や楽天銀行「自動振替」を活用',
    '米国発「Pay Yourself First（自分への先払い）」— 資産形成の鉄則',
  ],
  'これだけでは増えない。確保できるが運用しなければ意味がない。',
  '月 1〜3万円'
);

// ============================================================
//  スライド 07: 手法②
// ============================================================
methodSlide('02', '固定費の徹底見直し', '〜一度やれば毎月ずっと効く"最高コスパ"の節約〜',
  [
    '携帯：大手キャリア（〜8,000円）→ MVNO・格安SIM（〜1,500円）で月6,500円削減',
    '生命保険：40〜60代は保障が過大になりやすい。見直しで月1万円以上削減も',
    'サブスク棚卸し：使っていないNetflix・音楽・ツールを洗い出し月2,000〜5,000円削減',
    '電気・ガス切り替え：新電力・ガス自由化で月1,000〜3,000円削減',
  ],
  '削れる限度がある。固定費はゼロにはできない。上限が見えてくる。',
  '月 5千〜3万円'
);

// ============================================================
//  スライド 08: 手法③
// ============================================================
methodSlide('03', '変動費の最適化', '〜ゼロにしなくていい。"漏れ"をふさぐだけ〜',
  [
    '食費：まとめ買い＋冷凍活用。業務スーパー・コストコで食材費30%削減も可能',
    '外食：QRコード決済ポイント還元デー（PayPay20%還元など）を活用',
    '趣味・娯楽：図書館・動画共有プランで代替。月2,000〜5,000円の削減',
    '米国「ノースペンドデー」実践：週1日を消費ゼロの日に。月4回で意識が変わる',
  ],
  '節約だけでは人生の豊かさと逆行する。限界があり、精神的コストも高い。',
  '月 3千〜1.5万円'
);

// ============================================================
//  スライド 09: 手法④
// ============================================================
methodSlide('04', 'キャッシュレス還元の集約', '〜すでに使っているお金を"ポイント"に変換〜',
  [
    'クレカを1〜2枚に集約。楽天カード（最大3%）・PayPayカード（最大5%）など高還元に統一',
    '公共料金・保険料・通販・ガソリンもすべてカード払いに変更する',
    '月30万円の支出 × 還元率2% ＝ 月6,000円 → 年72,000円の軍資金',
    '楽天ポイント投資・PayPay資産運用など「ポイントをそのまま投資」に回す設定も有効',
  ],
  '使った金額以上は還元されない。収入が増えるわけではなく、上限が決まっている。',
  '月 2千〜1万円'
);

// ============================================================
//  スライド 10: 手法⑤
// ============================================================
methodSlide('05', 'ポイントサイト・アフィリエイト活用', '〜日常のサービス申し込みを"収入"に変える〜',
  [
    'A8.net：国内最大ASP。ブログ不要で自己アフィリエイト可能。証券口座開設で1〜3万円',
    'ハピタス：クレカ発行・FX口座開設等の高額案件多数。1案件で数千〜数万ポイント獲得',
    'ポイントタウン：ゲーム・アンケート・ショッピングで毎日コツコツ獲得→換金',
    'モッピー：ネットショッピングのポイント二重取り。日常の買い物を全て経由するだけ',
  ],
  '案件数に上限がある。継続収入にはなりにくく、受動的な稼ぎには限界がある。',
  '月 3千〜3万円'
);

// ============================================================
//  スライド 11: 手法⑥
// ============================================================
methodSlide('06', '不用品売却・断捨離収益化', '〜眠っている「資産」を現金化〜',
  [
    'メルカリ：衣類・雑貨・書籍。写真1枚で出品できる手軽さ。初回断捨離で数万円も',
    'ヤフオク：ブランド品・レア品・趣味用品。メルカリより高値がつきやすい',
    'ハードオフ・買取専門店：即現金化。家電・楽器・カメラ類に特に有効',
    '「一度きり」だからこそ種銭作りに最適。副産物として衝動買い防止の意識改革も',
  ],
  '家の中の不用品は有限。売れるものを使い切ったら終わり。継続収入にはならない。',
  '初回 〜5万円'
);

// ============================================================
//  スライド 12: 手法⑦
// ============================================================
methodSlide('07', '副業・スキル収益化', '〜本業以外の「第二の収入源」を作る〜',
  [
    'クラウドワークス・ランサーズ：ライティング・データ入力・翻訳。月1〜5万円から',
    'ストアカ・ココナラ：職歴・趣味・専門知識を「教える商品」に変換。月1〜10万円',
    'せどり・転売：Amazonやメルカリへの仕入れ転売。月3〜20万円の実績者も多い',
    '長年の職歴・人脈・専門知識は「副業の武器」。まず月1〜2万円達成を目標に',
  ],
  '時間を売っている限り収入に上限がある。体力・時間の制約から逃れられない。',
  '月 1〜10万円'
);

// ============================================================
//  スライド 13: 手法⑧
// ============================================================
methodSlide('08', '税制優遇制度の活用', '〜払わなくていい税金を取り戻す〜',
  [
    'iDeCo：掛金が全額所得控除。年収600万円で月2万円拠出すれば年間約4〜6万円の節税',
    'ふるさと納税：実質2,000円で返礼品受取＋住民税控除。食費削減にもなる一石二鳥',
    '医療費控除：年間10万円超の医療費は確定申告で還付。意外と知らない人が多い',
    '副業がある場合は青色申告：最大65万円の特別控除。経費計上で課税所得を大幅に圧縮',
  ],
  '節税の上限は決まっている。制度の恩恵を使い切ってしまえば、それ以上は増えない。',
  '年 5〜20万円'
);

// ============================================================
//  スライド 14: 8手法の合計まとめ
// ============================================================
(function() {
  const s = pptx.addSlide();
  addBg(s, C.navy);
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: '100%', h: 0.06, fill: { color: C.gold } });

  s.addText('8手法を組み合わせると…', {
    x: 0.5, y: 0.2, w: 12.3, h: 0.7,
    fontSize: 28, bold: true, color: C.white, fontFace: FONT_JA
  });

  const rows = [
    { no: '01〜04', label: '節約・還元系（4手法合計）', amt: '月 1.5〜7万円',  color: C.accent },
    { no: '05〜07', label: '収入増加系（3手法合計）',   amt: '月 1.3〜13万円', color: C.green },
    { no: '08',     label: '節税系',                  amt: '年 5〜20万円',   color: C.orange },
  ];

  rows.forEach((r, i) => {
    const y = 1.1 + i * 1.2;
    s.addShape(pptx.ShapeType.rect, { x: 0.5, y, w: 12.3, h: 1.0, fill: { color: '243060' }, line: { color: r.color, pt: 2 } });
    s.addShape(pptx.ShapeType.rect, { x: 0.5, y, w: 1.0, h: 1.0, fill: { color: r.color } });
    s.addText(r.no, { x: 0.5, y: y + 0.2, w: 1.0, h: 0.6, fontSize: 14, bold: true, color: C.white, fontFace: FONT_EN, align: 'center' });
    s.addText(r.label, { x: 1.7, y: y + 0.2, w: 7.5, h: 0.6, fontSize: 18, color: C.white, fontFace: FONT_JA });
    s.addText(r.amt, { x: 9.3, y: y + 0.15, w: 3.3, h: 0.7, fontSize: 20, bold: true, color: C.gold, fontFace: FONT_JA, align: 'right' });
  });

  // 合計
  s.addShape(pptx.ShapeType.rect, { x: 0.5, y: 4.75, w: 12.3, h: 1.3, fill: { color: C.gold } });
  s.addText('合計でおよそ…', { x: 0.8, y: 4.85, w: 4.0, h: 0.5, fontSize: 16, color: C.navy, fontFace: FONT_JA });
  s.addText('月 3〜20万円 の軍資金を作ることが可能', {
    x: 0.8, y: 5.25, w: 12.0, h: 0.7,
    fontSize: 26, bold: true, color: C.navy, fontFace: FONT_JA
  });
  addFooter(s, '高速資産形成セミナー集客講義');
})();

// ============================================================
//  スライド 15: ブリッジ「その軍資金、どこに置く？」
// ============================================================
(function() {
  const s = pptx.addSlide();
  addBg(s, C.light);
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: '100%', h: 1.1, fill: { color: C.navy } });
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 1.08, w: '100%', h: 0.05, fill: { color: C.orange } });
  s.addText('さあ、次の問いです', { x: 0.5, y: 0.05, w: 12.3, h: 0.5, fontSize: 18, color: 'AABBCC', fontFace: FONT_JA });
  s.addText('貯めた軍資金を、どこに置きますか？', {
    x: 0.5, y: 0.45, w: 12.3, h: 0.7,
    fontSize: 30, bold: true, color: C.gold, fontFace: FONT_JA
  });

  const opts = [
    { label: '銀行預金', detail: '金利0.1%\n100万円で年1,000円', verdict: '論外', vc: C.red },
    { label: '長期積立\nインデックス', detail: '優良な方法\nただし成果まで20〜30年', verdict: '時間がかかる', vc: C.orange },
    { label: '高配当株\nREIT', detail: '年4〜5%配当\n月15万円には3,000〜4,500万円の元本が必要', verdict: '元本が足りない', vc: C.orange },
  ];

  opts.forEach((o, i) => {
    const x = 0.5 + i * 4.3;
    s.addShape(pptx.ShapeType.rect, { x, y: 1.4, w: 4.0, h: 3.4, fill: { color: C.white }, line: { color: 'DDDDDD', pt: 1 } });
    s.addShape(pptx.ShapeType.rect, { x, y: 1.4, w: 4.0, h: 0.7, fill: { color: C.accent } });
    s.addText(o.label, { x, y: 1.42, w: 4.0, h: 0.66, fontSize: 16, bold: true, color: C.white, fontFace: FONT_JA, align: 'center' });
    s.addText(o.detail, { x: x + 0.2, y: 2.2, w: 3.6, h: 1.8, fontSize: 14, color: C.darkgray, fontFace: FONT_JA, align: 'center' });
    s.addShape(pptx.ShapeType.rect, { x: x + 0.4, y: 4.1, w: 3.2, h: 0.55, fill: { color: o.vc } });
    s.addText('→ ' + o.verdict, { x: x + 0.4, y: 4.1, w: 3.2, h: 0.55, fontSize: 14, bold: true, color: C.white, fontFace: FONT_JA, align: 'center' });
  });

  s.addText('「もし、もっと早く・少ない元手で 月数万〜15万円 を得られる方法があったとしたら？」', {
    x: 0.4, y: 5.0, w: 12.5, h: 0.8,
    fontSize: 18, bold: true, color: C.orange, fontFace: FONT_JA, align: 'center'
  });
  addFooter(s, '高速資産形成セミナー集客講義');
})();

// ============================================================
//  スライド 16: 米国株オプション取引とは
// ============================================================
(function() {
  const s = pptx.addSlide();
  addBg(s, C.navy);
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: '100%', h: 0.06, fill: { color: C.gold } });

  s.addText('解決策：米国株オプション取引', {
    x: 0.5, y: 0.2, w: 12.3, h: 0.7,
    fontSize: 28, bold: true, color: C.white, fontFace: FONT_JA
  });
  s.addText('〜「買う・売る」の2択から、「権利を売買する」世界へ〜', {
    x: 0.5, y: 0.9, w: 12.3, h: 0.4,
    fontSize: 16, color: 'AABBCC', fontFace: FONT_JA
  });

  const points = [
    { icon: '01', title: '毎月収益が狙える', detail: 'オプションの時間的価値を"売る"ことで、株が動かなくても毎月プレミアムを受取れる' },
    { icon: '02', title: '長期投資と並行可能', detail: '保有株を使ったオプション戦略（カバードコール等）は既存の資産と組み合わせて活用できる' },
    { icon: '03', title: '月数万〜15万円の実績', detail: '元本100〜300万円から始めて月5〜15万円の収益を上げているメンバーが多数在籍' },
  ];

  points.forEach((p, i) => {
    const y = 1.5 + i * 1.5;
    s.addShape(pptx.ShapeType.rect, { x: 0.5, y, w: 12.3, h: 1.3, fill: { color: '243060' }, line: { color: C.gold, pt: 1 } });
    s.addShape(pptx.ShapeType.rect, { x: 0.5, y, w: 0.8, h: 1.3, fill: { color: C.gold } });
    s.addText(p.icon, { x: 0.5, y: y + 0.3, w: 0.8, h: 0.7, fontSize: 22, bold: true, color: C.navy, fontFace: FONT_EN, align: 'center' });
    s.addText(p.title, { x: 1.5, y: y + 0.1, w: 5.0, h: 0.6, fontSize: 20, bold: true, color: C.gold, fontFace: FONT_JA });
    s.addText(p.detail, { x: 1.5, y: y + 0.65, w: 11.0, h: 0.55, fontSize: 14, color: 'CCDDEE', fontFace: FONT_JA });
  });
  addFooter(s, '高速資産形成セミナー集客講義');
})();

// ============================================================
//  スライド 17: セミナー案内（CTA）
// ============================================================
(function() {
  const s = pptx.addSlide();
  addBg(s, C.navy);
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 5.8, w: '100%', h: 0.06, fill: { color: C.gold } });
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 5.86, w: '100%', h: 0.04, fill: { color: C.orange } });

  s.addText('次のステップ', { x: 0.5, y: 0.2, w: 12.3, h: 0.6, fontSize: 20, color: 'AABBCC', fontFace: FONT_JA, align: 'center' });

  s.addShape(pptx.ShapeType.rect, { x: 1.5, y: 0.85, w: 10.3, h: 1.2, fill: { color: C.gold } });
  s.addText('高速資産形成セミナー', {
    x: 1.5, y: 0.85, w: 10.3, h: 1.2,
    fontSize: 38, bold: true, color: C.navy, fontFace: FONT_JA, align: 'center'
  });

  s.addText('米国株オプション取引の詳細を無料でお伝えします', {
    x: 0.5, y: 2.15, w: 12.3, h: 0.6,
    fontSize: 20, color: C.white, fontFace: FONT_JA, align: 'center'
  });

  const details = [
    '✅  軍資金を作りながら「同時に増やす仕組み」の全体像',
    '✅  月5〜15万円を実現しているメンバーの実例紹介',
    '✅  オプション取引の基礎〜実践戦略まで丁寧に解説',
    '✅  個別相談・質疑応答あり（先着順）',
  ];

  details.forEach((d, i) => {
    s.addText(d, {
      x: 1.5, y: 2.95 + i * 0.52, w: 10.3, h: 0.48,
      fontSize: 16, color: C.white, fontFace: FONT_JA
    });
  });

  s.addShape(pptx.ShapeType.rect, { x: 2.5, y: 5.05, w: 8.3, h: 0.75, fill: { color: C.orange } });
  s.addText('▶  まずは無料セミナーにご参加ください', {
    x: 2.5, y: 5.05, w: 8.3, h: 0.75,
    fontSize: 20, bold: true, color: C.white, fontFace: FONT_JA, align: 'center'
  });
})();

// ============================================================
//  スライド 18: まとめ・ロードマップ
// ============================================================
(function() {
  const s = pptx.addSlide();
  addBg(s, C.light);
  titleBox(s, 'まとめ：今日から始める3ステップ');

  const steps = [
    {
      step: 'STEP 1', sub: '今すぐ（〜1ヶ月）', color: C.accent,
      items: ['先取り貯蓄の自動化', '固定費の見直し（スマホ・保険）', 'キャッシュレス1本化']
    },
    {
      step: 'STEP 2', sub: '仕組みを整える（1〜3ヶ月）', color: C.green,
      items: ['ふるさと納税 / iDeCo 申込', 'ポイントサイト登録・自己アフィ活用', '断捨離・不用品売却']
    },
    {
      step: 'STEP 3', sub: '収入を増やす（3ヶ月〜）', color: C.orange,
      items: ['副業スタート（小さく始める）', '軍資金の全額を投資口座に自動移動', '高速資産形成セミナーで「増やす」へ']
    },
  ];

  steps.forEach((st, i) => {
    const x = 0.4 + i * 4.3;
    s.addShape(pptx.ShapeType.rect, { x, y: 1.3, w: 4.0, h: 5.0, fill: { color: C.white }, line: { color: 'DDDDDD', pt: 1 } });
    s.addShape(pptx.ShapeType.rect, { x, y: 1.3, w: 4.0, h: 1.0, fill: { color: st.color } });
    s.addText(st.step, { x, y: 1.3, w: 4.0, h: 0.55, fontSize: 20, bold: true, color: C.white, fontFace: FONT_EN, align: 'center' });
    s.addText(st.sub, { x, y: 1.82, w: 4.0, h: 0.45, fontSize: 11, color: C.white, fontFace: FONT_JA, align: 'center' });

    st.items.forEach((item, j) => {
      const iy = 2.5 + j * 0.85;
      s.addShape(pptx.ShapeType.rect, { x: x + 0.2, y: iy, w: 0.45, h: 0.45, fill: { color: st.color } });
      s.addText('✓', { x: x + 0.2, y: iy, w: 0.45, h: 0.45, fontSize: 14, bold: true, color: C.white, fontFace: FONT_EN, align: 'center' });
      s.addText(item, { x: x + 0.75, y: iy, w: 3.1, h: 0.7, fontSize: 13, color: C.darkgray, fontFace: FONT_JA });
    });
  });

  s.addShape(pptx.ShapeType.rect, { x: 0.4, y: 6.15, w: 12.5, h: 0.55, fill: { color: C.navy } });
  s.addText('軍資金を作りながら、同時に増やす仕組みを手に入れる。それが「高速資産形成」です。', {
    x: 0.4, y: 6.15, w: 12.5, h: 0.55,
    fontSize: 14, bold: true, color: C.gold, fontFace: FONT_JA, align: 'center'
  });
  addFooter(s, '高速資産形成セミナー集客講義');
})();

// ============================================================
//  保存
// ============================================================
const outputPath = path.join('D:/dev/軍資金作成資料/docs', '軍資金作成講義_スライド.pptx');
pptx.writeFile({ fileName: outputPath })
  .then(() => console.log('✅ PowerPoint saved:', outputPath))
  .catch(err => console.error('❌ Error:', err));
