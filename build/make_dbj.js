const pptxgen = require("pptxgenjs");
const { applyTheme } = require("/root/.claude/skills/synced/fed45bdf-403d-427c-9870-2efc96bd5d85_ba498201-fc8b-4d1a-bc58-a10d7624274f/pptx/scripts/apply_theme.js");

const THEME = { name: "DBJ IT Shindan Proposal", headFontFace: "Yu Gothic", bodyFontFace: "Yu Gothic",
  colors: { dk1: "22313A", lt1: "FFFFFF", dk2: "0B4F5C", lt2: "EEF4F5", accent1: "0B7285", accent2: "E8A33D",
    accent3: "2E8B6A", accent4: "C8473D", accent5: "6B7C85", accent6: "BFD6DA", hlink: "0B7285", folHlink: "6B7C85" } };

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "IT診断サービス コンテンツ施策のご提案";
const C = pres.SchemeColor;
const FOOT = "IT診断サービス コンテンツ施策のご提案";

pres.defineSlideMaster({ title: "COVER", background: { color: THEME.colors.dk2 }, objects: [] });
pres.defineSlideMaster({ title: "CONTENT", background: { color: "FFFFFF" },
  objects: [{ text: { text: FOOT, options: { x: 0.6, y: 7.05, w: 8, h: 0.3, fontSize: 10, color: C.accent5, margin: 0 } } }],
  slideNumber: { x: 12.1, y: 7.05, w: 0.6, h: 0.3, fontSize: 10, color: C.accent5, align: "right" } });

function slide(title, section) {
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: section });
  s.addText(title, { x: 0.6, y: 0.4, w: 12.1, h: 0.8, fontSize: 28, bold: true, color: C.text2, margin: 0, valign: "middle", isTextBox: true, objectName: "Title" });
  return s;
}
function bullets(arr) { return arr.map((t, i) => ({ text: t, options: { bullet: true, breakLine: i < arr.length - 1 } })); }
function circle(s, x, y, d, label, color, fs) {
  s.addShape(pres.ShapeType.ellipse, { x, y, w: d, h: d, fill: { color }, line: { color }, objectName: "Badge" });
  s.addText(label, { x, y, w: d, h: d, fontSize: fs || 16, bold: true, color: C.background1, align: "center", valign: "middle", margin: 0, isTextBox: true });
}
function card(s, x, y, w, h, head, body, opt = {}) {
  s.addShape(pres.ShapeType.rect, { x, y, w, h, fill: { color: C.background2 }, line: { color: C.background2 }, objectName: "Card" });
  s.addText(head, { x: x + 0.25, y: y + 0.15, w: w - 0.5, h: 0.5, fontSize: opt.hs || 16, bold: true, color: opt.hc || C.accent1, margin: 0, valign: "middle", isTextBox: true });
  s.addText(body, { x: x + 0.25, y: y + 0.75, w: w - 0.5, h: h - 0.9, fontSize: opt.bs || 14, color: C.text1, margin: 0, valign: "top", isTextBox: true, paraSpaceAfter: 6 });
}

// 1 cover
pres.addSection({ title: "導入" });
let s = pres.addSlide({ masterName: "COVER", sectionTitle: "導入" });
s.addText("CONTENT MARKETING PROPOSAL", { x: 0.9, y: 1.4, w: 11, h: 0.4, fontSize: 14, bold: true, color: C.accent6, charSpacing: 4, margin: 0, isTextBox: true });
s.addText("IT診断サービスの訴求に向けた\nコンテンツ施策のご提案", { x: 0.9, y: 2.0, w: 11.5, h: 1.9, fontSize: 40, bold: true, color: C.background1, margin: 0, valign: "top", isTextBox: true });
s.addText("経営層向け「IT健康診断」チェックリスト（ホワイトペーパー）と、\n流入元となるコラム群の設計案", { x: 0.9, y: 4.2, w: 11.5, h: 1.0, fontSize: 18, color: C.accent6, margin: 0, valign: "top", isTextBox: true });
s.addText("ご提案資料（案）　2026年10月", { x: 0.9, y: 6.3, w: 11, h: 0.4, fontSize: 14, color: C.background1, margin: 0, isTextBox: true });

// 2 persona
pres.addSection({ title: "狙いと全体像" });
s = slide("経営層の3つの悩みを、IT診断への入口にする", "狙いと全体像");
s.addText("想定ペルソナ：経営層（関心の軸は、経営リスク・投資対効果・成長戦略）", { x: 0.6, y: 1.3, w: 12.1, h: 0.45, fontSize: 16, color: C.text1, margin: 0, isTextBox: true });
[["1", "投資", "「IT投資が経営にどう効いているか、説明できない」", "投資対効果を、経営の言葉で語れていない状態"],
 ["2", "リスク", "「IT起因の経営リスク（セキュリティ・BCP等）を把握できていない」", "止まったとき・漏れたときの影響が見えていない状態"],
 ["3", "競争力", "「DXで他社に後れを取っていないか不安」", "経営戦略とIT施策がつながっているか確信が持てない状態"]].forEach((c, i) => {
  const x = 0.6 + i * 4.1;
  s.addShape(pres.ShapeType.rect, { x, y: 2.0, w: 3.9, h: 4.6, fill: { color: C.background2 }, line: { color: C.background2 }, objectName: "Card" });
  circle(s, x + 0.3, 2.25, 0.7, c[0], C.accent1, 22);
  s.addText(c[1], { x: x + 1.15, y: 2.25, w: 2.5, h: 0.7, fontSize: 22, bold: true, color: C.text2, margin: 0, valign: "middle", isTextBox: true });
  s.addText(c[2], { x: x + 0.3, y: 3.2, w: 3.3, h: 1.9, fontSize: 16, bold: true, color: C.text1, margin: 0, valign: "top", isTextBox: true });
  s.addText(c[3], { x: x + 0.3, y: 5.2, w: 3.3, h: 1.1, fontSize: 14, color: C.accent5, margin: 0, valign: "top", isTextBox: true });
});
s.addText("※悩みは仮説です。営業・コンサルタントの方へのヒアリングで検証します。", { x: 0.6, y: 6.7, w: 12.1, h: 0.3, fontSize: 10, color: C.accent5, margin: 0, isTextBox: true });

// 3 flow
s = slide("コラムからWP、商談までを一本の導線でつなぐ", "狙いと全体像");
[["コラム", "検索・SNSで経営層に出会う", "指標：PV、WPへのクリック率"],
 ["ホワイトペーパー", "10分のセルフチェックで現在地を知る", "指標：ダウンロード数"],
 ["結果の読み合わせ", "NOだった項目を30分で一緒に確認", "指標：実施率"],
 ["IT診断サービス", "2〜3カ月の診断へ進む", "指標：商談化数"]].forEach((f, i) => {
  const x = 0.6 + i * 3.1;
  s.addShape(pres.ShapeType.rect, { x, y: 1.9, w: 2.8, h: 3.6, fill: { color: i === 3 ? C.accent1 : C.background2 }, line: { color: i === 3 ? C.accent1 : C.background2 }, objectName: "Step" });
  circle(s, x + 0.25, 2.15, 0.6, String(i + 1), i === 3 ? C.background1 : C.accent1, 18);
  if (i === 3) s.addText("4", { x: x + 0.25, y: 2.15, w: 0.6, h: 0.6, fontSize: 18, bold: true, color: C.accent1, align: "center", valign: "middle", margin: 0, isTextBox: true });
  const tc = i === 3 ? C.background1 : C.text1;
  s.addText(f[0], { x: x + 0.25, y: 3.0, w: 2.3, h: 0.8, fontSize: 18, bold: true, color: tc, margin: 0, valign: "top", isTextBox: true });
  s.addText(f[1], { x: x + 0.25, y: 3.85, w: 2.3, h: 0.9, fontSize: 14, color: tc, margin: 0, valign: "top", isTextBox: true });
  s.addText(f[2], { x: x + 0.25, y: 4.8, w: 2.3, h: 0.6, fontSize: 12, color: i === 3 ? C.accent6 : C.accent5, margin: 0, valign: "top", isTextBox: true });
  if (i < 3) s.addShape(pres.ShapeType.rightArrow, { x: x + 2.82, y: 3.5, w: 0.26, h: 0.4, fill: { color: C.accent5 }, line: { color: C.accent5 }, objectName: "Arrow" });
});
s.addText("営業が使える社内用の版（NO項目＝診断の仮説）も同時に用意し、問い合わせ前の会話にも使えるようにします。", { x: 0.6, y: 5.9, w: 12.1, h: 0.7, fontSize: 16, color: C.text2, bold: true, margin: 0, valign: "top", isTextBox: true });

// 4 mapping
s = slide("悩みとIT診断の8領域を対応づける", "狙いと全体像");
const hd = (t) => ({ text: t, options: { bold: true, color: "FFFFFF", fill: { color: THEME.colors.dk2 }, fontSize: 14, valign: "middle" } });
const cell = (t, o = {}) => ({ text: t, options: Object.assign({ fontSize: 14, valign: "middle", color: THEME.colors.dk1 }, o) });
s.addTable([
  [hd("経営層の悩み"), hd("対応する診断領域（サービスページ）"), hd("チェックリストの軸")],
  [cell("① IT投資が経営にどう効いているか説明できない"), cell("IT投資、IT戦略"), cell("投資（7問）", { bold: true })],
  [cell("② IT起因の経営リスクを把握できていない"), cell("ITリスク（セキュリティ・IT-BCP・内部統制）、IT調達（ベンダー依存）、ITオペレーション"), cell("リスク（7問）", { bold: true })],
  [cell("③ DXで他社に後れを取っていないか不安"), cell("IT戦略、IT組織/IT人材、アプリケーション、ITアーキテクチャ"), cell("競争力（7問）", { bold: true })],
], { x: 0.6, y: 1.6, w: 12.1, colW: [4.0, 5.9, 2.2], rowH: [0.55, 1.0, 1.3, 1.1], border: { type: "solid", pt: 1, color: "BFD6DA" }, fill: { color: "FFFFFF" } });
s.addText("結果が3軸で返るため、「どの悩みに、診断のどの領域で応えられるか」がそのまま営業トークになります。", { x: 0.6, y: 5.9, w: 12.1, h: 0.7, fontSize: 16, bold: true, color: C.text2, margin: 0, valign: "top", isTextBox: true });

// 5 white paper
pres.addSection({ title: "ホワイトペーパー" });
s = slide("ホワイトペーパー案：10分で分かるIT健康診断", "ホワイトペーパー");
s.addText("仮タイトル：「経営層のための IT健康診断チェックリスト」", { x: 0.6, y: 1.3, w: 12.1, h: 0.45, fontSize: 18, bold: true, color: C.accent1, margin: 0, isTextBox: true });
[["はじめに：なぜ今、経営がITの健康状態を問うべきか", "1p"], ["使い方：10分、YES/NOで回答", "1p"], ["チェックリスト：3軸×7問＝21問", "6〜8p"],
 ["スコアの見方と、よくある「危険パターン」", "2p"], ["結果別の次の一手：今すぐ着手できる3つの打ち手", "1〜2p"], ["専門家による診断のご案内と問い合わせ導線", "1p"]].forEach((r, i) => {
  const y = 2.0 + i * 0.78;
  circle(s, 0.6, y, 0.55, String(i + 1), C.accent1, 16);
  s.addText(r[0], { x: 1.35, y, w: 5.9, h: 0.55, fontSize: 15, color: C.text1, margin: 0, valign: "middle", isTextBox: true });
  s.addText(r[1], { x: 7.2, y, w: 0.9, h: 0.55, fontSize: 14, bold: true, color: C.accent5, margin: 0, valign: "middle", align: "right", isTextBox: true });
});
card(s, 8.6, 2.0, 4.1, 4.2, "仕様（案）", bullets(["ボリューム：12〜16ページ", "所要時間：約10分", "回答形式：YES/NO", "設問数：21問（3軸×7問）", "出力：軸別スコアと判定", "別版：営業用ヒアリングシート"]), { bs: 15 });

// 6 scoring
s = slide("チェックリストはYES/NOで答え、軸別に点数化する", "ホワイトペーパー");
[["21問", "3軸×7問"], ["10分", "回答の目安"], ["YES=1点", "NO=0点"]].forEach((k, i) => {
  const x = 0.6 + i * 4.1;
  s.addShape(pres.ShapeType.rect, { x, y: 1.45, w: 3.9, h: 1.5, fill: { color: C.background2 }, line: { color: C.background2 }, objectName: "Stat" });
  s.addText(k[0], { x, y: 1.5, w: 3.9, h: 0.95, fontSize: 40, bold: true, color: C.accent1, align: "center", valign: "middle", margin: 0, isTextBox: true });
  s.addText(k[1], { x, y: 2.4, w: 3.9, h: 0.45, fontSize: 14, color: C.text1, align: "center", valign: "middle", margin: 0, isTextBox: true });
});
s.addText("軸ごとの判定（各7点満点）", { x: 0.6, y: 3.25, w: 6, h: 0.45, fontSize: 18, bold: true, color: C.text2, margin: 0, isTextBox: true });
[["5〜7点", "良好", C.accent3], ["3〜4点", "要注意", C.accent2], ["0〜2点", "要対応", C.accent4]].forEach((j, i) => {
  const x = 0.6 + i * 2.1;
  s.addShape(pres.ShapeType.rect, { x, y: 3.8, w: 1.95, h: 1.4, fill: { color: j[2] }, line: { color: j[2] }, objectName: "Grade" });
  s.addText(j[0], { x, y: 3.9, w: 1.95, h: 0.6, fontSize: 20, bold: true, color: C.background1, align: "center", valign: "middle", margin: 0, isTextBox: true });
  s.addText(j[1], { x, y: 4.5, w: 1.95, h: 0.5, fontSize: 16, bold: true, color: C.background1, align: "center", valign: "middle", margin: 0, isTextBox: true });
});
s.addText("点数の区切りは暫定です。実際の診断案件の傾向をもとに調整します。", { x: 0.6, y: 5.4, w: 6.2, h: 0.5, fontSize: 12, color: C.accent5, margin: 0, valign: "top", isTextBox: true });
card(s, 7.0, 3.25, 5.7, 3.4, "設計のポイント", bullets(["YES/NOで迷わないよう、設問は「文書がある」「直近1年以内に見直した」など事実で判定できる書き方にする", "見せるのは合計点より軸別の結果。何が弱いかが分かる", "致命的な項目（IT-BCP未整備など）がNOなら、点数に関係なく「最優先の確認事項」として警告する"]), { bs: 14 });

// 7 sample questions
s = slide("設問のイメージ（全21問のうち9問）", "ホワイトペーパー");
[["投資", ["中長期のIT投資計画があり、直近1年以内に経営会議で承認・見直しされている", "主要なIT投資について、投資前に期待効果を数値で定め、投資後に結果を確認している", "IT予算を「維持・運用」と「新規・変革」に分けて説明できる"]],
 ["リスク", ["重要システムごとに、停止時の事業影響と許容停止時間が文書で整理されている", "IT-BCP（復旧手順）があり、直近2年以内に訓練または見直しを行った", "主要システムの運用が、特定のベンダーや担当者なしでも継続できる"]],
 ["競争力", ["経営課題とIT施策の対応関係を、1枚の資料で説明できる", "IT人材に求める役割・スキルが定義され、育成または採用の計画がある", "業務システムの改善・刷新の優先順位が、経営視点の基準で決まっている"]]].forEach((c, i) => {
  const x = 0.6 + i * 4.1;
  s.addShape(pres.ShapeType.rect, { x, y: 1.45, w: 3.9, h: 5.2, fill: { color: C.background2 }, line: { color: C.background2 }, objectName: "Card" });
  s.addText(c[0], { x: x + 0.25, y: 1.6, w: 3.4, h: 0.5, fontSize: 20, bold: true, color: C.accent1, margin: 0, valign: "middle", isTextBox: true });
  c[1].forEach((q, j) => {
    const y = 2.3 + j * 1.45;
    s.addShape(pres.ShapeType.rect, { x: x + 0.2, y, w: 3.5, h: 1.3, fill: { color: "FFFFFF" }, line: { color: "FFFFFF" }, objectName: "Question" });
    s.addText(q, { x: x + 0.3, y: y + 0.05, w: 3.3, h: 1.2, fontSize: 13, color: C.text1, margin: 0, valign: "middle", isTextBox: true });
  });
});
s.addText("すべて「YES／NO」で回答します。", { x: 0.6, y: 6.7, w: 8, h: 0.3, fontSize: 10, color: C.accent5, margin: 0, isTextBox: true });

// 8 columns
pres.addSection({ title: "流入コラム" });
s = slide("流入コラム案：出口は「チェックリストで現在地を確認」", "流入コラム");
const rows = [
  ["#", "悩み", "仮タイトル", "狙う検索意図"],
  ["1", "投資", "IT投資の効果を経営に説明できない本当の理由と、投資対効果の考え方", "IT投資 効果測定、IT投資 ROI"],
  ["2", "投資", "取締役会・経営会議で承認されるIT投資計画の作り方", "IT投資計画、IT中期計画"],
  ["3", "リスク", "経営者が押さえるべきITリスク5分類", "ITリスク 経営、IT統制"],
  ["4", "リスク", "IT-BCPとは？「止まったらどうなるか」を経営が答えられるか", "IT-BCP とは、システム停止 事業影響"],
  ["5", "競争力", "DXで他社に後れを取っているか？経営層向け自己チェックの観点", "DX 遅れ、DX 自社 評価"],
  ["6", "競争力", "IT戦略と経営戦略がつながらない会社に共通する3つの症状", "IT戦略 策定、IT戦略 経営戦略"],
  ["α", "全体", "「IT診断」とは何か：やること、期間、費用感の考え方", "IT診断 とは、IT診断 コンサル"],
];
s.addTable(rows.map((r, i) => r.map((t, j) => i === 0
  ? { text: t, options: { bold: true, color: "FFFFFF", fill: { color: THEME.colors.dk2 }, fontSize: 13, valign: "middle", align: j < 2 ? "center" : "left" } }
  : { text: t, options: { fontSize: 13, valign: "middle", color: THEME.colors.dk1, align: j < 2 ? "center" : "left", bold: j === 1 } }
)), { x: 0.6, y: 1.4, w: 12.1, colW: [0.6, 1.2, 6.6, 3.7], rowH: [0.5, 0.58, 0.58, 0.58, 0.58, 0.58, 0.58, 0.58], border: { type: "solid", pt: 1, color: "BFD6DA" }, fill: { color: "FFFFFF" } });
s.addText("各コラムの末尾に、該当軸の設問を2〜3問だけ抜粋して掲載し、WPの中身を想像できるようにします。検索意図は想定で、キーワード調査で検証します。", { x: 0.6, y: 6.2, w: 12.1, h: 0.7, fontSize: 13, color: C.text2, bold: true, margin: 0, valign: "top", isTextBox: true });

// 9 roadmap
pres.addSection({ title: "進め方" });
s = slide("進め方：リスク軸のコラムから先行して拡充する", "進め方");
[["準備", "要件整理・設問確定", ["キーワード調査", "設問21問と判定基準の確定（営業・コンサルタントの監修）", "統計・出典の確認方針"]],
 ["先行公開", "コラム#3、#1", ["経営層が反応しやすいリスク・投資から着手", "末尾から「近日公開のWP」へ事前登録"]],
 ["WP公開", "WP制作・公開 / コラム#5、#4", ["WP本編と営業用版の制作", "ダウンロードフォームと結果の読み方メール"]],
 ["拡充・改善", "コラム#2、#6、#α", ["残りのコラムを公開", "指標を見て設問・導線を改善"]]].forEach((p, i) => {
  const x = 0.6 + i * 3.1;
  s.addShape(pres.ShapeType.rect, { x, y: 1.5, w: 2.9, h: 4.6, fill: { color: C.background2 }, line: { color: C.background2 }, objectName: "Phase" });
  circle(s, x + 0.25, 1.75, 0.6, String(i + 1), C.accent1, 18);
  s.addText(p[0], { x: x + 1.0, y: 1.75, w: 1.8, h: 0.6, fontSize: 18, bold: true, color: C.text2, margin: 0, valign: "middle", isTextBox: true });
  s.addText(p[1], { x: x + 0.25, y: 2.6, w: 2.4, h: 0.8, fontSize: 14, bold: true, color: C.accent1, margin: 0, valign: "top", isTextBox: true });
  s.addText(bullets(p[2]), { x: x + 0.25, y: 3.5, w: 2.4, h: 2.8, fontSize: 13, color: C.text1, margin: 0, valign: "top", isTextBox: true, paraSpaceAfter: 6 });
});
s.addText("スケジュールの期間は、ご体制・ご予算の確認後にお出しします。", { x: 0.6, y: 6.65, w: 10, h: 0.3, fontSize: 10, color: C.accent5, margin: 0, isTextBox: true });

// 10 sales + KPI
s = slide("営業での活用と、見るべき指標", "進め方");
card(s, 0.6, 1.5, 5.95, 4.2, "営業での使い方", bullets(["営業用の版：回答済みの結果を見ながら、NOの項目を診断の仮説として話せる", "結果の読み合わせ（30分）を、問い合わせ前の入口にする", "展示会・セミナー・既存顧客への案内でも、同じチェックリストを配布できる"]), { bs: 16 });
s.addText("指標の見方（ファネル）", { x: 6.85, y: 1.5, w: 5.8, h: 0.5, fontSize: 18, bold: true, color: C.text2, margin: 0, valign: "middle", isTextBox: true });
[["コラム：PV・WPへのクリック率", 5.85], ["WP：ダウンロード数", 5.0], ["読み合わせ：実施率", 4.2], ["IT診断：商談化数", 3.4]].forEach((f, i) => {
  const y = 2.2 + i * 1.05;
  const w = f[1];
  const x = 6.85 + (5.85 - w) / 2;
  s.addShape(pres.ShapeType.rect, { x, y, w, h: 0.85, fill: { color: i === 3 ? C.accent1 : C.accent6 }, line: { color: i === 3 ? C.accent1 : C.accent6 }, objectName: "Funnel" });
  s.addText(f[0], { x, y, w, h: 0.85, fontSize: 14, bold: true, color: i === 3 ? C.background1 : C.text1, align: "center", valign: "middle", margin: 0, isTextBox: true });
});
s.addText("目標値は、現状の数値を伺ってから設定します。", { x: 6.85, y: 6.45, w: 5.85, h: 0.3, fontSize: 10, color: C.accent5, margin: 0, isTextBox: true });

// 11 next
s = slide("ご確認事項と次のステップ", "進め方");
card(s, 0.6, 1.5, 5.95, 4.3, "ご確認したいこと", bullets(["対象企業の規模（中堅か大企業か）", "設問・判定基準を監修いただける方", "ダウンロードフォームの取得項目と、営業への引き継ぎ方法", "コラムの掲載先と、現在のアクセス状況", "公的データを引用する際の出典の扱い"]), { bs: 16 });
card(s, 6.75, 1.5, 5.95, 4.3, "次のステップ", bullets(["本資料へのご意見・優先順位のご確認", "キーワード調査の実施（コラムの優先順位づけ）", "チェックリスト全21問と判定ルールの草案作成", "最初のコラム1本の構成案と本文ドラフト", "体制・スケジュール・お見積りのご提示"]), { bs: 16, hc: C.text2 });

(async () => {
  const out = "../dbj_it_shindan_proposal.pptx";
  await pres.writeFile({ fileName: out });
  await applyTheme(out, THEME);
  console.log("done");
})();
