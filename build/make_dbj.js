const pptxgen = require("pptxgenjs");
const { applyTheme } = require("/root/.claude/skills/synced/fed45bdf-403d-427c-9870-2efc96bd5d85_ba498201-fc8b-4d1a-bc58-a10d7624274f/pptx/scripts/apply_theme.js");

const THEME = { name: "DBJ IT Shindan Proposal", headFontFace: "Yu Gothic", bodyFontFace: "Yu Gothic",
  colors: { dk1: "333333", lt1: "FFFFFF", dk2: "1F3A5F", lt2: "F4F6F9", accent1: "1F3A5F", accent2: "5B7399",
    accent3: "D5DCE8", accent4: "333333", accent5: "6B7280", accent6: "D9DDE3", hlink: "1F3A5F", folHlink: "6B7280" } };

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "IT診断サービス コンテンツ施策のご提案";
const C = pres.SchemeColor;
const FOOT = "IT診断サービス コンテンツ施策のご提案";

pres.defineSlideMaster({ title: "COVER", background: { color: "FFFFFF" }, objects: [] });
pres.defineSlideMaster({ title: "CONTENT", background: { color: "FFFFFF" },
  objects: [{ text: { text: FOOT, options: { x: 0.6, y: 7.05, w: 8, h: 0.3, fontSize: 10, color: C.accent5, margin: 0 } } }],
  slideNumber: { x: 12.1, y: 7.05, w: 0.6, h: 0.3, fontSize: 10, color: C.accent5, align: "right" } });

function slide(title, section) {
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: section });
  s.addText(title, { x: 0.6, y: 0.4, w: 12.1, h: 0.8, fontSize: 28, bold: true, color: C.text1, margin: 0, valign: "middle", isTextBox: true, objectName: "Title" });
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
s.addText("CONTENT MARKETING PROPOSAL", { x: 0.9, y: 1.4, w: 11, h: 0.4, fontSize: 14, bold: true, color: C.accent1, charSpacing: 4, margin: 0, isTextBox: true });
s.addText("IT診断サービスの訴求に向けた\nコンテンツ施策のご提案", { x: 0.9, y: 2.0, w: 11.5, h: 1.9, fontSize: 40, bold: true, color: C.text1, margin: 0, valign: "top", isTextBox: true });
s.addText("経営層向け「IT健康診断」チェックリスト（ホワイトペーパー）と、\n流入元となるコラム群の設計案", { x: 0.9, y: 4.2, w: 11.5, h: 1.0, fontSize: 18, color: C.accent5, margin: 0, valign: "top", isTextBox: true });
s.addText("ご提案資料（案）　2026年10月", { x: 0.9, y: 6.3, w: 11, h: 0.4, fontSize: 14, color: C.text1, margin: 0, isTextBox: true });

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

// 3 usage
s = slide("コンテンツの使い方：Webと営業、2つの入口からIT診断へ", "狙いと全体像");
const lanes = [
  ["Web", "検索・メルマガ", [["コラム／メルマガ", "コラムは検索で、メルマガは既存リードに届ける"], ["WPダウンロード", "10分のセルフチェックで現在地を知る"], ["問い合わせ", "結果をもとに問い合わせ、または営業からフォロー"]]],
  ["営業", "商談・営業トーク", [["商談・営業トーク", "課題のヒアリング中に、WPを紹介する"], ["WPを一緒に記入", "営業用の版で、その場でチェックする"], ["結果の読み合わせ", "NOだった項目を30分で一緒に確認"]]],
];
lanes.forEach((ln, li) => {
  const y = 1.6 + li * 2.1;
  s.addShape(pres.ShapeType.rect, { x: 0.6, y, w: 1.7, h: 1.9, fill: { color: C.accent1 }, line: { color: C.accent1 }, objectName: "Lane" });
  s.addText([{ text: ln[0], options: { fontSize: 22, bold: true, breakLine: true } }, { text: ln[1], options: { fontSize: 12 } }], { x: 0.6, y, w: 1.7, h: 1.9, color: C.background1, align: "center", valign: "middle", margin: 0.05, isTextBox: true });
  ln[2].forEach((b, i) => {
    const x = 2.55 + i * 2.7;
    s.addShape(pres.ShapeType.rect, { x, y, w: 2.35, h: 1.9, fill: { color: C.background2 }, line: { color: C.background2 }, objectName: "Step" });
    s.addText(b[0], { x: x + 0.2, y: y + 0.15, w: 1.95, h: 0.6, fontSize: 16, bold: true, color: C.text1, margin: 0, valign: "middle", isTextBox: true });
    s.addText(b[1], { x: x + 0.2, y: y + 0.8, w: 1.95, h: 1.0, fontSize: 13, color: C.text1, margin: 0, valign: "top", isTextBox: true });
    if (i < 2) s.addShape(pres.ShapeType.rightArrow, { x: x + 2.4, y: y + 0.78, w: 0.27, h: 0.4, fill: { color: C.accent5 }, line: { color: C.accent5 }, objectName: "Arrow" });
  });
});
s.addShape(pres.ShapeType.rightArrow, { x: 10.45, y: 3.45, w: 0.3, h: 0.4, fill: { color: C.accent5 }, line: { color: C.accent5 }, objectName: "Arrow" });
s.addShape(pres.ShapeType.rect, { x: 10.8, y: 1.6, w: 1.93, h: 4.0, fill: { color: C.accent1 }, line: { color: C.accent1 }, objectName: "Goal" });
s.addText([{ text: "IT診断\nサービス", options: { fontSize: 20, bold: true, breakLine: true } }, { text: "2〜3カ月の診断へ", options: { fontSize: 13 } }], { x: 10.8, y: 1.6, w: 1.93, h: 4.0, color: C.background1, align: "center", valign: "middle", margin: 0.1, isTextBox: true });
s.addText("同じホワイトペーパーを、Webでは集客に、営業の場では会話の道具に使います。営業用の版（NO項目＝診断の仮説）も用意します。", { x: 0.6, y: 5.95, w: 12.1, h: 0.8, fontSize: 16, bold: true, color: C.text2, margin: 0, valign: "top", isTextBox: true });

// 4 mapping
s = slide("悩みとIT診断の8領域を対応づける", "狙いと全体像");
const hd = (t) => ({ text: t, options: { bold: true, color: "FFFFFF", fill: { color: THEME.colors.dk2 }, fontSize: 14, valign: "middle" } });
const cell = (t, o = {}) => ({ text: t, options: Object.assign({ fontSize: 14, valign: "middle", color: THEME.colors.dk1 }, o) });
s.addTable([
  [hd("経営層の悩み"), hd("対応する診断領域（サービスページ）"), hd("チェックリストの軸")],
  [cell("① IT投資が経営にどう効いているか説明できない"), cell("IT投資、IT戦略"), cell("投資（7問）", { bold: true })],
  [cell("② IT起因の経営リスクを把握できていない"), cell("ITリスク（セキュリティ・IT-BCP・内部統制）、IT調達（ベンダー依存）、ITオペレーション"), cell("リスク（7問）", { bold: true })],
  [cell("③ DXで他社に後れを取っていないか不安"), cell("IT戦略、IT組織/IT人材、アプリケーション、ITアーキテクチャ"), cell("競争力（7問）", { bold: true })],
], { x: 0.6, y: 1.6, w: 12.1, colW: [4.0, 5.9, 2.2], rowH: [0.55, 1.0, 1.3, 1.1], border: { type: "solid", pt: 1, color: "D9DDE3" }, fill: { color: "FFFFFF" } });
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
card(s, 8.6, 2.0, 4.1, 4.2, "仕様（案）", bullets(["ボリューム：12〜16ページ", "所要時間：約10分", "回答形式：YES/NO", "設問数：21問（3軸×7問）", "出力：総合スコアと軸別の内訳", "別版：営業用ヒアリングシート"]), { bs: 15 });

// 6 scoring
s = slide("YES/NOで答え、総合スコアと内訳で見せる", "ホワイトペーパー");
[["21問", "3軸×7問"], ["10分", "回答の目安"], ["YES=1点", "NO=0点"]].forEach((k, i) => {
  const x = 0.6 + i * 4.1;
  s.addShape(pres.ShapeType.rect, { x, y: 1.45, w: 3.9, h: 1.5, fill: { color: C.background2 }, line: { color: C.background2 }, objectName: "Stat" });
  s.addText(k[0], { x, y: 1.5, w: 3.9, h: 0.95, fontSize: 40, bold: true, color: C.accent1, align: "center", valign: "middle", margin: 0, isTextBox: true });
  s.addText(k[1], { x, y: 2.4, w: 3.9, h: 0.45, fontSize: 14, color: C.text1, align: "center", valign: "middle", margin: 0, isTextBox: true });
});
s.addText("総合スコアの判定（21点満点）", { x: 0.6, y: 3.2, w: 6, h: 0.45, fontSize: 18, bold: true, color: C.text2, margin: 0, isTextBox: true });
[["17〜21点", "良好", C.accent3, C.text1], ["11〜16点", "要注意", C.accent2, C.background1], ["0〜10点", "要対応", C.accent1, C.background1]].forEach((j, i) => {
  const x = 0.6 + i * 2.1;
  s.addShape(pres.ShapeType.rect, { x, y: 3.75, w: 1.95, h: 1.2, fill: { color: j[2] }, line: { color: j[2] }, objectName: "Grade" });
  s.addText(j[0], { x, y: 3.8, w: 1.95, h: 0.6, fontSize: 20, bold: true, color: j[3], align: "center", valign: "middle", margin: 0, isTextBox: true });
  s.addText(j[1], { x, y: 4.35, w: 1.95, h: 0.5, fontSize: 16, bold: true, color: j[3], align: "center", valign: "middle", margin: 0, isTextBox: true });
});
s.addShape(pres.ShapeType.rect, { x: 0.6, y: 5.15, w: 6.2, h: 1.1, fill: { color: "FFFFFF" }, line: { color: C.accent6, width: 1 }, objectName: "Example" });
s.addText([{ text: "結果の見え方（例）", options: { bold: true, color: C.accent1, breakLine: true } }, { text: "総合 12点／21点（要注意）。内訳は、投資 3/7、リスク 5/7、競争力 4/7", options: { color: C.text1 } }], { x: 0.8, y: 5.2, w: 5.8, h: 1.0, fontSize: 14, margin: 0, valign: "middle", isTextBox: true });
s.addText("点数の区切りは暫定です。実際の診断案件の傾向をもとに調整します。", { x: 0.6, y: 6.35, w: 6.2, h: 0.4, fontSize: 12, color: C.accent5, margin: 0, valign: "top", isTextBox: true });
card(s, 7.0, 3.2, 5.7, 3.55, "設計のポイント", bullets(["YES/NOで迷わないよう、設問は「文書がある」「直近1年以内に見直した」など事実で判定できる書き方にする", "まず総合スコアで現在地を一言で伝え、軸別の内訳で弱い所を示す", "致命的な項目（IT-BCP未整備など）がNOなら、点数に関係なく「最優先の確認事項」として警告する"]), { bs: 14 });

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
  ["#", "悩み", "仮タイトル", "狙うキーワード（月間検索数）"],
  ["1", "投資", "IT投資の効果を経営に説明できない本当の理由と、投資対効果の考え方", "IT投資 250 ／ IT投資とは 150 ／ IT投資 事例 150"],
  ["2", "投資", "取締役会・経営会議で承認されるIT投資計画の作り方", "IT投資管理 70 ／ IT投資予算 60"],
  ["3", "リスク", "経営者が押さえるべきITリスク5分類", "IT統制 400 ／ IT全般統制 250 ／ ITリスク 100"],
  ["4", "リスク", "IT-BCPとは？「止まったらどうなるか」を経営が答えられるか", "IT-BCP 700 ／ BCP システム 100"],
  ["5", "競争力", "DXで他社に後れを取っているか？経営層向け自己チェックの観点", "DX 遅れ 100 ／ DX推進指標 100 ／ DX 日本 遅れ 40"],
  ["6", "競争力", "IT戦略と経営戦略がつながらない会社に共通する3つの症状", "IT戦略 500 ／ IT戦略 策定 10"],
  ["α", "全体", "「IT診断」とは何か：やること、期間、費用感の考え方", "IT診断 60 ／ IT健康診断 10"],
];
s.addTable(rows.map((r, i) => r.map((t, j) => i === 0
  ? { text: t, options: { bold: true, color: "FFFFFF", fill: { color: THEME.colors.dk2 }, fontSize: 13, valign: "middle", align: j < 2 ? "center" : "left" } }
  : { text: t, options: { fontSize: 13, valign: "middle", color: THEME.colors.dk1, align: j < 2 ? "center" : "left", bold: j === 1 } }
)), { x: 0.6, y: 1.4, w: 12.1, colW: [0.5, 1.0, 5.3, 5.3], rowH: [0.5, 0.58, 0.58, 0.58, 0.58, 0.58, 0.58, 0.58], border: { type: "solid", pt: 1, color: "D9DDE3" }, fill: { color: "FFFFFF" } });
s.addText("検索数はAhrefs（日本、2026年10月時点）。実施可否ではなく、タイトルとキーワードの選定に使います。各コラムの末尾に該当軸の設問を2〜3問抜粋し、WPの中身を想像できるようにします。", { x: 0.6, y: 6.2, w: 12.1, h: 0.7, fontSize: 13, color: C.text2, bold: true, margin: 0, valign: "top", isTextBox: true });

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
