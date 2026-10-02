const pptxgen = require("pptxgenjs");
const { applyTheme } = require("/root/.claude/skills/synced/fed45bdf-403d-427c-9870-2efc96bd5d85_ba498201-fc8b-4d1a-bc58-a10d7624274f/pptx/scripts/apply_theme.js");
const THEME = { name:"SBCS Proposal", headFontFace:"Yu Gothic", bodyFontFace:"Yu Gothic",
 colors:{dk1:"333333",lt1:"FFFFFF",dk2:"1F3A5F",lt2:"F4F6F9",accent1:"1F3A5F",accent2:"6B7280",accent3:"D9DDE3",accent4:"333333",accent5:"1F3A5F",accent6:"6B7280",hlink:"1F3A5F",folHlink:"6B7280"}};
const pres = new pptxgen(); pres.layout="LAYOUT_WIDE"; pres.theme={headFontFace:THEME.headFontFace,bodyFontFace:THEME.bodyFontFace};
pres.title="コラム制作・ホワイトペーパーデザイン ご提案・お見積り";
const C = pres.SchemeColor;
const FOOT="コラム制作・ホワイトペーパーデザイン ご提案・お見積り";
pres.defineSlideMaster({title:"COVER",background:{color:"FFFFFF"},objects:[],slideNumber:undefined});
pres.defineSlideMaster({title:"CONTENT",background:{color:"FFFFFF"},
 objects:[{text:{text:FOOT,options:{x:0.6,y:7.05,w:8,h:0.3,fontSize:9,color:"6B7280",margin:0}}}],
 slideNumber:{x:12.1,y:7.05,w:0.6,h:0.3,fontSize:9,color:"6B7280",align:"right"},
 });
function slide(title){
  const s=pres.addSlide({masterName:"CONTENT"});
  s.addText(title,{x:0.6,y:0.4,w:12.1,h:0.7,fontSize:28,bold:true,color:C.text1,margin:0,isTextBox:true,objectName:"Title"});
  return s;
}
function card(s,x,y,w,h,head,body,opt={}){
  s.addShape(pres.ShapeType.rect,{x,y,w,h,fill:{color:C.background2},line:{color:C.background2},objectName:"Card"});
  s.addText(head,{x:x+0.25,y:y+0.15,w:w-0.5,h:0.45,fontSize:opt.hs||16,bold:true,color:C.accent1,margin:0,isTextBox:true});
  s.addText(body,{x:x+0.25,y:y+0.65,w:w-0.5,h:h-0.8,fontSize:opt.bs||13,color:C.text1,margin:0,valign:"top",isTextBox:true,paraSpaceAfter:4});
}
function bullets(arr){return arr.map((t,i)=>({text:t,options:{bullet:true,breakLine:i<arr.length-1}}));}

// 1 cover
let s=pres.addSlide({masterName:"COVER"});
s.addText("PROPOSAL & ESTIMATE",{x:0.9,y:1.5,w:11,h:0.4,fontSize:14,bold:true,color:C.accent1,charSpacing:4,margin:0,isTextBox:true});
s.addText("Google関連ブログ コラム制作\n＋ ホワイトペーパーデザイン\nご提案・お見積り",{x:0.9,y:2.0,w:11.5,h:2.2,fontSize:38,bold:true,color:C.text1,margin:0,valign:"top",isTextBox:true});
s.addText("対象サイト：SB C&S ライセンスカウンター Google ブログ\nhttps://licensecounter.jp/google/blog/",{x:0.9,y:4.5,w:11,h:0.8,fontSize:14,color:C.text1,margin:0,isTextBox:true});
s.addShape(pres.ShapeType.line,{x:0.9,y:5.6,w:11.5,h:0,line:{color:C.accent1,width:1.5}});
s.addText("ご提出先：SB C&S株式会社 御中\nご提出日：2026年　月　日\nご提出元：株式会社イノーバ",{x:0.9,y:5.8,w:11,h:1.0,fontSize:13,color:C.text1,margin:0,isTextBox:true});

// 2 overview
s=slide("1. ご依頼内容とご提案の概要");
const ov=[["コラム制作","10本。ご指定のテーマに沿って制作。既存ブログのトーン＆マナーに合わせて執筆します。"],
["ホワイトペーパー","1本（デザインのみ・20ページ前後）。ご支給の骨子（本文・図版イメージ）を図版を含めてデザイン化します。"],
["納品物","コラム：本文＋アイキャッチ1枚\nホワイトペーパー：完成デザインデータ（形式は要ご相談）"],
["進め方","制作に必要な情報をご共有いただいたうえで着手。コラムは3本・4本・3本の3タームで五月雨に進行します。"]];
ov.forEach((o,i)=>card(s,0.6+(i%2)*6.15,1.5+Math.floor(i/2)*2.7,5.95,2.4,o[0],o[1],{bs:14}));

// 3 tone & manner
s=slide("2. トーン＆マナーとデザイン方針");
card(s,0.6,1.5,5.95,4.9,"既存コラムのトーン＆マナー（サイト確認結果）",bullets(["見出しは「【】」による強調や、「3選」「5ステップ」など数字を使った具体的な表現","読者別（IT管理者向け／パートナー向け／製品解説）に切り分けた構成","アイキャッチはストックフォトを使用した、教育・ビジネス双方に親和性のあるビジュアル"]),{bs:14});
card(s,6.75,1.5,5.95,2.2,"デザイン方針（共通）",bullets(["白背景・濃いグレー文字で、シンプルに","表や図版などで色を足す場合のみネイビーを使用"]),{bs:14});
// swatches
[["FFFFFF","白背景"],["333333","濃いグレー文字"],["1F3A5F","ネイビー（表・図版）"]].forEach((c,i)=>{
  s.addShape(pres.ShapeType.rect,{x:6.75+i*2.05,y:4.1,w:1.85,h:1.3,fill:{color:c[0]},line:{color:"D9DDE3",width:1},objectName:"Swatch"});
  s.addText(c[1],{x:6.75+i*2.05,y:5.5,w:1.85,h:0.5,fontSize:12,color:C.text1,margin:0,isTextBox:true});
});

// 4 required info
s=slide("3. 共有いただきたい情報①　コラム（1本ごと）");
[["テーマ（対策キーワード）","例：Chromebook 初期設定／Gemini 活用 など"],
 ["ペルソナ","IT担当者（エンドユーザー）または パートナー\n※職種・課題・知識レベルがあれば併記"],
 ["読了後の理想の態度変容","例：製品の導入検討を始める／資料請求する／パートナー認定取得を検討する"]]
 .forEach((o,i)=>card(s,0.6+i*4.1,1.5,3.9,3.0,o[0],o[1],{bs:13}));
s.addText("必須項目",{x:0.6,y:1.15,w:3,h:0.3,fontSize:12,bold:true,color:C.accent2,margin:0,isTextBox:true});
s.addText("上記3項目は必須です。スプレッドシート形式のシートをご用意し、1本ごとにご記入いただく想定です。",{x:0.6,y:4.9,w:12.1,h:0.5,fontSize:14,color:C.text1,margin:0,isTextBox:true});

// 5 optional info table
s=slide("3. 共有いただきたい情報②　あると精度が上がる情報");
const H={bold:true,color:"FFFFFF",fill:{color:"1F3A5F"},fontSize:12};
const rows=[["誘導先（CTA）","記事末尾で誘導したいページ・資料・お問い合わせ先のURL"],["訴求ポイント・伝えたい事実","強調したい機能・メリット、数値・事例、最新のアップデート情報"],["参考資料・一次情報","製品資料、既存記事、Google公式情報、ヒアリング可能な担当者"],["競合・参考記事","意識している競合記事、目指したい記事のイメージ"],["文字数・構成の目安","既存コラムに合わせる／ご希望の文字数"],["NG表現・表記ルール","用語の統一ルール、使用禁止表現、Google／SB C&S側のレギュレーション"],["内部リンク・関連記事","文中や末尾で紹介したい自社ページ・過去記事"],["確認フロー・監修者","確認者、確認の想定日数、監修／法務・Google側確認の有無"],["公開予定日","希望公開日、季節イベント等との連動"],["アイキャッチの希望","イメージ・使用可能なストックフォト契約の有無、掲載サイズ"]];
s.addTable([[{text:"項目",options:H},{text:"内容",options:H}],...rows.map(r=>[{text:r[0],options:{bold:true,fontSize:12,color:"333333"}},{text:r[1],options:{fontSize:12,color:"333333"}}])],
 {x:0.6,y:1.4,w:12.1,colW:[3.6,8.5],rowH:0.46,border:{type:"solid",pt:0.75,color:"D9DDE3"},valign:"middle",margin:[0,0.12,0,0.12],objectName:"InfoTable"});

// 6 WP info
s=slide("3. 共有いただきたい情報③　ホワイトペーパー");
[["骨子データ","本文テキスト、図版イメージ（PowerPoint／手書き可）、ページ構成案"],
 ["ブランド素材","ロゴ、レギュレーション（使用可能な色・フォント）、掲載必須の注意書き"],
 ["仕様・納品形式","判型（A4縦／横など）、ページ数、PDF／編集可能データのご希望"]]
 .forEach((o,i)=>card(s,0.6+i*4.1,1.5,3.9,3.2,o[0],o[1],{bs:14}));

// 7 schedule
s=slide("4. ご提案スケジュール");
const X0=2.9, WW=0.58, Y0=1.35;
for(let m=0;m<4;m++){
  s.addShape(pres.ShapeType.rect,{x:X0+m*4*WW,y:Y0,w:4*WW,h:0.35,fill:{color:"1F3A5F"},line:{color:"FFFFFF",width:1}});
  s.addText(`${m+1}ヶ月目`,{x:X0+m*4*WW,y:Y0,w:4*WW,h:0.35,fontSize:11,bold:true,color:"FFFFFF",align:"center",valign:"middle",margin:0,isTextBox:true});
}
for(let w=0;w<16;w++){
  s.addText(`${w%4+1}W`,{x:X0+w*WW,y:Y0+0.37,w:WW,h:0.28,fontSize:9,color:"6B7280",align:"center",valign:"middle",margin:0,isTextBox:true});
}
const steps=[[0,1],[1,1],[2,2],[4,1],[5,2],[7,1]];
const wpSteps=[[7,1],[8,1],[9,2],[11,1],[12,1],[13,1]];
const lanes=[["第1ターム\nコラム3本",0,steps,"1F3A5F"],["第2ターム\nコラム4本",4,steps,"1F3A5F"],["第3ターム\nコラム3本",8,steps,"1F3A5F"],["ホワイトペーパー\nデザイン",0,wpSteps,"6B7280"]];
lanes.forEach((l,i)=>{
  const y=Y0+0.85+i*0.8;
  s.addShape(pres.ShapeType.rect,{x:0.6,y:y,w:X0-0.6+16*WW,h:0.65,fill:{color:i%2?"FFFFFF":"F4F6F9"},line:{color:"F4F6F9",width:0}});
  s.addText(l[0],{x:0.7,y:y,w:2.1,h:0.65,fontSize:12,bold:true,color:"333333",valign:"middle",margin:0,isTextBox:true});
  l[2].forEach((st,k)=>{
    const x=X0+(l[1]+st[0])*WW;
    s.addShape(pres.ShapeType.rect,{x:x+0.02,y:y+0.1,w:st[1]*WW-0.04,h:0.45,fill:{color:l[3]},line:{color:l[3],width:0},objectName:"Bar"});
    s.addText(String(k+1),{x:x+0.02,y:y+0.1,w:st[1]*WW-0.04,h:0.45,fontSize:12,bold:true,color:"FFFFFF",align:"center",valign:"middle",margin:0,isTextBox:true});
  });
});
const ly=Y0+0.85+4*0.8+0.1;
s.addText([{text:"コラム（1ターム＝2ヶ月）　",options:{bold:true,color:"1F3A5F"}},{text:"①アウトライン制作 → ②アウトライン確認 → ③初稿制作 → ④初稿確認 → ⑤修正原稿制作 → ⑥ご確認・校了"}],{x:0.6,y:ly,w:12.1,h:0.4,fontSize:12,color:"333333",margin:0,isTextBox:true});
s.addText([{text:"ホワイトペーパー　",options:{bold:true,color:"6B7280"}},{text:"①表紙デザイン2パターン提示 → ②確認・選定FB → ③中面デザイン（2週間）→ ④全体デザイン確認 → ⑤修正 → ⑥ご確認・校了"}],{x:0.6,y:ly+0.4,w:12.1,h:0.4,fontSize:12,color:"333333",margin:0,isTextBox:true});
s.addText("※ 週数は目安です。ご確認期間、骨子・必要情報のご共有時期により前後します。開始時期は決定後に日付入りで再提出します。",{x:0.6,y:ly+0.85,w:12.1,h:0.3,fontSize:10,color:"6B7280",margin:0,isTextBox:true});

// 8 estimate
s=slide("5. お見積り");
[["A　コラム10本のみ","70,000円 × 10本","700,000円"],["B　コラム10本＋ホワイトペーパーデザイン1本","60,000円 × 10本 ＋ 300,000円","900,000円"]].forEach((p,i)=>{
  const x=0.6+i*6.15;
  s.addShape(pres.ShapeType.rect,{x,y:1.5,w:5.95,h:3.4,fill:{color:C.background2},line:{color:C.background2},objectName:"PlanCard"});
  s.addText("プラン"+p[0].charAt(0),{x:x+0.3,y:1.7,w:3,h:0.3,fontSize:12,bold:true,color:C.accent2,margin:0,isTextBox:true});
  s.addText(p[0].slice(2),{x:x+0.3,y:2.05,w:5.4,h:0.9,fontSize:18,bold:true,color:C.text1,margin:0,valign:"top",isTextBox:true});
  s.addText(p[1],{x:x+0.3,y:3.0,w:5.4,h:0.4,fontSize:14,color:C.text1,margin:0,isTextBox:true});
  s.addText(p[2],{x:x+0.3,y:3.6,w:5.4,h:0.9,fontSize:40,bold:true,color:C.accent1,margin:0,isTextBox:true});
});
s.addText("プランBでは、ホワイトペーパーと同時にご発注いただくことで、コラム単価を70,000円から60,000円に優遇しています（プランAとの差額は200,000円）。",{x:0.6,y:5.2,w:12.1,h:0.6,fontSize:13,color:C.text1,margin:0,isTextBox:true});
s.addText("※ 金額は税別表記です。（消費税の扱いは提出前に要確認）",{x:0.6,y:5.9,w:12.1,h:0.3,fontSize:10,color:"6B7280",margin:0,isTextBox:true});

// 9 scope
s=slide("6. お見積りの範囲と前提条件");
card(s,0.6,1.5,5.95,4.9,"お見積りの範囲",bullets(["コラム：本文＋アイキャッチ1枚を納品（アウトライン・初稿・修正原稿・校了までの工程を含む）","ホワイトペーパー：骨子を元にした、図版を含む20ページ前後のデザイン（表紙2パターン提示、中面デザイン、修正、校了まで）"]),{bs:14});
card(s,6.75,1.5,5.95,4.9,"含まれないもの・前提条件",bullets(["ホワイトペーパーの骨子（本文・図版イメージ）はご支給いただく前提","ストックフォトの購入費用、ホワイトペーパー用の有償素材費は別途ご相談","ブログへの入稿・公開作業は含まれません","必要情報のご共有後に着手。ご確認が遅れた場合はスケジュールが後ろ倒しになります","修正は各ターム・各工程で規定回数を想定。大幅な構成変更・ページ数追加は別途お見積り"]),{bs:13});

pres.writeFile({fileName:"proposal.pptx"}).then(()=>applyTheme("proposal.pptx",THEME)).then(()=>console.log("ok"));
