// 参考資料：学校評価の見直しについて（日比野中との比較・新旧対照）
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, AlignmentType, Table, TableRow, TableCell,
  WidthType, BorderStyle, ShadingType, VerticalAlign, Footer, PageNumber,
} = require("docx");

const FONT = { ascii: "ＭＳ 明朝", hAnsi: "ＭＳ 明朝", eastAsia: "ＭＳ 明朝", cs: "ＭＳ 明朝" };
const GOTHIC = { ascii: "ＭＳ ゴシック", hAnsi: "ＭＳ ゴシック", eastAsia: "ＭＳ ゴシック", cs: "ＭＳ ゴシック" };
const PAGE_W = 11906, MARGIN = 851, TEXT_W = PAGE_W - MARGIN * 2;
const T = (text, o = {}) => new TextRun({ text, font: o.gothic ? GOTHIC : FONT, size: o.size || 21, bold: o.bold });
const P = (runs, o = {}) => new Paragraph({
  alignment: o.align || AlignmentType.LEFT, keepNext: o.keepNext, indent: o.indent,
  spacing: { before: o.before || 0, after: o.after ?? 20, line: o.line || 276 },
  children: (Array.isArray(runs) ? runs : [runs]).map(r => typeof r === "string" ? T(r, o) : r),
});
const H = (text) => P(T(text, { bold: true, gothic: true, size: 22 }), { before: 160, after: 60, keepNext: true });
const border = { style: BorderStyle.SINGLE, size: 4, color: "000000" };
const borders = { top: border, bottom: border, left: border, right: border };
const cp = (text, o = {}) => P(T(text, { size: o.size || 18, bold: o.bold, gothic: o.gothic }), { align: o.align, after: 0, line: 250 });
const cell = (text, w, o = {}) => new TableCell({
  width: { size: w, type: WidthType.DXA }, borders, verticalAlign: VerticalAlign.CENTER,
  shading: o.fill ? { type: ShadingType.CLEAR, color: "auto", fill: o.fill } : undefined,
  margins: { top: 30, bottom: 30, left: 70, right: 70 },
  children: String(text).split("\n").map(t => cp(t, o)),
});
const table = (widths, header, body) => new Table({
  width: { size: TEXT_W, type: WidthType.DXA }, columnWidths: widths,
  rows: [
    new TableRow({ tableHeader: true, cantSplit: true, children: header.map((h, i) => cell(h, widths[i], { fill: "D9D9D9", bold: true, gothic: true, align: AlignmentType.CENTER })) }),
    ...body.map(r => new TableRow({ cantSplit: true, children: r.map((c, i) => cell(c, widths[i], { align: i === 0 && widths[0] < 900 ? AlignmentType.CENTER : AlignmentType.LEFT })) })),
  ],
});

const children = [
  P("参考資料", { align: AlignmentType.RIGHT, after: 120 }),
  P(T("学校評価の見直しについて（日比野中学校との比較・新旧対照）", { bold: true, size: 28, gothic: true }), { align: AlignmentType.CENTER, after: 160 }),

  H("１　日比野中学校の資料から取り入れた点"),
  table([3000, 3000, 4204], ["日比野中学校の工夫", "本校（令和７年度）", "令和８年度の対応"], [
    ["保護者と生徒に、同じ番号で同じ趣旨の質問をしている（Q1～Q13）。", "三者に同じ趣旨で聞く質問は一部のみ（生徒８問・保護者９問）。", "三者で同じ番号・同じ趣旨の質問を増やした（保護者14問・生徒15問）。受け止めの差が分かる。"],
    ["努力点の目標が適切か、生徒に意識させているかを職員に聞いている（[8][9]）。努力点に沿った授業づくりの質問がある（[23]～[25]）。", "努力点を直接評価する質問がない（第１回ＮＣＳ会議で指摘）。", "★の質問を９問設けた（うち新設７問）。努力点が目標として適切かを職員に聞く質問（質問1）を入れ、推進計画の柱（心理的安全性・対話の設定・協働的な探究・自己調整学習・主体性の発揮・組織的な授業改善）と汐中トークに対応させた。"],
    ["回答の選択肢を冒頭に明記している。", "様式に選択肢の記載がない。", "三者共通の４段階を提案文に明記した。"],
    ["職員の自由記述欄を部ごとに設け、関連する番号を文頭に入れる。", "自由記述欄がない。", "同じ形式を採用した（生徒にも自由記述を設けた）。"],
    ["実施期間を日付で示し、Formsで回答する。", "「12月」「１月末」のみの記載。", "日付を明記し、三者ともFormsで実施する。"],
    ["保護者への依頼文（記書き・二次元コード）を用意している。", "依頼文は提案に含まれていない。", "別紙２として依頼文（案）を作成した。"],
    ["結果を保護者に伝えると依頼文に明記している。", "次年度の学校説明会で説明する。", "３月に学校だより・ホームページで公表することを加えた。"],
    ["いじめ対応、情報発信、キャリア、ＩＣＴ、不登校、ＳＣ連携などを幅広く聞いている。", "一部のみ。", "本校の分掌に合わせて加えた（質問12・14・18・19・22・23）。"],
    ["道徳・人権週間の質問がある（[28][29]）。", "質問がない。", "日比野中の文のまま、教職員の質問16・17として加えた。"],
    ["健康や安全（感染症予防など）を保護者に聞いている（Q4）。", "事故・災害時の対応のみ聞いている。", "保護者の質問26として加えた（「検温」は除いた）。"],
  ]),

  H("２　取り入れなかった点"),
  table([3000, 7204], ["日比野中学校の内容", "理由"], [
    ["研究開発学校としての質問（５段階、Q14～19・[79]～[85]）", "本校は研究開発学校ではないため。「なかまと学び合う」「自分で考えて学ぶ」という趣旨は★の質問に含めた。"],
    ["職員の質問78問", "回答の負担が大きいため。本校は34問とし、重なる質問は統合した。"],
    ["補助教材などの細かな質問", "質問数を抑えるため見送った。必要があれば各部で追加を検討する。"],
  ]),

  H("３　前年度の文言から直したところ"),
  ...[
    "・保護者向けの質問は「生徒」を「お子様」にした（継続の質問を含む。R7-1・4・6・11・17）。",
    "・「学校生活が送ることができるように」→「学校生活を送ることができるように」（R7-1　教職員）",
    "・「有効活用している」→「適切に運用している」（R7-8）",
    "・「正しい職業観」→「勤労観・職業観」（R7-9）",
    "・「本人が自立できるように、服装・頭髪などの校則をきちんと守っている」→「学校のきまりの意味を理解し、守って生活している」（R7-11　生徒）",
    "・「文化祭」→「音楽発表会」（R7-17　保護者・生徒）",
    "・R7-19（心身の健康に関する指導）とR7-20（事故や災害時の対応）を１問にまとめた（教職員）。",
    "・「適切な就学指導」→「適切な就学相談」、「障害に応じた」→「障害の状態に応じた」（R7-21　現在の用語に合わせた）",
    "・「校舎の設備管理活用」→「校舎の設備管理・活用」（R7-23）",
    "・様式上部のメモ「努力点を評価できる質問を追加する！（ＮＳＣ会議より）」は削除し、提案文の見直し点に書いた。",
  ].map(t => P(t, { indent: { left: 420, hanging: 210 } })),

  H("４　新旧対照表（令和７年度の28項目の行き先）"),
  table([700, 4904, 1200, 900, 2500], ["R7", "令和７年度の項目（教職員）", "R7対象", "R8", "扱い"], [
    ["1", "生き生きと楽しく学校生活を送れる教育環境の整備", "教・保・生", "9", "継続（教職員は誤字修正、保護者は「お子様」）"],
    ["2", "問題・悩みを気軽に相談できる雰囲気", "教・保・生", "10", "継続（同文）・★に位置づけ"],
    ["3", "各種会議が情報交換と課題検討の場として機能", "教", "11", "継続（同文）"],
    ["4", "なかまとの対話を単元内に位置づけ、指導方法を改善", "教・保・生", "3", "継続・★（教職員は推進計画に合わせ修正、保護者は「お子様」）"],
    ["5", "評価の観点や規準を定め、学習の振り返りを援助", "教・保・生", "－", "削除"],
    ["6", "総合的な学習の時間が「生きる力」を育む", "教・保・生", "13", "継続（保護者は「（福祉・進路等）」を削除、「お子様」）"],
    ["7", "福祉・ボランティアに関する学習の工夫", "教", "－", "削除（13に含める）"],
    ["8", "図書館の有効活用", "教", "15", "継続（「適切に運用している」に修正）"],
    ["9", "生き方を考え、職業観を身に付ける系統的な指導", "教", "18", "継続（文言修正）・保護者と生徒を新設"],
    ["10", "適性に応じたきめ細かい進路指導", "教", "－", "削除"],
    ["11", "職員の共通理解のもとでの校則指導", "教・保・生", "20", "継続（保護者は「お子様」、生徒は文言修正）"],
    ["12", "問題行動に組織的に対応できる体制", "教", "22", "14と統合"],
    ["13", "きまりや校則が適切か話し合う機会", "教", "21", "継続（「生徒とともに」を加筆）・生徒を新設"],
    ["14", "関係機関との緊密な連携", "教", "22", "12と統合"],
    ["15", "避難訓練の工夫・改善", "教", "24", "継続（同文）"],
    ["16", "生徒会活動の支援", "教・生", "－", "削除（6に含める）"],
    ["17", "学校行事の工夫・改善", "教・保・生", "25", "継続（保護者・生徒は「音楽発表会」、保護者は「お子様」）"],
    ["18", "部活動に参加できる環境", "教", "－", "削除"],
    ["19", "心身の健康に関する指導の連携", "教", "26", "20と統合"],
    ["20", "事故や災害時の適切な対応", "教・保", "26", "19と統合（保護者は感染症予防の質問に差し替え）"],
    ["21", "適切な就学指導と教育課程の編成", "教", "27", "継続（用語修正）"],
    ["22", "特別支援学級と通常学級の交流", "教", "28", "継続（同文）"],
    ["23～27", "設備管理、安全点検、個人情報、ゴミ減量、予算", "教", "29～33", "継続（23のみ「・」を補う）"],
    ["28", "ＰＴＡと学校の連携", "教・保", "34", "継続（同文）"],
  ]),
  P("新設した質問：1・2・4～8（★努力点）、12（情報発信）、14（ＩＣＴ）、16（道徳）、17（人権）、19（いじめ）、23（不登校・校内居場所）", { before: 80 }),
];

const doc = new Document({
  styles: { default: { document: { run: { font: FONT, size: 21 } } } },
  sections: [{
    properties: { page: { size: { width: PAGE_W, height: 16838 }, margin: { top: 1021, bottom: 1021, left: MARGIN, right: MARGIN, footer: 400 } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: ["－ ", PageNumber.CURRENT, " －"], font: FONT, size: 18 })] })] }) },
    children,
  }],
});
const out = path.join(__dirname, "..", "参考_学校評価の見直し（日比野中との比較・新旧対照）.docx");
Packer.toBuffer(doc).then(b => { fs.writeFileSync(out, b); console.log("Wrote", out); });
