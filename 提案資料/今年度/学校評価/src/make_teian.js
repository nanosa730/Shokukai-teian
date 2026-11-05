// 令和８年度 学校評価の実施について（提案）＋別紙１ 評価項目一覧
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, AlignmentType, Table, TableRow, TableCell,
  WidthType, BorderStyle, ShadingType, VerticalAlign, VerticalMergeType, PageBreak,
  Footer, PageNumber,
} = require("docx");
const DATA = require("./items.js");
const R7 = require("./r7.js");
const RED = "FF0000"; // 昨年度からの変更点は赤字

const FONT = { ascii: "ＭＳ 明朝", hAnsi: "ＭＳ 明朝", eastAsia: "ＭＳ 明朝", cs: "ＭＳ 明朝" };
const GOTHIC = { ascii: "ＭＳ ゴシック", hAnsi: "ＭＳ ゴシック", eastAsia: "ＭＳ ゴシック", cs: "ＭＳ ゴシック" };
const PAGE_W = 11906, MARGIN = 851, TEXT_W = PAGE_W - MARGIN * 2; // 10204

const T = (text, o = {}) => new TextRun({ text, font: o.gothic ? GOTHIC : FONT, size: o.size || 21, bold: o.bold, color: o.red ? RED : undefined });
// «…» で囲んだ部分を赤字にする
const rich = (s, o = {}) => s.split(/(«[^»]*»)/).filter(Boolean).map(seg => seg.startsWith("«") ? T(seg.slice(1, -1), { ...o, red: true }) : T(seg, o));
// 昨年度の文と比べ、増えた・変わった文字を «» で囲む（文字単位のLCS）
function markDiff(oldS, newS) {
  const a = [...oldS.replace(/\s/g, "")], b = [...newS];
  const n = a.length, m = b.length;
  const dp = Array.from({ length: n + 1 }, () => new Array(m + 1).fill(0));
  for (let i = n - 1; i >= 0; i--) for (let j = m - 1; j >= 0; j--)
    dp[i][j] = a[i] === b[j] ? dp[i + 1][j + 1] + 1 : Math.max(dp[i + 1][j], dp[i][j + 1]);
  const keep = new Array(m).fill(false);
  for (let i = 0, j = 0; i < n && j < m;) {
    if (a[i] === b[j]) { keep[j] = true; i++; j++; } else if (dp[i + 1][j] >= dp[i][j + 1]) i++; else j++;
  }
  if (dp[0][0] / Math.max(m, 1) < 0.6) return `«${newS}»`; // 大きく書き換えた文は全体を赤字
  // 赤字にはさまれた２文字以下の一致は赤字に含める（細切れ防止）
  for (let j = 0; j < m;) {
    if (!keep[j]) { j++; continue; }
    let k = j; while (k < m && keep[k]) k++;
    if (k - j <= 2 && j > 0 && k < m) for (let x = j; x < k; x++) keep[x] = false;
    j = k;
  }
  let out = "", red = false;
  b.forEach((ch, j) => { if (!keep[j] !== red) { out += red ? "»" : "«"; red = !red; } out += ch; });
  return red ? out + "»" : out;
}
const P = (runs, o = {}) => new Paragraph({
  alignment: o.align || AlignmentType.LEFT,
  spacing: { before: o.before || 0, after: o.after ?? 20, line: o.line || 276 },
  indent: o.indent,
  keepNext: o.keepNext,
  children: (Array.isArray(runs) ? runs : [runs]).flatMap(r => typeof r === "string" ? rich(r, o) : [r]),
});
const H = (text) => P(T(text, { bold: true, gothic: true, size: 22 }), { before: 110, after: 30, keepNext: true });
// ぶら下げインデント付きの項目（(1) や ・ など）
const L = (head, body, lv = 1) => P([T(head), ...rich(body)], { indent: { left: 210 * (lv + 1), hanging: 210 * (head.length >= 3 ? 1.5 : 1) } });

const border = { style: BorderStyle.SINGLE, size: 4, color: "000000" };
const borders = { top: border, bottom: border, left: border, right: border };
const cell = (children, w, o = {}) => new TableCell({
  width: { size: w, type: WidthType.DXA }, borders,
  verticalAlign: o.valign || VerticalAlign.CENTER,
  verticalMerge: o.vmerge, columnSpan: o.span,
  shading: o.fill ? { type: ShadingType.CLEAR, color: "auto", fill: o.fill } : undefined,
  margins: { top: 30, bottom: 30, left: 70, right: 70 },
  children: Array.isArray(children) ? children : [children],
});
let KEEP = false; // true の間に作る表内段落は次の行と同じページに置く
const cp = (text, o = {}) => P(rich(text, { size: o.size || 19, bold: o.bold, gothic: o.gothic }), { align: o.align, after: 0, line: 260, keepNext: o.keepNext ?? KEEP });

// ---------- 件数 ----------
let total = 0; const cnt = { 教: 0, 保: 0, 生: 0 };
for (const s of DATA) for (const it of s.items) { total++; for (const r of it.rows) cnt[r[0]]++; }

// ---------- 1枚目：提案文 ----------
const page1 = [
  P("令和８年«１１月２６日（木）»", { align: AlignmentType.RIGHT, after: 0 }),
  P("«第９回職員会議»", { align: AlignmentType.RIGHT, after: 120 }),
  P(T("令和８年度　学校評価の実施について（提案）", { bold: true, size: 30, gothic: true }), { align: AlignmentType.CENTER, after: 120 }),

  H("１　ねらい"),
  P("　«本年度の教育活動を、教職員・保護者・生徒の三者の視点から振り返り、成果と課題を明らかにして次年度の改善につなげる。特に、学校教育の努力点「学び合い、高め合い、動き出す汐中生」の達成状況を把握できるようにする。»"),

  H("２　今年度の主な見直し点"),
  L("(1)", "«努力点を評価する質問（★）を新設する。推進計画の柱（心理的安全性・対話の設定・協働的な探究・自己調整学習・主体性の発揮・組織的な授業改善）と「汐中トーク」に対応させる。（第１回ＮＣＳ会議の「努力点を評価できる質問を」という意見を反映）»"),
  L("(2)", "«三者に同じ番号・同じ趣旨で質問し、教職員・保護者・生徒の受け止めの差が分かるようにする。»"),
  L("(3)", "«前年度と比べる質問は文言を変えずに残す（別紙１の「継続」）。昨年度の「授業への意欲・なかまとの対話」「学習評価の振り返り」「学校行事」「相談のしやすさ」の推移を追えるようにする。»"),
  L("(4)", `«生徒・保護者の質問を増やす（生徒８→${cnt.生}問、保護者９→${cnt.保}問）。いじめへの対応、情報発信、キャリア、ＩＣＴ活用などを加える。»`),
  L("(5)", "«教職員の質問に「業務改善」「不登校・校内居場所」「学年チームでの授業改善」を加え、重なる質問は統合する。部ごとに自由記述欄を設ける。»"),
  L("(6)", "«実施時期を早め、年度末部会反省①（１月２８日）の前に結果を共有する。»"),

  H("３　実施方法"),
];

const W3 = [1000, 1900, 3200, 4104];
const r3 = (a, b, c, d, head) => new TableRow({ cantSplit: true, children: [
  cell(cp(a, { align: AlignmentType.CENTER, bold: head, gothic: head }), W3[0], { fill: head ? "E7E6E6" : undefined }),
  cell(cp(b, { align: AlignmentType.CENTER, bold: head, gothic: head }), W3[1], { fill: head ? "E7E6E6" : undefined }),
  cell(cp(c, { align: head ? AlignmentType.CENTER : AlignmentType.LEFT, bold: head, gothic: head }), W3[2], { fill: head ? "E7E6E6" : undefined }),
  cell(cp(d, { align: head ? AlignmentType.CENTER : AlignmentType.LEFT, bold: head, gothic: head }), W3[3], { fill: head ? "E7E6E6" : undefined }),
]});
page1.push(new Table({ width: { size: TEXT_W, type: WidthType.DXA }, columnWidths: W3, rows: [
  r3("対　象", "質問数", "実施期間", "方　法", true),
  r3("保護者", `«${cnt.保}問＋自由記述»`, "«11月30日（月）～12月14日（月）»", "«依頼文（別紙２）を配信メールで送り、Formsで回答。保護者会（12月2日～8日）でも呼びかける。»"),
  r3("生　徒", `«${cnt.生}問＋自由記述»`, "«12月14日（月）～12月18日（金）»", "«学活・短学活の時間に、学習者用端末からFormsで回答する。»"),
  r3("教職員", `«${cnt.教}問＋自由記述»`, "«12月14日（月）～１月８日（金）»", "«チームズの連絡メモからFormsで回答する。»"),
]}));
page1.push(
  L("※", "«いずれも無記名とする。保護者・生徒は、はじめに学年を選んでから回答する（学年別に集計するため）。»", 0),
  L("※", "«教職員の自由記述は部ごとに設け、関連する番号を文頭に入力する。»", 0),

  H("４　回答の選択肢（三者共通）"),
  P("　«１　とても当てはまる　　２　当てはまる　　３　あまり当てはまらない　　４　全く当てはまらない»"),
  P("　«「とても当てはまる」「当てはまる」の合計を肯定的回答とし、その割合で比べる。»"),

  H("５　結果の集計と活用"),
);
const W5 = [3300, 6904];
const r5 = (a, b) => new TableRow({ cantSplit: true, children: [
  cell(cp(a), W5[0]), cell(cp(b), W5[1]),
]});
page1.push(new Table({ width: { size: TEXT_W, type: WidthType.DXA }, columnWidths: W5, rows: [
  r5("«１月12日～22日»", "集計«（肯定的回答の割合、前年度との比較、三者の比較、学年別）»"),
  r5("«１月26日（火）まで»", "集計結果を«チームズで»全職員に共有する。"),
  r5("«１月28日（木）・２月１日（月）»", "年度末部会反省«①・②»　«部ごとに»結果をもとに改善«策»を話し合う。"),
  r5("«２月２日（火）・２月４日（木）»", "«学校評議員会・第２回ＮＣＳ会議で結果を報告する。»"),
  r5("«２月９日（火）»", "年度末全体反省会で、«次年度の重点を確かめる。»"),
  r5("«２月»", "«★の結果を、努力点の「研究のまとめ」の成果と課題に生かす。»"),
  r5("«３月»", "«学校だより・ホームページで保護者・地域に結果を公表する。»令和９年度の学校説明会で«も»説明する。"),
]}));
page1.push(
  H("６　資料"),
  P("　別紙１　令和８年度　学校評価　質問項目一覧"),
  P("　«別紙２　保護者への依頼文（案）»"),
);

// ---------- 別紙１：質問項目一覧 ----------
const W = [560, 760, 7284, 1600]; // 番号 対象 項目 備考
const LABEL = { 教: "教職員", 保: "保護者", 生: "生　徒" };
const head = () => new TableRow({ tableHeader: true, cantSplit: true, children: [
  cell(cp("番号", { align: AlignmentType.CENTER, gothic: true, bold: true }), W[0], { fill: "D9D9D9" }),
  cell(cp("対象", { align: AlignmentType.CENTER, gothic: true, bold: true }), W[1], { fill: "D9D9D9" }),
  cell(cp("質　問　項　目", { align: AlignmentType.CENTER, gothic: true, bold: true }), W[2], { fill: "D9D9D9" }),
  cell(cp("備考", { align: AlignmentType.CENTER, gothic: true, bold: true }), W[3], { fill: "D9D9D9" }),
]});
const rows = [head()];
let no = 0;
for (const s of DATA) {
  KEEP = true;
  rows.push(new TableRow({ cantSplit: true, children: [
    cell(cp(s.sec.startsWith("学校教育") ? `【${s.sec}】` : `【${s.sec}】`, { gothic: true, bold: true }), TEXT_W, { span: 4, fill: s.sec.startsWith("学校教育") ? "FFF2CC" : "EDEDED" }),
  ]}));
  for (const it of s.items) {
    no++;
    const fill = it.tag ? "FFFBEA" : undefined;
    const notes = [];
    KEEP = it.rows.length > 1;
    if (it.tag) notes.push(cp(it.tag, { size: 17, bold: true }));
    notes.push(cp(it.note, { size: 17 }));
    if (it.extra) notes.push(cp(it.extra, { size: 15 }));
    else if (it.same && it.same.length && it.same.length < it.rows.length) notes.push(cp(`（${it.same.map(x => LABEL[x].replace("　", "")).join("・")}は前年度と同文）`, { size: 15 }));
    else if (it.same && it.same.length === it.rows.length) notes.push(cp("（前年度と同文）", { size: 15 }));
    it.rows.forEach((r, i) => {
      const first = i === 0;
      KEEP = i < it.rows.length - 1;
      rows.push(new TableRow({ cantSplit: true, children: [
        cell(first ? cp(String(no), { align: AlignmentType.CENTER }) : cp(""), W[0], { vmerge: it.rows.length > 1 ? (first ? VerticalMergeType.RESTART : VerticalMergeType.CONTINUE) : undefined, fill }),
        cell(cp(LABEL[r[0]], { align: AlignmentType.CENTER, size: 18 }), W[1], { fill }),
        cell(cp(R7[`${it.r7}-${r[0]}`] ? markDiff(R7[`${it.r7}-${r[0]}`], r[1]) : `«${r[1]}»`), W[2], { fill }),
        cell(first ? notes : cp(""), W[3], { vmerge: it.rows.length > 1 ? (first ? VerticalMergeType.RESTART : VerticalMergeType.CONTINUE) : undefined, fill, valign: VerticalAlign.CENTER }),
      ]}));
    });
  }
  if (s === DATA[DATA.length - 1]) break;
}
// 自由記述
KEEP = true;
rows.push(new TableRow({ cantSplit: true, children: [
  cell(cp("【自由記述】", { gothic: true, bold: true }), TEXT_W, { span: 4, fill: "EDEDED" }),
]}));
[["教", "各部の最後に設ける。改善点や意見を、関連する番号を文頭に付けて入力する。（例）【12】〇〇という理由で、〇〇にすると生徒の負担が減ると思う。"],
 ["保", "汐路中学校をよりよくするためのご意見・ご提案がありましたら、ご記入ください。"],
 ["生", "汐路中学校をよりよくするために、先生に伝えたいことがあれば書きましょう。"]].forEach((r, i) => {
  KEEP = i < 2;
  rows.push(new TableRow({ cantSplit: true, children: [
    cell(i === 0 ? cp("－", { align: AlignmentType.CENTER }) : cp(""), W[0], { vmerge: i === 0 ? VerticalMergeType.RESTART : VerticalMergeType.CONTINUE }),
    cell(cp(LABEL[r[0]], { align: AlignmentType.CENTER, size: 18 }), W[1]),
    cell(cp(`«${r[1]}»`), W[2]),
    cell(i === 0 ? cp("新規（形式）", { size: 17 }) : cp(""), W[3], { vmerge: i === 0 ? VerticalMergeType.RESTART : VerticalMergeType.CONTINUE }),
  ]}));
});

const bessi1 = [
  new Paragraph({ children: [new PageBreak()] }),
  P([T("別紙１", { gothic: true, bold: true })], { after: 0 }),
  P(T("令和８年度　学校評価　質問項目一覧", { bold: true, size: 28, gothic: true }), { align: AlignmentType.CENTER, after: 60 }),
  P(`質問数　教職員 ${cnt.教}問　保護者 ${cnt.保}問　生徒 ${cnt.生}問（いずれも自由記述を除く）`, { align: AlignmentType.CENTER, after: 40, size: 19 }),
  P("★…努力点を評価する質問　　継続…前年度と比べる質問　　（R7-○）…令和７年度の番号", { align: AlignmentType.CENTER, after: 0, size: 19 }),
  P("«赤字»…令和７年度から変更・追加した箇所（削除した項目・語句は参考資料を参照）", { align: AlignmentType.CENTER, after: 120, size: 19 }),
  new Table({ width: { size: TEXT_W, type: WidthType.DXA }, columnWidths: W, rows }),
];

const doc = new Document({
  styles: { default: { document: { run: { font: FONT, size: 21 } } } },
  sections: [{
    properties: { page: { size: { width: PAGE_W, height: 16838 }, margin: { top: 1021, bottom: 1021, left: MARGIN, right: MARGIN, footer: 400 } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: ["－ ", PageNumber.CURRENT, " －"], font: FONT, size: 18 })] })] }) },
    children: [...page1, ...bessi1],
  }],
});
// 正本は手で直したWord（../R8_学校評価の実施について（提案）.docx）。上書きしないよう _generated に出力する
const out = path.join(__dirname, "_generated", "R8_学校評価の実施について（提案）_自動生成.docx");
fs.mkdirSync(path.dirname(out), { recursive: true });
Packer.toBuffer(doc).then(b => { fs.writeFileSync(out, b); console.log("Wrote", out, total, cnt); });
