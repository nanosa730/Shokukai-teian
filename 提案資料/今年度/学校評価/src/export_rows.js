// 別紙１の表の行データ（赤字は «» で表す）をJSONで出力する。正本Wordの表を組み直すときに使う。
const fs = require("fs");
const src = fs.readFileSync(__dirname + "/make_teian.js", "utf8");
const i = src.indexOf("function markDiff"); const j = src.indexOf("\n}", i) + 2;
eval(src.slice(i, j));
const DATA = require("./items.js"), R7 = require("./r7.js");
const LABEL = { 教: "教職員", 保: "保護者", 生: "生　徒" };
const out = []; let no = 0; const cnt = { 教: 0, 保: 0, 生: 0 };
for (const s of DATA) {
  out.push({ type: "sec", text: `【${s.sec}】`, star: s.sec.startsWith("学校教育") });
  for (const it of s.items) {
    no++;
    const notes = [];
    if (it.tag) notes.push(["tag", it.tag]);
    notes.push(["note", it.note]);
    if (it.extra) notes.push(["small", it.extra]);
    else if (it.same && it.same.length && it.same.length < it.rows.length) notes.push(["small", `（${it.same.map(x => LABEL[x].replace("　", "")).join("・")}は前年度と同文）`]);
    else if (it.same && it.same.length === it.rows.length) notes.push(["small", "（前年度と同文）"]);
    it.rows.forEach((r, k) => {
      cnt[r[0]]++;
      const key = `${it.r7}-${r[0]}`;
      out.push({ type: "row", no, star: !!it.tag, first: k === 0, last: k === it.rows.length - 1, n: it.rows.length,
        label: LABEL[r[0]], text: R7[key] ? markDiff(R7[key], r[1]) : `«${r[1]}»`, notes: k === 0 ? notes : [] });
    });
  }
}
DATA.free.forEach((r, k) => out.push({ type: "free", first: k === 0, last: k === DATA.free.length - 1, label: LABEL[r[0]], text: `«${r[1]}»` }));
for (const r of out) if (r.text) r.text = r.text.replace("«適切に運»用", "«適切に運用»");
fs.writeFileSync(process.argv[2], JSON.stringify({ rows: out, cnt, total: no }, null, 1));
console.log(no, cnt);
