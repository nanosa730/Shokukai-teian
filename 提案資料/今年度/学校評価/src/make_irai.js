// 保護者への案内文（記書き形式）Word生成スクリプト
// 案件ごとに title / intro / body / items / outputPath を書き換えて使用する。
const {
  Document, Packer, Paragraph, TextRun, AlignmentType, ImageRun
} = require("docx");
const fs = require("fs");

// ==== ここから案件ごとに書き換える ====
const issueDate = "令和8年11月30日";
const title = "学校教育に関する保護者アンケートについて";
const intro = "　いつも本校の教育活動にご理解・ご協力をいただき、誠にありがとうございます。";
const body =
  "　さて、本校では、本年度の教育活動を振り返り、よりよい学校づくりに生かすため、毎年、保護者の皆様にアンケートへのご協力をお願いしております。" +
  "今年度は、学校教育の努力点「学び合い、高め合い、動き出す汐中生」の取組についてもお伺いします。\n" +
  "　つきましては、下記のとおりアンケートを実施いたしますので、ご協力をお願いいたします。";
const items = [
  { label: "回答期間", text: "11月30日（月）～12月14日（月）" },
  { label: "回答方法", text: "Microsoft Formsを用いたオンラインアンケート（無記名）\n下の二次元コード、または配信メールに記載のURLからご回答ください。お子様のタブレット端末からもご回答いただけます。\n【二次元コード・URL貼付欄】" },
  { label: "回答について", text: "お子様お一人につき１回ご回答ください。兄弟姉妹が在籍している場合は、それぞれのお子様についてご回答ください。\nはじめにお子様の学年を選んでから、質問（16問と自由記述）にお答えください。５分程度で回答できます。" },
  { label: "結果のお知らせ", text: "結果は、３月に学校だより・ホームページでお知らせし、今後の教育活動の改善に生かしてまいります。" },
];
const outputPath = require("path").join(__dirname, "..", "別紙２_保護者への依頼文（案）.docx");
// ==== ここまで案件ごとに書き換える ====

// 全角数字（項目番号）
const zenkakuDigits = ["０","１","２","３","４","５","６","７","８","９"];
function toZenkakuNumber(n) {
  return String(n)
    .split("")
    .map((d) => zenkakuDigits[Number(d)])
    .join("");
}

function textParagraphs(text, opts = {}) {
  return text.split("\n").map(
    (line) =>
      new Paragraph({
        alignment: opts.alignment || AlignmentType.LEFT,
        spacing: { after: 100 },
        children: [new TextRun({ text: line, size: 22 })],
      })
  );
}

const children = [];

// 1. 発行日（右揃え）
children.push(
  new Paragraph({
    alignment: AlignmentType.RIGHT,
    spacing: { after: 200 },
    children: [new TextRun({ text: "　" + issueDate + "　", size: 22 })],
  })
);

// 2. 宛名（左揃え）
children.push(
  new Paragraph({
    alignment: AlignmentType.LEFT,
    spacing: { after: 200 },
    children: [new TextRun({ text: "　保護者　様　", size: 22 })],
  })
);

// 2-3. 差出人（右揃え）
children.push(
  new Paragraph({
    alignment: AlignmentType.RIGHT,
    spacing: { after: 0 },
    children: [new TextRun({ text: "　名古屋市立汐路中学校長　", size: 22 })],
  })
);
children.push(
  new Paragraph({
    alignment: AlignmentType.RIGHT,
    spacing: { after: 300 },
    children: [new TextRun({ text: "　鈴　木　暁　子　", size: 22 })],
  })
);

// 4. タイトル（中央揃え）
children.push(
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 300 },
    children: [new TextRun({ text: title, bold: true, size: 28 })],
  })
);

// 5. 挨拶
children.push(...textParagraphs(intro));

// 6. 本文
children.push(...textParagraphs(body));
children.push(new Paragraph({ text: "", spacing: { after: 100 } }));

// 7. 「記」（中央揃え）
children.push(
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 300 },
    children: [new TextRun({ text: "記", bold: true, size: 24 })],
  })
);

// 8. 項目（全角数字＋半角スペース、「、」は付けない）
items.forEach((item, idx) => {
  const num = toZenkakuNumber(idx + 1);
  children.push(
    new Paragraph({
      alignment: AlignmentType.LEFT,
      spacing: { after: 50 },
      children: [
        new TextRun({ text: `${num}　${item.label}`, bold: true, size: 22 }),
      ],
    })
  );
  item.text.split("\n").forEach((line) => {
    children.push(
      new Paragraph({
        alignment: AlignmentType.LEFT,
        indent: { left: 400 },
        spacing: { after: 150 },
        children: [new TextRun({ text: line, size: 22 })],
      })
    );
  });
  // QRコードなどの画像があれば項目の直後に挿入
  if (item.imagePath && fs.existsSync(item.imagePath)) {
    children.push(
      new Paragraph({
        alignment: AlignmentType.LEFT,
        indent: { left: 400 },
        spacing: { after: 200 },
        children: [
          new ImageRun({
            type: "png",
            data: fs.readFileSync(item.imagePath),
            transformation: {
              width: item.imageWidth || 100,
              height: item.imageHeight || 100,
            },
          }),
        ],
      })
    );
  }
});

const doc = new Document({
  sections: [
    {
      properties: {
        page: { size: { width: 11906, height: 16838 } }, // A4
      },
      children,
    },
  ],
});

Packer.toBuffer(doc).then((buf) => {
  fs.mkdirSync(require("path").dirname(outputPath), { recursive: true });
  fs.writeFileSync(outputPath, buf);
  console.log("Wrote:", outputPath);
});
