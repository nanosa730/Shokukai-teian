import sys,re
sys.path.insert(0,'/home/user/Shokukai-teian/tools')
from wordtool import Doc
d=Doc('a000.docx')
ps=d.paras()
tpl=ps[22].group(0)   # 210 の段落をひな形にする
assert '校務分掌の業務改善' in tpl
tab='<w:r w:rsidRPr="003934A9"><w:rPr><w:rFonts w:ascii="UD デジタル 教科書体 NP" w:eastAsia="UD デジタル 教科書体 NP"/><w:szCs w:val="21"/></w:rPr><w:tab/></w:r>'
ITEMS=[('・教職員評価（記入例）について',int(sys.argv[1])),('・職場風土改革月間の取り組みについて',int(sys.argv[2])),
       ('・長期休業中の在宅勤務について',int(sys.argv[3])),('・防災体制（職員安否）について',int(sys.argv[4]))]
new=''
for text,ntab in ITEMS:
    p=re.sub(r'w14:paraId="[^"]*" w14:textId="[^"]*" ','',tpl)
    p=p.replace('<w:t>2</w:t>','<w:t></w:t>',1).replace('<w:t>1</w:t>','<w:t></w:t>',1)
    p=p.replace('<w:t>0　校務分掌の業務改善について</w:t>',f'<w:t xml:space="preserve">　　{text}</w:t>')
    # タブの数を調整（ひな形は単独タブ3つ＋タブ付き【教頭】）
    p=p.replace(tab*3,tab*ntab)
    new+=p
m=d.paras()[23]
assert '0　その他' in m.group(0)
d.x=d.x[:m.end()]+new+d.x[m.end():]
d.save('out/000_R8_第9回運営委員会・職員会議レジュメ.docx',pdf=True)
