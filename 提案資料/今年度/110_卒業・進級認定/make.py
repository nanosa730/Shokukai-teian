import sys; import os; sys.path.insert(0,os.path.join(os.path.dirname(__file__),'../../../tools'))
from wordtool import Doc
# 110
d=Doc('a110.docx')
d.runs(0,{2:'８'})
rows={1:[17,18,22,13,None,19,21,18,17,16,18,17,196],
      2:[18,18,22,13,None,19,21,18,17,16,18,17,197],
      3:[18,18,22,13,None,19,21,18,17,16,18,5,185]}
starts={1:18,2:32,3:46}
for g,vals in rows.items():
    p=starts[g]
    for v in vals:
        if v is None: p+=1; continue  # 8月は空欄（段落番号が飛ぶ）
        # 1つ目のランに数字、残りは空
        import re
        m=d.paras()[p]; n=len(re.findall(r'<w:t(?: [^>]*)?>',m.group(0)))
        ch={0:str(v)}; ch.update({i:'' for i in range(1,n)})
        d.runs(p,ch); p+=1
d.runs(61,{6:'水'})
d.runs(62,{5:'５'})
d.runs(63,{6:'木'})
d.runs(65,{2:'１',3:'４６'})
d.runs(68,{3:'６',5:'金'})
d.runs(69,{3:'４'})
d.runs(70,{3:'\u3000８'})
d.runs(72,{2:'１',3:'７',4:'９'})
d.runs(73,{2:'１',3:'８',4:'０'})
d.ud_font(); d.save('out/110_R8_卒業・進級認定について.docx',pdf=True)
# 111
d=Doc('a111.docx')
d.replace('令和７年度','令和８年度')
d.replace('３年生　１８８日（１２月末まで　１４８日）','３年生　１８５日（１２月末まで　１４６日）')
d.ud_font(); d.save('out/111_R8_卒業認定資料.docx',pdf=True)
# 112
d=Doc('a112.docx')
d.replace('令和７年度','令和８年度')
d.replace('１年生　１９８日（２月末　１８２日）','１年生　１９６日（２月末　１７９日）')
d.replace('２年生　１９９日（２月末　１８３日）','２年生　１９７日（２月末　１８０日）')
d.ud_font(); d.save('out/112_R8_進級認定資料.docx',pdf=True)
