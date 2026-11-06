import openpyxl,copy
from openpyxl.worksheet.properties import PageSetupProperties
exec(open('final.py').read().replace("print('残る重なり',bad)\n",'').split('for bu in B: print')[0])
FULL=[('校長','校　　　長'),('教頭','教　　　頭'),('佐野','佐野　嘉昭'),('渡邉貴','渡邉　貴志'),('鈴木','鈴木　雄詞'),('谷口','谷口　友美'),
('名倉','名倉　理王'),('大島','大島　裕美子'),('髙田','髙田　智大'),('藤井','藤井　優子'),('安井','安井　理沙'),('松田','松田　香織'),
('幸島','幸島　彩子'),('朝日','朝日　啓輔'),('長谷川','長谷川　俊雄'),('仲野','仲野　園恵'),('菅田','菅田　拡'),('山田','山田　佳世子'),
('石飛','石飛　典子'),('西脇','西脇　司'),('石垣','石垣　瑞穂'),('馬場','馬場　真美'),('柿家','柿家　悠一郎'),('松村','松村　奈津子'),
('春日井','春日井　誠'),
('城','城　智哉'),('浅野','浅野　洋太'),('髙橋咲','髙橋　咲月'),('酒井','酒井　美侑'),('樽見','樽見　宗'),('浅田','浅田　美紀'),
('野々村','野々村　一輝'),('渡邉華','渡邉　華子'),('福井','福井　咲'),
('坪田','坪田　郁子'),('佐藤','佐藤　樹'),('大野','大野　まや'),('三輪','三輪　人美'),('稲田','稲田　典子')]
COLS=['教務','校務・事務','特活','特別支援','保健','学習・進路','生徒指導','キャリア']
UD='UD デジタル 教科書体 N'
wb=openpyxl.load_workbook('a121.xlsx')
src=wb['R7一覧 (2)']
ws=wb.copy_worksheet(src); ws.title='R8一覧'
wb.move_sheet(ws, offset=-(len(wb.sheetnames)-1))
wb.active=0
for s in wb.worksheets: s.sheet_view.tabSelected = (s is ws)
ws.unmerge_cells('C44:K47')
n=len(FULL); first=6; last=first+n-1; cnt=last+1; note=cnt+2
# 行のスタイルを人の行（41行目）からコピー
for r in range(first,note+4):
    for c in range(1,24):
        ws.cell(r,c).value=None
for r in range(first,last+1):
    ws.row_dimensions[r].height=src.row_dimensions[41].height
    for c in range(2,12):
        ws.cell(r,c)._style=copy.copy(src.cell(41,c)._style)
for c in range(2,12):
    ws.cell(cnt,c)._style=copy.copy(src.cell(42,c)._style)
ws.row_dimensions[cnt].height=src.row_dimensions[42].height
for r in range(cnt+1,note+4):
    ws.row_dimensions[r].height=src.row_dimensions[44].height if r==cnt+1 else None
    for c in range(2,12): ws.cell(r,c)._style=copy.copy(src.cell(44,c)._style)
# 見出し
ws['D3']='1/28（木）'; ws['H3']='2/1（月）'
ws['D4']='14:40～'; ws['F4']='15:20～'; ws['H4']='14:35～'; ws['J4']='15:15～'
for i,b in enumerate(COLS): ws.cell(5,4+i).value=b
for i,(k,full) in enumerate(FULL):
    r=first+i
    ws.cell(r,2).value=f'{i+1:02d}'; ws.cell(r,3).value=full
    for j,b in enumerate(COLS):
        v=B[b].get(k)
        if v: ws.cell(r,4+j).value=v
for j,col in enumerate('CDEFGHIJK'):
    ws[f'{col}{cnt}']=f'=COUNTA({col}{first}:{col}{last})'
ws[f'C{note}']=src['C44'].value
ws[f'C{note}']._style=copy.copy(src['C44']._style)
ws.merge_cells(f'C{note}:K{note+3}')
# L列より右（R7の作業用メモ）は消す
for r in range(1,60):
    for c in range(12,24): ws.cell(r,c).value=None
# フォントをUDに
for row in ws.iter_rows(min_row=1,max_row=note+3,max_col=11):
    for c in row:
        f=copy.copy(c.font); f.name=UD; c.font=f
ws.print_area=f'A2:L{note+3}'
ws.sheet_properties.pageSetUpPr=PageSetupProperties(fitToPage=True)
ws.page_setup.fitToWidth=1; ws.page_setup.fitToHeight=1; ws.page_setup.paperSize=9; ws.page_setup.orientation='portrait'
wb.save('out121.xlsx')
print(first,last,cnt,note)
