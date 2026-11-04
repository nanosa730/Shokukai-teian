"""Word(.docx)の文字差し替え・UDフォント化・PDF書き出し。

使い方（Pythonから）:
    from wordtool import Doc
    d = Doc('in.docx')
    d.runs(61, {6: '水'})          # 段落61の7番目のランを「水」に
    d.replace('令和７', '令和８')    # 1つのランの中にある文字列を置換（全体で1回だけあることを確認）
    d.ud_font()                     # 全文を UD デジタル 教科書体 N に（行グリッドは外す）
    d.save('out.docx', pdf=True)    # PDFも同じ名前で書き出す
段落番号は dump() で確認する。
"""
import re, zipfile, subprocess, os, shutil, tempfile

UD = 'UD デジタル 教科書体 N'
T = r'(<w:t(?: [^>]*)?>)([^<]*)(</w:t>)'


class Doc:
    def __init__(self, path):
        z = zipfile.ZipFile(path)
        self.names = z.namelist()
        self.data = {n: z.read(n) for n in self.names}
        self.x = self.data['word/document.xml'].decode()

    def paras(self):
        return list(re.finditer(r'<w:p\b[^>]*?(?<!/)>.*?</w:p>', self.x, re.S))

    def dump(self):
        for i, m in enumerate(self.paras()):
            ts = [t[1] for t in re.findall(T, m.group(0))]
            if ''.join(ts).strip():
                print(i, ts)

    def runs(self, idx, changes):
        m = self.paras()[idx]
        k = [-1]

        def f(a):
            k[0] += 1
            return a.group(1) + changes.get(k[0], a.group(2)) + a.group(3)
        p = re.sub(T, f, m.group(0))
        self.x = self.x[:m.start()] + p + self.x[m.end():]

    def replace(self, old, new, count=1):
        n = self.x.count(old)
        assert n == count, f'「{old}」が{n}か所'
        self.x = self.x.replace(old, new)

    def ud_font(self, font=UD):
        def fix(xml):
            def r(m):
                tag = re.sub(r'w:(ascii|hAnsi|eastAsia|cs)="[^"]*"',
                             lambda a: f'w:{a.group(1)}="{font}"', m.group(0))
                tag = re.sub(r'\s*w:(ascii|hAnsi|eastAsia|cs)Theme="[^"]*"', '', tag)
                for k in ('ascii', 'hAnsi', 'eastAsia'):
                    if f'w:{k}=' not in tag:
                        tag = tag.replace('<w:rFonts', f'<w:rFonts w:{k}="{font}"', 1)
                return tag
            return re.sub(r'<w:rFonts\b[^>]*/?>', r, xml)
        self.x = fix(self.x)
        for n in ('word/styles.xml', 'word/numbering.xml'):
            if n in self.data:
                s = fix(self.data[n].decode())
                if n == 'word/styles.xml':
                    s = re.sub(r'<w:rPrDefault>\s*<w:rPr>(?!.*?<w:rFonts)',
                               f'<w:rPrDefault><w:rPr><w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:eastAsia="{font}" w:cs="{font}"/>', s, 1, re.S)
                self.data[n] = s.encode()
        # UDフォントはグリッドに合わせると行間・字間が広がって崩れるので、グリッドを外す
        self.x = re.sub(r'(<w:docGrid\b[^>]*?)\s+w:type="[^"]*"', r'\1', self.x)
        self.x = re.sub(r'(<w:docGrid\b[^>]*?)\s+w:charSpace="[^"]*"', r'\1', self.x)

    def save(self, path, pdf=False):
        self.data['word/document.xml'] = self.x.encode()
        with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as o:
            for n in self.names:
                o.writestr(n, self.data[n])
        if pdf:
            to_pdf(path)


def to_pdf(path):
    out = os.path.dirname(os.path.abspath(path))
    subprocess.run(['soffice', '--headless', '--convert-to', 'pdf', '--outdir', out, path],
                   check=True, capture_output=True)
    return os.path.splitext(path)[0] + '.pdf'


def doc_to_docx(path, outdir):
    subprocess.run(['soffice', '--headless', '--convert-to', 'docx', '--outdir', outdir, path],
                   check=True, capture_output=True)
    return os.path.join(outdir, os.path.splitext(os.path.basename(path))[0] + '.docx')
