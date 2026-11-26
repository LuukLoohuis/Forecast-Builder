import sys, re
sys.path.insert(0, '/tmp/claude-0/-home-user-Forecast-Builder/67551728-ddc3-553b-bb67-64b56d7ff89a/scratchpad/tools')
from lxml import etree
from xlsxlib import Book
N='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
def info(bk):
    st=etree.fromstring(bk.data['xl/styles.xml'])
    fmts={nf.get('numFmtId'):nf.get('formatCode') for nf in st.iter(N+'numFmt')}
    fonts=st.find(N+'fonts').findall(N+'font'); fills=st.find(N+'fills').findall(N+'fill')
    xfs=st.find(N+'cellXfs').findall(N+'xf')
    def d(i):
        x=xfs[int(i)]; f=fonts[int(x.get('fontId',0))]; fl=fills[int(x.get('fillId',0))]
        fg=fl.find('.//'+N+'fgColor'); col=f.find(N+'color')
        return dict(fmt=fmts.get(x.get('numFmtId'),x.get('numFmtId')), bold=f.find(N+'b') is not None, sz=f.find(N+'sz').get('val') if f.find(N+'sz') is not None else None,
                    color=col.get('rgb') if col is not None else None, fill=fg.get('rgb') if fg is not None else None, wrap=(x.find(N+'alignment').get('wrapText') if x.find(N+'alignment') is not None else None))
    return d
if __name__=='__main__':
    bk=Book(sys.argv[1]); d=info(bk)
    for a in sys.argv[2:]:
        sh,ref=a.split('!'); s=bk.sheet(sh); sid=s.style(ref); print(a,sid,d(sid) if sid else None)
