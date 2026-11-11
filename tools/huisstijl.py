import re,sys,zipfile
src,dst=sys.argv[1],sys.argv[2]
NAVY_OLD,NAVY='14284B','002060'; GEEL_OLD,GEEL='F5B800','F2B800'; SUB_OLD='2E4A7D'
def styles(s):
    fonts_m=re.search(r'<fonts count="(\d+)">(.*?)</fonts>',s,re.S)
    fonts=re.findall(r'<font>.*?</font>|<font/>',fonts_m.group(2),re.S)
    fills_m=re.search(r'<fills count="\d+">(.*?)</fills>',s,re.S)
    fills=re.findall(r'<fill>.*?</fill>',fills_m.group(1),re.S)
    navy_fill=[i for i,f in enumerate(fills) if 'fgColor rgb="00'+NAVY_OLD in f][0]
    # kopbanden: marineblauw vlak met witte/gele letters -> geel vlak met blauwe letters
    xfs_m=re.search(r'<cellXfs count="\d+">(.*?)</cellXfs>',s,re.S)
    newfont={}
    def fix(m):
        xf=m.group(0)
        if re.search(r'fillId="%d"'%navy_fill,xf):
            fid=int(re.search(r'fontId="(\d+)"',xf).group(1))
            f=fonts[fid]
            if re.search(r'<color rgb="00(FFFFFF|%s)"/>'%GEEL_OLD,f):
                if fid not in newfont:
                    fonts.append(re.sub(r'<color rgb="00\w{6}"/>','<color rgb="00%s"/>'%NAVY,f)); newfont[fid]=len(fonts)-1
                xf=xf.replace('fontId="%d"'%fid,'fontId="%d"'%newfont[fid])
        return xf
    new_xfs=re.sub(r'<xf [^>]*?(/>|>.*?</xf>)',fix,xfs_m.group(1),flags=re.S)
    s=s.replace(xfs_m.group(1),new_xfs)
    s=s.replace(fonts_m.group(0),'<fonts count="%d">%s</fonts>'%(len(fonts),''.join(fonts)))
    # vlakkleuren
    nf=fills[navy_fill].replace(NAVY_OLD,GEEL)
    s=s.replace(fills[navy_fill],nf)
    s=s.replace('00'+SUB_OLD,'00'+NAVY)
    # gekozen kolom (voorwaardelijke opmaak): was geel met blauwe letters -> blauw met witte letters
    s=s.replace('<dxf><font><b val="1"/><color rgb="0014284B"/></font><fill><patternFill patternType="solid"><fgColor rgb="00F5B800"/><bgColor rgb="00F5B800"/>',
                '<dxf><font><b val="1"/><color rgb="00FFFFFF"/></font><fill><patternFill patternType="solid"><fgColor rgb="00002060"/><bgColor rgb="00002060"/>')
    return s.replace('00'+NAVY_OLD,'00'+NAVY).replace('00'+GEEL_OLD,'00'+GEEL)
zi=zipfile.ZipFile(src); zo=zipfile.ZipFile(dst,'w',zipfile.ZIP_DEFLATED)
for it in zi.infolist():
    d=zi.read(it.filename); n=it.filename
    if n=='xl/styles.xml': d=styles(d.decode('utf8')).encode('utf8')
    elif n.startswith(('xl/charts/','xl/drawings/')) and n.endswith('.xml'):
        d=d.decode('utf8').replace('"'+NAVY_OLD+'"','"'+NAVY+'"').replace('"'+GEEL_OLD+'"','"'+GEEL+'"').encode('utf8')
    elif n.startswith('xl/worksheets/sheet'):
        d=d.replace(b'tabColor rgb="00'+NAVY_OLD.encode(),b'tabColor rgb="00'+NAVY.encode()).replace(b'tabColor rgb="00'+GEEL_OLD.encode(),b'tabColor rgb="00'+GEEL.encode())
    zo.writestr(it,d)
zo.close()
