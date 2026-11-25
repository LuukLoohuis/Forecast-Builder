"""Minimal OLE/CFB reader + MS-OVBA decompression: dumps VBA module source. usage: readvba.py vbaProject.bin outdir"""
import sys, struct, os
def cfb(data):
    ss=1<<struct.unpack_from('<H',data,30)[0]; mss=1<<struct.unpack_from('<H',data,32)[0]
    ndif=struct.unpack_from('<I',data,0x2C)[0]; dirstart=struct.unpack_from('<I',data,0x30)[0]
    cutoff=struct.unpack_from('<I',data,0x38)[0]; minifat_start=struct.unpack_from('<I',data,0x3C)[0]
    difat=list(struct.unpack_from('<109I',data,0x4C)); dstart=struct.unpack_from('<I',data,0x44)[0]
    sec=lambda i: data[(i+1)*ss:(i+2)*ss]
    while dstart not in (0xFFFFFFFE,0xFFFFFFFF):
        s=sec(dstart); v=struct.unpack('<%dI'%(ss//4),s); difat+=v[:-1]; dstart=v[-1]
    fat=[]
    for d in difat:
        if d>=0xFFFFFFFA: continue
        fat+=struct.unpack('<%dI'%(ss//4),sec(d))
    def chain(st):
        out=[]
        while st<0xFFFFFFFA: out.append(st); st=fat[st]
        return out
    rd=lambda st: b''.join(sec(i) for i in chain(st))
    dirs=rd(dirstart); ents=[]
    for i in range(len(dirs)//128):
        e=dirs[i*128:(i+1)*128]; nl=struct.unpack_from('<H',e,64)[0]
        name=e[:max(0,nl-2)].decode('utf-16le'); typ=e[66]
        l,r,c=struct.unpack_from('<III',e,68); st=struct.unpack_from('<I',e,116)[0]; sz=struct.unpack_from('<Q',e,120)[0]
        ents.append(dict(name=name,typ=typ,l=l,r=r,c=c,st=st,sz=sz))
    root=ents[0]; ministream=rd(root['st'])
    minifat=[]
    for s in chain(minifat_start): minifat+=struct.unpack('<%dI'%(ss//4),sec(s))
    def read(e):
        if e['sz']<cutoff:
            out=b''; st=e['st']
            while st<0xFFFFFFFA: out+=ministream[st*mss:(st+1)*mss]; st=minifat[st]
            return out[:e['sz']]
        return rd(e['st'])[:e['sz']]
    # build paths
    paths={}
    def walk(i,prefix):
        if i==0xFFFFFFFF: return
        e=ents[i]; walk(e['l'],prefix); walk(e['r'],prefix)
        p=prefix+'/'+e['name']; paths[p]=e
        if e['typ'] in (1,5): walk(e['c'],p)
    walk(root['c'],'')
    return paths,read
def decompress(d):
    assert d[0]==1; out=bytearray(); i=1
    while i<len(d):
        hdr=struct.unpack_from('<H',d,i)[0]; size=(hdr&0xFFF)+3; comp=hdr>>15; end=i+size; i+=2
        if not comp: out+=d[i:i+4096]; i=end; continue
        start=len(out)
        while i<end:
            flags=d[i]; i+=1
            for b in range(8):
                if i>=end: break
                if not flags&(1<<b): out.append(d[i]); i+=1
                else:
                    tok=struct.unpack_from('<H',d,i)[0]; i+=2
                    dl=len(out)-start; bc=max((dl-1).bit_length(),4); lm=0xFFFF>>bc
                    ln=(tok&lm)+3; off=(tok>>(16-bc))+1
                    for _ in range(ln): out.append(out[-off])
    return bytes(out)
data=open(sys.argv[1],'rb').read(); paths,read=cfb(data)
dirst=decompress(read(paths['/VBA/dir']))
# parse dir for module names + offsets
i=0; mods=[]; cur={}
while i<len(dirst):
    rid,sz=struct.unpack_from('<HI',dirst,i); i+=6
    if rid==9: sz=6
    val=dirst[i:i+sz]; i+=sz
    if rid==0x19: cur={'name':val.decode('latin1')}
    elif rid==0x1A: cur['stream']=val.decode('latin1')
    elif rid==0x31: cur['off']=struct.unpack('<I',val)[0]
    elif rid==0x2B: mods.append(cur)
os.makedirs(sys.argv[2],exist_ok=True)
for m in mods:
    raw=read(paths['/VBA/'+m['stream']]); src=decompress(raw[m['off']:]).decode('cp1252','replace')
    open(os.path.join(sys.argv[2],m['name']+'.bas'),'w').write(src)
    print(m['name'],len(src))
