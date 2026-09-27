"""Portable native-PPTX catalog, search and page extraction. Python 3 stdlib only."""
import argparse
import copy
import hashlib
import json
import posixpath
import re
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from xml.etree import ElementTree as E

BASE = Path(__file__).resolve().parents[1]
NS = {'p':'http://schemas.openxmlformats.org/presentationml/2006/main',
      'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
      'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
      'rel':'http://schemas.openxmlformats.org/package/2006/relationships',
      'ct':'http://schemas.openxmlformats.org/package/2006/content-types',
      'p14':'http://schemas.microsoft.com/office/powerpoint/2010/main'}
for prefix in ('p','a','r','p14'): E.register_namespace(prefix, NS[prefix])
FIELDS = {'模板':'template','原型':'prototype','关系':'relation','结构':'structure',
          '适用':'use_for','槽位':'slots','注意':'avoid','分组':'group','来源':'source'}

def resolve(owner, target):
    return target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join(posixpath.dirname(owner),target))

def relpart(owner):
    if not owner:return '_rels/.rels'
    return posixpath.join(posixpath.dirname(owner),'_rels',posixpath.basename(owner)+'.rels')

def rels(z, owner):
    p=relpart(owner)
    return E.fromstring(z.read(p)) if p in z.namelist() else []

def slide_parts(z):
    root=E.fromstring(z.read('ppt/presentation.xml'))
    links={r.get('Id'):r for r in rels(z,'ppt/presentation.xml')}
    rows=[]
    for entry in root.find('p:sldIdLst',NS):
        link=links[entry.get('{'+NS['r']+'}id')]
        rows.append((entry,resolve('ppt/presentation.xml',link.get('Target'))))
    return root,rows

def catalog(deck):
    output=[];ids=set()
    with ZipFile(deck) as z:
        if z.testzip():raise ValueError('PPTX ZIP integrity failure')
        _,slides=slide_parts(z)
        for page,(_,part) in enumerate(slides,1):
            links=[r for r in rels(z,part) if r.get('Type','').endswith('/notesSlide')]
            if len(links)!=1:raise ValueError(f'Page {page}: expected one notes part')
            notes=E.fromstring(z.read(resolve(part,links[0].get('Target'))))
            body=next((s for s in notes.findall('.//p:sp',NS) if s.find('.//p:ph[@type="body"]',NS) is not None),None)
            if body is None:raise ValueError(f'Page {page}: missing note body')
            values={}
            for p in body.findall('p:txBody/a:p',NS):
                line=''.join(p.itertext()).strip()
                if not line:continue
                key,sep,value=line.partition('：')
                if key in FIELDS and sep:values[FIELDS[key]]=value
            missing=set(FIELDS.values())-set(values)
            if missing:raise ValueError(f'Page {page}: missing {sorted(missing)}')
            tid,sep,name=values.pop('template').partition('｜')
            if not re.fullmatch(r'TPL-[A-Z0-9-]+',tid) or not sep:raise ValueError(f'Invalid template ID: {tid}')
            if tid in ids:raise ValueError('Duplicate template ID: '+tid)
            ids.add(tid)
            pid,_,pname=values.pop('prototype').partition('｜')
            if pid!='无原型' and not re.fullmatch(r'T\d{2}',pid):raise ValueError('Invalid prototype ID: '+pid)
            s=E.fromstring(z.read(part))
            shape_names=[n.get('name','') for n in s.findall('.//p:cNvPr',NS)]
            output.append({'id':tid,'page':page,'name':name,'prototype':pid,'prototype_name':pname,**values,
                           'shape_count':len(s.findall('.//p:sp',NS)),
                           'table_count':len(s.findall('.//a:tbl',NS)),
                           'image_count':len(s.findall('.//p:pic',NS)),
                           'named_image_areas':sum('图片区' in n or n.startswith('留白') for n in shape_names)})
    return {'schema_version':'1.0','deck_file':deck.name,'deck_sha256':hashlib.sha256(deck.read_bytes()).hexdigest(),'slides':output}

def save_json(path,data,replace=False):
    path=Path(path);path.parent.mkdir(parents=True,exist_ok=True)
    with path.open('w' if replace else 'x',encoding='utf-8') as f:json.dump(data,f,ensure_ascii=False,indent=2)

def search(data,relation='',prototype='',query='',top=8,group=''):
    out=[]
    for r in data['slides']:
        if r['prototype']=='无原型':continue
        if relation and relation not in r['relation']:continue
        if prototype and prototype!=r['prototype']:continue
        if group and group not in r['group']:continue
        target=' '.join(r[k] for k in ['name','relation','structure','use_for','slots','prototype_name'])
        terms=[t for t in re.split(r'[\s，,、；;]+',query) if t]
        score=sum(12 if term in target else sum(term[i:i+2] in target for i in range(len(term)-1)) for term in terms)
        if query and not score:continue
        out.append({'match_score':score,**r})
    return sorted(out,key=lambda r:(-r['match_score'],r['page']))[:top]

def extract(deck,ids,output):
    output=Path(output)
    if output.exists():raise ValueError('Output already exists; choose a new filename')
    if len(ids)!=len(set(ids)):raise ValueError('Use unique IDs; duplicate slides later in your presentation editor')
    lookup={r['id']:r for r in catalog(deck)['slides']}
    missing=set(ids)-set(lookup)
    if missing:raise ValueError('Unknown IDs: '+', '.join(sorted(missing)))
    with ZipFile(deck) as z:
        pres,slides=slide_parts(z)
        selected=[slides[lookup[tid]['page']-1] for tid in ids]
        sldlist=pres.find('p:sldIdLst',NS);sldlist[:]=[copy.deepcopy(row[0]) for row in selected]
        # Section membership from the full library is invalid after selection/reordering.
        for extlist in pres.findall('p:extLst',NS):
            for extension in list(extlist):
                if extension.find('p14:sectionLst',NS) is not None:extlist.remove(extension)
            if not len(extlist):pres.remove(extlist)
        pr=E.fromstring(z.read('ppt/_rels/presentation.xml.rels'))
        selected_rids={r[0].get('{'+NS['r']+'}id') for r in selected}
        for link in list(pr):
            if link.get('Type','').endswith('/slide') and link.get('Id') not in selected_rids:pr.remove(link)
        E.register_namespace('',NS['rel'])
        overrides={'ppt/presentation.xml':E.tostring(pres,encoding='utf-8',xml_declaration=True),
                   'ppt/_rels/presentation.xml.rels':E.tostring(pr,encoding='utf-8',xml_declaration=True)}
        if 'docProps/app.xml' in z.namelist():
            app=E.fromstring(z.read('docProps/app.xml'))
            for e in app.iter():
                if e.tag.rsplit('}',1)[-1]=='Slides':e.text=str(len(ids))
            overrides['docProps/app.xml']=E.tostring(app,encoding='utf-8',xml_declaration=True)
        # Copy only parts reachable from package roots, preserving native objects.
        needed=set();pending=['']
        while pending:
            owner=pending.pop()
            if owner and owner in needed:continue
            if owner:needed.add(owner)
            rp=relpart(owner)
            if rp not in z.namelist():continue
            needed.add(rp)
            rr=E.fromstring(overrides.get(rp,z.read(rp)))
            for link in rr:
                if link.get('TargetMode')=='External':continue
                target=resolve(owner,link.get('Target'))
                if target not in z.namelist():raise ValueError('Missing related part: '+target)
                if target not in needed:pending.append(target)
        ct=E.fromstring(z.read('[Content_Types].xml'))
        for entry in list(ct):
            if entry.tag.endswith('Override') and entry.get('PartName','').lstrip('/') not in needed:ct.remove(entry)
        E.register_namespace('',NS['ct'])
        overrides['[Content_Types].xml']=E.tostring(ct,encoding='utf-8',xml_declaration=True);needed.add('[Content_Types].xml')
        output.parent.mkdir(parents=True,exist_ok=True)
        with ZipFile(output,'x',compression=ZIP_DEFLATED) as dest:
            for part in sorted(needed):dest.writestr(part,overrides.get(part,z.read(part)))
    extracted=catalog(output)
    if [r['id'] for r in extracted['slides']]!=ids:raise ValueError('Extracted slide order mismatch')
    return {'output':str(output),'ids':ids,'slides':len(ids)}

def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--deck',type=Path,default=BASE/'assets/template-library.pptx')
    sub=parser.add_subparsers(dest='command',required=True)
    ix=sub.add_parser('index');ix.add_argument('--output',required=True);ix.add_argument('--replace',action='store_true')
    ss=sub.add_parser('search');ss.add_argument('--relation',default='');ss.add_argument('--prototype',default='');ss.add_argument('--query',default='');ss.add_argument('--group',default='');ss.add_argument('--top',type=int,default=8)
    ex=sub.add_parser('extract');ex.add_argument('--ids',required=True);ex.add_argument('--output',required=True)
    sub.add_parser('validate')
    args=parser.parse_args()
    if args.command=='extract':result=extract(args.deck,[s.strip() for s in args.ids.split(',') if s.strip()],args.output)
    else:
        data=catalog(args.deck)
        if args.command=='index':save_json(args.output,data,args.replace);result={'output':args.output,'slides':len(data['slides'])}
        elif args.command=='search':result=search(data,args.relation,args.prototype,args.query,args.top,args.group)
        else:
            cached=BASE/'assets/catalog.json'
            if args.deck.resolve()==(BASE/'assets/template-library.pptx').resolve() and cached.exists():
                if json.loads(cached.read_text('utf-8'))!=data:raise ValueError('Cached catalog differs from deck; regenerate index')
            result={'status':'pass','slides':len(data['slides']),'prototypes':len({r['prototype'] for r in data['slides'] if r['prototype']!='无原型'}),'tables':sum(r['table_count'] for r in data['slides'])}
    print(json.dumps(result,ensure_ascii=False,indent=2))

if __name__=='__main__':
    try:main()
    except (ValueError,KeyError,FileExistsError) as e:raise SystemExit(str(e))
