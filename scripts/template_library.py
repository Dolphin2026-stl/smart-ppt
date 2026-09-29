"""Index complete user-supplied PPTX files recursively under templates/."""
import argparse,hashlib
from pathlib import Path
from pptx import Presentation
from common import read_json,write_json
ROOT=Path(__file__).resolve().parents[1]

def template_roots():
    base=ROOT/'templates';base.mkdir(exist_ok=True)
    roots=[base]
    config=ROOT/'template_roots.local.json'
    if config.exists():
        for value in read_json(config).get('template_roots',[]):
            path=Path(value).expanduser().resolve()
            if not path.is_dir():raise ValueError(f'Template directory does not exist: {path}')
            if path not in roots:roots.append(path)
    return roots
def digest(path):return hashlib.sha256(Path(path).read_bytes()).hexdigest()
def index_library():
    base=ROOT/'templates';items=[]
    for root in template_roots():
        for p in sorted(root.rglob('*.pptx')):
            if p.name.startswith('~$') or p.is_symlink():continue
            label=p.relative_to(ROOT).as_posix() if p.is_relative_to(ROOT) else str(p.resolve())
            prs=Presentation(p)
            items.append({'id':'tpl-'+hashlib.sha256(label.encode()).hexdigest()[:12],'name':p.name,'path':label,'slides':len(prs.slides),'masters':len(prs.slide_masters),'sha256':digest(p),'bytes':p.stat().st_size})
    write_json(base/'library.json',{'policy':'Complete files; all pages retained until explicitly edited.','templates':items})
    return items

def list_library():
    for x in index_library():print(x['id'],x['slides'],'slides',x['path'])
def select(template_id,output):
    catalog=ROOT/'templates/library.json'
    if not catalog.exists():index_library()
    matches=[x for x in read_json(catalog)['templates'] if x['id']==template_id]
    if len(matches)!=1:raise ValueError('Unknown template ID; run list to refresh the index')
    item=matches[0];source=(ROOT/item['path']).resolve()
    if not any(source.is_relative_to(root.resolve()) for root in template_roots()):raise ValueError('Template path is outside configured roots')
    if digest(source)!=item['sha256']:raise ValueError('Template changed since indexing; run list again')
    from full_deck_engine import generate_full
    return generate_full(source,{},output)
if __name__=='__main__':
    p=argparse.ArgumentParser(description=__doc__);sub=p.add_subparsers(dest='action',required=True)
    sub.add_parser('list');sub.add_parser('index')
    a=sub.add_parser('select');a.add_argument('template_id');a.add_argument('output')
    a=p.parse_args()
    if a.action=='select':select(a.template_id,a.output)
    else:list_library()
