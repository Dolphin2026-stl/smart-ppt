"""Release packager: excludes user-provided template files; excludes local test outputs."""
from pathlib import Path
import hashlib,zipfile
ROOT=Path(__file__).resolve().parents[1]
def allowed(rel):
    parts=rel.parts
    if '__pycache__' in parts or rel.suffix=='.pyc':return False
    if len(parts)==1:return rel.name in {'SKILL.md','README.md','README.zh-CN.md','requirements.txt','.gitignore','index.html'}
    if parts[0]=='scripts':return rel.suffix in {'.py','.ps1'}
    if parts[0]=='references':return rel.suffix=='.md'
    if parts[0]=='tests':return rel.suffix=='.py'
    if parts[0]=='docs':return rel.suffix in {'.md','.png','.gif','.json'}
    if parts[0]=='templates':return rel.as_posix() in {'templates/.gitkeep','templates/README.md'}
    if parts[0]=='assets':return (len(parts)==3 and parts[1]=='styles' and rel.suffix=='.json') or rel.as_posix()=='assets/fonts/mapping.json' or rel.name=='.gitkeep'
    if parts[:2]==('examples','demo-deck'):return rel.suffix.lower() in {'.pptx','.json','.md','.png'}
    return False

def main():
    output=ROOT.parent/'smart-ppt-public.zip';manifest=[]
    with zipfile.ZipFile(output,'w',zipfile.ZIP_DEFLATED) as z:
        for path in sorted(ROOT.rglob('*')):
            if path.is_file() and allowed(path.relative_to(ROOT)):
                rel=path.relative_to(ROOT);z.write(path,'smart-ppt/'+rel.as_posix());manifest.append(rel.as_posix())
    with zipfile.ZipFile(output) as z:
        if z.testzip() is not None:raise RuntimeError('ZIP CRC failure')
        if any('/local/' in n for n in z.namelist()):raise RuntimeError('Private files included')
    digest=hashlib.sha256(output.read_bytes()).hexdigest();output.with_suffix('.zip.sha256').write_text(digest+'  '+output.name+'\n',encoding='ascii')
    output.with_suffix('.manifest.txt').write_text('\n'.join(manifest)+'\n',encoding='utf-8')
    print(f'{output}: {len(manifest)} files, {output.stat().st_size} bytes, SHA-256 {digest}')
if __name__=='__main__':main()
