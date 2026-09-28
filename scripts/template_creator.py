"""Classify assets, quantize reference colors, author a reusable placeholder template."""
import argparse
from pathlib import Path
from PIL import Image
from common import *
from style_engine import load_style,build_template

def scan_assets(folder):
    root=Path(folder)
    if not root.is_dir():raise ValueError('assets directory does not exist')
    groups={k:[] for k in ('backgrounds','logos','icons','palettes','fonts','unclassified')}
    aliases={'background':'backgrounds','背景':'backgrounds','logo':'logos','标识':'logos','icon':'icons','图标':'icons','palette':'palettes','配色':'palettes','font':'fonts','字体':'fonts'}
    for p in sorted(root.rglob('*')):
        if not p.is_file() or p.name.startswith('.'):continue
        ext=p.suffix.lower()
        if ext in ('.ttf','.otf','.woff','.woff2'):key='fonts'
        elif ext in ('.png','.jpg','.jpeg','.webp','.bmp','.tif','.tiff'):
            parts=p.relative_to(root).as_posix().lower();key=next((k for k in groups if k!='unclassified' and k in parts),None)
            key=key or next((v for k,v in aliases.items() if k in parts),'unclassified')
        else:continue
        groups[key].append(str(p.resolve()))
    return groups

def palette(path,count=5):
    with Image.open(path) as im:
        im=im.convert('RGBA');canvas=Image.new('RGBA',im.size,'white');canvas.alpha_composite(im);im=canvas.convert('RGB');im.thumbnail((160,160))
        q=im.quantize(colors=count,method=Image.Quantize.MEDIANCUT);pal=q.getpalette()
        return ['%02X%02X%02X'%tuple(pal[idx*3:idx*3+3]) for _,idx in sorted(q.getcolors(),reverse=True)]

def create(assets,output,font,base='academic'):
    groups=scan_assets(assets)
    if not groups['logos']:raise ValueError('Provide a raster logo in assets/logos/')
    refs=groups['palettes'] or groups['backgrounds']
    if not refs:raise ValueError('Provide a background or palette reference image')
    style=load_style(base);colors=palette(refs[0]);style['fonts']['title']=font;style['fonts']['body']=font
    # Choose darkest extracted color for readable white-background text and strongest non-gray accent.
    rgb=lambda h:tuple(int(h[i:i+2],16) for i in (0,2,4))
    dark=min(colors,key=lambda h:sum(rgb(h)));accent=max(colors,key=lambda h:max(rgb(h))-min(rgb(h)))
    style['colors'].update(primary=dark,accent=accent)
    brand={'background':groups['backgrounds'][0] if groups['backgrounds'] else None,'logo':groups['logos'][0]}
    mapping=build_template(style,output,brand)
    write_json(str(output)+'.assets.json',{'assets':groups,'palette':colors,'layouts':mapping,'font_preference':font,'font_files_embedded':False,'icons_used':False})
    return mapping
if __name__=='__main__':
    p=argparse.ArgumentParser(description=__doc__);p.add_argument('assets');p.add_argument('output');p.add_argument('--font',required=True);p.add_argument('--base-style',default='academic');a=p.parse_args();create(a.assets,a.output,a.font,a.base_style)
