"""Reproduce public examples using only self-authored demo assets."""
from pathlib import Path
from PIL import Image,ImageDraw,ImageFont
from common import *
from style_engine import create_deck,build_template,load_style
from template_engine import generate
from template_creator import create
from validate_output import validate,markdown
ROOT=Path(__file__).resolve().parents[1]
def main():
    d=ROOT/'examples/demo-deck';d.mkdir(parents=True,exist_ok=True)
    assets=d/'creator-assets'
    for k in ('logos','palettes'):(assets/k).mkdir(parents=True,exist_ok=True)
    # Original typographic demo mark and color swatch, not a school logo.
    im=Image.new('RGB',(500,200),'#FFFFFF');draw=ImageDraw.Draw(im)
    try:f=ImageFont.truetype('arial.ttf',90)
    except OSError:f=ImageFont.load_default(size=90)
    draw.text((22,42),'SMART',fill='#432354',font=f);im.save(assets/'logos/smart.png')
    im=Image.new('RGB',(500,100));draw=ImageDraw.Draw(im)
    for i,c in enumerate(['#432354','#815A86','#BCADBF','#FCFAF6','#302538']):draw.rectangle([i*100,0,(i+1)*100,100],fill=c)
    im.save(assets/'palettes/purple.png')
    content=read_json(d/'content.json')
    create_deck('tech-blue',content,d/'style-demo.pptx')
    layout=build_template(load_style('academic'),d/'demo-template.pptx')
    plan={'mode':'layouts','allowed_layouts':list(layout.values()),'slides':[dict(x,layout=layout[x['kind']]) for x in content['slides']]}
    write_json(d/'template-plan.json',plan);generate(d/'demo-template.pptx',plan,d/'template-demo.pptx')
    layout=create(assets,d/'creator-template.pptx','Microsoft YaHei')
    generate(d/'creator-template.pptx',plan,d/'creator-demo.pptx')
    for name in ['template-demo','style-demo','creator-demo']:
        report=validate(d/(name+'.pptx'));write_json(d/(name+'.validation.json'),report)
        (d/(name+'.validation.md')).write_text(markdown(report),encoding='utf-8')
        print(name,'slides',report['slide_count'],'pass',report['passed'],'issues',len(report['issues']))
if __name__=='__main__':main()
