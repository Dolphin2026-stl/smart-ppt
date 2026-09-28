"""Build README images from PowerPoint renders; reconstruct an honest 20s walkthrough."""
from pathlib import Path
import shutil,subprocess,sys
from PIL import Image,ImageDraw,ImageFont
ROOT=Path(__file__).resolve().parents[1]
def main():
    out=ROOT/'docs/assets';out.mkdir(parents=True,exist_ok=True);frames=[]
    for mode in ('template','style','creator'):
        slides=sorted((ROOT/'local/renders'/mode).glob('*.PNG'))
        if len(slides)!=3:raise RuntimeError(f'Render all three {mode} slides first')
        shutil.copyfile(slides[0],out/'showcase'/f'{mode}.png')
    shutil.copyfile(out/'showcase/template.png',out/'hero-screenshot.png')
    try:font=ImageFont.truetype('consola.ttf',21);title=ImageFont.truetype('consolab.ttf',32)
    except OSError:font=ImageFont.load_default(size=21);title=ImageFont.load_default(size=32)
    d=ROOT/'examples/demo-deck'
    commands=[
        ('1 / Inspect template',[sys.executable,'scripts/layout_inspector.py',str(d/'demo-template.pptx'),'--table']),
        ('2 / Fill selected layouts',[sys.executable,'scripts/template_engine.py',str(d/'demo-template.pptx'),str(d/'template-plan.json'),str(ROOT/'local/gif-demo.pptx')]),
        ('3 / Validate output',[sys.executable,'scripts/validate_output.py',str(ROOT/'local/gif-demo.pptx')])]
    for label,args in commands:
        result=subprocess.run(args,cwd=ROOT,capture_output=True,text=True,encoding='utf-8',check=True)
        canvas=Image.new('RGB',(1200,675),'#08192F');draw=ImageDraw.Draw(canvas)
        draw.text((40,35),'smart-ppt  /  '+label,font=title,fill='#44DDDA')
        draw.text((40,100),'Recorded command result (reconstructed walkthrough)',font=font,fill='#A2B6C9')
        lines=result.stdout.splitlines()[:13]
        for i,line in enumerate(lines):draw.text((40,158+i*29),line[:86],font=font,fill='#EEF6FF')
        draw.text((40,625),'Actual commands + PowerPoint renders; not a live screen recording',font=font,fill='#A2B6C9')
        frames.append(canvas)
    final=Image.open(out/'hero-screenshot.png').convert('RGB').resize((1200,675));frames.append(final)
    frames[0].save(out/'demo-workflow.gif',save_all=True,append_images=frames[1:],duration=[5000]*4,loop=0,optimize=True)
    print('Generic showcase generated')
if __name__=='__main__':main()
