import zipfile
import re
files=['data/plantillas/plantilla_contenido_4.pptx','data/pptx-parts/contenido_invgate.asj.pptx']
for fn in files:
    print('FILE', fn)
    with zipfile.ZipFile(fn) as z:
        slides=[name for name in z.namelist() if name.startswith('ppt/slides/slide')]
        for slide in slides:
            xml=z.read(slide).decode('utf-8')
            print(' SLIDE', slide)
            print(' PICS', len(re.findall(r'<p:pic', xml)))
            print(' BLIPS', len(re.findall(r'<a:blip', xml)))
            print(' TEXTS:')
            for match in re.finditer(r'<a:t>(.*?)</a:t>', xml, re.DOTALL):
                txt = match.group(1).strip()
                if txt:
                    print('  ', repr(txt))
            print('---')
