import zipfile
import re
files=[
    'libreoffice-python/data/plantillas/plantilla_invgate.pptx',
    'libreoffice-python/data/plantillas/plantilla_buenas_practicas.pptx',
    'libreoffice-python/data/plantillas/plantilla_portada.pptx',
    'libreoffice-python/data/plantillas/plantilla_contenido_4.pptx',
]
for fn in files:
    print('FILE', fn)
    with zipfile.ZipFile(fn) as z:
        slides=[name for name in z.namelist() if name.startswith('ppt/slides/slide')]
        for slide in slides:
            print(' SLIDE', slide)
            xml = z.read(slide).decode('utf-8')
            for match in re.finditer(r'<a:t>(.*?)</a:t>', xml, re.DOTALL):
                text = match.group(1).strip()
                if text:
                    print('  TEXT', repr(text))
            for match in re.finditer(r'<p:ph(.*?)/>', xml):
                print('  PH', match.group(1))
    print()
