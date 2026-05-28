import zipfile, os
phrase = 'El periodo se caracterizó por la resolución de una alta demanda de consultas técnicas y administrativas'
base = r'C:\Users\cmolina\Desktop\varios\k8s-n8n\data\pptx-parts'
for root, dirs, files in os.walk(base):
    for f in files:
        if f.lower().endswith('.pptx'):
            fn = os.path.join(root, f)
            with zipfile.ZipFile(fn) as z:
                for name in z.namelist():
                    if name.endswith('.xml'):
                        data = z.read(name).decode('utf-8', errors='ignore')
                        if phrase in data:
                            print('FOUND', fn, name)
