"""Teste independente do Corel: PDF -> Ghostscript -> chapas CMYK/spot.
Requer Python, reportlab, Pillow e Ghostscript instalado.
Uso: python TestarChapas.py [caminho-do-gswin64c.exe-ou-gs]
"""
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path
from PIL import Image
from reportlab.pdfgen import canvas
from reportlab.lib.colors import CMYKColorSep

executavel = sys.argv[1] if len(sys.argv) > 1 else shutil.which('gs')
if not executavel:
    raise SystemExit('Informe o caminho do Ghostscript.')
verificacoes = 0

def exigir(valor, descricao):
    global verificacoes
    verificacoes += 1
    if not valor:
        raise AssertionError(descricao)

with tempfile.TemporaryDirectory(prefix='teste_chapas_') as pasta:
    pasta = Path(pasta)
    pdf = pasta / 'chapas.pdf'
    c = canvas.Canvas(str(pdf), pagesize=(360, 180))
    c.setFillColorCMYK(1, 0, 0, 0)
    c.rect(0, 0, 90, 180, fill=1, stroke=0)
    c.setFillColor(CMYKColorSep(0, 1, .6, 0, spotName='PANTONE 186 C'))
    c.rect(90, 0, 90, 180, fill=1, stroke=0)
    c.setFillColor(CMYKColorSep(.2, .3, .8, .1, spotName='PANTONE 871 C'))
    c.rect(180, 0, 90, 180, fill=1, stroke=0)
    c.setFillColorCMYK(0, 0, 0, 1)
    c.rect(270, 0, 90, 180, fill=1, stroke=0)
    # Faixa inferior: dois Pantones sobrepostos a ciano; sobreimpressao so a esquerda.
    c.setFillColorCMYK(1, 0, 0, 0)
    c.rect(0, 0, 180, 45, fill=1, stroke=0)
    c.setFillOverprint(True)
    c.setOverprintMask(True)
    c.setFillColor(CMYKColorSep(0, 1, .6, 0, spotName='PANTONE 186 C'))
    c.rect(0, 0, 90, 45, fill=1, stroke=0)
    c.setFillOverprint(False)
    c.setFillColor(CMYKColorSep(0, 1, .6, 0, spotName='PANTONE 186 C', density=.5))
    c.rect(90, 0, 90, 45, fill=1, stroke=0)
    c.save()
    resultado = subprocess.run([executavel, '-dSAFER', '-dBATCH', '-dNOPAUSE', '-sDEVICE=tiffsep',
        '-dPrintSpotCMYK=true', '-dMaxSpots=60', '-dFirstPage=1', '-dLastPage=1', '-dUseCropBox',
        '-sCompression=none', '-dOverprint=/enable', '-r72', '-sOutputFile='+str(pasta/'placas.tif'), str(pdf)],
        capture_output=True, text=True, timeout=120, check=True)
    exigir('PANTONE 186 C' in resultado.stdout + resultado.stderr, 'metadados Pantone 186 C')
    exigir('PANTONE 871 C' in resultado.stdout + resultado.stderr, 'metadados Pantone 871 C')
    esperado = {'Cyan':[0,255,255,255], 'Magenta':[255]*4, 'Yellow':[255]*4,
        'Black':[255,255,255,0], 'PANTONE 186 C':[255,0,255,255], 'PANTONE 871 C':[255,255,0,255]}
    chapas = {}
    for nome, valores in esperado.items():
        with Image.open(pasta/('placas('+nome+').tif')) as im:
            exigir(im.mode == 'L' and im.size == (360,180), 'formato/dimensoes '+nome)
            exigir([im.getpixel((x,90)) for x in (45,135,225,315)] == valores, 'cobertura independente '+nome)
            chapas[nome] = [im.getpixel((x,160)) for x in (45,135)]
    exigir(chapas['Cyan'][0] == 0 and chapas['PANTONE 186 C'][0] == 0, 'sobreimpressao preserva ciano e Pantone')
    exigir(chapas['Cyan'][1] == 255, 'knockout remove ciano sob Pantone')
    exigir(abs(chapas['PANTONE 186 C'][1]-127) <= 1, 'reticula Pantone 50 por cento')
    exigir(len(list(pasta.glob('placas*.tif'))) == 7, 'quatro processos, duas especiais e composto')
print(f'PASSOU: {verificacoes} verificacoes de chapas, sobreimpressao, knockout e reticula.')
