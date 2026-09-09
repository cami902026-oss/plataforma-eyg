# -*- coding: utf-8 -*-
"""
Genera CAMBIOS.md: la bitácora de la plataforma, en español y por fecha.

POR QUÉ ES GENERADO Y NO ESCRITO A MANO. Una bitácora que hay que acordarse de
llenar se abandona a la tercera semana. Los mensajes de commit ya cuentan qué
cambió y por qué, así que esto los recoge y los ordena. Se regenera cuando se
quiera y siempre dice la verdad de lo que está desplegado.

QUÉ DEJA POR FUERA. El repo también guarda datos: cada cotización, solicitud o
minuta que alguien guarda deja su propio commit. Esos viven en `data/`, así que
aquí se piden únicamente las rutas de código y quedan afuera solos.

USO:  python scripts/generar_cambios.py          (desde la raíz del repo)
      python scripts/generar_cambios.py --copia "C:/Users/Lenovo/Desktop/Cambios_Plataforma.txt"
"""
import re
import subprocess
import sys
import io
import os

# Las rutas donde vive el código. Todo lo demás del repo son datos.
CODIGO = [
    'Index.html', 'service-worker.js', 'apps-script.gs', 'energy-proxy.gs',
    'email-to-cotiz.gs', 'email-to-oc.gs', 'whatsapp-bot.gs', 'claude-proxy.gs',
    'github-proxy.gs', 'pdf-studio.html', 'scripts', 'netlify', '.github',
]

# Lo poco que igual se cuela tocando código: subidas a mano por la web de
# GitHub, el inventario automático de las 2:30 PM y mensajes que no dicen nada.
RUIDO = re.compile(
    u'^(Add files via upload|Update [\\w.]+$|Merge |data:|test$|test |corre$'
    u'|Auto-update|\U0001F504 Inventario|\u26a1 Deploy|feat: nueva solicitud'
    u'|feat: nueva cotizaci|\U0001F4BE Backup|\U0001F4B8 Informe|\U0001F4CA Informe'
    u'|\U0001F4C4 Cotizaciones|\U0001F4E9 Solicitudes|\U0001F465 Clientes|Diag )'
)

CAB = u"""# \U0001F4CB Cambios de la plataforma E&G

Qué se le fue cambiando a la plataforma, lo más nuevo arriba. Cada línea es un
cambio que ya está en vivo: para verlo hay que recargar con **Ctrl + F5**.

> Este archivo lo genera `scripts/generar_cambios.py` desde el historial del
> repositorio. No se escribe a mano: si algo falta aquí, es que no se desplegó.

"""

MESES = [u'', u'enero', u'febrero', u'marzo', u'abril', u'mayo', u'junio', u'julio',
         u'agosto', u'septiembre', u'octubre', u'noviembre', u'diciembre']


def fecha_larga(iso):
    a, m, d = iso.split('-')
    return u'%d de %s de %s' % (int(d), MESES[int(m)], a)


def main():
    salida = subprocess.check_output(
        ['git', 'log', '--format=%h\x1f%ad\x1f%s', '--date=short', '-40000', '--'] + CODIGO,
        stderr=subprocess.STDOUT)
    if not isinstance(salida, str):
        salida = salida.decode('utf-8', 'replace')

    por_dia = []
    visto = {}
    for linea in salida.split('\n'):
        if '\x1f' not in linea:
            continue
        sha, fecha, msg = linea.split('\x1f', 2)
        msg = msg.strip()
        if not msg or RUIDO.match(msg):
            continue
        # El mismo arreglo desplegado dos veces no se cuenta dos veces.
        clave = re.sub(r'^v\d+\s*[\u2014-]\s*', '', msg).lower()
        if clave in visto:
            continue
        visto[clave] = 1
        if not por_dia or por_dia[-1][0] != fecha:
            por_dia.append((fecha, []))
        por_dia[-1][1].append((sha, msg))

    partes = [CAB]
    for fecha, items in por_dia:
        partes.append(u'\n## %s\n\n' % fecha_larga(fecha))
        for sha, msg in items:
            partes.append(u'- %s  \n  <sub>`%s`</sub>\n' % (msg, sha))
    texto = u''.join(partes)

    with io.open('CAMBIOS.md', 'w', encoding='utf-8') as f:
        f.write(texto)
    print('CAMBIOS.md: %d dias, %d cambios' % (len(por_dia), len(visto)))

    if '--copia' in sys.argv:
        destino = sys.argv[sys.argv.index('--copia') + 1]
        plano = re.sub(r'\s*<sub>`\w+`</sub>', '', texto)
        plano = plano.replace(u'## ', u'').replace(u'# ', u'')
        plano = plano.replace(u'> ', u'').replace(u'**', u'').replace(u'  \n', u'\n')
        d = os.path.dirname(destino)
        if d and not os.path.isdir(d):
            os.makedirs(d)
        with io.open(destino, 'w', encoding='utf-8') as f:
            f.write(plano)
        print('copia en', destino)


if __name__ == '__main__':
    main()
