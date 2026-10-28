"""Ejecuta todas las pruebas: python tests/run_all.py"""
import os
import subprocess
import sys

AQUI = os.path.dirname(os.path.abspath(__file__))
PRUEBAS = ["roundtrip_docx.py", "integracion_docx.py", "smoke_recuadros.py", "pptx_export_test.py", "scroll_test.py", "ui_completa_test.py"]
env = dict(os.environ, PYTHONIOENCODING="utf-8")
fallos = []
for p in PRUEBAS:
    r = subprocess.run([sys.executable, os.path.join(AQUI, p)], env=env,
                       capture_output=True, text=True, encoding="utf-8", errors="replace")
    estado = "OK" if r.returncode == 0 else "FALLA"
    print(f"{estado:6} {p}")
    if r.returncode:
        fallos.append(p)
        print(r.stdout[-800:], r.stderr[-1500:])
sys.exit(1 if fallos else 0)