@echo off
REM Compila la aplicacion en un .exe portable (carpeta dist\)
pip install -r requirements.txt
pyinstaller --noconfirm --onefile --windowed --name ComitePolicial run.py
copy /Y .env.example dist\.env
if not exist dist\output mkdir dist\output
if not exist dist\plantillas mkdir dist\plantillas
echo.
echo Listo: dist\ComitePolicial.exe  (editar dist\.env para configurar)
pause
