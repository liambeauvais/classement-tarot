@echo off
REM Création d'un environnement virtuel
python -m venv venv

REM Activation de l'environnement et installation des dépendances
call venv\\Scripts\\activate
pip install --upgrade pip
pip install -r requirements.txt
pip install pyinstaller

REM Création du dossier de distribution
pyinstaller --name "Classement_Tarot" --windowed --onefile --hidden-import=reportlab --hidden-import=PIL --hidden-import=openpyxl --clean tarot_gui.py

echo.
echo L'exécutable se trouve dans le dossier dist\
pause
