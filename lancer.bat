@echo off
echo ========================================
echo   Impulse Academy - Generateur
echo ========================================
echo.
echo Installation des dependances...
pip install flask python-pptx openpyxl --quiet
echo.
echo Lancement du serveur...
echo Ouvrez votre navigateur sur : http://localhost:5000
echo (Appuyez sur Ctrl+C pour arreter)
echo.
python app.py
pause
