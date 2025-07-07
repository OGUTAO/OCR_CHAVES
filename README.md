# OCR_CHAVES
python da versão mais atualizada instalado na maquina

Bibliotecas:
python -m pip install --upgrade pyinstaller customtkinter pandas openpyxl python-dotenv requests
pip install opencv-python numpy
pip install google-generativeai
pip install Pillow

Rodar:
cd OneDrive\"PastaDoPrograma"\OCR_CHAVES (Se estiver no OneDrive)
cd C:\OCR_CHAVES (Se estiver no disco local)
python extrair_chaves.py

Criar App:
pyinstaller --name="SysKey" --windowed --icon="icone.ico" extrair_chaves.py

Criar App Sem Bibliotecas:
"Criar (Novo) no path"
where python
python -m pip install --upgrade pyinstaller customtkinter pandas openpyxl python-dotenv requests
pyinstaller --name="SysKey" --windowed --icon="icone.ico" extrair_chaves.py

PATH para criar:
C:\Users\Felipe\AppData\Local\Programs\Python\Python313\Scripts
C:\Users\Felipe\AppData\Local\Programs\Python\Python313\
