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

<img width="1919" height="1019" alt="image" src="https://github.com/user-attachments/assets/079a5f8f-78c6-4704-9d2a-60d1c77d3393" />
