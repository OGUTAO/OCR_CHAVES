import os
import google.generativeai as genai
from dotenv import load_dotenv

# Carrega a chave da API do ficheiro .env
load_dotenv()
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")

try:
    # Configura a API com a sua chave
    genai.configure(api_key=GEMINI_API_KEY)

    print("Modelos de IA generativa disponíveis para a sua chave:")
    print("-" * 50)

    # Pede à API para listar todos os modelos
    for m in genai.list_models():
        # Verifica se o modelo suporta o método de geração de conteúdo
        if 'generateContent' in m.supported_generation_methods:
            print(f"Nome do Modelo: {m.name}")

    print("-" * 50)
    print("\nInstrução: Copie um dos nomes de modelo acima (por exemplo, 'gemini-1.5-flash') e cole no seu ficheiro 'extrair_chaves.py' na linha 51.")

except Exception as e:
    print(f"\nOcorreu um erro ao tentar listar os modelos: {e}")
    print("Verifique se a sua chave de API no ficheiro .env está correta e se a sua ligação à internet está a funcionar.")