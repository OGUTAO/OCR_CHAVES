import os
import requests
import pandas as pd
import re
from dotenv import load_dotenv
import google.generativeai as genai
from PIL import Image

load_dotenv() 

OCR_API_KEY = os.getenv("OCR_API_KEY")
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")

OCR_API_URL = "https://api.ocr.space/parse/image"

DIRETORIO_IMAGENS = 'imagens_chaves'

REGEX_PADRAO_CHAVE_COMPLETA = r'([A-Z0-9]{5}-[A-Z0-9]{5}-[A-Z0-9]{5}-[A-Z0-9]{5}-[A-Z0-9]{5})'

# --- FUNÇÕES DE EXTRAÇÃO DE TEXTO ---

def extrair_texto_com_ocr_space(caminho_imagem: str) -> str | None:

    nome_arquivo = os.path.basename(caminho_imagem)
    print(f"    Enviando '{nome_arquivo}' para a API OCR.space...")
    payload = {
        'apikey': OCR_API_KEY, 'language': 'eng', 'scale': True, 'OCREngine': 2,
    }
    try:
        with open(caminho_imagem, 'rb') as f:
            response = requests.post(
                OCR_API_URL,
                files={'file': (nome_arquivo, f, 'image/jpeg')},
                data=payload, timeout=60
            )
        response.raise_for_status()
        result = response.json()
        if result.get("IsErroredOnProcessing"):
            print(f"    ❌ Erro retornado pela API OCR.space: {result.get('ErrorMessage')}")
            return None
        if result.get("ParsedResults") and result["ParsedResults"][0].get("ParsedText"):
            print("    ✅ Texto extraído com sucesso via OCR.space.")
            return result["ParsedResults"][0]["ParsedText"]
        else:
            print("    ⚠️ A API OCR.space não retornou texto.")
            return None
    except Exception as e:
        print(f"    ❌ Erro na chamada da API OCR.space: {e}")
        return None

def extrair_texto_com_gemini(caminho_imagem: str) -> str | None:
    nome_arquivo = os.path.basename(caminho_imagem)
    print(f"    Enviando '{nome_arquivo}' para a API Gemini...")
    try:
        genai.configure(api_key=GEMINI_API_KEY)
        model = genai.GenerativeModel('gemini-1.5-flash-latest')
        img = Image.open(caminho_imagem)
        

        prompt = """
        Analise esta imagem de uma etiqueta de software e transcreva TODO o texto visível.
        Preserve a estrutura e as quebras de linha o máximo possível.
        Não adicione comentários, explicações ou formatação extra.
        Apenas retorne o texto contido na imagem.
        """
        
        response = model.generate_content([prompt, img])
        print("    ✅ Texto extraído com sucesso via Gemini.")
        return response.text
        
    except Exception as e:
        print(f"    ❌ Erro na chamada da API Gemini: {e}")
        return None


def extrair_chaves_do_texto(texto_bruto: str) -> list:
    """Usa Regex para encontrar todas as chaves em um bloco de texto."""
    if not texto_bruto: return []
    texto_formatado = re.sub(r'[\s\t\r\n]+', ' ', texto_bruto.upper())
    chaves_encontradas = re.findall(REGEX_PADRAO_CHAVE_COMPLETA, texto_formatado)
    return list(set(chaves_encontradas))

def processar_diretorio(diretorio_img: str, metodo_ocr: str) -> tuple[list, int]:
    if not os.path.isdir(diretorio_img):
        print(f"❌ Diretório de imagens não encontrado: {diretorio_img}")
        return [], 0

    imagens = [f for f in os.listdir(diretorio_img) if f.lower().endswith(('.png', '.jpg', '.jpeg'))]
    if not imagens:
        print(f"⚠️ Nenhuma imagem encontrada em '{diretorio_img}'")
        return [], 0

    todas_chaves_encontradas = []
    for nome_img in imagens:
        print(f"\n🖼️ Processando {nome_img}...")
        caminho_completo = os.path.join(diretorio_img, nome_img)
        
        texto_da_api = None
        if metodo_ocr == '1':
            texto_da_api = extrair_texto_com_ocr_space(caminho_completo)
        elif metodo_ocr == '2':
            texto_da_api = extrair_texto_com_gemini(caminho_completo)

        if texto_da_api:
            chaves = extrair_chaves_do_texto(texto_da_api)
            if chaves:
                for chave in chaves:
                    print(f"  🔑 Chave Encontrada: {chave}")
                    todas_chaves_encontradas.append({'Imagem': nome_img, 'Chave': chave})
            else:
                print("  ⚠️ Nenhuma chave no formato correto foi encontrada no texto retornado.")

    return todas_chaves_encontradas, len(imagens)

def salvar_em_excel(dados: list, nome_arquivo: str = 'chaves_extraidas_final.xlsx'):

    if not dados: return
    df = pd.DataFrame(dados)
    try:
        df.to_excel(nome_arquivo, index=False)
        print(f"\n✅ Chaves extraídas com sucesso! Arquivo salvo como '{nome_arquivo}'.")
    except Exception as e:
        print(f"\n❌ Erro ao salvar o arquivo Excel: {e}")

if __name__ == "__main__":
    
    metodo_escolhido = ''
    while metodo_escolhido not in ['1', '2']:
        metodo_escolhido = input("""
Escolha o método para extração de texto:
[1] Simples
[2] Avançado
Digite 1 ou 2: """)

    if metodo_escolhido == '1' and (not OCR_API_KEY):
        print("‼️ ERRO CRÍTICO: Chave da API do OCR.space (OCR_API_KEY) não encontrada no arquivo .env!")
    elif metodo_escolhido == '2' and (not GEMINI_API_KEY):
        print("‼️ ERRO CRÍTICO: Chave da API do Gemini (GEMINI_API_KEY) não encontrada no arquivo .env!")
    else:
        resultados, total_imagens = processar_diretorio(DIRETORIO_IMAGENS, metodo_escolhido)
        salvar_em_excel(resultados)

        print("\n" + "="*40)
        print("📊 RESUMO DA EXECUÇÃO")
        print("="*40)
        print(f"Método utilizado: {'OCR.space' if metodo_escolhido == '1' else 'Google Gemini'}")
        print(f"Total de imagens na pasta: {total_imagens}")
        
        if resultados:
            imagens_com_chaves = len(set(item['Imagem'] for item in resultados))
            total_chaves = len(resultados)
            print(f"Imagens com chaves encontradas: {imagens_com_chaves}")
            print(f"Total de chaves extraídas: {total_chaves}")
        else:
            print("Nenhuma chave foi encontrada em nenhuma imagem.")
        print("="*40)