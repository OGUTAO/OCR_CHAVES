import os
import requests
import pandas as pd
import re
from dotenv import load_dotenv
load_dotenv()

# --- CONFIGURAÇÕES ---
# ⚠️ COLE A SUA CHAVE DA API DO OCR.SPACE AQUI
OCR_API_KEY = os.getenv("OCR_API_KEY")  # Sua chave já está aqui
OCR_API_URL = "https://api.ocr.space/parse/image"

DIRETORIO_IMAGENS = 'imagens_chaves'

REGEX_PADRAO_CHAVE_COMPLETA = r'([A-Z0-9]{5}-[A-Z0-9]{5}-[A-Z0-9]{5}-[A-Z0-9]{5}-[A-Z0-9]{5})'

# --- FUNÇÕES PRINCIPAIS ---

def extrair_texto_com_api(caminho_imagem: str) -> str | None:
    
    nome_arquivo = os.path.basename(caminho_imagem)
    print(f"    Enviando '{nome_arquivo}' para a API OCR.space...")

    payload = {
        'apikey': OCR_API_KEY,
        'language': 'eng',          
        'detectOrientation': True,
        'scale': True,              
        'OCREngine': 2,             
    }

    try:
        with open(caminho_imagem, 'rb') as f:
            response = requests.post(
                OCR_API_URL,
                files={'file': (nome_arquivo, f, 'image/jpeg')},
                data=payload,
                timeout=60  
            )
        response.raise_for_status()
        result = response.json()

        if result.get("IsErroredOnProcessing"):
            print(f"    ❌ Erro retornado pela API: {result.get('ErrorMessage')}")
            return None

        if result.get("ParsedResults") and result["ParsedResults"][0].get("ParsedText"):
            print("    ✅ Texto extraído com sucesso.")
            return result["ParsedResults"][0]["ParsedText"]
        else:
            print("    ⚠️ A API não retornou texto para esta imagem.")
            return None

    except FileNotFoundError:
        print(f"    ❌ Arquivo não encontrado: {caminho_imagem}")
        return None
    except requests.exceptions.RequestException as e:
        print(f"    ❌ Erro de conexão ou na requisição para a API: {e}")
        return None
    except Exception as e:
        print(f"    ❌ Ocorreu um erro inesperado: {e}")
        return None

def extrair_chaves_do_texto(texto_bruto: str) -> list:

    if not texto_bruto:
        return []
    texto_formatado = re.sub(r'[\s\t\r\n]+', ' ', texto_bruto.upper())
    chaves_encontradas = re.findall(REGEX_PADRAO_CHAVE_COMPLETA, texto_formatado)
    return list(set(chaves_encontradas))

def processar_diretorio(diretorio_img: str) -> list:

    if not os.path.isdir(diretorio_img):
        print(f"❌ Diretório de imagens não encontrado: {diretorio_img}")
        return []

    imagens = [f for f in os.listdir(diretorio_img) if f.lower().endswith(('.png', '.jpg', '.jpeg'))]
    if not imagens:
        print(f"⚠️ Nenhuma imagem encontrada em '{diretorio_img}'")
        return []

    todas_chaves_encontradas = []
    for nome_img in imagens:
        print(f"\n🖼️ Processando {nome_img}...")
        caminho_completo = os.path.join(diretorio_img, nome_img)
        texto_da_api = extrair_texto_com_api(caminho_completo)

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
    
    if not dados:
        return

    df = pd.DataFrame(dados)
    try:
        df.to_excel(nome_arquivo, index=False)
        print(f"\n✅ Chaves extraídas com sucesso! Arquivo salvo como '{nome_arquivo}'.")
    except PermissionError:
        print(f"\n❌ ERRO DE PERMISSÃO: Feche o arquivo Excel '{nome_arquivo}' e tente novamente.")
    except Exception as e:
        print(f"\n❌ Erro ao salvar o arquivo Excel: {e}")

if __name__ == "__main__":
    if not OCR_API_KEY or OCR_API_KEY == "SUA_CHAVE_API_AQUI":
        print("‼️ ERRO CRÍTICO: A chave da API (OCR_API_KEY) não foi definida no script!")
    else:
        # Captura os resultados e o número de imagens processadas
        resultados, total_imagens = processar_diretorio(DIRETORIO_IMAGENS)
        
        # Salva os resultados no Excel
        salvar_em_excel(resultados)

        # --- RESUMO DETALHADO ADICIONADO ---
        print("\n" + "="*40)
        print("📊 RESUMO DA EXECUÇÃO")
        print("="*40)
        print(f"Total de imagens na pasta: {total_imagens}")
        
        if resultados:
            imagens_com_chaves = len(set(item['Imagem'] for item in resultados))
            total_chaves = len(resultados)
            print(f"Imagens com chaves encontradas: {imagens_com_chaves}")
            print(f"Total de chaves extraídas: {total_chaves}")
        else:
            print("Nenhuma chave foi encontrada em nenhuma imagem.")
        print("="*40)