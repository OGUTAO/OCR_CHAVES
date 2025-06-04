import os
import cv2
import easyocr
import pandas as pd
import re
import numpy as np

# Diretórios
diretorio_imagens = 'imagens_chaves'  # ATENÇÃO: Verifique se este é o nome correto da sua pasta
diretorio_debug = 'debug_output'
os.makedirs(diretorio_debug, exist_ok=True)

# Inicializar OCR
try:
    reader = easyocr.Reader(['en'], gpu=False, verbose=False)
    print("EasyOCR Reader inicializado.")
except Exception as e:
    print(f"Erro ao inicializar EasyOCR: {e}")
    exit()

CARACTERES_VALIDOS = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789-'
REGEX_PADRAO_CHAVE_COMPLETA = r'([A-Z0-9]{5}-[A-Z0-9]{5}-[A-Z0-9]{5}-[A-Z0-9]{5}-[A-Z0-9]{5})'

def mostrar_imagem(titulo, imagem):
    h, w = imagem.shape[:2]
    max_h, max_w = 700, 1000 # Ajuste para o tamanho da sua tela se necessário
    if h > max_h or w > max_w:
        scale = min(max_h/h, max_w/w)
        imagem_resized = cv2.resize(imagem, None, fx=scale, fy=scale, interpolation=cv2.INTER_AREA)
        cv2.imshow(titulo, imagem_resized)
    else:
        cv2.imshow(titulo, imagem)
    print(f"Pressione qualquer tecla na janela '{titulo}' para continuar...")
    cv2.waitKey(0)
    cv2.destroyAllWindows()

def preprocessar_imagem(caminho_imagem, salvar_debug=False, nome_imagem=''):
    img = cv2.imread(caminho_imagem)
    if img is None:
        print(f"⚠️ Imagem não encontrada ou não pôde ser lida: {caminho_imagem}")
        return None, None

    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    
    # CLAHE para melhorar o contraste local - pode ajudar antes do adaptiveThreshold
    clahe = cv2.createCLAHE(clipLimit=3.0, tileGridSize=(8, 8)) # Experimente clipLimit 2.0 a 4.0
    gray_contraste = clahe.apply(gray)

    # === PONTO DE AJUSTE CRÍTICO no adaptiveThreshold ===
    block_size_atual = 25 # Deve ser ímpar. Experimente 15, 21, 25, 31.
    c_atual = 2           # <<< PRÓXIMO TESTE: COMECE COM 2. 
                          # Se caracteres muito finos/quebrados, tente diminuir MAIS (ex: 1) ou aumentar blockSize.
                          # Se muito ruído de fundo (pontos pretos), AUMENTE C (ex: 3, 4, 5).
    
    # Alternativa de método de limiarização adaptativa:
    metodo_adaptativo = cv2.ADAPTIVE_THRESH_GAUSSIAN_C 
    # metodo_adaptativo = cv2.ADAPTIVE_THRESH_MEAN_C # Descomente para testar MEAN_C
    # =================================================

    bin_img = cv2.adaptiveThreshold(
        gray_contraste, # Usar a imagem com contraste melhorado
        255, 
        metodo_adaptativo, 
        cv2.THRESH_BINARY_INV, # Para texto escuro em fundo claro, resulta em texto branco/fundo preto
        blockSize=block_size_atual, 
        C=c_atual 
    )
    # Inverter para ter texto PRETO em fundo BRANCO, que EasyOCR geralmente prefere
    img_para_ocr = cv2.bitwise_not(bin_img)
    
    print(f"    Pré-processamento: CLAHE(clip={clahe.getClipLimit()}), adaptiveThreshold(method={'GAUSSIAN' if metodo_adaptativo == cv2.ADAPTIVE_THRESH_GAUSSIAN_C else 'MEAN'}, blockSize={block_size_atual}, C={c_atual})")

    # (Opcional) Leve operação morfológica para limpar pequenos ruídos
    # Use com MUITA cautela, pode danificar caracteres. Teste após ter uma boa binarização.
    # kernel_size = 2 # Experimente 2 ou 3
    # kernel_morph = cv2.getStructuringElement(cv2.MORPH_RECT, (kernel_size, kernel_size))
    # Se há "pimenta" (pontos pretos no texto branco na imagem `bin_img`):
    # bin_img_opened = cv2.morphologyEx(bin_img, cv2.MORPH_OPEN, kernel_morph, iterations=1)
    # img_para_ocr = cv2.bitwise_not(bin_img_opened) # Re-inverter
    # OU, se há "sal" (pontos brancos nos caracteres pretos na imagem `img_para_ocr`):
    # img_para_ocr = cv2.morphologyEx(img_para_ocr, cv2.MORPH_OPEN, kernel_morph, iterations=1) 
    # print(f"    Aplicada operação MORPH_OPEN (opcional) com kernel {kernel_size}x{kernel_size}")


    if salvar_debug:
        cv2.imwrite(os.path.join(diretorio_debug, f"dbg_0_gray_{nome_imagem}"), gray)
        cv2.imwrite(os.path.join(diretorio_debug, f"dbg_1_clahe_{nome_imagem}"), gray_contraste)
        cv2.imwrite(os.path.join(diretorio_debug, f"dbg_2_bin_inverted_{nome_imagem}"), bin_img) # Texto branco, fundo preto
        cv2.imwrite(os.path.join(diretorio_debug, f"dbg_3_ocr_ready_{nome_imagem}"), img_para_ocr) # Texto preto, fundo branco
    return img, img_para_ocr

def corrigir_e_formatar_texto_ocr(texto_bruto_do_bloco_ocr):
    if not texto_bruto_do_bloco_ocr: 
        return ""
    texto = str(texto_bruto_do_bloco_ocr).upper()
    mapa_correcoes_especificas = {
        'O': '0', 'I': '1', 'L': '1', 'S': '5', 'Z': '2',
        # Adicione mais baseado nos erros consistentes que você observa
        # Ex: 'B': '8' ou '8': 'B'
    }
    texto_corrigido = ""
    for char_original in texto:
        char_mapeado = mapa_correcoes_especificas.get(char_original, char_original)
        texto_corrigido += char_mapeado
    return texto_corrigido

def extrair_chaves(texto_global_corrigido):
    texto_upper = texto_global_corrigido.upper()
    texto_apenas_com_caracteres_validos = re.sub(r'[^A-Z0-9-]', '', texto_upper)
    chaves_encontradas = re.findall(REGEX_PADRAO_CHAVE_COMPLETA, texto_apenas_com_caracteres_validos)
    return list(set(chaves_encontradas))

def processar_imagem(caminho_imagem, salvar_debug=False, mostrar_debug=False):
    nome_imagem = os.path.basename(caminho_imagem)
    print(f"\n🖼️ Processando {nome_imagem}...")

    img_original, img_proc_para_ocr = preprocessar_imagem(caminho_imagem, salvar_debug, nome_imagem)

    if img_original is None or img_proc_para_ocr is None:
        print(f"  ⚠️ Falha ao pré-processar {nome_imagem}. Pulando.")
        return []

    if mostrar_debug:
        mostrar_imagem(f"Original - {nome_imagem}", img_original)
        mostrar_imagem(f"Pronta para OCR ({nome_imagem})", img_proc_para_ocr)

    # === PONTO DE AJUSTE EASYOCR ===
    LIMIAR_CONFIANCA_BLOCO_OCR = 0.15 # Pode reduzir um pouco para ver se algo mais passa, mas cuidado com lixo.
                                     # Se o pré-processamento melhorar, pode aumentar (0.2, 0.25).
    mag_ratio_atual = 2.0          # <<< PRÓXIMO TESTE: COMECE COM 2.0. Experimente 1.5 ou 2.5.
    text_threshold_atual = 0.15    # Limiar interno do EasyOCR. Experimente 0.1, 0.15, 0.2.
    # ==============================
    print(f"    EasyOCR: mag_ratio={mag_ratio_atual}, text_threshold={text_threshold_atual}, LIMIAR_CONFIANCA_BLOCO_OCR={LIMIAR_CONFIANCA_BLOCO_OCR}")

    resultado_ocr = reader.readtext(
        img_proc_para_ocr,
        detail=1,
        paragraph=False, 
        allowlist=CARACTERES_VALIDOS,
        text_threshold=text_threshold_atual, 
        low_text=text_threshold_atual,
        mag_ratio=mag_ratio_atual,    
        # x_ths=1.0, # Se o OCR "picar" texto que deveria ser contínuo, aumente (ex: 1.0, 1.5)
    )

    textos_corrigidos_confiaveis = []
    print(f"\n--- Detalhes do OCR para {nome_imagem} ---")
    if not resultado_ocr:
        print("  Nenhum texto detectado pelo OCR.")
    
    for i, item_ocr in enumerate(resultado_ocr):
        texto_detectado = item_ocr[1]
        confianca = item_ocr[2]
        
        print(f"  Bloco OCR {i+1}: '{texto_detectado}' (Confiança: {confianca:.2f})")

        if confianca < LIMIAR_CONFIANCA_BLOCO_OCR:
            print(f"      ↳ Bloco descartado (confiança {confianca:.2f} < {LIMIAR_CONFIANCA_BLOCO_OCR})")
            continue
        
        texto_bloco_corrigido = corrigir_e_formatar_texto_ocr(texto_detectado)
        if texto_bloco_corrigido != texto_detectado:
             print(f"      ↳ Bloco após correções: '{texto_bloco_corrigido}'")
        textos_corrigidos_confiaveis.append(texto_bloco_corrigido)

    texto_global_para_extrair = "".join(textos_corrigidos_confiaveis)
    
    chaves_encontradas_na_imagem = []
    if texto_global_para_extrair.strip():
        print(f"📝 Texto Global (corrigido, confiável, unido) para extração:\n'{texto_global_para_extrair}'\n")
        chaves_da_imagem = extrair_chaves(texto_global_para_extrair)

        if chaves_da_imagem:
            for c in chaves_da_imagem:
                print(f"  🔑 Chave Encontrada na Imagem: {c}")
            chaves_encontradas_na_imagem.extend(chaves_da_imagem)
        else:
            print("  ⚠️ Nenhuma chave no formato XXXXX-XXXXX-... encontrada no texto processado.")
    else:
        print("  ℹ️ Nenhum texto confiável acumulado para extração de chaves.")
        
    return list(set(chaves_encontradas_na_imagem))

def processar_diretorio(diretorio_img, salvar_debug=False, mostrar_debug=False):
    if not os.path.isdir(diretorio_img):
        print(f"❌ Diretório de imagens não encontrado: {diretorio_img}")
        return []
    imagens = [f for f in os.listdir(diretorio_img) if f.lower().endswith(('.png', '.jpg', '.jpeg'))]
    if not imagens:
        print(f"⚠️ Nenhuma imagem encontrada em: {diretorio_img}")
        return []
    todas_chaves = []
    for nome_img in imagens:
        caminho_completo = os.path.join(diretorio_img, nome_img)
        chaves = processar_imagem(caminho_completo, salvar_debug=salvar_debug, mostrar_debug=mostrar_debug)
        for chave_encontrada in chaves:
            todas_chaves.append({'Imagem': nome_img, 'Chave': chave_encontrada})
    return todas_chaves

def salvar_em_excel(dados, nome_arquivo='chaves_extraidas_final.xlsx'):
    if not dados:
        print("\nℹ️ Nenhum dado para salvar em Excel.")
        return
    if os.path.exists(nome_arquivo):
        try:
            os.remove(nome_arquivo)
            print(f"\n🗑️ Arquivo antigo {nome_arquivo} removido.")
        except PermissionError:
            print(f"❌ ERRO DE PERMISSÃO: Feche o arquivo '{nome_arquivo}' e tente novamente.")
            return
        except Exception as e_rem:
            print(f"❌ ERRO ao remover arquivo antigo '{nome_arquivo}': {e_rem}")
            return
    df = pd.DataFrame(dados)
    if df.empty:
        print("\nℹ️ DataFrame vazio, nenhum arquivo Excel gerado.")
        return
    try:
        df.to_excel(nome_arquivo, index=False)
        print(f"\n✅ Chaves extraídas com sucesso! Arquivo salvo como {nome_arquivo}.")
    except Exception as e:
        print(f"\n❌ Erro ao salvar o arquivo Excel '{nome_arquivo}': {e}")

if __name__ == "__main__":
    try:
        if not os.path.isdir(diretorio_imagens):
             print(f"‼️ ERRO CRÍTICO: A pasta de imagens '{diretorio_imagens}' não existe!")
             print(f"‼️ Por favor, crie a pasta ou corrija a variável 'diretorio_imagens' no script.")
        else:
            resultados = processar_diretorio(
                diretorio_imagens,
                salvar_debug=True,
                mostrar_debug=True  # Mantido True para visualização
            )
            if resultados:
                salvar_em_excel(resultados)
            else:
                print("\n⚠️ Nenhuma chave encontrada em nenhuma das imagens processadas.")
    except FileNotFoundError as e_fnf:
        print(f"\n❌ ERRO DE ARQUIVO/DIRETÓRIO: {e_fnf}")
    except Exception as e_main:
        print(f"\n❌ ERRO INESPERADO NO SCRIPT: {e_main}")