import os
import sys  # Importação necessária para encontrar o ícone
import pandas as pd
import re
from dotenv import load_dotenv
import google.generativeai as genai
from PIL import Image
import threading
import customtkinter as ctk
from tkinter import filedialog, messagebox

# --- CONFIGURAÇÃO INICIAL E CONSTANTES ---
load_dotenv()
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
NOME_ARQUIVO_EXCEL = 'chaves_extraidas_final.xlsx'


# --- NOVA FUNÇÃO AUXILIAR PARA ENCONTRAR ARQUIVOS NO APP ---
def resource_path(relative_path):
    """ Retorna o caminho absoluto para o recurso, funciona para desenvolvimento e para o executável do PyInstaller """
    try:
        # PyInstaller cria uma pasta temporária e armazena o caminho em sys._MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


# --- LÓGICA DE EXTRAÇÃO E ANÁLISE (MOTOR DO PROGRAMA) ---

def extrair_e_analisar_imagem(caminho_imagem: str) -> list:
    """Usa um prompt avançado para extrair chaves e a auto-análise de confiança da IA."""
    nome_arquivo = os.path.basename(caminho_imagem)
    print(f"\n🖼️  Analisando '{nome_arquivo}' com diagnóstico da IA...")
    try:
        genai.configure(api_key=GEMINI_API_KEY)
        model = genai.GenerativeModel('gemini-1.5-flash-latest')
        img = Image.open(caminho_imagem)
        
        prompt = """
        Você é um especialista em OCR para chaves de produto da Microsoft.
        Sua tarefa é transcrever TODAS as chaves de produto da imagem com precisão máxima.

        Para cada chave de produto encontrada, forneça DUAS informações em linhas separadas e consecutivas:
        1.  Uma linha começando com "CHAVE: " seguida da chave no formato XXXXX-XXXXX-XXXXX-XXXXX-XXXXX.
        2.  Uma linha começando com "VERIFICACAO: " contendo os caracteres que você considera mais difíceis de ler ou ambíguos nesta chave (ex: O, 0, B, 8, S, 5). Se a leitura for 100% clara para você, escreva "Nenhum".

        Exemplo de uma saída correta para uma imagem com duas chaves:
        CHAVE: GR79F-V4NGQ-RBGQK-X4RVV-PWF9C
        VERIFICACAO: G, B, Q
        CHAVE: NCKM6-93VT7-D64WF-2X9VK-MG9TT
        VERIFICACAO: Nenhum
        """
        
        response = model.generate_content([prompt, img])
        print(f"    ✅ Texto e análise recebidos com sucesso.")

        if not response.text: return []
        
        # Faz o parsing da resposta estruturada
        resultados = []
        linhas = response.text.strip().split('\n')
        for i, linha in enumerate(linhas):
            if linha.upper().startswith("CHAVE:"):
                chave = linha[len("CHAVE:"):].strip()
                verificacao = "Não informado" # Padrão
                if i + 1 < len(linhas) and linhas[i+1].upper().startswith("VERIFICACAO:"):
                    verificacao = linhas[i+1][len("VERIFICACAO:"):].strip()
                
                print(f"    🔑 Chave: {chave} (Diagnóstico da IA: {verificacao})")
                resultados.append({'Imagem': nome_arquivo, 'Chave': chave, 'Diagnostico': verificacao})
        return resultados

    except Exception as e:
        print(f"    ❌ Erro na chamada da API Gemini para '{nome_arquivo}': {e}")
        return []

# --- APLICAÇÃO COM INTERFACE GRÁFICA (PAINEL) ---

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("SysKey - Painel de Extração de Chaves Windows")
        self.geometry("1000x750")
        ctk.set_appearance_mode("dark")
        
        # --- LINHA ADICIONADA PARA DEFINIR O ÍCONE DA JANELA ---
        try:
            # Garante que o nome do seu ícone seja exatamente 'icone.ico'
            self.iconbitmap(resource_path("icone.ico"))
        except Exception as e:
            print(f"Aviso: Não foi possível carregar o ícone da janela: {e}")

        self.dados_chaves = []
        self.grid_columnconfigure(0, weight=1); self.grid_rowconfigure(1, weight=1); self.grid_rowconfigure(3, weight=0)
        
        # (O resto do __init__ permanece o mesmo)
        self.frame_botoes = ctk.CTkFrame(self); self.frame_botoes.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
        self.botao_carregar_arquivo = ctk.CTkButton(self.frame_botoes, text="Carregar Arquivo", command=self.iniciar_extracao_arquivo); self.botao_carregar_arquivo.pack(side="left", padx=5, pady=5)
        self.botao_carregar_pasta = ctk.CTkButton(self.frame_botoes, text="Carregar Pasta", command=self.iniciar_extracao_pasta); self.botao_carregar_pasta.pack(side="left", padx=5, pady=5)
        self.frame_rolavel = ctk.CTkScrollableFrame(self, label_text="Chaves Extraídas"); self.frame_rolavel.grid(row=1, column=0, padx=10, pady=0, sticky="nsew")
        self.botao_adicionar = ctk.CTkButton(self, text="Adicionar Chave Manualmente", command=lambda: self.adicionar_linha_chave()); self.botao_adicionar.grid(row=2, column=0, padx=10, pady=(10,5), sticky="ew")
        self.frame_analise = ctk.CTkFrame(self, height=150); self.frame_analise.grid(row=3, column=0, padx=10, pady=5, sticky="ew")
        self.frame_analise.grid_propagate(False); self.frame_analise.grid_columnconfigure(0, weight=1); self.frame_analise.grid_rowconfigure(0, weight=1)
        self.textbox_analise = ctk.CTkTextbox(self.frame_analise, wrap="word", state="disabled", font=("Consolas", 12)); self.textbox_analise.grid(row=0, column=0, sticky="nsew")
        self.frame_status_progresso = ctk.CTkFrame(self); self.frame_status_progresso.grid(row=4, column=0, padx=10, pady=5, sticky="ew")
        self.frame_status_progresso.grid_columnconfigure(0, weight=1)
        self.label_status = ctk.CTkLabel(self.frame_status_progresso, text="Pronto. Selecione um arquivo ou pasta para começar.", text_color="gray"); self.label_status.grid(row=0, column=0, padx=5, sticky="w")
        self.progressbar = ctk.CTkProgressBar(self.frame_status_progresso, orientation="horizontal")
        self.frame_salvar = ctk.CTkFrame(self); self.frame_salvar.grid(row=5, column=0, padx=10, pady=5, sticky="ew")
        self.botao_adicionar_excel = ctk.CTkButton(self.frame_salvar, text="Adicionar ao Excel", fg_color="sea green", hover_color="dark green", command=self.adicionar_ao_excel); self.botao_adicionar_excel.pack(side="left", padx=5, pady=5)
        self.botao_substituir_excel = ctk.CTkButton(self.frame_salvar, text="Substituir a Planilha Excel", command=self.substituir_em_excel); self.botao_substituir_excel.pack(side="left", padx=5, pady=5)
        
    def _bloquear_botoes(self, processando=True):
        estado = "disabled" if processando else "normal"
        self.botao_carregar_arquivo.configure(state=estado); self.botao_carregar_pasta.configure(state=estado)
        self.botao_adicionar_excel.configure(state=estado); self.botao_substituir_excel.configure(state=estado)
        self.botao_adicionar.configure(state=estado)
        
    def iniciar_extracao_base(self, target_func, target_arg):
        if not target_arg: self.label_status.configure(text="Operação cancelada."); return
        self.label_status.configure(text=f"Processando '{os.path.basename(target_arg)}'...")
        self.progressbar.grid(row=0, column=1, padx=10, sticky="ew"); self.progressbar.set(0)
        self._bloquear_botoes(True)
        self.atualizar_painel([])
        self.textbox_analise.configure(state="normal"); self.textbox_analise.delete("1.0", "end"); self.textbox_analise.configure(state="disabled")
        threading.Thread(target=target_func, args=(target_arg,), daemon=True).start()

    def iniciar_extracao_arquivo(self):
        caminho = filedialog.askopenfilename(title="Selecione um Arquivo de Imagem", filetypes=[("Arquivos de Imagem", "*.jpg *.jpeg *.png"), ("Todos os arquivos", "*.*")])
        self.iniciar_extracao_base(self._processar_arquivo_em_background, caminho)

    def iniciar_extracao_pasta(self):
        caminho = filedialog.askdirectory(title="Selecione a Pasta com as Imagens")
        self.iniciar_extracao_base(self._processar_pasta_em_background, caminho)
        
    def _processar_arquivo_em_background(self, caminho_arquivo):
        self.progressbar.set(0.5)
        resultados = extrair_e_analisar_imagem(caminho_arquivo)
        self.progressbar.set(1.0)
        self.after(0, self.atualizar_interface_completa, resultados, [caminho_arquivo])
        
    def _processar_pasta_em_background(self, caminho_pasta):
        imagens = [f for f in os.listdir(caminho_pasta) if f.lower().endswith(('.png', '.jpg', '.jpeg'))]
        total_imagens = len(imagens)
        if total_imagens == 0: self.after(0, self.atualizar_interface_completa, [], []); return
        
        resultados_finais = []
        for i, nome_img in enumerate(imagens):
            caminho_completo = os.path.join(caminho_pasta, nome_img)
            resultados_finais.extend(extrair_e_analisar_imagem(caminho_completo))
            self.progressbar.set((i + 1) / total_imagens)
        self.after(0, self.atualizar_interface_completa, resultados_finais, imagens)

    def atualizar_interface_completa(self, resultados, imagens_processadas):
        self.atualizar_painel(resultados)
        self.exibir_relatorio_de_revisao(resultados, imagens_processadas)
        self.label_status.configure(text=f"{len(resultados)} chaves carregadas. Prontas para revisão.")
        self.progressbar.grid_forget()
        self._bloquear_botoes(False)

    def atualizar_painel(self, resultados):
        for widget_info in self.dados_chaves:
            widget_info['frame'].destroy()
        self.dados_chaves.clear()
        for resultado in resultados:
            self.adicionar_linha_chave(
                imagem=resultado['Imagem'], 
                chave=resultado['Chave'], 
                diagnostico=resultado['Diagnostico']
            )

    def exibir_relatorio_de_revisao(self, resultados, imagens_processadas):
        chaves_para_revisao = [res for res in resultados if res['Diagnostico'].lower() != 'nenhum']
        
        relatorio = f"--- DIAGNÓSTICO DA IA ---\n"
        relatorio += f"Imagens Processadas: {len(imagens_processadas)} | Total de Chaves Encontradas: {len(resultados)}\n\n"
        
        if chaves_para_revisao:
            relatorio += f"🚨 {len(chaves_para_revisao)} CHAVE(S) MARCADAS PELA IA PARA REVISÃO MANUAL:\n"
            for item in chaves_para_revisao:
                relatorio += f"- {item['Chave']} (Imagem: {item['Imagem']})\n"
                relatorio += f"  Pontos de Dúvida da IA -> {item['Diagnostico']}\n"
        else:
            relatorio += "✅ Alta Confiança: A IA não reportou dúvidas em nenhuma das chaves extraídas."
        
        self.textbox_analise.configure(state="normal")
        self.textbox_analise.delete("1.0", "end")
        self.textbox_analise.insert("1.0", relatorio)
        self.textbox_analise.configure(state="disabled")

    def adicionar_linha_chave(self, imagem="MANUAL", chave="", diagnostico=None):
        frame_linha = ctk.CTkFrame(self.frame_rolavel); frame_linha.pack(fill="x", expand=True, padx=5, pady=2)
        
        if diagnostico is None or diagnostico.lower() == "não informado":
            cor_diagnostico = "gray"
            texto_diagnostico = "N/A"
        elif diagnostico.lower() == 'nenhum':
            cor_diagnostico = "light green"
            texto_diagnostico = "✅ Alta Confiança"
        else:
            cor_diagnostico = "orange"
            texto_diagnostico = f"⚠️ Revisar: {diagnostico}"

        label_imagem = ctk.CTkLabel(frame_linha, text=imagem, width=150); label_imagem.pack(side="left", padx=5)
        entry_chave = ctk.CTkEntry(frame_linha, width=350); entry_chave.insert(0, chave); entry_chave.pack(side="left", fill="x", expand=True, padx=5)
        label_diagnostico = ctk.CTkLabel(frame_linha, text=texto_diagnostico, text_color=cor_diagnostico, font=("Arial", 11, "bold"), width=150); label_diagnostico.pack(side="left", padx=5)
        botao_excluir = ctk.CTkButton(frame_linha, text="Excluir", width=80, fg_color="firebrick", hover_color="darkred", command=lambda f=frame_linha: self.excluir_linha_chave(f)); botao_excluir.pack(side="left", padx=5)
        self.dados_chaves.append({'frame': frame_linha, 'label': label_imagem, 'entry': entry_chave})
    
    def excluir_linha_chave(self, frame_a_excluir):
        for i, widget_info in enumerate(self.dados_chaves):
            if widget_info['frame'] == frame_a_excluir: self.dados_chaves.pop(i); break
        frame_a_excluir.destroy()
    def _coletar_dados_do_painel(self):
        if not self.dados_chaves: messagebox.showwarning("Aviso", "Não há chaves no painel para salvar."); return None
        dados_revisados = [{'Imagem': w['label'].cget("text"), 'Chave': w['entry'].get()} for w in self.dados_chaves]
        return pd.DataFrame(dados_revisados)
    def substituir_em_excel(self):
        df_novo = self._coletar_dados_do_painel();
        if df_novo is None: return
        caminho_arquivo_final = os.path.join(os.getcwd(), NOME_ARQUIVO_EXCEL)
        try:
            df_novo.to_excel(caminho_arquivo_final, index=False)
            messagebox.showinfo("Sucesso", f"As chaves foram salvas, substituindo o arquivo anterior!\n\nO arquivo será aberto a seguir.")
            if os.path.exists(caminho_arquivo_final): os.startfile(caminho_arquivo_final)
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro ao salvar o arquivo:\n{e}")
    def adicionar_ao_excel(self):
        df_novo = self._coletar_dados_do_painel()
        if df_novo is None: return
        caminho_arquivo_final = os.path.join(os.getcwd(), NOME_ARQUIVO_EXCEL)
        try:
            if os.path.exists(caminho_arquivo_final):
                df_antigo = pd.read_excel(caminho_arquivo_final)
                df_final = pd.concat([df_antigo, df_novo], ignore_index=True)
                df_final.drop_duplicates(subset=['Chave'], keep='first', inplace=True)
            else:
                df_final = df_novo
            df_final.to_excel(caminho_arquivo_final, index=False)
            messagebox.showinfo("Sucesso", f"As chaves foram adicionadas à planilha!\n\nO arquivo será aberto a seguir.")
            if os.path.exists(caminho_arquivo_final): os.startfile(caminho_arquivo_final)
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro ao salvar o arquivo:\n{e}")

if __name__ == "__main__":
    app = App()
    app.mainloop()