import os
import sys
import pandas as pd
import threading
import customtkinter as ctk
from tkinter import filedialog, messagebox, StringVar
from PIL import Image
from dotenv import load_dotenv
import google.generativeai as genai

# --- CONFIGURAÇÃO INICIAL E CONSTANTES ---
load_dotenv()
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
NOME_ARQUIVO_EXCEL = 'chaves_extraidas_final.xlsx'


# --- FUNÇÃO AUXILIAR PARA ENCONTRAR ARQUIVOS NO APP (PyInstaller) ---
def resource_path(relative_path):
    """ Retorna o caminho absoluto para o recurso, funciona para desenvolvimento e para o executável do PyInstaller """
    try:
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
        """
        
        response = model.generate_content([prompt, img])
        print(f"    ✅ Texto e análise recebidos com sucesso.")

        if not response.text: return []
        
        resultados = []
        linhas = response.text.strip().split('\n')
        for i, linha in enumerate(linhas):
            if linha.upper().startswith("CHAVE:"):
                chave = linha[len("CHAVE:"):].strip().upper() # Garante que a chave da IA já venha maiúscula
                verificacao = "Não informado"
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
        
        try:
            self.iconbitmap(resource_path("icone.ico"))
        except Exception as e:
            print(f"Aviso: Não foi possível carregar o ícone da janela: {e}")

        self.resultados_atuais = []
        self.widgets_linhas = []
        
        self.grid_columnconfigure(0, weight=1); self.grid_rowconfigure(1, weight=1); self.grid_rowconfigure(3, weight=0)
        
        self.frame_botoes = ctk.CTkFrame(self); self.frame_botoes.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
        self.botao_carregar_arquivo = ctk.CTkButton(self.frame_botoes, text="Carregar Arquivo", command=self.iniciar_extracao_arquivo); self.botao_carregar_arquivo.pack(side="left", padx=5, pady=5)
        self.botao_carregar_pasta = ctk.CTkButton(self.frame_botoes, text="Carregar Pasta", command=self.iniciar_extracao_pasta); self.botao_carregar_pasta.pack(side="left", padx=5, pady=5)
        self.botao_limpar = ctk.CTkButton(self.frame_botoes, text="Limpar Painel", command=self.limpar_tudo, fg_color="#585858", hover_color="#404040"); self.botao_limpar.pack(side="left", padx=5, pady=5)

        self.frame_rolavel = ctk.CTkScrollableFrame(self, label_text="Chaves Extraídas"); self.frame_rolavel.grid(row=1, column=0, padx=10, pady=0, sticky="nsew")
        self.botao_adicionar = ctk.CTkButton(self, text="Adicionar Chave Manualmente", command=self.handle_adicionar_manual); self.botao_adicionar.grid(row=2, column=0, padx=10, pady=(10,5), sticky="ew")
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
        self.botao_limpar.configure(state=estado)
        
    def iniciar_extracao_base(self, target_func, target_arg):
        if not target_arg: self.label_status.configure(text="Operação cancelada."); return
        self.label_status.configure(text=f"Processando '{os.path.basename(str(target_arg))}'...")
        self.progressbar.grid(row=0, column=1, padx=10, sticky="ew"); self.progressbar.set(0)
        self._bloquear_botoes(True)
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
        novos_resultados = extrair_e_analisar_imagem(caminho_arquivo)
        self.progressbar.set(1.0)
        self.after(0, self.atualizar_interface_completa, novos_resultados, [caminho_arquivo])
        
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

    def atualizar_interface_completa(self, novos_resultados, imagens_processadas):
        self._sincronizar_painel_com_dados()
        self.resultados_atuais.extend(novos_resultados)
        self._redesenhar_painel_completo()
        self.exibir_relatorio_de_revisao(self.resultados_atuais, imagens_processadas)
        self.label_status.configure(text=f"{len(self.resultados_atuais)} chaves carregadas. Prontas para revisão.")
        self.progressbar.grid_forget()
        self._bloquear_botoes(False)

    def _sincronizar_painel_com_dados(self):
        """Salva os valores atuais dos campos de entrada da UI na lista de dados 'self.resultados_atuais'."""
        for widget_info in self.widgets_linhas:
            idx = widget_info['index']
            if idx < len(self.resultados_atuais):
                valor_atual_entry = widget_info['entry'].get()
                self.resultados_atuais[idx]['Chave'] = valor_atual_entry

    def _redesenhar_painel_completo(self):
        """Limpa e redesenha todo o painel de chaves a partir da lista de dados."""
        for widget_info in self.widgets_linhas:
            widget_info['frame'].destroy()
        self.widgets_linhas.clear()
        
        for i, resultado in enumerate(self.resultados_atuais):
            self._desenhar_linha_chave(i, resultado)

    def exibir_relatorio_de_revisao(self, resultados, imagens_processadas):
        chaves_para_revisao = [res for res in resultados if res.get('Diagnostico', '').lower() not in ['nenhum', 'não informado', '']]
        relatorio = f"--- DIAGNÓSTICO DA IA ---\n"
        relatorio += f"Imagens Processadas nesta Rodada: {len(imagens_processadas)} | Total de Chaves no Painel: {len(resultados)}\n\n"
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

    def _desenhar_linha_chave(self, index, dados_linha):
        """Apenas desenha uma linha na interface com base nos dados fornecidos."""
        imagem = dados_linha.get('Imagem', 'MANUAL')
        chave = dados_linha.get('Chave', '')
        diagnostico = dados_linha.get('Diagnostico')

        frame_linha = ctk.CTkFrame(self.frame_rolavel); frame_linha.pack(fill="x", expand=True, padx=5, pady=2)
        
        # MUDANÇA: Adiciona um número de linha fixo para visualização
        label_numero_linha = ctk.CTkLabel(frame_linha, text=f"{index + 1}.", width=30, text_color="gray"); 
        label_numero_linha.pack(side="left", padx=(5, 0))

        cor_diagnostico, texto_diagnostico = "gray", "N/A"
        if diagnostico:
            if diagnostico.lower() == 'nenhum':
                cor_diagnostico, texto_diagnostico = "light green", "✅ Alta Confiança"
            elif diagnostico.lower() != 'não informado':
                cor_diagnostico, texto_diagnostico = "orange", f"⚠️ Revisar: {diagnostico}"

        label_imagem = ctk.CTkLabel(frame_linha, text=imagem, width=120, anchor="w"); label_imagem.pack(side="left", padx=5)
        
        string_var = StringVar()
        string_var.set(chave)
        def to_upper(*args):
            string_var.set(string_var.get().upper())
        string_var.trace("w", to_upper)
        entry_chave = ctk.CTkEntry(frame_linha, textvariable=string_var); entry_chave.pack(side="left", fill="x", expand=True, padx=5)
        
        label_diagnostico = ctk.CTkLabel(frame_linha, text=texto_diagnostico, text_color=cor_diagnostico, width=150); label_diagnostico.pack(side="left", padx=5)
        
        frame_acoes = ctk.CTkFrame(frame_linha, fg_color="transparent"); frame_acoes.pack(side="left", padx=5)
        
        botao_subir = ctk.CTkButton(frame_acoes, text="↑", width=30, command=lambda idx=index: self.mover_linha(idx, -1)); botao_subir.pack(side="left", padx=(0,2))
        botao_descer = ctk.CTkButton(frame_acoes, text="↓", width=30, command=lambda idx=index: self.mover_linha(idx, 1)); botao_descer.pack(side="left", padx=(0,5))
        botao_excluir = ctk.CTkButton(frame_acoes, text="Excluir", width=60, fg_color="firebrick", hover_color="darkred", command=lambda idx=index: self.excluir_linha_chave(idx)); botao_excluir.pack(side="left")
        
        self.widgets_linhas.append({'frame': frame_linha, 'label': label_imagem, 'entry': entry_chave, 'index': index})

    def handle_adicionar_manual(self):
        """Sincroniza os dados, adiciona uma nova entrada e redesenha a interface."""
        self._sincronizar_painel_com_dados()
        self.resultados_atuais.append({'Imagem': 'MANUAL', 'Chave': '', 'Diagnostico': 'Não informado'})
        self._redesenhar_painel_completo()

    def mover_linha(self, index_atual, direcao):
        """Sincroniza os dados, move um item na lista e redesenha a interface."""
        self._sincronizar_painel_com_dados()
        novo_index = index_atual + direcao
        if 0 <= novo_index < len(self.resultados_atuais):
            self.resultados_atuais[index_atual], self.resultados_atuais[novo_index] = self.resultados_atuais[novo_index], self.resultados_atuais[index_atual]
            self._redesenhar_painel_completo()

    def excluir_linha_chave(self, index_para_remover):
        """Sincroniza os dados, remove um item da lista e redesenha a interface."""
        self._sincronizar_painel_com_dados()
        if 0 <= index_para_remover < len(self.resultados_atuais):
            self.resultados_atuais.pop(index_para_remover)
            self._redesenhar_painel_completo()
            
    def limpar_tudo(self):
        self.resultados_atuais.clear()
        self._redesenhar_painel_completo()
        self.textbox_analise.configure(state="normal")
        self.textbox_analise.delete("1.0", "end")
        self.textbox_analise.configure(state="disabled")
        self.label_status.configure(text="Painel limpo. Pronto para começar.")

    def _coletar_dados_do_painel(self):
        """Coleta e VALIDA os dados da lista 'self.resultados_atuais' para salvar."""
        self._sincronizar_painel_com_dados() 
        
        if not self.resultados_atuais:
            messagebox.showwarning("Aviso", "Não há chaves no painel para salvar.")
            return None
        
        for i, r in enumerate(self.resultados_atuais):
            if not r['Chave'].strip():
                messagebox.showerror("Erro de Validação", 
                                     f"A chave na linha {i+1} está vazia.\n\n"
                                     "Por favor, preencha todas as chaves ou exclua as linhas em branco antes de salvar.")
                return None
            
        dados_para_salvar = [{'Imagem': r['Imagem'], 'Chave': r['Chave']} for r in self.resultados_atuais]
        return pd.DataFrame(dados_para_salvar)

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
                df_final.drop_duplicates(subset=['Chave'], keep='last', inplace=True)
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
