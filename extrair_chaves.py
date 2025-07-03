import os
import sys
import pandas as pd
import threading
import customtkinter as ctk
from tkinter import filedialog, messagebox, StringVar
from PIL import Image
from dotenv import load_dotenv
import google.generativeai as genai
import re


load_dotenv()
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
NOME_ARQUIVO_EXCEL = 'chaves_extraidas_final.xlsx'



def resource_path(relative_path):
    """ Retorna o caminho absoluto para o recurso, funciona para desenvolvimento e para o executável do PyInstaller """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


def validar_formato_chave(chave: str) -> bool:
    """Verifica se a chave segue o padrão XXXXX-XXXXX-XXXXX-XXXXX-XXXXX."""
    padrao = re.compile(r'^[A-Z0-9]{5}-[A-Z0-9]{5}-[A-Z0-9]{5}-[A-Z0-9]{5}-[A-Z0-9]{5}$')
    return bool(padrao.match(chave))


def extrair_chaves_da_imagem(caminho_imagem: str) -> list:
    """Usa um prompt simplificado para extrair apenas as chaves da imagem."""
    nome_arquivo = os.path.basename(caminho_imagem)
    print(f"\n🖼️  Extraindo chaves de '{nome_arquivo}'...")
    try:
        genai.configure(api_key=GEMINI_API_KEY)
        model = genai.GenerativeModel('gemini-1.5-flash-latest')
        img = Image.open(caminho_imagem)
        
        prompt = """
        Sua tarefa é transcrever TODAS as chaves de produto da imagem com a maior precisão possível.
        Liste cada chave encontrada em uma nova linha, começando com "CHAVE: ".
        
        Exemplo:
        CHAVE: GR79F-V4NGQ-RBGQK-X4RVV-PWF9C
        CHAVE: NCKM6-93VT7-D64WF-2X9VK-MG9TT
        """
        
        response = model.generate_content([prompt, img])
        print(f"    ✅ Texto recebido com sucesso.")

        if not response.text: return []
        
        resultados = []
        linhas = response.text.strip().split('\n')
        for linha in linhas:
            if linha.upper().startswith("CHAVE:"):
                chave = linha[len("CHAVE:"):].strip().upper()
                resultados.append({'Imagem': nome_arquivo, 'Chave': chave})
        return resultados

    except Exception as e:
        print(f"    ❌ Erro na chamada da API Gemini para '{nome_arquivo}': {e}")
        return []



class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        
        self.title("SysKey - Painel de Extração de Chaves")
        self.geometry("1100x700")
        ctk.set_appearance_mode("dark")
        
        try:
            self.iconbitmap(resource_path("icone.ico"))
        except Exception as e:
            print(f"Aviso: Não foi possível carregar o ícone da janela: {e}")

        
        self.resultados_atuais = []
        self.widgets_linhas = []
        self.title_font = ctk.CTkFont(family="Arial", size=18, weight="bold")
        self.main_font = ctk.CTkFont(family="Arial", size=12)
        self.status_font = ctk.CTkFont(family="Arial", size=11)

        
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        
        self.frame_superior = ctk.CTkFrame(self, corner_radius=0)
        self.frame_superior.grid(row=0, column=0, padx=10, pady=(10, 5), sticky="ew")
        self.frame_superior.grid_columnconfigure(3, weight=1) 

        self.botao_carregar_arquivo = ctk.CTkButton(self.frame_superior, text="Carregar Arquivo", command=self.iniciar_extracao_arquivo)
        self.botao_carregar_arquivo.grid(row=0, column=0, padx=5, pady=5)
        self.botao_carregar_pasta = ctk.CTkButton(self.frame_superior, text="Carregar Pasta", command=self.iniciar_extracao_pasta)
        self.botao_carregar_pasta.grid(row=0, column=1, padx=5, pady=5)
        self.botao_limpar = ctk.CTkButton(self.frame_superior, text="Limpar Painel", command=self.limpar_tudo, fg_color="#585858", hover_color="#404040")
        self.botao_limpar.grid(row=0, column=4, padx=5, pady=5)

        
        self.frame_rolavel = ctk.CTkScrollableFrame(self, label_text="Chaves Extraídas", label_font=self.title_font)
        self.frame_rolavel.grid(row=1, column=0, padx=10, pady=5, sticky="nsew")

        
        self.frame_inferior = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.frame_inferior.grid(row=2, column=0, padx=10, pady=5, sticky="ew")
        self.frame_inferior.grid_columnconfigure(0, weight=1)
        self.frame_inferior.grid_columnconfigure(1, weight=1)

        self.botao_adicionar = ctk.CTkButton(self.frame_inferior, text="Adicionar Chave Manualmente", command=self.handle_adicionar_manual)
        self.botao_adicionar.grid(row=0, column=0, columnspan=2, padx=0, pady=5, sticky="ew")
        
        self.frame_analise = ctk.CTkFrame(self.frame_inferior)
        self.frame_analise.grid(row=1, column=0, padx=(0, 5), pady=5, sticky="nsew")
        self.frame_analise.grid_rowconfigure(0, weight=1)
        self.frame_analise.grid_columnconfigure(0, weight=1)
        self.textbox_analise = ctk.CTkTextbox(self.frame_analise, wrap="word", state="disabled", font=self.main_font, border_spacing=5)
        self.textbox_analise.grid(row=0, column=0, sticky="nsew")

        self.frame_status_salvar = ctk.CTkFrame(self.frame_inferior)
        self.frame_status_salvar.grid(row=1, column=1, padx=(5, 0), pady=5, sticky="nsew")
        self.frame_status_salvar.grid_columnconfigure(0, weight=1)

        self.frame_status_progresso = ctk.CTkFrame(self.frame_status_salvar, fg_color="transparent")
        self.frame_status_progresso.pack(fill="x", expand=True, padx=5, pady=5)
        self.frame_status_progresso.grid_columnconfigure(0, weight=1)
        self.label_status = ctk.CTkLabel(self.frame_status_progresso, text="Pronto. Selecione um arquivo ou pasta para começar.", text_color="gray", font=self.status_font)
        self.label_status.grid(row=0, column=0, sticky="w")
        self.progressbar = ctk.CTkProgressBar(self.frame_status_progresso, orientation="horizontal")
        
        self.frame_salvar = ctk.CTkFrame(self.frame_status_salvar, fg_color="transparent")
        self.frame_salvar.pack(fill="x", expand=True, side="bottom", padx=5, pady=10)
        self.frame_salvar.grid_columnconfigure((0,1), weight=1)
        self.botao_adicionar_excel = ctk.CTkButton(self.frame_salvar, text="Adicionar ao Excel", fg_color="sea green", hover_color="dark green", command=self.adicionar_ao_excel)
        self.botao_adicionar_excel.grid(row=0, column=0, padx=(0,5), pady=5, sticky="ew")
        self.botao_substituir_excel = ctk.CTkButton(self.frame_salvar, text="Substituir a Planilha", command=self.substituir_em_excel)
        self.botao_substituir_excel.grid(row=0, column=1, padx=(5,0), pady=5, sticky="ew")
        
    def _bloquear_botoes(self, processando=True):
        estado = "disabled" if processando else "normal"
        self.botao_carregar_arquivo.configure(state=estado)
        self.botao_carregar_pasta.configure(state=estado)
        self.botao_adicionar_excel.configure(state=estado)
        self.botao_substituir_excel.configure(state=estado)
        self.botao_adicionar.configure(state=estado)
        self.botao_limpar.configure(state=estado)
        
    def iniciar_extracao_base(self, target_func, target_arg):
        if not target_arg: return
        self.label_status.configure(text=f"Processando '{os.path.basename(str(target_arg))}'...")
        self.progressbar.grid(row=0, column=1, padx=10, sticky="ew")
        self.progressbar.set(0)
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
        novos_resultados = extrair_chaves_da_imagem(caminho_arquivo)
        self.progressbar.set(1.0)
        self.after(0, self.atualizar_interface_completa, novos_resultados, [caminho_arquivo])
        
    def _processar_pasta_em_background(self, caminho_pasta):
        imagens = [f for f in os.listdir(caminho_pasta) if f.lower().endswith(('.png', '.jpg', '.jpeg'))]
        total_imagens = len(imagens)
        if total_imagens == 0: self.after(0, self.atualizar_interface_completa, [], []); return
        
        resultados_finais = []
        for i, nome_img in enumerate(imagens):
            caminho_completo = os.path.join(caminho_pasta, nome_img)
            resultados_finais.extend(extrair_chaves_da_imagem(caminho_completo))
            self.progressbar.set((i + 1) / total_imagens)
        self.after(0, self.atualizar_interface_completa, resultados_finais, imagens)

    def atualizar_interface_completa(self, novos_resultados, imagens_processadas):
        self._sincronizar_painel_com_dados()
        self.resultados_atuais.extend(novos_resultados)
        self._redesenhar_painel_completo()
        self.label_status.configure(text=f"{len(self.resultados_atuais)} chaves carregadas. {len(imagens_processadas)} imagens processadas.")
        self.progressbar.grid_forget()
        self._bloquear_botoes(False)

    def _sincronizar_painel_com_dados(self):
        for widget_info in self.widgets_linhas:
            idx = widget_info['index']
            if idx < len(self.resultados_atuais):
                valor_atual_entry = widget_info['entry'].get()
                self.resultados_atuais[idx]['Chave'] = valor_atual_entry

    def _redesenhar_painel_completo(self):
        for widget_info in self.widgets_linhas:
            widget_info['frame'].destroy()
        self.widgets_linhas.clear()
        
        for i, resultado in enumerate(self.resultados_atuais):
            self._desenhar_linha_chave(i, resultado)
        
        self._atualizar_relatorio_validacao()

    def _atualizar_relatorio_validacao(self):
        self._sincronizar_painel_com_dados()
        chaves_para_revisao = []
        for i, res in enumerate(self.resultados_atuais):
            if not validar_formato_chave(res.get('Chave', '')):
                chaves_para_revisao.append((i + 1, res))

        total_chaves = len(self.resultados_atuais)
        
        relatorio = f"--- RELATÓRIO DE VALIDAÇÃO ---\n"
        relatorio += f"Total de Chaves no Painel: {total_chaves}\n\n"
        if chaves_para_revisao:
            relatorio += f"🚨 {len(chaves_para_revisao)} CHAVE(S) COM FORMATO INVÁLIDO PARA REVISÃO:\n"
            for linha_num, item in chaves_para_revisao:
                relatorio += f"- Linha {linha_num}: '{item.get('Chave', '')}' (Imagem: {item.get('Imagem', 'N/A')})\n"
        else:
            relatorio += "✅ Formato Válido: Todas as chaves no painel seguem o padrão correto."
        
        self.textbox_analise.configure(state="normal")
        self.textbox_analise.delete("1.0", "end")
        self.textbox_analise.insert("1.0", relatorio)
        self.textbox_analise.configure(state="disabled")

    def _desenhar_linha_chave(self, index, dados_linha):
        imagem = dados_linha.get('Imagem', 'MANUAL')
        chave = dados_linha.get('Chave', '')

        frame_linha = ctk.CTkFrame(self.frame_rolavel, fg_color=("gray85", "gray20"))
        frame_linha.pack(fill="x", expand=True, padx=5, pady=2)
        
        def on_enter(e): frame_linha.configure(fg_color=("gray80", "gray25"))
        def on_leave(e): frame_linha.configure(fg_color=("gray85", "gray20"))
        frame_linha.bind("<Enter>", on_enter)
        frame_linha.bind("<Leave>", on_leave)

        label_numero_linha = ctk.CTkLabel(frame_linha, text=f"{index + 1}.", width=30, text_color="gray", font=self.main_font)
        label_numero_linha.pack(side="left", padx=(10, 5), pady=5)

        formato_valido = validar_formato_chave(chave)
        cor_diagnostico = "light green" if formato_valido else "orange"
        texto_diagnostico = "✅ Padrão Correto" if formato_valido else "⚠️ Verificar Formato"

        label_imagem = ctk.CTkLabel(frame_linha, text=imagem, width=120, anchor="w", font=self.main_font)
        label_imagem.pack(side="left", padx=5, pady=5)
        
        string_var = StringVar()
        string_var.set(chave)
        
        label_diagnostico = ctk.CTkLabel(frame_linha, text=texto_diagnostico, text_color=cor_diagnostico, width=150, anchor="w", font=self.status_font)
        
        def on_key_change(*args):
            texto_atual = string_var.get()
            string_var.set(texto_atual.upper())
            formato_valido_agora = validar_formato_chave(string_var.get())
            if formato_valido_agora:
                label_diagnostico.configure(text="✅ Padrão Correto", text_color="light green")
            else:
                label_diagnostico.configure(text="⚠️ Verificar Formato", text_color="orange")
            self.after(50, self._atualizar_relatorio_validacao)

        string_var.trace("w", on_key_change)
        
        entry_chave = ctk.CTkEntry(frame_linha, textvariable=string_var, font=self.main_font)
        entry_chave.pack(side="left", fill="x", expand=True, padx=5, pady=5)
        
        label_diagnostico.pack(side="left", padx=5, pady=5)
        
        frame_acoes = ctk.CTkFrame(frame_linha, fg_color="transparent")
        frame_acoes.pack(side="left", padx=5, pady=5)
        
        botao_subir = ctk.CTkButton(frame_acoes, text="↑", width=30, command=lambda idx=index: self.mover_linha(idx, -1))
        botao_subir.pack(side="left", padx=(0,2))
        botao_descer = ctk.CTkButton(frame_acoes, text="↓", width=30, command=lambda idx=index: self.mover_linha(idx, 1))
        botao_descer.pack(side="left", padx=(0,5))
        botao_excluir = ctk.CTkButton(frame_acoes, text="Excluir", width=60, fg_color="firebrick", hover_color="darkred", command=lambda idx=index: self.excluir_linha_chave(idx))
        botao_excluir.pack(side="left")
        
        self.widgets_linhas.append({'frame': frame_linha, 'label': label_imagem, 'entry': entry_chave, 'index': index})

    def handle_adicionar_manual(self):
        self._sincronizar_painel_com_dados()
        self.resultados_atuais.append({'Imagem': 'MANUAL', 'Chave': ''})
        self._redesenhar_painel_completo()

    def mover_linha(self, index_atual, direcao):
        self._sincronizar_painel_com_dados()
        novo_index = index_atual + direcao
        if 0 <= novo_index < len(self.resultados_atuais):
            self.resultados_atuais[index_atual], self.resultados_atuais[novo_index] = self.resultados_atuais[novo_index], self.resultados_atuais[index_atual]
            self._redesenhar_painel_completo()

    def excluir_linha_chave(self, index_para_remover):
        self._sincronizar_painel_com_dados()
        if 0 <= index_para_remover < len(self.resultados_atuais):
            self.resultados_atuais.pop(index_para_remover)
            self._redesenhar_painel_completo()
            
    def limpar_tudo(self):
        self.resultados_atuais.clear()
        self._redesenhar_painel_completo()
        self.label_status.configure(text="Painel limpo. Pronto para começar.")

    def _coletar_dados_do_painel(self):
        self._sincronizar_painel_com_dados() 
        
        if not self.resultados_atuais:
            messagebox.showwarning("Aviso", "Não há chaves no painel para salvar.")
            return None
        
        for i, r in enumerate(self.resultados_atuais):
            if not validar_formato_chave(r.get('Chave', '')):
                messagebox.showerror("Erro de Validação", 
                                     f"A chave na linha {i+1} não segue o formato padrão (XXXXX-XXXXX-XXXXX-XXXXX-XXXXX).\n\n"
                                     "Por favor, corrija a chave ou exclua a linha antes de salvar.")
                return None
            
        dados_para_salvar = [{'Imagem': r['Imagem'], 'Chave': r['Chave']} for r in self.resultados_atuais]
        return pd.DataFrame(dados_para_salvar)

    def substituir_em_excel(self):
        df_novo = self._coletar_dados_do_painel();
        if df_novo is None: return
        caminho_arquivo_final = os.path.join(os.getcwd(), NOME_ARQUIVO_EXCEL)
        try:
            
            df_novo.to_excel(caminho_arquivo_final, index=False, header=False)
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
                
                df_antigo = pd.read_excel(caminho_arquivo_final, header=None, names=['Imagem', 'Chave'])
                df_final = pd.concat([df_antigo, df_novo], ignore_index=True)
                df_final.drop_duplicates(subset=['Chave'], keep='last', inplace=True)
            else:
                df_final = df_novo
            
            df_final.to_excel(caminho_arquivo_final, index=False, header=False)
            messagebox.showinfo("Sucesso", f"As chaves foram adicionadas à planilha!\n\nO arquivo será aberto a seguir.")
            if os.path.exists(caminho_arquivo_final): os.startfile(caminho_arquivo_final)
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro ao salvar o arquivo:\n{e}")

if __name__ == "__main__":
    app = App()
    app.mainloop()
