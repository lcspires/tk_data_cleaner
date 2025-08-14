"""

Tkinter Application for Preparing and Converting Excel (.xls, .xlsx), CSV, and TXT Files into Formatted TXT Files

Main Features:

- Loads Excel, CSV, or TXT files.
- Allows selection and organization of the loaded file’s columns.
- Removes excess whitespace from all selected columns.
- Removes rows where the first column has fewer characters than a user-defined minimum.
- Removes duplicates based on the first column.
- Exports the final result as a semicolon-delimited TXT file.

Imported Modules:

- pandas for tabular data manipulation.
- tkinter for the graphical user interface.
- logic.py containing functions for data cleaning and filtering.
- tooltip.py for displaying help information in the interface.

Author: Lucas Ferreira
Date: 14/08/2025
"""

import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os
from logic import limpar_espacos_em_colunas, remover_duplicatas, filtrar_primeira_coluna_por_tamanho
from tooltip import ToolTip

def tentar_leitura_csv_variavel_sep(caminho_arquivo):
    """
    Attempts to read a CSV or TXT file using common delimiters (;, ,, \t).
    If the file has encoding issues, it tries both utf-8 and latin1.
    Returns a pandas DataFrame if successful, or raises an exception otherwise.
    
    Args: caminho_arquivo (str): Path to the CSV or TXT file.

    Returns: pd.DataFrame: Data read from the file.
    """
    sep_list = [';', ',', '\t']
    for sep in sep_list:
        try:
            return pd.read_csv(caminho_arquivo, sep=sep, quotechar='"', encoding='utf-8', on_bad_lines='skip')
        except UnicodeDecodeError:
            try:
                return pd.read_csv(caminho_arquivo, sep=sep, quotechar='"', encoding='latin1', on_bad_lines='skip')
            except Exception:
                continue
        except Exception:
            continue
    raise Exception("Não foi possível ler o arquivo com os separadores padrão")

class UnifiedApp:
    """
    Main application class for Tkinter-based preparation and conversion of Excel, CSV, or TXT files.
    """
    def __init__(self, root):
        """
        Initializes the application's interface and variables.
        
        Args: root (tk.Tk): Instance of the Tkinter main window.
        """
        self.root = root
        self.df = None
        self.total_espacos_removidos = 0
        self.linhas_removidas_por_tamanho = 0
        self.total_duplicatas_removidas = 0
        self._build_ui()

    def _build_ui(self):
        """
        Builds the application's graphical interface, including buttons, column list,
        field for the minimum number of digits in the first column, and an informational tooltip.
        """
        self.root.title("Preparador XLS/CSV → TXT")
        self.root.geometry("600x550")

        btn_select = tk.Button(self.root, text="Selecionar arquivo", command=self.selecionar_arquivo, font=("Arial", 12))
        btn_select.pack(pady=10)

        # Campo para definir mínimo de dígitos logo abaixo do botão
        frame_min_len = tk.Frame(self.root)
        frame_min_len.pack(pady=(0, 10))
        tk.Label(frame_min_len, text="Mínimo de dígitos na 1ª coluna:").pack(side=tk.LEFT)
        self.spin_min_len = tk.Spinbox(frame_min_len, from_=1, to=100, width=5)
        self.spin_min_len.pack(side=tk.LEFT, padx=5)
        self.spin_min_len.delete(0, "end")
        self.spin_min_len.insert(0, "6")  # padrão 6

        self.label_resultado = tk.Label(self.root, text="", font=("Arial", 10), wraplength=580, justify="center")
        self.label_resultado.pack(pady=(0, 5))

        self.icone_info = tk.Label(self.root, text="ℹ️", font=("Arial", 14), cursor="question_arrow")
        self.icone_info.pack_forget()

        texto_tooltip = (
            "Selecione um arquivo Excel (.xls, .xlsx), CSV ou TXT.\n"
            "Após o carregamento, organize as colunas e gere o arquivo TXT.\n"
            "O programa remove espaços extras e exclui linhas cuja primeira coluna "
            "tenha menos caracteres que o número definido acima."
        )
        self.tooltip = ToolTip(self.icone_info, texto_tooltip)

        self.listbox_colunas = tk.Listbox(self.root, selectmode=tk.EXTENDED, width=50, height=15)
        self.listbox_colunas.pack(pady=10)

        frame_botoes = tk.Frame(self.root)
        frame_botoes.pack()

        tk.Button(frame_botoes, text="↑ Mover para cima", command=self.mover_cima).grid(row=0, column=0, padx=10, pady=5)
        tk.Button(frame_botoes, text="↓ Mover para baixo", command=self.mover_baixo).grid(row=0, column=1, padx=10, pady=5)
        tk.Button(frame_botoes, text="✖ Remover selecionadas", command=self.remover_colunas, fg="red").grid(row=0, column=2, padx=10, pady=5)

        btn_generate = tk.Button(self.root, text="Gerar arquivo TXT", command=self.gerar_txt, font=("Arial", 12), bg="#4CAF50", fg="white")
        btn_generate.pack(pady=20)

    def selecionar_arquivo(self):
        """
        Opens a dialog for selecting an Excel, CSV, or TXT file.
        Performs the appropriate reading method depending on the file extension.
        Updates the column list displayed in the interface.
        """
        caminho_arquivo = filedialog.askopenfilename(
            title="Selecione um arquivo Excel, CSV ou TXT",
            filetypes=[
                ("Arquivos Excel", "*.xls *.xlsx"),
                ("Arquivos CSV", "*.csv"),
                ("Arquivos TXT", "*.txt"),
                ("Todos os arquivos", "*.*")
            ]
        )
        if not caminho_arquivo:
            self.label_resultado.config(text="Nenhum arquivo selecionado.")
            self.icone_info.pack_forget()
            return

        ext = os.path.splitext(caminho_arquivo)[1].lower()

        try:
            if ext in ['.xls', '.xlsx']:
                self.df = pd.read_excel(caminho_arquivo)
            elif ext in ['.csv', '.txt']:
                self.df = tentar_leitura_csv_variavel_sep(caminho_arquivo)
            else:
                raise Exception("Tipo de arquivo não suportado.")
            self.label_resultado.config(text=f"Arquivo carregado: {caminho_arquivo}")
            self._carregar_colunas(self.df.columns.tolist())
            self.icone_info.pack(pady=(0, 10))
            self.total_espacos_removidos = 0
            self.linhas_removidas_por_tamanho = 0
            self.total_duplicatas_removidas = 0
            self._atualizar_tooltip()
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao ler o arquivo:\n{e}")

    def _carregar_colunas(self, colunas):
        """
        Updates the Listbox to display the available columns for selection.
        
        Args: colunas (list of str): List of column names.
        """
        self.listbox_colunas.delete(0, tk.END)
        for col in colunas:
            self.listbox_colunas.insert(tk.END, col)

    def mover_cima(self):
        """
        Moves the selected columns up in the list, allowing reordering.
        """
        selecionados = self.listbox_colunas.curselection()
        for index in selecionados:
            if index == 0:
                continue
            texto = self.listbox_colunas.get(index)
            acima = self.listbox_colunas.get(index - 1)
            self.listbox_colunas.delete(index - 1, index)
            self.listbox_colunas.insert(index - 1, texto)
            self.listbox_colunas.insert(index, acima)
            self.listbox_colunas.selection_set(index - 1)
            self.listbox_colunas.selection_clear(index)

    def mover_baixo(self):
        """
        Moves the selected columns down in the list, allowing reordering.
        """
        selecionados = self.listbox_colunas.curselection()
        count = self.listbox_colunas.size()
        for index in reversed(selecionados):
            if index == count - 1:
                continue
            texto = self.listbox_colunas.get(index)
            abaixo = self.listbox_colunas.get(index + 1)
            self.listbox_colunas.delete(index, index + 1)
            self.listbox_colunas.insert(index, abaixo)
            self.listbox_colunas.insert(index + 1, texto)
            self.listbox_colunas.selection_set(index + 1)
            self.listbox_colunas.selection_clear(index)

    def remover_colunas(self):
        """
        Removes the selected columns from the list of columns to be processed.
        """
        selecionados = self.listbox_colunas.curselection()
        for index in reversed(selecionados):
            self.listbox_colunas.delete(index)

    def _atualizar_tooltip(self):
        """
        Updates the tooltip text to display a summary of the operations performed,
        such as the number of spaces removed, rows discarded due to length,
        and duplicates removed.
        """
        texto_base = "Limpeza de dados:"
        texto_espacos = f"\nEspaçamentos removidos: {self.total_espacos_removidos}"
        texto_tamanho = f"\nLinhas removidas (menores que {self.spin_min_len.get()} digitos): {self.linhas_removidas_por_tamanho}"
        texto_duplicatas = f"\nDuplicatas removidas: {self.total_duplicatas_removidas}"
        self.tooltip.texto = texto_base + texto_espacos + texto_tamanho + texto_duplicatas

    def gerar_txt(self):
        """
        Performs data cleaning, filtering by minimum length, and duplicate removal,
        and saves the resulting DataFrame as a semicolon-delimited TXT file.
        Displays success or error messages to the user.
        """
        if self.df is None:
            messagebox.showwarning("Aviso", "Nenhum arquivo carregado.")
            return

        colunas_selecionadas = list(self.listbox_colunas.get(0, tk.END))
        if not colunas_selecionadas:
            messagebox.showwarning("Aviso", "Selecione pelo menos uma coluna.")
            return

        try:
            min_len = int(self.spin_min_len.get())
            if min_len < 1:
                raise ValueError
        except ValueError:
            messagebox.showwarning("Aviso", "Digite um número inteiro válido para o mínimo de caracteres.")
            return

        primeira_col = colunas_selecionadas[0]

        # Cleans whitespace in all selected columns.
        df_filtrado, self.total_espacos_removidos = limpar_espacos_em_colunas(
            self.df[colunas_selecionadas].copy(), colunas_selecionadas
        )

        # Filters rows where the first column has fewer than min_len characters.
        df_filtrado, self.linhas_removidas_por_tamanho = filtrar_primeira_coluna_por_tamanho(
            df_filtrado, primeira_col, min_len=min_len
        )

        # Removes duplicates based on the first column.
        df_filtrado, self.total_duplicatas_removidas = remover_duplicatas(df_filtrado, primeira_col)

        self._atualizar_tooltip()

        caminho_save = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Arquivo TXT", "*.txt")],
            title="Salvar arquivo TXT"
        )
        if not caminho_save:
            return

        try:
            df_filtrado.to_csv(caminho_save, sep=';', index=False, encoding='utf-8-sig')
            messagebox.showinfo(
                "Sucesso",
                f"Arquivo salvo em:\n{caminho_save}\n"
                f"Linhas removidas por tamanho: {self.linhas_removidas_por_tamanho}"
            )
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao salvar o arquivo:\n{e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = UnifiedApp(root)
    root.mainloop()
