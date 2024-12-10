import customtkinter as ctk
import os
from PIL import Image
from tkinter import filedialog, messagebox

class ExcelFilter:
    def __init__(self):
        self.df = None
        self.headers = []
        
    def load_excel(self, filepath):
        try:
            import pandas as pd
            self.df = pd.read_excel(filepath)
            self.headers = self.df.columns.tolist()
            return True
        except:
            return False
            
    def get_unique_values(self, column):
        if self.df is not None and column in self.df.columns:
            return self.df[column].unique().tolist()
        return []
        
    def filter_and_save(self, column, value, output_dir):
        if self.df is not None:
            filtered_df = self.df[self.df[column] == value]
            output_file = os.path.join(output_dir, f"filtered_{column}_{value}.xlsx")
            filtered_df.to_excel(output_file, index=False)
            return output_file
        return None

class ExcelFilterGUI:
    def __init__(self):
        # Configurações do tema
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("green")
        
        # Configuração da janela principal
        self.root = ctk.CTk()
        self.root.title("Excel Filter Pro")
        self.root.geometry("1000x600")
        
        # Cores
        self.excel_green = "#217346"
        self.light_gray = "#F5F5F5"
        self.dark_gray = "#333333"
        
        # Carregar imagens
        self.load_images()
        
        # Configurar ícone da janela
        self.set_window_icon()
        
        self.create_gui()
        self.excel_filter = ExcelFilter()
    
    def load_images(self):
        # Diretório base de assets
        assets_dir = os.path.join(os.path.dirname(__file__), "assets")
        
        # Carregar logo
        self.logo = ctk.CTkImage(
            light_image=Image.open(os.path.join(assets_dir, "logo.png")),
            dark_image=Image.open(os.path.join(assets_dir, "logo.png")),
            size=(48, 48)  # Tamanho do logo na interface
        )
    
    def set_window_icon(self):
        # Carregar ícone da janela
        icon_path = os.path.join(
            os.path.dirname(__file__),
            "assets",
            "icon.png"
        )
        # Converter para formato PhotoImage
        from tkinter import PhotoImage
        icon = PhotoImage(file=icon_path)
        self.root.iconphoto(True, icon)
        
    def create_gui(self):
        # Frame principal
        self.main_frame = ctk.CTkFrame(self.root, fg_color="transparent")
        self.main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Header Frame
        header_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        header_frame.pack(fill="x", pady=(0, 20))
        
        # Logo
        logo_label = ctk.CTkLabel(
            header_frame,
            image=self.logo,
            text=""
        )
        logo_label.pack(side="left", padx=10)
        
        # Título
        title_label = ctk.CTkLabel(
            header_frame,
            text="Excel Filter Pro",
            font=ctk.CTkFont(size=24, weight="bold"),
            text_color=self.excel_green
        )
        title_label.pack(side="left", padx=10)
        
        # Frame para o arquivo
        self.file_frame = ctk.CTkFrame(self.main_frame)
        self.file_frame.pack(fill="x", padx=20, pady=(0, 20))
        
        self.file_entry = ctk.CTkEntry(
            self.file_frame,
            placeholder_text="Selecione um arquivo Excel...",
            width=400
        )
        self.file_entry.pack(side="left", padx=(10, 10), pady=10)
        
        self.browse_button = ctk.CTkButton(
            self.file_frame,
            text="Procurar",
            command=self.browse_file,
            width=100
        )
        self.browse_button.pack(side="left", padx=5)
        
        # Frame para as operações
        self.operations_frame = ctk.CTkFrame(self.main_frame)
        self.operations_frame.pack(fill="both", expand=True, padx=20)
        
        # Botões de operação
        operations = [
            ("Filtrar (Único)", self.filter_single),
            ("Filtrar (Múltiplo)", self.filter_multiple),
            ("Manter Colunas", self.keep_columns),
            ("Remover Colunas", self.remove_columns)
        ]
        
        for text, command in operations:
            btn = ctk.CTkButton(
                self.operations_frame,
                text=text,
                command=command,
                height=40,
                font=ctk.CTkFont(size=14)
            )
            btn.pack(pady=10, padx=20, fill="x")
            
    def browse_file(self):
        filename = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if filename:
            self.file_entry.delete(0, "end")
            self.file_entry.insert(0, filename)
            
    def get_save_path(self):
        return filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        
    def filter_single(self):
        if not self.load_file():
            return
            
        # Criar nova janela para filtros
        filter_window = ctk.CTkToplevel(self.root)
        filter_window.title("Filtro Único")
        filter_window.geometry("400x300")
        
        # Combobox para selecionar coluna
        column_var = ctk.StringVar()
        column_combo = ctk.CTkComboBox(
            filter_window,
            values=self.excel_filter.headers,
            variable=column_var,
            width=300
        )
        column_combo.pack(pady=20)
        
        # Quando uma coluna é selecionada, atualizar valores únicos
        def on_column_select(*args):
            values = self.excel_filter.get_unique_values(column_var.get())
            value_combo.configure(values=values)
            
        column_var.trace('w', on_column_select)
        
        # Combobox para selecionar valor
        value_combo = ctk.CTkComboBox(
            filter_window,
            values=[],
            width=300
        )
        value_combo.pack(pady=20)
        
        # Botão para aplicar filtro
        ctk.CTkButton(
            filter_window,
            text="Aplicar Filtro",
            command=lambda: self.apply_single_filter(
                column_var.get(),
                value_combo.get(),
                filter_window
            )
        ).pack(pady=20)

    def apply_single_filter(self, column, value, window):
        save_path = self.get_save_path()
        if save_path:
            output_file = self.excel_filter.filter_and_save(
                column, value, os.path.dirname(save_path)
            )
            messagebox.showinfo(
                "Sucesso",
                f"Arquivo filtrado salvo em:\n{output_file}"
            )
            window.destroy()
            
    def load_file(self):
        filepath = self.file_entry.get()
        if not filepath:
            messagebox.showerror("Erro", "Selecione um arquivo Excel primeiro!")
            return False
        
        if not self.excel_filter.load_excel(filepath):
            messagebox.showerror("Erro", "Erro ao carregar o arquivo Excel!")
            return False
            
        return True
        
    def filter_multiple(self):
        # TODO: Implementar filtro múltiplo
        pass
        
    def keep_columns(self):
        # TODO: Implementar manter colunas
        pass
        
    def remove_columns(self):
        # TODO: Implementar remover colunas
        pass
        
    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    app = ExcelFilterGUI()
    app.run()
