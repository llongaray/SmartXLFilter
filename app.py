from InquirerPy import inquirer
from InquirerPy.base.control import Choice
import pandas as pd
import os

class ExcelFilter:
    def __init__(self):
        self.df = None
        self.filepath = None
        self.headers = None
        
    def load_excel(self, filepath):
        """Carrega o arquivo Excel e extrai os cabeçalhos"""
        try:
            self.filepath = filepath
            self.df = pd.read_excel(filepath)
            self.headers = list(self.df.columns)
            return True
        except Exception as e:
            print(f"Erro ao carregar arquivo: {e}")
            return False

    def get_unique_values(self, column):
        """Retorna valores únicos de uma coluna específica"""
        return self.df[column].unique().tolist()

    def filter_and_save(self, column, value, output_path):
        """Filtra o DataFrame e salva em novo arquivo"""
        filtered_df = self.df[self.df[column] == value]
        output_file = os.path.join(output_path, f'filtered_{os.path.basename(self.filepath)}')
        filtered_df.to_excel(output_file, index=False)
        return output_file

    def filter_and_save_multiple(self, filters, output_path):
        """Filtra o DataFrame com múltiplos critérios e salva em novo arquivo"""
        filtered_df = self.df.copy()
        for column, value in filters.items():
            filtered_df = filtered_df[filtered_df[column] == value]
        
        output_file = os.path.join(output_path, f'filtered_{os.path.basename(self.filepath)}')
        filtered_df.to_excel(output_file, index=False)
        return output_file

    def get_unique_values_filtered(self, column, current_filters):
        """Retorna valores únicos de uma coluna com filtros aplicados"""
        filtered_df = self.df.copy()
        for col, val in current_filters.items():
            filtered_df = filtered_df[filtered_df[col] == val]
        return filtered_df[column].unique().tolist()

    def keep_columns(self, columns, output_path):
        """Mantém apenas as colunas selecionadas"""
        filtered_df = self.df[columns]
        output_file = os.path.join(output_path, f'kept_columns_{os.path.basename(self.filepath)}')
        filtered_df.to_excel(output_file, index=False)
        return output_file

    def remove_columns(self, columns, output_path):
        """Remove as colunas selecionadas"""
        filtered_df = self.df.drop(columns=columns)
        output_file = os.path.join(output_path, f'removed_columns_{os.path.basename(self.filepath)}')
        filtered_df.to_excel(output_file, index=False)
        return output_file

def filter_single_excel():
    filter_system = ExcelFilter()
    
    excel_path = inquirer.text(
        message="Digite o caminho do arquivo Excel:"
    ).execute()
    
    if not filter_system.load_excel(excel_path):
        return
    
    selected_header = inquirer.select(
        message="Selecione o cabeçalho para filtrar:",
        choices=filter_system.headers
    ).execute()
    
    unique_values = filter_system.get_unique_values(selected_header)
    
    selected_value = inquirer.select(
        message=f"Selecione o valor para filtrar em '{selected_header}':",
        choices=unique_values
    ).execute()
    
    output_dir = inquirer.text(
        message="Digite o caminho para salvar o arquivo filtrado:"
    ).execute()
    
    output_file = filter_system.filter_and_save(selected_header, selected_value, output_dir)
    print(f"\nArquivo filtrado salvo em: {output_file}")

def filter_multiple_excel():
    filter_system = ExcelFilter()
    
    excel_path = inquirer.text(
        message="Digite o caminho do arquivo Excel:"
    ).execute()
    
    if not filter_system.load_excel(excel_path):
        return
    
    filters = {}
    while True:
        # Pergunta se quer adicionar mais um filtro
        should_continue = inquirer.confirm(
            message="Deseja adicionar um filtro?",
            default=True
        ).execute()
        
        if not should_continue:
            break
            
        # Seleciona o cabeçalho
        selected_header = inquirer.select(
            message="Selecione o cabeçalho para filtrar:",
            choices=filter_system.headers
        ).execute()
        
        # Obtém valores únicos considerando filtros anteriores
        unique_values = filter_system.get_unique_values_filtered(selected_header, filters)
        
        if not unique_values:
            print("Não há valores disponíveis com os filtros atuais.")
            break
            
        # Seleciona o valor
        selected_value = inquirer.select(
            message=f"Selecione o valor para filtrar em '{selected_header}':",
            choices=unique_values
        ).execute()
        
        filters[selected_header] = selected_value
    
    if filters:
        output_dir = inquirer.text(
            message="Digite o caminho para salvar o arquivo filtrado:"
        ).execute()
        
        output_file = filter_system.filter_and_save_multiple(filters, output_dir)
        print(f"\nArquivo filtrado salvo em: {output_file}")

def select_columns():
    """Função auxiliar para selecionar múltiplas colunas"""
    filter_system = ExcelFilter()
    
    excel_path = inquirer.text(
        message="Digite o caminho do arquivo Excel:"
    ).execute()
    
    if not filter_system.load_excel(excel_path):
        return None, None
    
    selected_columns = []
    while True:
        should_continue = inquirer.confirm(
            message="Deseja selecionar uma coluna?",
            default=True
        ).execute()
        
        if not should_continue:
            break
        
        remaining_columns = [col for col in filter_system.headers if col not in selected_columns]
        if not remaining_columns:
            print("Todas as colunas já foram selecionadas.")
            break
            
        selected_header = inquirer.select(
            message="Selecione a coluna:",
            choices=remaining_columns
        ).execute()
        
        selected_columns.append(selected_header)
        
    return filter_system, selected_columns

def keep_selected_columns():
    filter_system, selected_columns = select_columns()
    
    if not filter_system or not selected_columns:
        return
    
    output_dir = inquirer.text(
        message="Digite o caminho para salvar o arquivo:"
    ).execute()
    
    output_file = filter_system.keep_columns(selected_columns, output_dir)
    print(f"\nArquivo salvo com as colunas selecionadas em: {output_file}")

def remove_selected_columns():
    filter_system, selected_columns = select_columns()
    
    if not filter_system or not selected_columns:
        return
    
    output_dir = inquirer.text(
        message="Digite o caminho para salvar o arquivo:"
    ).execute()
    
    output_file = filter_system.remove_columns(selected_columns, output_dir)
    print(f"\nArquivo salvo sem as colunas selecionadas em: {output_file}")

def main():
    while True:
        choice = inquirer.select(
            message="Selecione uma opção:",
            choices=[
                Choice("1", "Filtrar Excel (único)"),
                Choice("2", "Filtrar Excel (múltiplo)"),
                Choice("3", "Manter colunas selecionadas"),
                Choice("4", "Remover colunas selecionadas"),
                Choice("5", "Sair")
            ]
        ).execute()
        
        if choice == "1":
            filter_single_excel()
        elif choice == "2":
            filter_multiple_excel()
        elif choice == "3":
            keep_selected_columns()
        elif choice == "4":
            remove_selected_columns()
        else:
            print("Programa encerrado!")
            break

if __name__ == "__main__":
    main()
