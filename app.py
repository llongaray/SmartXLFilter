from InquirerPy import inquirer
from InquirerPy.base.control import Choice
import pandas as pd
import os
from rich import print
from rich.progress import track
import time
import logging

# Configuração do logger
logging.basicConfig(filename='app.log', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

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
        print("\n[bold yellow]╔══ Iniciando Filtragem Múltipla ══╗[/bold yellow]\n")
        
        filtered_df = self.df.copy()
        total_inicial = len(filtered_df)
        
        steps = len(filters)
        step_size = 100 // steps
        
        for column, value in filters.items():
            for _ in track(range(step_size), description=f"[cyan]Aplicando filtro para {column}...[/cyan]"):
                time.sleep(0.01)
            filtered_df = filtered_df[filtered_df[column] == value]
        
        output_file = os.path.join(output_path, f'filtered_{os.path.basename(self.filepath)}')
        filtered_df.to_excel(output_file, index=False)
        
        print("\n[bold green]╔══ Resumo da Operação ══╗[/bold green]")
        print(f"[white]► Registros originais:[/white]    {total_inicial:,}")
        print(f"[white]► Registros após filtros:[/white] {len(filtered_df):,}")
        print(f"[white]► Registros filtrados:[/white]    {total_inicial - len(filtered_df):,}")
        print(f"\n[bold green]✓ Processo concluído com sucesso![/bold green]")
        print(f"[dim]📁 Arquivo salvo em: {output_file}[/dim]\n")
        
        return output_file

    def get_unique_values_filtered(self, column, current_filters):
        """Retorna valores únicos de uma coluna com filtros aplicados"""
        filtered_df = self.df.copy()
        for col, val in current_filters.items():
            filtered_df = filtered_df[filtered_df[col] == val]
        return filtered_df[column].unique().tolist()

    def keep_columns(self, columns, output_path):
        """Mantém apenas as colunas selecionadas"""
        print("\n[bold yellow]╔══ Iniciando Seleção de Colunas ══╗[/bold yellow]\n")
        
        total_colunas = len(self.df.columns)
        
        for _ in track(range(100), description="[cyan]Processando colunas...[/cyan]"):
            time.sleep(0.01)
        
        filtered_df = self.df[columns].copy()
        output_file = os.path.join(output_path, f'kept_columns_{os.path.basename(self.filepath)}')
        filtered_df.to_excel(output_file, index=False)
        
        print("\n[bold green]╔══ Resumo da Operação ══╗[/bold green]")
        print(f"[white]► Total de colunas original:[/white] {total_colunas:,}")
        print(f"[white]► Colunas mantidas:[/white]        {len(columns):,}")
        print(f"[white]► Colunas removidas:[/white]       {total_colunas - len(columns):,}")
        print(f"\n[bold green]✓ Processo concluído com sucesso![/bold green]")
        print(f"[dim]📁 Arquivo salvo em: {output_file}[/dim]\n")
        
        return output_file

    def remove_columns(self, columns, output_path):
        """Remove as colunas selecionadas"""
        print("\n[bold yellow]╔══ Iniciando Remoção de Colunas ══╗[/bold yellow]\n")
        
        total_colunas = len(self.df.columns)
        
        for _ in track(range(100), description="[cyan]Processando colunas...[/cyan]"):
            time.sleep(0.01)
        
        filtered_df = self.df.drop(columns=columns).copy()
        output_file = os.path.join(output_path, f'removed_columns_{os.path.basename(self.filepath)}')
        filtered_df.to_excel(output_file, index=False)
        
        print("\n[bold green]╔══ Resumo da Operação ══╗[/bold green]")
        print(f"[white]► Total de colunas original:[/white] {total_colunas:,}")
        print(f"[white]► Colunas removidas:[/white]        {len(columns):,}")
        print(f"[white]► Colunas restantes:[/white]        {len(filtered_df.columns):,}")
        print(f"\n[bold green]✓ Processo concluído com sucesso![/bold green]")
        print(f"[dim]📁 Arquivo salvo em: {output_file}[/dim]\n")
        
        return output_file

    def filter_numeric_greater_than(self, column, value, output_path):
        """Filtra valores numéricos maiores que o valor especificado"""
        filtered_df = self.df[self.df[column] > value]
        output_file = os.path.join(output_path, f'numeric_filtered_{os.path.basename(self.filepath)}')
        filtered_df.to_excel(output_file, index=False)
        return output_file

    def filter_numeric_between(self, column, min_value, max_value, output_path):
        """Filtra valores numéricos entre dois valores"""
        filtered_df = self.df[(self.df[column] >= min_value) & (self.df[column] <= max_value)]
        output_file = os.path.join(output_path, f'numeric_filtered_{os.path.basename(self.filepath)}')
        filtered_df.to_excel(output_file, index=False)
        return output_file

    def is_numeric_column(self, column):
        """Verifica se uma coluna é numérica"""
        return pd.api.types.is_numeric_dtype(self.df[column])

    @staticmethod
    def unify_excel_files(directory_path, output_path):
        """Unifica arquivos Excel baseado no CPF"""
        all_files = [f for f in os.listdir(directory_path) if f.endswith(('.xlsx', '.xls'))]
        if not all_files:
            print("Nenhum arquivo Excel encontrado no diretório.")
            return None

        dfs = []
        for file in all_files:
            df = pd.read_excel(os.path.join(directory_path, file))
            if 'CPF' not in df.columns:
                print(f"Arquivo {file} não contém a coluna 'CPF'. Ignorando...")
                continue
            dfs.append(df)

        if not dfs:
            print("Nenhum arquivo válido encontrado.")
            return None

        unified_df = pd.concat(dfs, ignore_index=True)
        unified_df = unified_df.drop_duplicates(subset=['CPF'], keep='first')
        
        output_file = os.path.join(output_path, 'unified_excel.xlsx')
        unified_df.to_excel(output_file, index=False)
        return output_file

    def normalize_cpf(self, cpf):
        """Normaliza o CPF removendo caracteres especiais e espaços"""
        # Converte para string primeiro
        cpf_str = str(cpf)
        return ''.join(filter(str.isdigit, cpf_str))

    def unify_excel_files_with_cpf(self, base_file_path, second_file_path, base_cpf_column, second_cpf_column, output_path):
        """Unifica dois arquivos Excel baseado no CPF"""
        print("\n[bold yellow]╔═�� Iniciando Unificação por CPF ══╗[/bold yellow]\n")
        
        base_df = pd.read_excel(base_file_path)
        second_df = pd.read_excel(second_file_path)
        total_base = len(base_df)
        total_second = len(second_df)

        # Normaliza os CPFs
        for _ in track(range(33), description="[cyan]Normalizando CPFs do arquivo base...[/cyan]"):
            time.sleep(0.01)
        base_df[base_cpf_column] = base_df[base_cpf_column].apply(self.normalize_cpf)
        
        for _ in track(range(33), description="[cyan]Normalizando CPFs do segundo arquivo...[/cyan]"):
            time.sleep(0.01)
        second_df[second_cpf_column] = second_df[second_cpf_column].apply(self.normalize_cpf)
        
        # Realiza o merge
        for _ in track(range(34), description="[cyan]Unificando arquivos...[/cyan]"):
            time.sleep(0.01)
        merged_df = pd.merge(base_df, second_df, left_on=base_cpf_column, right_on=second_cpf_column, how='inner')
        
        output_file = os.path.join(output_path, 'unified_by_cpf.xlsx')
        merged_df.to_excel(output_file, index=False)
        
        print("\n[bold green]╔══ Resumo da Operação ══╗[/bold green]")
        print(f"[white]► Registros no arquivo base:[/white]    {total_base:,}")
        print(f"[white]► Registros no segundo arquivo:[/white] {total_second:,}")
        print(f"[white]► Registros após unificação:[/white]    {len(merged_df):,}")
        print(f"\n[bold green]✓ Processo concluído com sucesso![/bold green]")
        print(f"[dim]📁 Arquivo salvo em: {output_file}[/dim]\n")
        
        return output_file

    def filter_cpf_removal(self, base_file_path, removal_file_path, base_cpf_column, removal_cpf_column, output_path):
        """Remove do arquivo base os CPFs que existem no arquivo de remoção"""
        print("\n[bold yellow]╔══ Iniciando Remoção de CPFs ══╗[/bold yellow]\n")
        
        base_df = pd.read_excel(base_file_path)
        removal_df = pd.read_excel(removal_file_path)
        total_base = len(base_df)
        
        # Normaliza os CPFs
        for _ in track(range(33), description="[cyan]Normalizando CPFs do arquivo base...[/cyan]"):
            time.sleep(0.01)
        base_df[base_cpf_column] = base_df[base_cpf_column].apply(self.normalize_cpf)
        
        for _ in track(range(33), description="[cyan]Normalizando CPFs do arquivo de remoção...[/cyan]"):
            time.sleep(0.01)
        removal_df[removal_cpf_column] = removal_df[removal_cpf_column].apply(self.normalize_cpf)
        
        # Remove as linhas
        for _ in track(range(34), description="[cyan]Removendo CPFs...[/cyan]"):
            time.sleep(0.01)
        filtered_df = base_df[~base_df[base_cpf_column].isin(removal_df[removal_cpf_column])].copy()
        
        # Formata os CPFs
        filtered_df[base_cpf_column] = filtered_df[base_cpf_column].apply(self.format_cpf)
        
        output_file = os.path.join(output_path, f'cpf_filtered_{os.path.basename(base_file_path)}')
        filtered_df.to_excel(output_file, index=False)
        
        print("\n[bold green]╔══ Resumo da Operação ══╗[/bold green]")
        print(f"[white]► Registros originais:[/white]    {total_base:,}")
        print(f"[white]► Registros após remoção:[/white] {len(filtered_df):,}")
        print(f"[white]► Registros removidos:[/white]    {total_base - len(filtered_df):,}")
        print(f"\n[bold green]✓ Processo concluído com sucesso![/bold green]")
        print(f"[dim]📁 Arquivo salvo em: {output_file}[/dim]\n")
        
        return output_file

    

    def filter_cpf_duplicates(self, file_path, cpf_column, output_path):
        """Remove CPFs duplicados mantendo apenas a primeira ocorrência"""
        print("\n[bold yellow]╔══ Iniciando Remoção de Duplicatas ══╗[/bold yellow]\n")
        
        df = pd.read_excel(file_path)
        total = len(df)
        
        # Normaliza os CPFs
        for _ in track(range(50), description="[cyan]Normalizando CPFs...[/cyan]"):
            time.sleep(0.01)
        df[cpf_column] = df[cpf_column].apply(self.normalize_cpf)
        
        # Remove duplicatas
        for _ in track(range(50), description="[cyan]Removendo duplicatas...[/cyan]"):
            time.sleep(0.01)
        filtered_df = df.drop_duplicates(subset=[cpf_column], keep='first').copy()
        
        # Formata os CPFs
        filtered_df[cpf_column] = filtered_df[cpf_column].apply(self.format_cpf)
        
        output_file = os.path.join(output_path, f'unique_cpf_{os.path.basename(file_path)}')
        filtered_df.to_excel(output_file, index=False)
        
        duplicatas = total - len(filtered_df)
        
        print("\n[bold green]╔══ Resumo da Operação ══╗[/bold green]")
        print(f"[white]► Registros originais:[/white]    {total:,}")
        print(f"[white]► Registros únicos:[/white]      {len(filtered_df):,}")
        print(f"[white]► Duplicatas removidas:[/white]  {duplicatas:,}")
        print(f"\n[bold green]✓ Processo concluído com sucesso![/bold green]")
        print(f"[dim]📁 Arquivo salvo em: {output_file}[/dim]\n")
        
        return output_file

    def format_cpf(self, cpf):
        """Formata o CPF para ter 11 dígitos, adicionando zeros à esquerda se necessário"""
        # Primeiro normaliza o CPF para ter apenas dígitos
        cpf_clean = self.normalize_cpf(cpf)
        # Adiciona zeros à esquerda se necessário para ter 11 dígitos
        return cpf_clean.zfill(11)

def filter_single_excel():
    filter_system = ExcelFilter()
    
    print("\n[bold yellow]╔══ Iniciando Filtro Único ══╗[/bold yellow]\n")
    
    excel_path = inquirer.text(
        message="Digite o caminho do arquivo Excel:"
    ).execute()
    
    if not filter_system.load_excel(excel_path):
        print("[bold red]✗ Erro ao carregar arquivo![/bold red]\n")
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
    
    total_registros = len(filter_system.df)
    
    for _ in track(range(100), description="[cyan]Aplicando filtro...[/cyan]"):
        time.sleep(0.01)
    
    filtered_df = filter_system.df[filter_system.df[selected_header] == selected_value].copy()
    output_file = filter_system.filter_and_save(selected_header, selected_value, output_dir)
    
    print("\n[bold green]╔══ Resumo da Operação ══╗[/bold green]")
    print(f"[white]► Registros originais:[/white]    {total_registros:,}")
    print(f"[white]► Registros filtrados:[/white]    {len(filtered_df):,}")
    print(f"[white]► Registros removidos:[/white]    {total_registros - len(filtered_df):,}")
    print(f"\n[bold green]✓ Processo concluído com sucesso![/bold green]")
    print(f"[dim]📁 Arquivo salvo em: {output_file}[/dim]\n")

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

def filter_numeric():
    """Função para filtrar valores numéricos"""
    filter_system = ExcelFilter()
    
    print("\n[bold yellow]╔══ Iniciando Filtro Numérico ══╗[/bold yellow]\n")
    
    excel_path = inquirer.text(
        message="Digite o caminho do arquivo Excel:"
    ).execute()
    
    if not filter_system.load_excel(excel_path):
        print("[bold red]✗ Erro ao carregar arquivo![/bold red]\n")
        return

    # Filtra apenas colunas numéricas
    numeric_columns = [col for col in filter_system.headers if filter_system.is_numeric_column(col)]
    if not numeric_columns:
        print("[bold red]✗ Não há colunas numéricas neste arquivo![/bold red]\n")
        return

    selected_header = inquirer.select(
        message="Selecione a coluna numérica para filtrar:",
        choices=numeric_columns
    ).execute()

    filter_type = inquirer.select(
        message="Selecione o tipo de filtro:",
        choices=[
            Choice("1", "Maior que"),
            Choice("2", "Entre valores")
        ]
    ).execute()

    output_dir = inquirer.text(
        message="Digite o caminho para salvar o arquivo filtrado:"
    ).execute()

    total_registros = len(filter_system.df)

    if filter_type == "1":
        value = float(inquirer.text(
            message="Digite o valor mínimo:"
        ).execute())
        
        for _ in track(range(100), description="[cyan]Aplicando filtro...[/cyan]"):
            time.sleep(0.01)
            
        filtered_df = filter_system.df[filter_system.df[selected_header] > value].copy()
        output_file = filter_system.filter_numeric_greater_than(selected_header, value, output_dir)
        
        print("\n[bold green]╔══ Resumo da Operação ══╗[/bold green]")
        print(f"[white]► Registros originais:[/white]    {total_registros:,}")
        print(f"[white]► Registros > {value}:[/white]    {len(filtered_df):,}")
        print(f"[white]► Registros removidos:[/white]    {total_registros - len(filtered_df):,}")
        
    else:
        min_value = float(inquirer.text(
            message="Digite o valor mínimo:"
        ).execute())
        max_value = float(inquirer.text(
            message="Digite o valor máximo:"
        ).execute())
        
        for _ in track(range(100), description="[cyan]Aplicando filtro...[/cyan]"):
            time.sleep(0.01)
            
        filtered_df = filter_system.df[(filter_system.df[selected_header] >= min_value) & 
                                     (filter_system.df[selected_header] <= max_value)].copy()
        output_file = filter_system.filter_numeric_between(selected_header, min_value, max_value, output_dir)
        
        print("\n[bold green]╔══ Resumo da Operação ══╗[/bold green]")
        print(f"[white]► Registros originais:[/white]    {total_registros:,}")
        print(f"[white]► Registros entre {min_value} e {max_value}:[/white]    {len(filtered_df):,}")
        print(f"[white]► Registros removidos:[/white]    {total_registros - len(filtered_df):,}")

    print(f"\n[bold green]✓ Processo concluído com sucesso![/bold green]")
    print(f"[dim]📁 Arquivo salvo em: {output_file}[/dim]\n")

def unify_excel_files():
    """Função para unificar arquivos Excel"""
    print("\n[bold yellow]╔══ Iniciando Unificação de Arquivos ══╗[/bold yellow]\n")
    print("[white]► Requisitos: os arquivos precisam ter colunas com mesmo nome[/white]")
    print("[white]► Coluna obrigatória: 'CPF'[/white]\n")
    
    directory_path = inquirer.text(
        message="Digite o caminho da pasta com os arquivos Excel:"
    ).execute()
    
    if not os.path.isdir(directory_path):
        print("[bold red]✗ Diretório inválido![/bold red]\n")
        return

    output_dir = inquirer.text(
        message="Digite o caminho para salvar o arquivo unificado:"
    ).execute()

    for _ in track(range(100), description="[cyan]Unificando arquivos...[/cyan]"):
        time.sleep(0.01)

    output_file = ExcelFilter.unify_excel_files(directory_path, output_dir)
    
    if output_file:
        print("\n[bold green]✓ Processo concluído com sucesso![/bold green]")
        print(f"[dim]📁 Arquivo salvo em: {output_file}[/dim]\n")
    else:
        print("[bold red]✗ Erro ao unificar arquivos![/bold red]\n")

def unify_excel_files_with_cpf():
    """Função para unificar arquivos Excel com base no CPF"""
    filter_system = ExcelFilter()

    base_file_path = inquirer.text(
        message="Digite o caminho do arquivo base (.xlsx):"
    ).execute()

    if not filter_system.load_excel(base_file_path):
        return

    # Seleciona a coluna de CPF do arquivo base
    base_cpf_column = inquirer.select(
        message="Selecione a coluna de CPF do arquivo base:",
        choices=filter_system.headers
    ).execute()

    second_file_path = inquirer.text(
        message="Digite o caminho do segundo arquivo (.xlsx):"
    ).execute()

    if not filter_system.load_excel(second_file_path):
        return

    # Seleciona a coluna de CPF do segundo arquivo
    second_cpf_column = inquirer.select(
        message="Selecione a coluna de CPF do segundo arquivo:",
        choices=filter_system.headers
    ).execute()

    output_dir = inquirer.text(
        message="Digite o caminho para salvar o arquivo unificado:"
    ).execute()

    output_file = filter_system.unify_excel_files_with_cpf(base_file_path, second_file_path, base_cpf_column, second_cpf_column, output_dir)
    print(f"\nArquivo unificado salvo em: {output_file}")

def filter_cpf_removal():
    """Função para remover CPFs de um arquivo base que existem em outro arquivo"""
    filter_system = ExcelFilter()
    
    # Arquivo base
    base_file_path = inquirer.text(
        message="Digite o caminho do arquivo base (.xlsx):"
    ).execute()
    
    if not filter_system.load_excel(base_file_path):
        return
        
    # Seleciona coluna CPF do arquivo base
    base_cpf_column = inquirer.select(
        message="Selecione a coluna de CPF do arquivo base:",
        choices=filter_system.headers
    ).execute()
    
    # Arquivo de remoção
    removal_file_path = inquirer.text(
        message="Digite o caminho do arquivo com CPFs a serem removidos (.xlsx):"
    ).execute()
    
    if not filter_system.load_excel(removal_file_path):
        return
        
    # Seleciona coluna CPF do arquivo de remoção
    removal_cpf_column = inquirer.select(
        message="Selecione a coluna de CPF do arquivo de remoção:",
        choices=filter_system.headers
    ).execute()
    
    output_dir = inquirer.text(
        message="Digite o caminho para salvar o arquivo filtrado:"
    ).execute()
    
    output_file = filter_system.filter_cpf_removal(base_file_path, removal_file_path, 
                                                 base_cpf_column, removal_cpf_column, output_dir)
    print(f"\nArquivo filtrado salvo em: {output_file}")

def filter_cpf_duplicates():
    """Função para remover CPFs duplicados"""
    filter_system = ExcelFilter()
    
    file_path = inquirer.text(
        message="Digite o caminho do arquivo (.xlsx):"
    ).execute()
    
    if not filter_system.load_excel(file_path):
        return
        
    cpf_column = inquirer.select(
        message="Selecione a coluna de CPF:",
        choices=filter_system.headers
    ).execute()
    
    output_dir = inquirer.text(
        message="Digite o caminho para salvar o arquivo filtrado:"
    ).execute()
    
    output_file = filter_system.filter_cpf_duplicates(file_path, cpf_column, output_dir)
    print(f"\nArquivo com CPFs únicos salvo em: {output_file}")

def filter_phone_numbers_csv():
    """Função para verificar números de telefone com prefixo '55' e exatamente 11 dígitos em arquivos CSV."""
    print("\n[bold yellow]╔══ Iniciando Filtro de Números de Telefone (CSV) ══╗[/bold yellow]\n")

    # Solicita o caminho do arquivo CSV
    csv_path = inquirer.text(
        message="Digite o caminho do arquivo CSV:"
    ).execute()

    try:
        # Carrega o arquivo CSV em um DataFrame
        df = pd.read_csv(csv_path)
    except Exception as e:
        print(f"[bold red]✗ Erro ao carregar arquivo CSV: {e}[/bold red]\n")
        return

    # Lista os cabeçalhos e solicita ao usuário para selecionar a coluna de telefone
    selected_header = inquirer.select(
        message="Selecione a coluna que contém os números de telefone:",
        choices=df.columns.tolist()
    ).execute()

    # Inicializa contadores
    total_registros = len(df)
    registros_removidos = 0

    # Processa cada linha e remove as inválidas
    indices_to_drop = []
    for index in track(df.index, description="[cyan]Filtrando registros...[cyan]", total=total_registros):
        value = str(df.at[index, selected_header]).strip()

        # Remove caracteres não numéricos
        clean_value = ''.join(filter(str.isdigit, value))

        # Verifica se o número é válido
        if len(clean_value) != 13 or not clean_value.startswith('55'):
            indices_to_drop.append(index)
            registros_removidos += 1

    # Remove os índices coletados
    df.drop(indices_to_drop, inplace=True)

    # Calcula total de registros após remoção
    registros_restantes = len(df)

    # Exibe resumo da operação
    print("\n[bold green]╔══ Resumo da Operação ══╗[/bold green]")
    print(f"[white]► Registros originais:[/white]    {total_registros:,}")
    print(f"[white]► Registros removidos:[/white]    {registros_removidos:,}")
    print(f"[white]► Registros restantes:[/white]   {registros_restantes:,}")

    # Solicita o diretório de saída
    output_dir = inquirer.text(
        message="Digite o caminho para salvar o arquivo filtrado:"
    ).execute()

    # Salva o arquivo alterado com prefixo no nome
    output_file = os.path.join(output_dir, f'filtro_cel_num_{os.path.basename(csv_path)}')
    try:
        df.to_csv(output_file, index=False)
        print(f"\n[bold green]✓ Processo concluído com sucesso![bold green]")
        print(f"[dim]📁 Arquivo salvo em: {output_file}[dim]\n")
    except Exception as e:
        print(f"[bold red]✗ Erro ao salvar o arquivo: {e}[/bold red]\n")

def adjust_cpfs_to_11_digits():
    """Função para ajustar CPFs em um arquivo Excel, garantindo que todos tenham 11 dígitos."""
    filter_system = ExcelFilter()

    print("\n[bold yellow]╔══ Iniciando Ajuste de CPFs ══╗[/bold yellow]\n")

    # Solicita o caminho do arquivo Excel
    excel_path = inquirer.text(
        message="Digite o caminho do arquivo Excel:"
    ).execute()

    if not filter_system.load_excel(excel_path):
        print("[bold red]✗ Erro ao carregar arquivo![/bold red]\n")
        return

    # Lista os cabeçalhos e solicita ao usuário para selecionar a coluna de CPF
    selected_header = inquirer.select(
        message="Selecione a coluna que contém os CPFs:",
        choices=filter_system.headers
    ).execute()

    # Converte a coluna para string antes de normalizar os CPFs
    filter_system.df[selected_header] = filter_system.df[selected_header].astype(str)

    # Inicializa contadores
    total_registros = len(filter_system.df)
    registros_normalizados = 0

    # Ajusta os CPFs na coluna selecionada
    for index in track(filter_system.df.index, description="[cyan]Ajustando CPFs...[cyan]", total=total_registros):
        value = str(filter_system.df.at[index, selected_header]).strip()

        # Remove caracteres não numéricos
        clean_value = ''.join(filter(str.isdigit, value))

        # Ajusta para 11 dígitos adicionando zeros à esquerda
        if clean_value and len(clean_value) <= 11:
            normalized_cpf = clean_value.zfill(11)
            filter_system.df.at[index, selected_header] = normalized_cpf
            registros_normalizados += 1

    # Exibe resumo da operação
    print("\n[bold green]╔══ Resumo da Operação ══╗[/bold green]")
    print(f"[white]► Registros originais:[/white]       {total_registros:,}")
    print(f"[white]► CPFs ajustados:[/white]          {registros_normalizados:,}")

    # Solicita o diretório de saída
    output_dir = inquirer.text(
        message="Digite o caminho para salvar o arquivo com CPFs ajustados:"
    ).execute()

    # Salva o arquivo alterado com prefixo no nome
    output_file = os.path.join(output_dir, f'cpfs_ajustados_{os.path.basename(excel_path)}')
    try:
        filter_system.df.to_excel(output_file, index=False)
        print(f"\n[bold green]✓ Processo concluído com sucesso![bold green]")
        print(f"[dim]📁 Arquivo salvo em: {output_file}[dim]\n")
    except Exception as e:
        print(f"[bold red]✗ Erro ao salvar o arquivo: {e}[/bold red]\n")

def format_money_column():
    """Função para formatar uma coluna de valores para o padrão monetário"""
    filter_system = ExcelFilter()
    
    print("\n[bold yellow]╔══ Iniciando Formatação Monetária ══╗[/bold yellow]\n")
    
    excel_path = inquirer.text(
        message="Digite o caminho do arquivo Excel:"
    ).execute()
    
    if not filter_system.load_excel(excel_path):
        print("[bold red]✗ Erro ao carregar arquivo![/bold red]\n")
        return
    
    selected_column = inquirer.select(
        message="Selecione a coluna de valores:",
        choices=filter_system.headers
    ).execute()
    
    total_registros = len(filter_system.df)
    formatted_count = 0
    
    # Formata os valores para o padrão monetário
    for index, value in filter_system.df[selected_column].items():
        try:
            # Converte o valor para string e formata para o padrão monetário
            formatted_value = f"{int(value) / 100:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            filter_system.df.at[index, selected_column] = formatted_value
            formatted_count += 1
        except ValueError:
            continue
    
    output_dir = inquirer.text(
        message="Digite o caminho para salvar o arquivo formatado:"
    ).execute()
    
    output_file = os.path.join(output_dir, f'format_money_{os.path.basename(excel_path)}')
    filter_system.df.to_excel(output_file, index=False)
    
    print("\n[bold green]╔══ Resumo da Operação ══╗[/bold green]")
    print(f"[white]► Total de registros processados:[/white] {total_registros:,}")
    print(f"[white]► Valores formatados:[/white] {formatted_count:,}")
    print(f"\n[bold green]✓ Processo concluído com sucesso![/bold green]")
    print(f"[dim]📁 Arquivo salvo em: {output_file}[/dim]\n")

def filter_and_format_rgs():
    """Função para filtrar RGS inválidos e formatar os válidos para 10 dígitos."""

    print("\n[bold yellow]╔══ Iniciando Filtragem e Formatação de RGs ══╗[/bold yellow]\n")

    # Recebe o arquivo base
    base_file_path = inquirer.text(
        message="Digite o caminho do arquivo base (.xlsx):"
    ).execute()

    try:
        base_df = pd.read_excel(base_file_path)
    except Exception as e:
        print(f"[bold red]✗ Erro ao carregar o arquivo base: {e}[/bold red]\n")
        return

    # Seleciona a coluna de RG no arquivo base
    rg_column = inquirer.select(
        message="Selecione a coluna de RG:",
        choices=base_df.columns.tolist()
    ).execute()

    print("\n[cyan]Verificando e filtrando RGs inválidos...[/cyan]")
    for _ in track(range(100), description="[cyan]Processando RGs...[/cyan]"):
        time.sleep(0.01)

    # Verifica RGs inválidos
    def is_valid_rg(value):
        if pd.isna(value):  # Verifica valores nulos
            return False
        value = str(value).strip()
        if not value.isdigit():  # Verifica se contém apenas dígitos
            return False
        if len(value) < 5:  # Verifica se tem pelo menos 5 dígitos
            return False
        return True

    # Filtra os registros válidos e inválidos
    base_df['RG_VALIDO'] = base_df[rg_column].apply(is_valid_rg)
    invalid_rgs = base_df[~base_df['RG_VALIDO']].copy()
    valid_rgs = base_df[base_df['RG_VALIDO']].copy()

    # Formata os RGs válidos para 10 dígitos
    print("\n[cyan]Formatando RGs válidos para 10 dígitos...[/cyan]")
    valid_rgs[rg_column] = valid_rgs[rg_column].astype(str).str.zfill(10)

    # Exibe resumo da operação
    total_registros = len(base_df)
    total_validos = len(valid_rgs)
    total_invalidos = len(invalid_rgs)

    print("\n[bold green]╔══ Resumo da Operação ══╗[/bold green]")
    print(f"[white]► Registros originais:[/white]    {total_registros:,}")
    print(f"[white]► RGs válidos:[/white]           {total_validos:,}")
    print(f"[white]► RGs inválidos:[/white]         {total_invalidos:,}")

    # Recebe o diretório de saída
    output_dir = inquirer.text(
        message="Digite o caminho para salvar os arquivos filtrados:"
    ).execute()

    # Salva os arquivos filtrados
    valid_output_file = os.path.join(output_dir, f'valid_rgs_{os.path.basename(base_file_path)}')
    invalid_output_file = os.path.join(output_dir, f'invalid_rgs_{os.path.basename(base_file_path)}')

    try:
        valid_rgs.drop(columns=['RG_VALIDO'], inplace=True)
        invalid_rgs.drop(columns=['RG_VALIDO'], inplace=True)

        valid_rgs.to_excel(valid_output_file, index=False)
        invalid_rgs.to_excel(invalid_output_file, index=False)

        print("\n[bold green]✓ Processo concluído com sucesso![/bold green]")
        print(f"[dim]📁 RGs válidos salvos em: {valid_output_file}[/dim]")
        print(f"[dim]📁 RGs inválidos salvos em: {invalid_output_file}[/dim]\n")

    except Exception as e:
        print(f"[bold red]✗ Erro ao salvar os arquivos: {e}[/bold red]\n")

def filter_remove_by_name():
    """Função para filtrar e remover registros por nome."""

    print("\n[bold yellow]╔══ Iniciando Filtragem e Remoção por Nome ══╗[/bold yellow]\n")

    # Recebe o arquivo base
    base_file_path = inquirer.text(
        message="Digite o caminho do arquivo base (.xlsx):"
    ).execute()

    try:
        base_df = pd.read_excel(base_file_path)
    except Exception as e:
        print(f"[bold red]✗ Erro ao carregar o arquivo base: {e}[/bold red]\n")
        return

    # Seleciona a coluna de nome no arquivo base
    base_name_column = inquirer.select(
        message="Selecione a coluna de NOME no arquivo base:",
        choices=base_df.columns.tolist()
    ).execute()

    # Recebe o arquivo de blacklist
    blacklist_file_path = inquirer.text(
        message="Digite o caminho do arquivo de blacklist (.xlsx):"
    ).execute()

    try:
        blacklist_df = pd.read_excel(blacklist_file_path)
    except Exception as e:
        print(f"[bold red]✗ Erro ao carregar o arquivo de blacklist: {e}[/bold red]\n")
        return

    # Seleciona a coluna de nome no arquivo de blacklist
    blacklist_name_column = inquirer.select(
        message="Selecione a coluna de NOME no arquivo de blacklist:",
        choices=blacklist_df.columns.tolist()
    ).execute()

    # Converte os nomes para caixa alta
    print("\n[cyan]Normalizando nomes para caixa alta...[/cyan]")
    for _ in track(range(100), description="[cyan]Processando...[/cyan]"):
        time.sleep(0.01)

    base_df[base_name_column] = base_df[base_name_column].str.upper().fillna("")
    blacklist_df[blacklist_name_column] = blacklist_df[blacklist_name_column].str.upper().fillna("")

    # Filtra os registros
    print("\n[cyan]Removendo registros encontrados na blacklist...[/cyan]")
    blacklist_names = set(blacklist_df[blacklist_name_column].tolist())
    filtered_df = base_df[~base_df[base_name_column].isin(blacklist_names)].copy()

    # Exibe resumo da operação
    total_registros = len(base_df)
    registros_removidos = total_registros - len(filtered_df)

    print("\n[bold green]╔══ Resumo da Operação ══╗[/bold green]")
    print(f"[white]► Registros originais:[/white]    {total_registros:,}")
    print(f"[white]► Registros removidos:[/white]    {registros_removidos:,}")
    print(f"[white]► Registros restantes:[/white]   {len(filtered_df):,}")

    # Recebe o diretório de saída
    output_dir = inquirer.text(
        message="Digite o caminho para salvar o arquivo filtrado:"
    ).execute()

    # Salva o arquivo filtrado com o prefixo
    output_file = os.path.join(output_dir, f'filtra_name_remove_{os.path.basename(base_file_path)}')
    try:
        filtered_df.to_excel(output_file, index=False)
        print(f"\n[bold green]✓ Processo concluído com sucesso![/bold green]")
        print(f"[dim]📁 Arquivo salvo em: {output_file}[/dim]\n")
    except Exception as e:
        print(f"[bold red]✗ Erro ao salvar o arquivo: {e}[/bold red]\n")

def format_numbers_with_prefix():
    """Função para adicionar o prefixo '55' a números com 11 dígitos"""
    filter_system = ExcelFilter()
    
    excel_path = inquirer.text(
        message="Digite o caminho do arquivo Excel (.xlsx):"
    ).execute()
    
    if not filter_system.load_excel(excel_path):
        print("[bold red]✗ Erro ao carregar arquivo![/bold red]\n")
        return
    
    # Seleciona a coluna de números
    numeric_columns = [col for col in filter_system.headers if filter_system.is_numeric_column(col)]
    if not numeric_columns:
        print("[bold red]✗ Não há colunas numéricas neste arquivo![/bold red]\n")
        return

    selected_column = inquirer.select(
        message="Selecione a coluna de números:",
        choices=numeric_columns
    ).execute()

    # Adiciona '55' aos números com 11 dígitos
    total_numbers = len(filter_system.df)
    formatted_count = 0

    for index, value in filter_system.df[selected_column].items():
        if len(str(value)) == 11:
            filter_system.df.at[index, selected_column] = f'55{value}'
            formatted_count += 1

    output_dir = inquirer.text(
        message="Digite o caminho para salvar o arquivo formatado:"
    ).execute()

    output_file = os.path.join(output_dir, f'num_format_{os.path.basename(excel_path)}')
    filter_system.df.to_excel(output_file, index=False)

    print("\n[bold green]╔══ Resumo da Operação ══╗[/bold green]")
    print(f"[white]► Total de números processados:[/white] {total_numbers:,}")
    print(f"[white]► Números formatados com prefixo '55':[/white] {formatted_count:,}")
    print(f"[dim]📁 Arquivo salvo em: {output_file}[/dim]\n")

def main():
    while True:
        choice = inquirer.select(
            message="Selecione uma categoria:",
            choices=[
                Choice("1", "Filtros Únicos"),
                Choice("2", "Filtros Múltiplos"),
                Choice("3", "Remoções"),
                Choice("4", "Adições/Unificações"),
                Choice("5", "Formatações"),
                Choice("6", "Sair")
            ]
        ).execute()

        if choice == "1":
            filtros_unicos()
        elif choice == "2":
            filtros_multiplos()
        elif choice == "3":
            remocoes()
        elif choice == "4":
            adicoes_unificacoes()
        elif choice == "5":
            formatacoes()
        elif choice == "6":
            print("Programa encerrado!")
            break

def filtros_unicos():
    while True:
        choice = inquirer.select(
            message="Selecione um filtro único:",
            choices=[
                Choice("1", "Filtrar Excel (único)"),
                Choice("2", "Filtrar valores numéricos"),
                Choice("3", "Voltar")
            ]
        ).execute()

        if choice == "1":
            filter_single_excel()
        elif choice == "2":
            filter_numeric()
        elif choice == "3":
            break

def filtros_multiplos():
    while True:
        choice = inquirer.select(
            message="Selecione um filtro múltiplo:",
            choices=[
                Choice("1", "Filtrar Excel (múltiplo)"),
                Choice("2", "Voltar")
            ]
        ).execute()

        if choice == "1":
            filter_multiple_excel()
        elif choice == "2":
            break

def remocoes():
    while True:
        choice = inquirer.select(
            message="Selecione uma remoção:",
            choices=[
                Choice("1", "Filtrar CPF - Remoção"),
                Choice("2", "Filtrar e remover por nome"),
                Choice("3", "Remover números fixos e células vazias"),
                Choice("4", "Voltar")
            ]
        ).execute()

        if choice == "1":
            filter_cpf_removal()
        elif choice == "2":
            filter_remove_by_name()
        elif choice == "3":
            filter_phone_numbers_csv()
        elif choice == "4":
            break

def adicoes_unificacoes():
    while True:
        choice = inquirer.select(
            message="Selecione uma adição ou unificação:",
            choices=[
                Choice("1", "Unificar arquivos Excel"),
                Choice("2", "Unificar arquivos Excel com base no CPF"),
                Choice("3", "Voltar")
            ]
        ).execute()

        if choice == "1":
            unify_excel_files()
        elif choice == "2":
            unify_excel_files_with_cpf()
        elif choice == "3":
            break

def formatacoes():
    while True:
        choice = inquirer.select(
            message="Selecione uma formatação:",
            choices=[
                Choice("1", "Ajustar CPFs para 11 dígitos"),
                Choice("2", "Formatar coluna de valores para padrão monetário"),
                Choice("3", "Formatar Números com Prefixo '55'"),
                Choice("4", "Filtrar e formatar RGs"),
                Choice("5", "Voltar")
            ]
        ).execute()

        if choice == "1":
            adjust_cpfs_to_11_digits()
        elif choice == "2":
            format_money_column()
        elif choice == "3":
            format_numbers_with_prefix()
        elif choice == "4":
            filter_and_format_rgs()
        elif choice == "5":
            break

if __name__ == "__main__":
    main()
