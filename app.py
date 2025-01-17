from InquirerPy import inquirer
from InquirerPy.base.control import Choice
import requests
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

def format_values_to_money():
    """
    Formata valores de uma coluna para o formato monetário (123400 -> 1234,00).
    """
    print("\n[bold yellow]╔══ Iniciando Formatação Monetária ══╗[/bold yellow]\n")

    # Recebe o caminho do arquivo
    file_path = inquirer.text(
        message="Digite o caminho do arquivo Excel (.xlsx):"
    ).execute()

    try:
        df = pd.read_excel(file_path)
    except Exception as e:
        print(f"[bold red]✗ Erro ao carregar o arquivo: {e}[bold red]\n")
        return

    # Seleciona a coluna de valores
    selected_column = inquirer.select(
        message="Selecione a coluna com os valores a formatar:",
        choices=df.columns.tolist()
    ).execute()

    print("\n[cyan]Formatando valores...[/cyan]")

    # Formata os valores para o padrão monetário
    def format_money(value):
        try:
            # Divide por 100 e converte para string no formato monetário
            formatted_value = f"{int(value) / 100:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            return formatted_value
        except (ValueError, TypeError):
            return value  # Retorna o valor original se não for possível formatar

    # Aplica a formatação
    for _ in track(range(100), description="[cyan]Processando valores...[/cyan]"):
        df[selected_column] = df[selected_column].apply(format_money)

    # Pergunta o diretório para salvar
    output_dir = inquirer.text(
        message="Digite o caminho para salvar o arquivo formatado:"
    ).execute()

    # Adiciona o prefixo ao nome do arquivo de saída
    output_file = os.path.join(output_dir, f"format_money_{os.path.basename(file_path)}")

    try:
        df.to_excel(output_file, index=False)
        print(f"\n[bold green]✓ Processo concluído com sucesso![bold green]")
        print(f"[dim]📁 Arquivo salvo em: {output_file}[dim]\n")
    except Exception as e:
        print(f"[bold red]✗ Erro ao salvar o arquivo: {e}[bold red]\n")

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

def formatar_coluna_data():
    """Função para formatar colunas de data em um arquivo Excel."""
    print("\n[bold yellow]╔══ Iniciando Formatação de Datas ══╗[/bold yellow]\n")

    # Recebe o arquivo
    file_path = inquirer.text(
        message="Digite o caminho do arquivo Excel (.xlsx):"
    ).execute()

    try:
        df = pd.read_excel(file_path)
    except Exception as e:
        print(f"[bold red]✗ Erro ao carregar o arquivo: {e}[/bold red]\n")
        return

    # Seleciona a coluna de data
    date_column = inquirer.select(
        message="Selecione a coluna de data:",
        choices=df.columns.tolist()
    ).execute()

    print("\n[cyan]Formatando dados...[/cyan]")

    for _ in track(range(100), description="[cyan]Processando...[/cyan]"):
        pass

    # Converte as datas para o formato dd/MM/YYYY
    try:
        df[date_column] = pd.to_datetime(df[date_column], errors='coerce').dt.strftime('%d/%m/%Y')
    except Exception as e:
        print(f"[bold red]✗ Erro ao formatar as datas: {e}[/bold red]\n")
        return

    # Pergunta o diretório para salvar
    output_dir = inquirer.text(
        message="Digite o caminho para salvar o arquivo formatado:"
    ).execute()

    output_file = os.path.join(output_dir, f"data_formatada_{os.path.basename(file_path)}")

    try:
        df.to_excel(output_file, index=False)
        print(f"\n[bold green]✓ Processo concluído com sucesso![bold green]")
        print(f"[dim]📁 Arquivo salvo em: {output_file}[dim]\n")
    except Exception as e:
        print(f"[bold red]✗ Erro ao salvar o arquivo: {e}[/bold red]\n")


def adicionar_dados_por_cpf():
    """Função para adicionar colunas de dados de um arquivo Excel a outro com base no CPF."""
    print("\n[bold yellow]╔══ Iniciando Junção por CPF ══╗[/bold yellow]\n")

    # Recebe o primeiro arquivo (base)
    base_file_path = inquirer.text(
        message="Digite o caminho do arquivo base (.xlsx):"
    ).execute()

    try:
        base_df = pd.read_excel(base_file_path)
    except Exception as e:
        print(f"[bold red]✗ Erro ao carregar o arquivo base: {e}[/bold red]\n")
        return

    # Seleciona a coluna de CPF no arquivo base
    base_cpf_column = inquirer.select(
        message="Selecione a coluna de CPF no arquivo base:",
        choices=base_df.columns.tolist()
    ).execute()

    # Recebe o segundo arquivo (dados a adicionar)
    second_file_path = inquirer.text(
        message="Digite o caminho do segundo arquivo (.xlsx):"
    ).execute()

    try:
        second_df = pd.read_excel(second_file_path)
    except Exception as e:
        print(f"[bold red]✗ Erro ao carregar o segundo arquivo: {e}[/bold red]\n")
        return

    # Seleciona a coluna de CPF no segundo arquivo
    second_cpf_column = inquirer.select(
        message="Selecione a coluna de CPF no segundo arquivo:",
        choices=second_df.columns.tolist()
    ).execute()

    print("\n[cyan]Normalizando CPFs...[/cyan]")
    for _ in track(range(100), description="[cyan]Processando CPFs...[/cyan]"):
        pass

    # Normaliza os CPFs
    base_df[base_cpf_column] = base_df[base_cpf_column].astype(str).str.zfill(11)
    second_df[second_cpf_column] = second_df[second_cpf_column].astype(str).str.zfill(11)

    # Identifica os CPFs em comum
    print("\n[cyan]Filtrando apenas CPFs em comum...[/cyan]")
    common_cpfs = base_df[base_cpf_column].isin(second_df[second_cpf_column])
    base_df = base_df[common_cpfs]

    # Filtra os dados do segundo arquivo para os CPFs em comum
    merged_df = pd.merge(
        base_df,
        second_df,
        left_on=base_cpf_column,
        right_on=second_cpf_column,
        suffixes=("", "_from_second")
    )

    # Remove a coluna CPF duplicada do segundo arquivo
    merged_df.drop(columns=[second_cpf_column], inplace=True)

    # Contagem de registros
    total_cpfs_base = len(base_df)
    total_cpfs_second = len(second_df)
    total_cpfs_common = len(merged_df)

    print("\n[bold green]╔══ Resumo da Operação ══╗[/bold green]")
    print(f"[white]► CPFs no arquivo base:[/white]       {total_cpfs_base:,}")
    print(f"[white]► CPFs no segundo arquivo:[/white]   {total_cpfs_second:,}")
    print(f"[white]► CPFs em comum:[/white]            {total_cpfs_common:,}")

    # Pergunta o diretório para salvar
    output_dir = inquirer.text(
        message="Digite o caminho para salvar o arquivo atualizado:"
    ).execute()

    # Define o caminho do arquivo de saída
    output_file = os.path.join(output_dir, f'juncao_cpfs_{os.path.basename(base_file_path)}')

    try:
        merged_df.to_excel(output_file, index=False)
        print(f"\n[bold green]✓ Processo concluído com sucesso![bold green]")
        print(f"[dim]📁 Arquivo salvo em: {output_file}[dim]\n")
    except Exception as e:
        print(f"[bold red]✗ Erro ao salvar o arquivo: {e}[/bold red]\n")


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

def filter_agencies():
    """Função para filtrar agências bancárias com base em critérios específicos."""
    
    print("\n[bold yellow]╔══ Iniciando Filtro de Agências ══╗[/bold yellow]\n")

    # Recebe o arquivo Excel
    file_path = inquirer.text(
        message="Digite o caminho do arquivo Excel (.xlsx):"
    ).execute()

    try:
        df = pd.read_excel(file_path)
    except Exception as e:
        print(f"[bold red]✗ Erro ao carregar o arquivo: {e}[/bold red]\n")
        return

    # Seleciona a coluna de agência
    agency_column = inquirer.select(
        message="Selecione a coluna de agência bancária:",
        choices=df.columns.tolist()
    ).execute()

    print("\n[cyan]Filtrando agências...[/cyan]")

    # Critérios de filtragem
    initial_count = len(df)
    filtered_df = df[df[agency_column].astype(str).str.len() >= 4]
    filtered_df = filtered_df[filtered_df[agency_column].notnull()]

    final_count = len(filtered_df)
    removed_count = initial_count - final_count

    # Resumo da operação
    print("\n[bold green]╔══ Resumo da Operação ══╗[/bold green]")
    print(f"[white]► Registros originais:[/white]    {initial_count:,}")
    print(f"[white]► Registros removidos:[/white]    {removed_count:,}")
    print(f"[white]► Registros restantes:[/white]   {final_count:,}")

    # Pergunta o diretório para salvar
    output_dir = inquirer.text(
        message="Digite o caminho para salvar o arquivo filtrado:"
    ).execute()

    # Salva o arquivo filtrado
    output_file = os.path.join(output_dir, f"filtro_agencias_{os.path.basename(file_path)}")

    try:
        filtered_df.to_excel(output_file, index=False)
        print(f"\n[bold green]✓ Processo concluído com sucesso![bold green]")
        print(f"[dim]📁 Arquivo salvo em: {output_file}[dim]\n")
    except Exception as e:
        print(f"[bold red]✗ Erro ao salvar o arquivo: {e}[bold red]\n")


def map_columns_and_merge():
    """Função para mapear colunas de um modelo e preencher com dados de outro arquivo."""

    # Recebe o arquivo modelo
    print("\n[bold yellow]╔══ Iniciando Mapeamento de Colunas ══╗[/bold yellow]\n")
    model_file_path = inquirer.text(
        message="Digite o caminho do arquivo modelo (.xlsx):"
    ).execute()

    try:
        model_df = pd.read_excel(model_file_path)
        model_columns = model_df.columns.tolist()
    except Exception as e:
        print(f"[bold red]✗ Erro ao carregar o arquivo modelo: {e}[/bold red]\n")
        return

    if not model_columns:
        print("[bold red]✗ O arquivo modelo não possui cabeçalhos![bold red]\n")
        return

    # Recebe o arquivo com dados
    data_file_path = inquirer.text(
        message="Digite o caminho do arquivo de dados (.xlsx):"
    ).execute()

    try:
        data_df = pd.read_excel(data_file_path)
        data_columns = data_df.columns.tolist()
    except Exception as e:
        print(f"[bold red]✗ Erro ao carregar o arquivo de dados: {e}[/bold red]\n")
        return

    if not data_columns:
        print("[bold red]✗ O arquivo de dados não possui cabeçalhos![bold red]\n")
        return

    # Inicializa o DataFrame de saída com as mesmas colunas do modelo
    output_df = pd.DataFrame(columns=model_columns)

    # Mapeamento das colunas
    column_mapping = {}
    used_columns = set()
    print("\n[cyan]Mapeie as colunas do arquivo modelo com as do arquivo de dados:[/cyan]\n")

    for model_col in model_columns:
        available_columns = [col for col in data_columns if col not in used_columns] + ["Ignorar"]
        mapped_column = inquirer.select(
            message=f"Selecione a coluna correspondente para '{model_col}' no arquivo de dados:",
            choices=available_columns,
        ).execute()

        if mapped_column != "Ignorar":
            column_mapping[model_col] = mapped_column
            used_columns.add(mapped_column)

    # Preenchendo o DataFrame de saída com os dados mapeados
    for model_col, data_col in column_mapping.items():
        output_df[model_col] = data_df[data_col]

    # Exibindo resumo
    print("\n[bold green]╔══ Resumo da Operação ══╗[/bold green]")
    print(f"[white]► Linhas no arquivo modelo:[/white]       {len(model_df):,}")
    print(f"[white]► Linhas no arquivo de dados:[/white]    {len(data_df):,}")
    print(f"[white]► Colunas no arquivo modelo:[/white]     {len(model_columns):,}")
    print(f"[white]► Colunas no arquivo de dados:[/white]   {len(data_columns):,}")

    # Pergunta o diretório para salvar
    output_dir = inquirer.text(
        message="Digite o caminho para salvar o arquivo resultante:"
    ).execute()

    output_file = os.path.join(output_dir, f"resultado_{os.path.basename(model_file_path)}")

    try:
        output_df.to_excel(output_file, index=False)
        print(f"\n[bold green]✓ Processo concluído com sucesso![bold green]")
        print(f"[dim]📁 Arquivo salvo em: {output_file}[dim]\n")
    except Exception as e:
        print(f"[bold red]✗ Erro ao salvar o arquivo: {e}[bold red]\n")

def validate_address_number():
    """Valida números de endereço e preenche células vazias com 0, converte para texto no final."""
    print("\n[bold yellow]╔══ Iniciando Validação de Números de Endereço ══╗[/bold yellow]\n")

    # Recebe o caminho do arquivo
    file_path = inquirer.text(
        message="Digite o caminho do arquivo Excel (.xlsx):"
    ).execute()

    try:
        df = pd.read_excel(file_path)
    except Exception as e:
        print(f"[bold red]✗ Erro ao carregar o arquivo: {e}[/bold red]\n")
        return

    # Seleciona a coluna de números de endereço
    column_name = inquirer.select(
        message="Selecione a coluna que contém os números de endereço:",
        choices=df.columns.tolist()
    ).execute()

    print("\n[cyan]Validando números de endereço...[/cyan]")

    # Processando e preenchendo valores vazios
    for _ in track(range(100), description="[cyan]Processando...[/cyan]"):
        pass

    try:
        df[column_name] = df[column_name].fillna(0)
        df[column_name] = df[column_name].apply(lambda x: int(str(x).strip()) if str(x).strip().isdigit() else 0)
    except Exception as e:
        print(f"[bold red]✗ Erro durante a validação: {e}[/bold red]\n")
        return

    # Convertendo todas as células para texto
    df[column_name] = df[column_name].astype(str)

    # Exibindo resumo
    total_linhas = len(df)
    linhas_vazias = (df[column_name] == "0").sum()

    print("\n[bold green]╔══ Resumo da Validação ══╗[/bold green]")
    print(f"[white]► Total de linhas no arquivo:[/white] {total_linhas:,}")
    print(f"[white]► Linhas vazias na coluna:[/white]   {linhas_vazias:,}")

    # Pergunta o diretório para salvar
    output_dir = inquirer.text(
        message="Digite o caminho para salvar o arquivo validado:"
    ).execute()

    output_file = os.path.join(output_dir, f"validated_address_numbers_{os.path.basename(file_path)}")

    try:
        df.to_excel(output_file, index=False)
        print(f"\n[bold green]✓ Processo concluído com sucesso![bold green]")
        print(f"[dim]📁 Arquivo salvo em: {output_file}[dim]\n")
    except Exception as e:
        print(f"[bold red]✗ Erro ao salvar o arquivo: {e}[bold red]\n")

def delete_rows_with_empty_cells():
    """Remove linhas de um arquivo Excel onde a célula na coluna selecionada está vazia."""
    print("\n[bold yellow]╔══ Iniciando Remoção de Linhas com Células Vazias ══╗[/bold yellow]\n")

    # Recebe o caminho do arquivo
    file_path = inquirer.text(
        message="Digite o caminho do arquivo Excel (.xlsx):"
    ).execute()

    try:
        df = pd.read_excel(file_path)
    except Exception as e:
        print(f"[bold red]✗ Erro ao carregar o arquivo: {e}[/bold red]\n")
        return

    # Seleciona a coluna para verificar células vazias
    column_name = inquirer.select(
        message="Selecione a coluna para verificar células vazias:",
        choices=df.columns.tolist()
    ).execute()

    print("\n[cyan]Removendo linhas com células vazias...[/cyan]")

    try:
        # Remove linhas com células vazias na coluna selecionada
        initial_row_count = len(df)
        df = df.dropna(subset=[column_name])
        final_row_count = len(df)
        removed_rows = initial_row_count - final_row_count
    except Exception as e:
        print(f"[bold red]✗ Erro durante a remoção: {e}[/bold red]\n")
        return

    # Exibindo resumo
    print("\n[bold green]╔══ Resumo da Remoção ══╗[/bold green]")
    print(f"[white]► Total de linhas no arquivo original:[/white] {initial_row_count:,}")
    print(f"[white]► Linhas removidas:[/white]                 {removed_rows:,}")
    print(f"[white]► Total de linhas no arquivo final:[/white] {final_row_count:,}")

    # Pergunta o diretório para salvar
    output_dir = inquirer.text(
        message="Digite o caminho para salvar o arquivo atualizado:"
    ).execute()

    output_file = os.path.join(output_dir, f"rows_removed_{os.path.basename(file_path)}")

    try:
        df.to_excel(output_file, index=False)
        print(f"\n[bold green]✓ Processo concluído com sucesso![bold green]")
        print(f"[dim]📁 Arquivo salvo em: {output_file}[dim]\n")
    except Exception as e:
        print(f"[bold red]✗ Erro ao salvar o arquivo: {e}[bold red]\n")

def format_benefit_file():
    """Formata as colunas de sexo e tipo_beneficio em um arquivo Excel."""
    print("\n[bold yellow]╔══ Iniciando Formatação de Benefício ══╗[/bold yellow]\n")

    # Recebe o arquivo
    file_path = inquirer.text(
        message="Digite o caminho do arquivo Excel (.xlsx):"
    ).execute()

    try:
        df = pd.read_excel(file_path)
    except Exception as e:
        print(f"[bold red]✗ Erro ao carregar o arquivo: {e}[/bold red]\n")
        return

    # Seleciona a coluna de sexo
    sexo_column = inquirer.select(
        message="Selecione a coluna de sexo:",
        choices=df.columns.tolist()
    ).execute()

    # Seleciona a coluna de tipo_beneficio
    beneficio_column = inquirer.select(
        message="Selecione a coluna de tipo_beneficio:",
        choices=df.columns.tolist()
    ).execute()

    print("\n[cyan]Formatando dados...[/cyan]")

    # Formata a coluna de sexo
    for _ in track(range(50), description="[cyan]Formatando coluna de sexo...[/cyan]"):
        pass

    df[sexo_column] = df[sexo_column].replace({'M': 'Masculino', 'F': 'Feminino'})

    # Formata a coluna de tipo_beneficio
    for _ in track(range(50), description="[cyan]Formatando coluna de tipo_beneficio...[/cyan]"):
        pass

    df[beneficio_column] = df[beneficio_column].astype(str).str[:2]

    # Pergunta o diretório para salvar
    output_dir = inquirer.text(
        message="Digite o caminho para salvar o arquivo formatado:"
    ).execute()

    output_file = os.path.join(output_dir, f"format_benf_{os.path.basename(file_path)}")

    try:
        df.to_excel(output_file, index=False)
        print(f"\n[bold green]✓ Processo concluído com sucesso![bold green]")
        print(f"[dim]📁 Arquivo salvo em: {output_file}[dim]\n")
    except Exception as e:
        print(f"[bold red]✗ Erro ao salvar o arquivo: {e}[/bold red]\n")

def format_agency_column():
    """
    Formata uma coluna de agência, removendo o último dígito para agências com dois ou mais dígitos,
    substituindo valores vazios, nulos ou iguais a '0' por '1', e salvando o arquivo com prefixo 'agencia_format_'.
    """
    print("\n[bold yellow]╔══ Iniciando Formatação de Agências ══╗[/bold yellow]\n")

    # Recebe o caminho do arquivo
    file_path = inquirer.text(
        message="Digite o caminho do arquivo Excel (.xlsx):"
    ).execute()

    try:
        df = pd.read_excel(file_path)
    except Exception as e:
        print(f"[bold red]✗ Erro ao carregar o arquivo: {e}[bold red]\n")
        return

    # Seleciona a coluna de agência
    agency_column = inquirer.select(
        message="Selecione a coluna de agência:",
        choices=df.columns.tolist()
    ).execute()

    print("\n[cyan]Formatando valores da coluna de agência...[/cyan]")

    # Função para formatar os valores da coluna de agência
    def format_agency(value):
        if pd.isna(value) or str(value).strip() in ('', '0'):
            return '1'  # Substituir valores vazios, nulos ou iguais a 0 por '1'
        value = str(value).strip()  # Remove espaços
        if len(value) > 1:  # Se o valor tiver dois ou mais dígitos, remove o último dígito
            return value[:-3]
        return value

    # Aplica a formatação e conta alterações
    total_rows = len(df)
    original_column = df[agency_column].astype(str).copy()  # Copia os valores originais como string
    df[agency_column] = df[agency_column].apply(format_agency)
    modified_rows = (original_column != df[agency_column]).sum()  # Conta as linhas modificadas

    # Resumo da operação
    print("\n[bold green]╔══ Resumo da Formatação ══╗[/bold green]")
    print(f"[white]► Total de linhas no arquivo original:[/white] {total_rows:,}")
    print(f"[white]► Linhas modificadas:[/white] {modified_rows:,}")

    # Pergunta o diretório para salvar
    output_dir = inquirer.text(
        message="Digite o caminho para salvar o arquivo atualizado:"
    ).execute()

    # Define o caminho do arquivo de saída
    output_file = os.path.join(output_dir, f"agencia_format_{os.path.basename(file_path)}")

    # Salva o arquivo atualizado
    try:
        df.to_excel(output_file, index=False)
        print(f"\n[bold green]✓ Processo concluído com sucesso![bold green]")
        print(f"[dim]📁 Arquivo salvo em: {output_file}[dim]\n")
    except Exception as e:
        print(f"[bold red]✗ Erro ao salvar o arquivo: {e}[bold red]\n")

def validate_and_format_cep():
    """Valida, verifica existência e busca detalhes de CEPs usando a API OpenCEP."""
    print("\n[bold yellow]╔══ Iniciando Validação e Busca de CEP ══╗[/bold yellow]\n")

    # Recebe o caminho do arquivo
    file_path = inquirer.text(
        message="Digite o caminho do arquivo Excel (.xlsx):"
    ).execute()

    try:
        df = pd.read_excel(file_path)
    except Exception as e:
        print(f"[bold red]✗ Erro ao carregar o arquivo: {e}[bold red]\n")
        return

    # Seleciona as colunas necessárias
    cep_column = inquirer.select(
        message="Selecione a coluna que contém os CEPs:",
        choices=df.columns.tolist()
    ).execute()

    endereco_column = inquirer.select(
        message="Selecione a coluna de Endereço:",
        choices=df.columns.tolist()
    ).execute()

    bairro_column = inquirer.select(
        message="Selecione a coluna de Bairro:",
        choices=df.columns.tolist()
    ).execute()

    cidade_column = inquirer.select(
        message="Selecione a coluna de Cidade:",
        choices=df.columns.tolist()
    ).execute()

    estado_column = inquirer.select(
        message="Selecione a coluna de Estado:",
        choices=df.columns.tolist()
    ).execute()

    print("\n[cyan]Validando CEPs...[/cyan]")

    # Validação inicial do CEP
    def validate_cep(value):
        if pd.isna(value):
            return None
        cep = str(value).strip().replace("-", "")
        if len(cep) != 8 or not cep.isdigit():
            return None
        return cep

    # Aplicar validação
    df[cep_column] = df[cep_column].apply(validate_cep)

    # Remove linhas com CEP inválido
    initial_row_count = len(df)
    df_invalid = df[df[cep_column].isna()].copy()
    df = df.dropna(subset=[cep_column]).copy()

    print(f"[bold green]✓ Linhas removidas devido a CEPs inválidos: {len(df_invalid)}[/bold green]\n")

    # Fase 1: Verificar se o CEP existe
    print("[cyan]Verificando a existência dos CEPs...[/cyan]")

    def check_cep_exists(cep):
        try:
            response = requests.get(f"https://opencep.com/v1/{cep}.json", timeout=5)
            if response.status_code == 200:
                data = response.json()
                return "erro" not in data  # Retorna True se o CEP existir
            return False
        except Exception as e:
            logging.warning(f"Erro ao verificar CEP {cep}: {e}")
            return False

    # Adiciona uma nova coluna para marcar CEPs existentes
    df["EXISTE"] = False
    for index in track(df.index, description="[cyan]Verificando CEPs...[/cyan]"):
        cep = df.at[index, cep_column]
        df.at[index, "EXISTE"] = check_cep_exists(cep)
        print(f"Verificando linha {index + 1}/{len(df)} - CEP: {cep}")

    # Fase 2: Obter detalhes dos CEPs existentes
    print("[cyan]Buscando detalhes dos CEPs existentes...[/cyan]")
    valid_indices = []

    def fetch_cep_details(cep):
        try:
            response = requests.get(f"https://opencep.com/v1/{cep}.json")
            if response.status_code == 200:
                return response.json()
        except Exception as e:
            logging.warning(f"Erro ao buscar detalhes do CEP {cep}: {e}")
            return None

    for index in track(df.index, description="[cyan]Processando CEPs existentes...[/cyan]"):
        if df.at[index, "EXISTE"]:
            cep = df.at[index, cep_column]
            address_data = fetch_cep_details(cep)
            if address_data:
                valid_indices.append(index)
                df.at[index, endereco_column] = address_data.get("logradouro", df.at[index, endereco_column])
                df.at[index, bairro_column] = address_data.get("bairro", df.at[index, bairro_column])
                df.at[index, cidade_column] = address_data.get("localidade", df.at[index, cidade_column])
                df.at[index, estado_column] = address_data.get("uf", df.at[index, estado_column])
                print(f"Detalhes obtidos para CEP: {cep}")
            else:
                print(f"Falha ao buscar detalhes para o CEP: {cep}")

    # Remove CEPs não existentes do DataFrame
    df_invalid = pd.concat([df_invalid, df[~df["EXISTE"]]])
    df_valid = df.loc[valid_indices].copy()
    df.drop(columns=["EXISTE"], inplace=True)

    # Resumo final
    print("\n[bold green]╔══ Resumo Final ══╗[/bold green]")
    print(f"[white]► Total de linhas no arquivo original:[/white] {initial_row_count:,}")
    print(f"[white]► CEPs válidos encontrados e detalhados:[/white] {len(df_valid):,}")
    print(f"[white]► Linhas removidas (CEPs inválidos ou inexistentes):[/white] {len(df_invalid):,}")

    # Pergunta o diretório para salvar
    output_dir = inquirer.text(
        message="Digite o caminho para salvar os arquivos:"
    ).execute()

    # Caminhos para os arquivos de saída
    valid_output_file = os.path.join(output_dir, f"cep_validos_{os.path.basename(file_path)}")
    invalid_output_file = os.path.join(output_dir, f"cep_invalidos_{os.path.basename(file_path)}")

    # Salva os arquivos
    try:
        df_valid.to_excel(valid_output_file, index=False)
        df_invalid.to_excel(invalid_output_file, index=False)
        print(f"\n[bold green]✓ Arquivos salvos com sucesso![bold green]")
        print(f"[dim]📁 Arquivo com CEPs válidos salvo em: {valid_output_file}[dim]")
        print(f"[dim]📁 Arquivo com CEPs inválidos salvo em: {invalid_output_file}[dim]\n")
    except Exception as e:
        print(f"[bold red]✗ Erro ao salvar os arquivos: {e}[bold red]\n")

def validador_de_bancos():
    """
    Valida colunas de banco, agência e conta.
    Remove linhas que não atendem aos critérios:
    - Banco: 1 a 3 dígitos
    - Agência: 1 a 4 dígitos
    - Conta: Não pode ter letras, espaços ou estar vazia.
    """
    print("\n[bold yellow]╔══ Iniciando Validação de Banco ══╗[/bold yellow]\n")

    # Recebe o caminho do arquivo
    file_path = inquirer.text(
        message="Digite o caminho do arquivo Excel (.xlsx):"
    ).execute()

    try:
        df = pd.read_excel(file_path)
    except Exception as e:
        print(f"[bold red]✗ Erro ao carregar o arquivo: {e}[bold red]\n")
        return

    # Seleciona as colunas necessárias
    banco_column = inquirer.select(
        message="Selecione a coluna de Banco:",
        choices=df.columns.tolist()
    ).execute()

    agencia_column = inquirer.select(
        message="Selecione a coluna de Agência:",
        choices=df.columns.tolist()
    ).execute()

    conta_column = inquirer.select(
        message="Selecione a coluna de Conta:",
        choices=df.columns.tolist()
    ).execute()

    print("\n[cyan]Validando dados...[/cyan]")

    # Funções de validação
    def is_valid_banco(value):
        return str(value).isdigit() and 1 <= len(str(value)) <= 3

    def is_valid_agencia(value):
        return str(value).isdigit() and 1 <= len(str(value)) <= 4

    def is_valid_conta(value):
        return str(value).isdigit() and len(str(value)) > 0

    # Inicializa contadores
    initial_row_count = len(df)

    # Aplica validação para todas as colunas e filtra as linhas inválidas
    df["VALIDO"] = df[banco_column].apply(is_valid_banco) & \
                   df[agencia_column].apply(is_valid_agencia) & \
                   df[conta_column].apply(is_valid_conta)

    df_invalid = df[~df["VALIDO"]].copy()  # Linhas inválidas
    df = df[df["VALIDO"]].copy()           # Linhas válidas

    # Remove a coluna auxiliar "VALIDO"
    df.drop(columns=["VALIDO"], inplace=True)
    df_invalid.drop(columns=["VALIDO"], inplace=True)

    # Resumo da validação
    linhas_invalidas = len(df_invalid)
    linhas_validas = len(df)

    print("\n[bold green]╔══ Resumo da Validação ══╗[/bold green]")
    print(f"[white]► Linhas originais:[/white]    {initial_row_count:,}")
    print(f"[white]► Linhas válidas:[/white]      {linhas_validas:,}")
    print(f"[white]► Linhas inválidas:[/white]    {linhas_invalidas:,}")

    # Pergunta o diretório para salvar
    output_dir = inquirer.text(
        message="Digite o caminho para salvar os arquivos filtrados:"
    ).execute()

    # Adiciona prefixo aos nomes dos arquivos de saída
    valid_output_file = os.path.join(output_dir, f"filtrar_bank_validos_{os.path.basename(file_path)}")
    invalid_output_file = os.path.join(output_dir, f"filtrar_bank_invalidos_{os.path.basename(file_path)}")

    # Salva os arquivos
    try:
        df.to_excel(valid_output_file, index=False)
        df_invalid.to_excel(invalid_output_file, index=False)
        print(f"\n[bold green]✓ Arquivos salvos com sucesso![bold green]")
        print(f"[dim]📁 Arquivo com dados válidos salvo em: {valid_output_file}[dim]")
        print(f"[dim]📁 Arquivo com dados inválidos salvo em: {invalid_output_file}[dim]\n")
    except Exception as e:
        print(f"[bold red]✗ Erro ao salvar os arquivos: {e}[bold red]\n")

def validate_sex_column():
    """Valida a coluna de sexo, convertendo 'M' e 'F' para 'Masculino' e 'Feminino',
    removendo linhas com valores inválidos."""
    print("\n[bold yellow]╔══ Iniciando Validação da Coluna de Sexo ══╗[/bold yellow]\n")

    # Recebe o caminho do arquivo
    file_path = inquirer.text(
        message="Digite o caminho do arquivo Excel (.xlsx):"
    ).execute()

    try:
        df = pd.read_excel(file_path)
    except Exception as e:
        print(f"[bold red]✗ Erro ao carregar o arquivo: {e}[/bold red]\n")
        return

    # Seleciona a coluna de sexo
    column_name = inquirer.select(
        message="Selecione a coluna que contém os valores de sexo:",
        choices=df.columns.tolist()
    ).execute()

    print("\n[cyan]Validando a coluna de sexo...[/cyan]")

    # Processando os valores
    valid_sex_values = {"M": "Masculino", "F": "Feminino"}
    try:
        df[column_name] = df[column_name].apply(lambda x: valid_sex_values.get(str(x).strip(), x))

        # Filtra as linhas válidas
        valid_rows = df[column_name].isin(["Masculino", "Feminino"])
        filtered_df = df[valid_rows].copy()

        invalid_rows_count = len(df) - len(filtered_df)

    except Exception as e:
        print(f"[bold red]✗ Erro durante a validação: {e}[/bold red]\n")
        return

    # Exibindo resumo
    total_linhas = len(df)
    linhas_validas = len(filtered_df)

    print("\n[bold green]╔══ Resumo da Validação ══╗[/bold green]")
    print(f"[white]► Total de linhas no arquivo:[/white] {total_linhas:,}")
    print(f"[white]► Linhas válidas:[/white] {linhas_validas:,}")
    print(f"[white]► Linhas removidas:[/white] {invalid_rows_count:,}")

    # Pergunta o diretório para salvar
    output_dir = inquirer.text(
        message="Digite o caminho para salvar o arquivo validado:"
    ).execute()

    output_file = os.path.join(output_dir, f"validated_sex_column_{os.path.basename(file_path)}")

    try:
        filtered_df.to_excel(output_file, index=False)
        print(f"\n[bold green]✓ Processo concluído com sucesso![bold green]")
        print(f"[dim]📁 Arquivo salvo em: {output_file}[dim]\n")
    except Exception as e:
        print(f"[bold red]✗ Erro ao salvar o arquivo: {e}[bold red]\n")

def extract_ddd_and_number():
    """Função para extrair DDD e número de uma coluna de celular."""
    print("\n[bold yellow]╔══ Iniciando Extração de DDD e Número ══╗[/bold yellow]\n")

    # Recebe o arquivo
    file_path = inquirer.text(
        message="Digite o caminho do arquivo Excel (.xlsx):"
    ).execute()

    try:
        df = pd.read_excel(file_path)
    except Exception as e:
        print(f"[bold red]✗ Erro ao carregar o arquivo: {e}[/bold red]\n")
        return

    # Seleciona a coluna de números de celular
    phone_column = inquirer.select(
        message="Selecione a coluna que contém os números de celular (DDD+Número):",
        choices=df.columns.tolist()
    ).execute()

    # Seleciona a coluna de saída para DDD
    ddd_column = inquirer.select(
        message="Selecione a coluna onde será inserido o DDD extraído:",
        choices=df.columns.tolist()
    ).execute()

    # Inicializa contadores
    total_registros = len(df)
    registros_validos = 0
    registros_invalidos = 0

    # Processa cada linha e separa o DDD do número
    def process_phone(value):
        nonlocal registros_validos, registros_invalidos
        if pd.isna(value):
            registros_invalidos += 1
            return None, None

        value = str(value).strip()
        if len(value) == 11 and value.isdigit():
            registros_validos += 1
            return value[:2], value[2:]
        else:
            registros_invalidos += 1
            return None, None

    print("\n[cyan]Processando números...[/cyan]")

    df[ddd_column], df[phone_column] = zip(*df[phone_column].apply(process_phone))

    # Exibe resumo da operação
    print("\n[bold green]╔══ Resumo da Operação ══╗[/bold green]")
    print(f"[white]► Registros totais:[/white]    {total_registros:,}")
    print(f"[white]► Registros válidos:[/white]   {registros_validos:,}")
    print(f"[white]► Registros inválidos:[/white] {registros_invalidos:,}")

    # Pergunta o diretório para salvar
    output_dir = inquirer.text(
        message="Digite o caminho para salvar o arquivo atualizado:"
    ).execute()

    output_file = os.path.join(output_dir, f"extracted_number_ddd_{os.path.basename(file_path)}")

    try:
        df.to_excel(output_file, index=False)
        print(f"\n[bold green]✓ Processo concluído com sucesso![bold green]")
        print(f"[dim]📁 Arquivo salvo em: {output_file}[dim]\n")
    except Exception as e:
        print(f"[bold red]✗ Erro ao salvar o arquivo: {e}[bold red]\n")

def whitelist_blacklist_removal():
    """
    Remove linhas do arquivo base que possuem números contidos no arquivo de blacklist.
    """
    print("\n[bold yellow]╔══ Remoção de Linhas com Números na Blacklist ══╗[/bold yellow]\n")

    # Recebe o caminho do arquivo base
    base_file_path = inquirer.text(
        message="Digite o caminho do arquivo base (.xlsx):"
    ).execute()

    try:
        base_df = pd.read_excel(base_file_path)
    except Exception as e:
        print(f"[bold red]✗ Erro ao carregar o arquivo base: {e}[bold red]\n")
        return

    # Seleciona a coluna de números no arquivo base
    base_number_column = inquirer.select(
        message="Selecione a coluna de números do arquivo base:",
        choices=base_df.columns.tolist()
    ).execute()

    # Recebe o caminho do arquivo de blacklist
    blacklist_file_path = inquirer.text(
        message="Digite o caminho do arquivo de blacklist (.xlsx):"
    ).execute()

    try:
        blacklist_df = pd.read_excel(blacklist_file_path)
    except Exception as e:
        print(f"[bold red]✗ Erro ao carregar o arquivo de blacklist: {e}[bold red]\n")
        return

    # Seleciona a coluna de números no arquivo de blacklist
    blacklist_number_column = inquirer.select(
        message="Selecione a coluna de números do arquivo de blacklist:",
        choices=blacklist_df.columns.tolist()
    ).execute()

    print("\n[cyan]Removendo números contidos na blacklist...[/cyan]")

    # Converte a coluna de números da blacklist em um conjunto para busca rápida
    blacklist_numbers = set(blacklist_df[blacklist_number_column].astype(str).str.strip())

    # Filtra as linhas no arquivo base
    initial_row_count = len(base_df)
    base_df["VALIDO"] = ~base_df[base_number_column].astype(str).str.strip().isin(blacklist_numbers)

    valid_df = base_df[base_df["VALIDO"]].drop(columns=["VALIDO"]).copy()  # Linhas válidas
    invalid_df = base_df[~base_df["VALIDO"]].drop(columns=["VALIDO"]).copy()  # Linhas removidas

    # Resumo da remoção
    linhas_removidas = len(invalid_df)
    linhas_restantes = len(valid_df)

    print("\n[bold green]╔══ Resumo da Remoção ══╗[/bold green]")
    print(f"[white]► Total de linhas no arquivo base:[/white] {initial_row_count:,}")
    print(f"[white]► Linhas removidas:[/white]              {linhas_removidas:,}")
    print(f"[white]► Linhas restantes:[/white]             {linhas_restantes:,}")

    # Pergunta o diretório para salvar os arquivos
    output_dir = inquirer.text(
        message="Digite o caminho para salvar os arquivos filtrados:"
    ).execute()

    # Adiciona prefixos aos nomes dos arquivos de saída
    valid_output_file = os.path.join(output_dir, f"whitelist_{os.path.basename(base_file_path)}")
    invalid_output_file = os.path.join(output_dir, f"blacklist_{os.path.basename(base_file_path)}")

    # Salva os arquivos
    try:
        valid_df.to_excel(valid_output_file, index=False)
        invalid_df.to_excel(invalid_output_file, index=False)
        print(f"\n[bold green]✓ Arquivos salvos com sucesso![bold green]")
        print(f"[dim]📁 Arquivo com números válidos salvo em: {valid_output_file}[dim]")
        print(f"[dim]📁 Arquivo com números removidos salvo em: {invalid_output_file}[dim]\n")
    except Exception as e:
        print(f"[bold red]✗ Erro ao salvar os arquivos: {e}[bold red]\n")

def filter_num_nine():
    """
    Formata números de celular adicionando o dígito '9' após o DDD em números de 12 dígitos.
    Remove linhas com números que não possuem 12 ou 13 dígitos.
    """
    print("\n[bold yellow]╔══ Formatação de Números com '9' ══╗[/bold yellow]\n")
    print("[bold cyan]Observação: Certifique-se de que os números estejam no formato correto, começando com '55' seguido do DDD e número.[/bold cyan]\n")

    # Recebe o caminho do arquivo
    file_path = inquirer.text(
        message="Digite o caminho do arquivo Excel (.xlsx):"
    ).execute()

    try:
        df = pd.read_excel(file_path)
    except Exception as e:
        print(f"[bold red]✗ Erro ao carregar o arquivo: {e}[bold red]\n")
        return

    # Seleciona a coluna de números
    number_column = inquirer.select(
        message="Selecione a coluna de números de celular:",
        choices=df.columns.tolist()
    ).execute()

    print("\n[cyan]Formatando números...[/cyan]")

    # Função para verificar e corrigir números
    def format_number(value):
        try:
            value = str(value).strip()
            if len(value) == 12:  # Número com 12 dígitos (faltando o 9)
                return value[:4] + "9" + value[4:]
            elif len(value) == 13:  # Número já no formato correto
                return value
            return None  # Número inválido
        except Exception:
            return None

    # Aplica a formatação e filtra números inválidos
    initial_row_count = len(df)
    df[number_column] = df[number_column].apply(format_number)

    df_invalid = df[df[number_column].isna()].copy()  # Números inválidos
    df = df.dropna(subset=[number_column]).copy()     # Números válidos

    # Resumo da formatação
    linhas_invalidas = len(df_invalid)
    linhas_validas = len(df)

    print("\n[bold green]╔══ Resumo da Formatação ══╗[/bold green]")
    print(f"[white]► Linhas originais:[/white] {initial_row_count:,}")
    print(f"[white]► Números formatados:[/white] {linhas_validas:,}")
    print(f"[white]► Linhas removidas (números inválidos):[/white] {linhas_invalidas:,}")

    # Pergunta o diretório para salvar
    output_dir = inquirer.text(
        message="Digite o caminho para salvar os arquivos formatados:"
    ).execute()

    # Adiciona prefixo aos nomes dos arquivos de saída
    valid_output_file = os.path.join(output_dir, f"filtrer_num_nine_validos_{os.path.basename(file_path)}")
    invalid_output_file = os.path.join(output_dir, f"filtrer_num_nine_invalidos_{os.path.basename(file_path)}")

    # Salva os arquivos
    try:
        df.to_excel(valid_output_file, index=False)
        df_invalid.to_excel(invalid_output_file, index=False)
        print(f"\n[bold green]✓ Arquivos salvos com sucesso![bold green]")
        print(f"[dim]📁 Arquivo com números válidos salvo em: {valid_output_file}[dim]")
        print(f"[dim]📁 Arquivo com números inválidos salvo em: {invalid_output_file}[dim]\n")
    except Exception as e:
        print(f"[bold red]✗ Erro ao salvar os arquivos: {e}[bold red]\n")

def filter_back_age():
    """
    Valida banco, agência e conta e remove linhas que atendem aos critérios de remoção:
    - Contém letras
    - Contém espaços
    - Está vazio
    - É igual a zero
    """
    print("\n[bold yellow]╔══ Iniciando Validação de Banco, Agência e Conta ══╗[/bold yellow]\n")

    # Recebe o caminho do arquivo
    file_path = inquirer.text(
        message="Digite o caminho do arquivo Excel (.xlsx):"
    ).execute()

    try:
        df = pd.read_excel(file_path)
    except Exception as e:
        print(f"[bold red]✗ Erro ao carregar o arquivo: {e}[bold red]\n")
        return

    # Seleciona as colunas necessárias
    banco_column = inquirer.select(
        message="Selecione a coluna de Banco:",
        choices=df.columns.tolist()
    ).execute()

    agencia_column = inquirer.select(
        message="Selecione a coluna de Agência:",
        choices=df.columns.tolist()
    ).execute()

    conta_column = inquirer.select(
        message="Selecione a coluna de Conta:",
        choices=df.columns.tolist()
    ).execute()

    print("\n[cyan]Validando dados...[/cyan]")

    # Função de validação
    def is_invalid(value):
        """Verifica se o valor contém letras, espaços, está vazio ou é igual a zero."""
        if pd.isna(value) or str(value).strip() == "" or str(value).strip() == "0":
            return True
        if any(char.isalpha() for char in str(value)) or " " in str(value):
            return True
        return False

    # Aplica a validação e filtra as linhas inválidas
    initial_row_count = len(df)
    df["VALIDO"] = ~(
        df[banco_column].apply(is_invalid) |
        df[agencia_column].apply(is_invalid) |
        df[conta_column].apply(is_invalid)
    )

    df_invalid = df[~df["VALIDO"]].copy()  # Linhas inválidas
    df = df[df["VALIDO"]].copy()           # Linhas válidas

    # Remove a coluna auxiliar "VALIDO"
    df.drop(columns=["VALIDO"], inplace=True)
    df_invalid.drop(columns=["VALIDO"], inplace=True)

    # Resumo da validação
    linhas_invalidas = len(df_invalid)
    linhas_validas = len(df)

    print("\n[bold green]╔══ Resumo da Validação ══╗[/bold green]")
    print(f"[white]► Linhas originais:[/white] {initial_row_count:,}")
    print(f"[white]► Linhas válidas:[/white]   {linhas_validas:,}")
    print(f"[white]► Linhas inválidas:[/white] {linhas_invalidas:,}")

    # Pergunta o diretório para salvar
    output_dir = inquirer.text(
        message="Digite o caminho para salvar os arquivos filtrados:"
    ).execute()

    # Adiciona prefixo aos nomes dos arquivos de saída
    valid_output_file = os.path.join(output_dir, f"filter_back_age_validos_{os.path.basename(file_path)}")
    invalid_output_file = os.path.join(output_dir, f"filter_back_age_invalidos_{os.path.basename(file_path)}")

    # Salva os arquivos
    try:
        df.to_excel(valid_output_file, index=False)
        df_invalid.to_excel(invalid_output_file, index=False)
        print(f"\n[bold green]✓ Arquivos salvos com sucesso![bold green]")
        print(f"[dim]📁 Arquivo com dados válidos salvo em: {valid_output_file}[dim]")
        print(f"[dim]📁 Arquivo com dados inválidos salvo em: {invalid_output_file}[dim]\n")
    except Exception as e:
        print(f"[bold red]✗ Erro ao salvar os arquivos: {e}[bold red]\n")

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
                Choice("6", "Mapeamento de Colunas"),
                Choice("7", "Formatação de Datas"),
                Choice("8", "Buscar e Validar CEPs"),
                Choice("9", "Sair")
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
            map_columns_and_merge()
        elif choice == "7":
            formatar_coluna_data()
        elif choice == "8":
            validate_and_format_cep()
        elif choice == "9":
            print("Programa encerrado!")
            break

def filtros_unicos():
    while True:
        choice = inquirer.select(
            message="Selecione um filtro único:",
            choices=[
                Choice("1", "Filtrar Excel (único)"),
                Choice("2", "Filtrar valores numéricos"),
                Choice("3", "Extração de DDD e Números"),
                Choice("4", "Filtrar Agências"),
                Choice("5", "Validador de Bancos"),
                Choice("6", "Validador Banco, Agência e Conta"),  # Novo filtro adicionado
                Choice("7", "Voltar")
            ]
        ).execute()

        if choice == "1":
            filter_single_excel()
        elif choice == "2":
            filter_numeric()
        elif choice == "3":
            extract_ddd_and_number()
        elif choice == "4":
            filter_agencies()
        elif choice == "5":
            validador_de_bancos()
        elif choice == "6":
            filter_back_age()  # Chamada para a nova função
        elif choice == "7":
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
                Choice("4", "Remover Linhas com Células Vazias"),
                Choice("5", "Remover Números da Blacklist"),  # Nova funcionalidade adicionada
                Choice("6", "Voltar")
            ]
        ).execute()

        if choice == "1":
            filter_cpf_removal()
        elif choice == "2":
            filter_remove_by_name()
        elif choice == "3":
            filter_phone_numbers_csv()
        elif choice == "4":
            delete_rows_with_empty_cells()
        elif choice == "5":
            whitelist_blacklist_removal()  # Chamada para a nova função
        elif choice == "6":
            break


def adicoes_unificacoes():
    while True:
        choice = inquirer.select(
            message="Selecione uma adição ou unificação:",
            choices=[
                Choice("1", "Unificar arquivos Excel"),
                Choice("2", "Unificar arquivos Excel com base no CPF"),
                Choice("3", "Adicionar dados de CPFs entre arquivos"),
                Choice("4", "Voltar")
            ]
        ).execute()

        if choice == "1":
            unify_excel_files()
        elif choice == "2":
            unify_excel_files_with_cpf()
        elif choice == "3":
            adicionar_dados_por_cpf()
        elif choice == "4":
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
                Choice("5", "Formatar Benefícios"),
                Choice("6", "Validar Número de Endereço"),
                Choice("7", "Validar Coluna de Sexo"),
                Choice("8", "Formatar Coluna de Agência"),
                Choice("9", "Formatar Números de Celular sem '9'"),  # Nova funcionalidade adicionada
                Choice("10", "Voltar")
            ]
        ).execute()

        if choice == "1":
            adjust_cpfs_to_11_digits()
        elif choice == "2":
            format_values_to_money()
        elif choice == "3":
            format_numbers_with_prefix()
        elif choice == "4":
            filter_and_format_rgs()
        elif choice == "5":
            format_benefit_file()
        elif choice == "6":
            validate_address_number()
        elif choice == "7":
            validate_sex_column()
        elif choice == "8":
            format_agency_column()
        elif choice == "9":
            filter_num_nine()  # Chamada para a nova função
        elif choice == "10":
            break


if __name__ == "__main__":
    main()