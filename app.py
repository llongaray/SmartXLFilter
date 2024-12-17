from InquirerPy import inquirer
from InquirerPy.base.control import Choice
import pandas as pd
import os
from rich import print
from rich.progress import track
import time

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
        print("\n[bold yellow]╔══ Iniciando Unificação por CPF ══╗[/bold yellow]\n")
        
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

def main():
    while True:
        choice = inquirer.select(
            message="Selecione uma opção:",
            choices=[
                Choice("1", "Filtrar Excel (único)"),
                Choice("2", "Filtrar Excel (múltiplo)"),
                Choice("3", "Manter colunas selecionadas"),
                Choice("4", "Remover colunas selecionadas"),
                Choice("5", "Filtrar valores numéricos"),
                Choice("6", "Unificar arquivos Excel"),
                Choice("7", "Unificar arquivos Excel com base no CPF"),
                Choice("8", "Filtrar CPF - Remoção"),
                Choice("9", "Filtrar CPF - Duplicidade"),
                Choice("10", "Sair")
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
        elif choice == "5":
            filter_numeric()
        elif choice == "6":
            unify_excel_files()
        elif choice == "7":
            unify_excel_files_with_cpf()
        elif choice == "8":
            filter_cpf_removal()
        elif choice == "9":
            filter_cpf_duplicates()
        elif choice == "10":
            print("Programa encerrado!")
            break

if __name__ == "__main__":
    main()
