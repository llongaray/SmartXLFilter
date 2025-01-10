# Guia de Uso - SmartXLFilter

## Introdução
Este guia detalha o passo a passo para utilizar cada funcionalidade do programa SmartXLFilter, que auxilia na manipulação de arquivos Excel com foco em filtragem, remoção, adições e formatações.

---

## Executando o Programa
1. Abra seu terminal ou prompt de comando.
2. Navegue até o diretório onde o programa está localizado.
3. Execute o comando:
   ```bash
   python app.py
   ```

---

## Menu Principal
O menu é dividido em cinco categorias principais:

1. **Filtros Únicos**
2. **Filtros Múltiplos**
3. **Remoções**
4. **Adições/Unificações**
5. **Formatações**
6. **Sair**

Escolha uma opção para acessar as funcionalidades relacionadas.

---

## Filtros Únicos

### 1. Filtrar Excel (Único)
- **Uso:** Filtra registros de um arquivo Excel com base em um valor de uma coluna.
- **Passos:**
  1. Informe o caminho do arquivo Excel.
  2. Escolha a coluna para aplicar o filtro.
  3. Selecione o valor desejado.
  4. Informe o diretório onde salvar o arquivo filtrado.

### 2. Filtrar Valores Numéricos
- **Uso:** Filtra valores numéricos maiores que ou entre limites especificados.
- **Passos:**
  1. Informe o caminho do arquivo Excel.
  2. Escolha a coluna numérica.
  3. Escolha o tipo de filtro (maior que ou entre valores).
  4. Informe o valor ou os limites do filtro.
  5. Informe o diretório onde salvar o arquivo filtrado.

---

## Filtros Múltiplos

### Filtrar Excel (Múltiplo)
- **Uso:** Aplica diversos filtros em cascata.
- **Passos:**
  1. Informe o caminho do arquivo Excel.
  2. Selecione múltiplas colunas e valores para filtrar.
  3. Informe o diretório onde salvar o arquivo filtrado.

---

## Remoções

### 1. Filtrar CPF - Remoção
- **Uso:** Remove registros de um arquivo base cujos CPFs estão em outro arquivo.
- **Passos:**
  1. Informe o arquivo base e selecione a coluna de CPF.
  2. Informe o arquivo de remoção e selecione a coluna de CPF.
  3. Informe o diretório onde salvar o arquivo filtrado.

### 2. Filtrar e Remover por Nome
- **Uso:** Remove registros de um arquivo base cujos nomes estão em uma blacklist.
- **Passos:**
  1. Informe o arquivo base e selecione a coluna de nomes.
  2. Informe o arquivo de blacklist e selecione a coluna de nomes.
  3. Informe o diretório onde salvar o arquivo filtrado.

### 3. Remover Números Fixos e Células Vazias
- **Uso:** Remove números de telefone fixo ou células vazias em arquivos CSV.
- **Passos:**
  1. Informe o caminho do arquivo CSV.
  2. Selecione a coluna de números.
  3. Informe o diretório onde salvar o arquivo filtrado.

---

## Adições/Unificações

### 1. Unificar Arquivos Excel
- **Uso:** Combina múltiplos arquivos Excel em um único arquivo, removendo duplicatas.
- **Passos:**
  1. Informe o diretório contendo os arquivos Excel.
  2. Informe o diretório onde salvar o arquivo unificado.

### 2. Unificar Arquivos com Base no CPF
- **Uso:** Combina dois arquivos Excel, unificando registros com base no CPF.
- **Passos:**
  1. Informe o arquivo base e selecione a coluna de CPF.
  2. Informe o segundo arquivo e selecione a coluna de CPF.
  3. Informe o diretório onde salvar o arquivo unificado.

---

## Formatações

### 1. Ajustar CPFs para 11 Dígitos
- **Uso:** Normaliza e ajusta CPFs para o padrão de 11 dígitos.
- **Passos:**
  1. Informe o arquivo Excel e selecione a coluna de CPFs.
  2. Informe o diretório onde salvar o arquivo ajustado.

### 2. Formatar Coluna de Valores para Padrão Monetário
- **Uso:** Formata valores em uma coluna para o padrão monetário brasileiro (R$).
- **Passos:**
  1. Informe o arquivo Excel e selecione a coluna de valores.
  2. Informe o diretório onde salvar o arquivo formatado.

### 3. Formatar Números com Prefixo '55'
- **Uso:** Adiciona o prefixo '55' a números com 11 dígitos.
- **Passos:**
  1. Informe o arquivo Excel e selecione a coluna de números.
  2. Informe o diretório onde salvar o arquivo formatado.

### 4. Filtrar e Formatar RGs
- **Uso:** Filtra RGs inválidos e ajusta os válidos para o formato de 10 dígitos.
- **Passos:**
  1. Informe o arquivo base e selecione a coluna de RGs.
  2. Informe o diretório onde salvar os arquivos de RGs válidos e inválidos.

---

## Observações
- **Cuidado com arquivos originais:** Sempre utilize cópias dos arquivos para evitar perda de dados.
- **Nomes de colunas:** Certifique-se de que as colunas selecionadas contêm os dados esperados.
- **Erros:** Caso encontre algum erro, verifique o log gerado no arquivo `app.log`.

---

Para mais informações, consulte a documentação oficial ou entre em contato com o suporte.

