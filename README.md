# SmartXLFilter

SmartXLFilter é uma ferramenta Python interativa para manipulação inteligente de arquivos Excel, oferecendo uma interface de linha de comando (CLI) intuitiva para filtrar e gerenciar dados de planilhas.

## Funcionalidades

- **Filtro Único**: Filtra dados baseado em uma única coluna e valor
- **Filtros Múltiplos**: Aplica múltiplos filtros em cascata
- **Gerenciamento de Colunas**: 
  - Manter apenas colunas específicas
  - Remover colunas selecionadas
- **Interface Interativa**: Utiliza prompts interativos para guiar o usuário
- **Preservação de Dados**: Sempre gera novos arquivos, mantendo o original intacto

## Pré-requisitos

- Python 3.6 ou superior
- pandas
- InquirerPy

## Instalação

1. Clone o repositório
2. Instale as dependências necessárias

## Como Usar

1. Execute o programa
2. Selecione uma das opções disponíveis:
   - Filtrar Excel (único)
   - Filtrar Excel (múltiplo)
   - Manter colunas selecionadas
   - Remover colunas selecionadas
3. Siga as instruções interativas na tela

## Exemplos de Uso

### Filtro Único
- Ideal para filtrar dados por uma categoria específica
- Ex.: Filtrar vendas por região

### Filtros Múltiplos
- Permite combinar vários critérios de filtro
- Ex.: Filtrar vendas por região E por produto

### Gerenciamento de Colunas
- Útil para extrair ou remover informações específicas
- Ex.: Manter apenas colunas relevantes para análise

## Boas Práticas

1. Sempre mantenha backup dos arquivos originais
2. Verifique o caminho dos arquivos antes de executar
3. Certifique-se de ter permissões de escrita no diretório de saída

## Contribuindo

Contribuições são bem-vindas! Por favor, sinta-se à vontade para submeter pull requests.

1. Faça um Fork do projeto
2. Crie sua Feature Branch
3. Commit suas mudanças
4. Push para a Branch
5. Abra um Pull Request

## Licença

Este projeto está licenciado sob a Licença MIT - veja o arquivo LICENSE para detalhes.

## Suporte

Para reportar bugs ou sugerir novas funcionalidades, por favor abra uma issue no GitHub.

## Autor

Leonardo Longaray dos Santos

---

**Nota**: Este projeto foi desenvolvido com o objetivo de simplificar a manipulação de arquivos Excel através de uma interface amigável e intuitiva.