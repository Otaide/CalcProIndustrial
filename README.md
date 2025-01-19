# CalcProIndustrial
Ao longo da candeia de suprimento de um esteira de produção, temos a fase do abastecimento de insumos baseado em ordens de produção (Op's). Para atender essas solicitações,o setor Almoxarifado precisa de uma ferramenta que permita calcular as quantidades de cada matéria-prima. Para tanto,a presente ferramenta se apresenta  como solução prática.

# Calculadora de Fórmulas - Documentação Completa

## 1. Visão Geral
A Calculadora de Fórmulas é um aplicativo desktop desenvolvido em Python para auxiliar no cálculo e gerenciamento de fórmulas químicas/industriais. O sistema permite calcular quantidades, manter histórico e exportar resultados.

## 2. Funcionalidades Principais

### 2.1 Calculadora Principal
- Pesquisa de fórmulas por número
- Cálculo automático baseado no peso informado
- Visualização dos resultados em tempo real
- Soma automática dos kg calculados
- Filtro de resultados por descrição

### 2.2 Histórico
- Registro automático dos cálculos realizados
- Visualização do histórico completo
- Edição de registros salvos
- Exclusão de registros
- Exportação para Excel
- Filtro de resultados detalhados
- Soma dos kg filtrados

## 3. Interface do Usuário

### 3.1 Tela Principal
- **Barra superior**:
  - Campo para nome da programação
  - Botão "Histórico"
  - Botão "Informações"
- **Seção de Fórmulas**:
  - Campo de pesquisa por número
  - Lista de fórmulas disponíveis
  - Checkbox para seleção
  - Campo para entrada de peso
- **Seção de Resultados**:
  - Tabela com resultados calculados
  - Filtro por descrição
  - Soma total dos kg

### 3.2 Tela de Histórico
- Lista de registros salvos
- Resultados detalhados
- Botões de ação (Editar, Excluir, Exportar)
- Filtro de resultados
- Soma dos kg filtrados

## 4. Banco de Dados

### 4.1 Arquivo `bd.csv`
- Armazena as fórmulas base
- **Estrutura**:
  - Fórmula (número identificador)
  - Descrição
  - Kg (quantidade base)
  - Tipo
  - Observação

### 4.2 Banco SQLite (`historico.db`)
- Armazena o histórico de cálculos
- **Tabelas**:
  - `historico` (dados gerais)
  - `resultados` (resultados detalhados)

## 5. Como Usar

### 5.1 Realizando Cálculos
1. Digite o nome da programação (opcional)
2. Pesquise a fórmula pelo número ou navegue na lista
3. Selecione as fórmulas desejadas
4. Digite o peso para cada fórmula selecionada
5. Os resultados serão calculados automaticamente

### 5.2 Usando o Histórico
1. Clique no botão "Histórico"
2. Navegue pelos registros salvos
3. Use os filtros para encontrar registros específicos
4. Exporte os resultados quando necessário

### 5.3 Exportando Resultados
1. Selecione o registro desejado
2. Clique no botão "Exportar"
3. Escolha o local para salvar o arquivo Excel
4. O arquivo será gerado com formatação profissional
5. Vídeo:

https://github.com/user-attachments/assets/3e0b6648-c771-40b3-a8bd-10b7b112ee6b


     

## 6. Requisitos Técnicos
- Python 3.x
- Bibliotecas:
  - `tkinter`
  - `sqlite3`
  - `openpyxl`
  - `datetime`
  - `csv`

## 7. Arquivos do Projeto
- `calculadora.py` (programa principal)
- `historico.py` (módulo de histórico)
- `bd.csv` (banco de dados de fórmulas)
- `historico.db` (banco de dados SQLite)
- `documentacao.docx` (este documento)
- `index.html` (página de atualização)

## 8. Manutenção

### 8.1 Atualizando Fórmulas
1. Edite o arquivo `bd.csv`
2. Mantenha a estrutura das colunas
3. Use o separador correto (`,")
4. Reinicie o programa para aplicar as mudanças

### 8.2 Backup
- Faça backup regular do arquivo `historico.db`
- Mantenha uma cópia do `bd.csv`
- Armazene os backups em local seguro

## 9. Suporte e Contato
Para suporte técnico ou dúvidas:
- **Email**: Otaide03@gmail.com
- **Telefone**: (85) 98734-3543

## 10. Atualizações e Versões
**Versão atual**: 2.0
- Histórico completo de cálculos
- Exportação para Excel
- Filtros avançados
- Interface profissional
