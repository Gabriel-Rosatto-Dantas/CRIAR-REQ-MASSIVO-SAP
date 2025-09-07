# Automação SAP - Requisição de Compra

Este aplicativo automatiza o processo de criação de requisições de compra no SAP, utilizando dados de uma planilha Google Sheets.

## Versões Disponíveis

- **APP.py**: Versão original com interface de console
- **APP_GUI.py**: Nova versão com interface gráfica amigável

## Requisitos

- Python 3.6 ou superior
- Bibliotecas Python: pandas, win32com, gspread, tkinter, PIL
- Acesso ao SAP GUI com permissões para criar requisições
- Arquivo de credenciais do Google Sheets (credentials.json)

## Como Usar a Interface Gráfica

1. Execute o arquivo `APP_GUI.py`
2. Na aba "Configurações":
   - Configure o caminho do SAP Logon
   - Defina o sistema SAP
   - Insira suas credenciais SAP (opcional)
   - Configure o acesso ao Google Sheets
   - Clique em "Salvar Configurações"
3. Na aba "Automação":
   - Clique em "Iniciar Automação" para começar o processo
   - O progresso será exibido na área de log
   - Use o botão "Parar Automação" para interromper o processo

## Estrutura da Planilha

A planilha do Google Sheets deve conter as seguintes colunas:

- **PN**: Número da peça
- **QTD**: Quantidade
- **ORIGEM**: Centro de origem
- **DATA REMESSA**: Data de remessa (formato DD/MM/AAAA)
- **DESTINO**: Centro de destino
- **TEXTO**: Descrição do item
- **DEPOSITO**: Depósito
- **Validador**: Índice para validação (0 para iniciar um novo lote)
- **Status**: Status do processamento (deixar em branco para itens novos)
- **REQUISIÇÃO**: Número da requisição criada (preenchido automaticamente)

## Funcionalidades

- Interface gráfica amigável
- Configuração simplificada
- Visualização em tempo real do progresso
- Processamento automático de lotes
- Validação de itens antes da criação da requisição
- Atualização automática da planilha com resultados