# Google-Sheets-and-Docs-Merge
Este repositório contém um código .gs feito no Apps Script para criar uma mala direta a partir de uma planilha do Google Sheets e um modelo de documento no Google Docs. O script substitui marcadores de posição no documento com os dados da planilha, gerando múltiplos documentos personalizados.

# Instruções:

1) Adicione os dados à planilha do Google Sheets (pode ser via Google Forms).
2) Defina o modelo de documento no Google Docs com marcadores de posição {{NOME_VARIAVEL}}.
3) Configure as variáveis do script para apontar para a planilha e o modelo de documento.
4) Execute a função criarMalaDireta() para gerar os documentos personalizados.

# Detalhes do Código:

Variáveis da Planilha de Banco de Dados: As variáveis app, spreadsheet e sheet são configuradas para acessar a planilha especificada.
Variável do arquivo de Mala Direta: A variável docsMalaDireta é configurada para acessar o documento que será transformado na mala direta.
Variável do arquivo de Template: A variável docsTemplate é configurada para acessar o modelo de documento utilizado para os contratos.
Função criarMalaDireta(): Esta função realiza a mala direta, substituindo os marcadores de posição no modelo de documento com os dados da planilha.

Este repositório fornece uma solução prática para automatizar o processo de geração de documentos personalizados a partir de dados armazenados em uma planilha do Google Sheets.
