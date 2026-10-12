Automação de Documentos — Google Apps Script
Objetivo

Automatizar a localização de notas fiscais e boletos em arquivos PDF armazenados no Google Drive, vinculando os documentos aos respectivos lançamentos em uma planilha do Google Sheets.

Funcionalidades
Processamento dos lançamentos por data.
Busca de PDFs em pastas organizadas por dia.
Identificação de notas fiscais, boletos e faturas.
Correspondência por número da fatura, nome, valor e CNPJ.
Inserção de links dos documentos nas colunas J e K.
Registro de situações que precisam de conferência na coluna Q.
Exibição de um resumo ao final do processamento.
Tecnologias
Google Apps Script
Google Sheets
Google Drive
Git e GitHub para versionamento do código
Utilização
Abra a planilha vinculada ao projeto.
Acesse o menu 📎 Documentos.
Selecione Processar um dia.
Informe a data no formato DD.MM.AAAA.
Confira os links inseridos e os resultados do processamento.
Segurança

O código pode depender de identificadores de pastas e permissões do Google Drive. Não publique informações confidenciais, credenciais ou identificadores internos em um repositório público.

Observação

O GitHub armazena o código-fonte. A execução da automação continua sendo realizada pelo Google Apps Script.
