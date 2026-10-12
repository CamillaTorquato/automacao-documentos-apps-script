**Automação de Documentos — Google Apps Script**

**Objetivo**

Automatizar a localização de notas fiscais e boletos em arquivos PDF armazenados no Google Drive, vinculando os documentos aos respectivos lançamentos em uma planilha do Google Sheets.

**Funcionalidades**

- Processamento dos lançamentos por data.
- Busca de PDFs em pastas organizadas por dia.
- Identificação de notas fiscais, boletos e faturas.
- Correspondência por número da fatura, nome, valor e CNPJ.
- Inserção de links dos documentos nas colunas correspondentes.
- Registro de situações que precisam de conferência na coluna designada.
- Exibição de um resumo ao final do processamento.
  
**Tecnologias**
- Google Apps Script
- Google Sheets
- Google Drive

**Utilização**
1. Abra a planilha vinculada ao projeto.
2. Acesse o menu 📎 Documentos.
3. Selecione Processar um dia.
4. Informe a data no formato DD.MM.AAAA.
5. Confira os links inseridos e os resultados do processamento.
   

**Observações**

O código pode depender de identificadores de pastas e permissões do Google Drive.  
A execução da automação é realizada pelo Google Apps Script.
