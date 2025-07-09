# Importador de Feriados

Este projeto é um programa em C# para importar feriados de planilhas Excel e inserir/atualizar em um banco de dados DB2.

## Estrutura

- `Program.cs` - Ponto de entrada da aplicação, orquestra os serviços.
- `DbService.cs` - Classe para acesso ao banco DB2 (comentada por enquanto).
- `DbServiceSQLite.cs` - Implementação para testes locais com banco SQLite.
- `ExcelReader.cs` - (Futuro) Leitura e parsing das planilhas Excel.
