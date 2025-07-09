# Importador de Feriados

Este projeto foi desenvolvido durante o meu estágio com o objetivo de **automatizar a tarefa de inserção de feriados no sistema interno da organização**.

## 📌 Objetivo

Automatizar o processo de importação de feriados por meio da leitura de uma arquivo Excel local com planilhas para feriados nacionais, estaduais e municipais. O programa atualiza automaticamente as tabelas TB_FERIADO e TB_FERIADO_LOCALIDADE (tabela que relaciona um feriado municipal a uma cidade) do banco de dados da organização com essas informações.

## ⚙️ Funcionamento

- Desenvolvido como uma **classe utilitária** (`ImportadorFeriados.cs`).
- É chamado dentro do sistema principal da organização.
- Realiza a **leitura de um arquivo Excel local**.
- Processa os dados e insere/atualiza os feriados no banco de dados.

## 📁 Estrutura do Projeto

```
📦 ImportadorFeriados
┣ 📂Config # Configurações
┣ 📂Data # Conexão e acesso ao banco de dados
┣ 📂Models # Modelos de dados utilizados
┣ 📂Services # Leitor de planilhas e serviços de importação
┣ 📂Utils # Funções utilitárias (ex: removedor de acentos)
┣ 📜.gitignore # Arquivos e pastas ignoradas pelo Git
┣ 📜ImportadorFeriados.cs # Classe principal que orquestra a importação
````

## 🛠️ Tecnologias Utilizadas

- C# (.NET) – Linguagem e plataforma principal utilizada no desenvolvimento da aplicação.
- ODBC – Para conexão com o banco de dados e execução de operações SQL.
- ClosedXML – Utilizado para leitura e manipulação de arquivos Excel (.xlsx), facilitando a extração dos feriados da planilha.
- Microsoft.Extensions.Configuration – Para leitura e gerenciamento de configurações do sistema de forma centralizada, com suporte a arquivos JSON.
- SQL – Para consultas e manipulação de dados no banco relacional da organização.

## 📌 Observações

- O programa foi desenvolvido como parte de uma demanda real da organização.
- A planilha de feriados segue um formato padronizado (colunas: dia, mês, ano, descrição, localidade, etc.).
- O código está modularizado e comentado para facilitar manutenção e reutilização.
