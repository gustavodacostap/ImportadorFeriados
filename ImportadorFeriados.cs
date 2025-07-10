using ImportadorFeriados.Config;
using ImportadorFeriados.Data;
using ImportadorFeriados.Services;
using ImportadorFeriados.Utils;
using Microsoft.Extensions.Configuration;
using System.Text;

namespace ImportadorFeriados;

public class ImportadorFeriadosClasse
{
    private readonly ImportadorFeriadosConfig _config;

    // Extrai as configurações de uma seção "ImportadorFeriados" no appsettings.json e cria um obj configuração
    public ImportadorFeriadosClasse(IConfiguration configuration) =>
        _config = configuration
            .GetSection("ImportadorFeriados")
            .Get<ImportadorFeriadosConfig>() ?? throw new Exception("Configuração 'ImportadorFeriados' não encontrada.");
    public void Importar()
    {
        // Monta a connection string para conectar ao banco DB2
        string connStr = $"Driver={{IBM DB2 ODBC DRIVER}};Database={_config.NomeDoBanco};Hostname={_config.Hostname};Port={_config.Porta};Uid={_config.Usuario};Pwd={_config.Senha};";

        var dbService = new DbService(connStr, _config.Schema);

        // Lê os dados do Excel (feriados nacionais/estaduais e municipais)
        var excelReader = new ExcelReader(_config.CaminhoArquivoExcel);
        var feriadosNE = excelReader.LerFeriadosNacionaisEstaduais();
        var feriadosMunicipais = excelReader.LerFeriadosMunicipais();

        // Concatena todos os feriados
        var todos = feriadosNE
            .FromNacionaisEstaduais(_config.Usuario)
            .Concat(feriadosMunicipais.FromMunicipais(_config.Usuario));

        // Listas auxiliares para armazenar os resultados e depois montar o arquivo de log
        var feriadosInseridos = new List<string>();
        var feriadosJaExistentes = new List<string>();
        var feriadosVinculados = new List<string>();
        var feriadosIgnorados = new List<string>();
        var localidadesNaoEncontradas = new List<string>();

        var logBuilder = new StringBuilder(); // StringBuilder para gerar log em arquivo

        int total = todos.Count();
        int atual = 0;

        // Loop sobre todos os feriados lidos do Excel
        foreach (var feriado in todos)
        {
            atual++;
            LogUtils.MostrarBarraProgresso(atual, total);

            // Tenta inserir ou localizar o feriado no banco
            var resultado = dbService.InserirFeriado(feriado);

            if (resultado.FeriadoJaExistia)
            {
                if (resultado.CdFeriado.HasValue && resultado.LocNu.HasValue)
                {
                    // Já existia, mas **não tinha vínculo com a cidade** -> agora vinculou
                    feriadosVinculados.Add($"{feriado.DS_FERIADO} - {feriado.DIA_FERIADO}/{feriado.MES_FERIADO}/{feriado.ANO_FERIADO} - {feriado.Cidade}");
                    dbService.InserirFeriadoLocalidade(resultado.CdFeriado.Value, resultado.LocNu.Value, feriado.USU_INCL);
                }
                else
                {
                    // Já existia **e já tinha vínculo** (ou não é municipal)
                    string info = $"{feriado.DS_FERIADO} - {feriado.DIA_FERIADO}/{feriado.MES_FERIADO}/{feriado.ANO_FERIADO}";
                    if (!string.IsNullOrWhiteSpace(feriado.Cidade))
                        info += $" - {feriado.Cidade}";
                    feriadosIgnorados.Add(info);
                }

                continue;
            }

            // Novo feriado: adiciona à lista de inseridos
            string novoLog = $"{feriado.DS_FERIADO} - {feriado.DIA_FERIADO}/{feriado.MES_FERIADO}/{feriado.ANO_FERIADO}";
            if (!string.IsNullOrWhiteSpace(feriado.Cidade))
                novoLog += $" - {feriado.Cidade}";
            feriadosInseridos.Add(novoLog);

            // Se a localidade foi identificada com sucesso, insere o vínculo
            if (resultado.LocNu.HasValue)
            {
                dbService.InserirFeriadoLocalidade(resultado.CdFeriado!.Value, resultado.LocNu.Value, feriado.USU_INCL);
            }
            else if (!string.IsNullOrWhiteSpace(feriado.Cidade))
            {
                // Se cidade foi informada mas não foi encontrada no banco, adiciona na lista de falhas
                localidadesNaoEncontradas.Add(feriado.Cidade);
            }
        }

        Console.WriteLine("\nImportação finalizada com sucesso.");

        // Montando o arquivo log
        logBuilder.AppendLine("====== FERIADOS INSERIDOS (NOVOS) ======");
        foreach (var f in feriadosInseridos)
            logBuilder.AppendLine(f);

        logBuilder.AppendLine("\n====== FERIADOS JÁ EXISTIAM, MAS FORAM VINCULADOS À CIDADE ======");
        foreach (var f in feriadosVinculados)
            logBuilder.AppendLine(f);

        logBuilder.AppendLine("\n====== FERIADOS IGNORADOS (QUE JÁ EXISTIAM, COM VÍNCULO OU SEM) ======");
        foreach (var f in feriadosIgnorados)
            logBuilder.AppendLine(f);

        if (localidadesNaoEncontradas.Count != 0)
        {
            logBuilder.AppendLine("\n====== FERIADOS MUNICIPAIS COM LOCALIDADES NÃO ENCONTRADAS ======");
        }
        foreach (var cidade in localidadesNaoEncontradas.Distinct())
            logBuilder.AppendLine($"! Localidade não encontrada: {cidade}");

        // Cria pasta Logs ao lado do executável, se não existir
        var pastaLogs = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Logs");
        Directory.CreateDirectory(pastaLogs);

        // Caminho do arquivo com timestamp
        var caminhoLog = Path.Combine(pastaLogs, $"log_feriados_{DateTime.Now:yyyyMMdd_HHmmss}.txt");

        // Escreve o conteúdo do log
        File.WriteAllText(caminhoLog, logBuilder.ToString());

        Console.WriteLine($"\nLog salvo em: {caminhoLog}");
    }
}
