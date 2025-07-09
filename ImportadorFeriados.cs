using ImportadorFeriados.Data;
using ImportadorFeriados.Services;
using ImportadorFeriados.Config;
using ImportadorFeriados.Utils;
using Microsoft.Extensions.Configuration;

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

        // Apenas para debug
        var feriadosInseridos = new List<string>();
        var feriadosJaExistentes = new List<string>();
        var feriadosVinculados = new List<string>();
        var feriadosIgnorados = new List<string>();
        var localidadesNaoEncontradas = new List<string>();

        // Loop sobre todos os feriados lidos do Excel
        foreach (var feriado in todos)
        {
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
                    feriadosIgnorados.Add($"{feriado.DS_FERIADO} - {feriado.DIA_FERIADO}/{feriado.MES_FERIADO}/{feriado.ANO_FERIADO}");
                }

                continue;
            }

            // Novo feriado: adiciona à lista de inseridos
            feriadosInseridos.Add($"{feriado.DS_FERIADO} - {feriado.DIA_FERIADO}/{feriado.MES_FERIADO}/{feriado.ANO_FERIADO}");

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

        // Debug
        Console.WriteLine("\n====== FERIADOS INSERIDOS (NOVOS) ======");
        foreach (var f in feriadosInseridos)
            Console.WriteLine(f);

        Console.WriteLine("\n====== FERIADOS JÁ EXISTIAM, MAS FORAM VINCULADOS À CIDADE ======");
        foreach (var f in feriadosVinculados)
            Console.WriteLine(f);

        Console.WriteLine("\n====== FERIADOS IGNORADOS (QUE JÁ EXISTIAM, COM VÍNCULO OU SEM) ======");
        foreach (var f in feriadosIgnorados)
            Console.WriteLine(f);

        if (localidadesNaoEncontradas.Count != 0)
        {
            Console.WriteLine("\n====== FERIADOS MUNICIPAIS COM LOCALIDADES NÃO ENCONTRADAS ======");
        }
        foreach (var cidade in localidadesNaoEncontradas.Distinct())
            Console.WriteLine($"! Localidade não encontrada: {cidade}");
    }
}
