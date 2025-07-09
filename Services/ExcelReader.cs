using ClosedXML.Excel;
using ImportadorFeriados.Models;
using ImportadorFeriados.Utils;

namespace ImportadorFeriados.Services
{
    public class ExcelReader
    {
        private readonly string _caminhoArquivo;

        /// <summary>
        /// Construtor que recebe o caminho do arquivo Excel a ser lido.
        /// </summary>
        public ExcelReader(string caminhoArquivo) =>
            (_caminhoArquivo) = (caminhoArquivo);

        /// <summary>
        /// Abre a planilha com o nome informado no arquivo Excel.
        /// Lança exceção se a planilha não existir.
        /// </summary>
        public IXLWorksheet AbrirPlanilha(string nomePlanilha)
        {
            XLWorkbook workbook;

            try
            {
                workbook = new XLWorkbook(_caminhoArquivo);
            }
            catch (IOException)
            {
                throw new IOException($"Erro ao abrir o arquivo Excel. Verifique se o arquivo '{_caminhoArquivo}' está fechado antes de rodar o programa.");
            }

            if (!workbook.Worksheets.Any(p => p.Name == nomePlanilha))
                throw new Exception($"Planilha '{nomePlanilha}' não encontrada!");

            return workbook.Worksheet(nomePlanilha);
        }

        /// <summary>
        /// Lê os feriados nacionais e estaduais da planilha e retorna uma lista de objetos FeriadoNacionalEstadual.
        /// </summary>
        public List<FeriadoNacionalEstadual> LerFeriadosNacionaisEstaduais()
        {
            var lista = new List<FeriadoNacionalEstadual>();
            var planilha = AbrirPlanilha("FERIADOS NACIONAIS E ESTADUAIS");

            // Itera sobre as linhas usadas (com conteúdo), ignorando o cabeçalho
            foreach (var linha in planilha.RowsUsed().Skip(1))
            {
                var feriado = new FeriadoNacionalEstadual();

                // Coluna 1: Data do feriado
                if (linha.Cell(1).TryGetValue(out DateTime data))
                { 
                    feriado.Data = data;
                    Console.Write(data.ToString("dd/MM/yyyy") + " - ");
                }
                else
                {
                    Console.WriteLine($"[Aviso] Linha {linha.RowNumber()}: DATA inválida ou vazia. Linha ignorada.");
                    continue;
                }

                // Coluna 3: Descrição do feriado
                if (linha.Cell(3).TryGetValue(out string descricao))
                {
                    feriado.Descricao = descricao;
                    Console.Write($"{descricao} - ");

                    // Verifica se é uma ponte (ignora maiúsculas/minúsculas)
                    if (descricao.Contains("ponte", StringComparison.OrdinalIgnoreCase))
                    {
                        feriado.Tipo = FeriadoGuids.TIPO_PONTE;
                        Console.Write("PONTE - ");
                    }                        
                    else
                    {
                        feriado.Tipo = FeriadoGuids.TIPO_FERIADO;
                        Console.Write("FERIADO - ");
                    }
                }

                // Coluna 4: Abrangência do feriado
                if (linha.Cell(4).TryGetValue(out string abrangencia))
                {
                    abrangencia = abrangencia.Trim();
                    if (abrangencia.Contains("NACIONAL", StringComparison.OrdinalIgnoreCase))
                    {
                        feriado.Abrangencia = FeriadoGuids.ABRANG_NACIONAL;
                        feriado.UFCode = 0;
                        Console.Write("NACIONAL - ");
                    }
                    else
                    {
                        feriado.Abrangencia = FeriadoGuids.ABRANG_ESTADUAL;
                        feriado.UFCode = 19;
                        Console.Write("ESTADUAL - ");
                    }                       
                }

                // Coluna 5: Periodicidade do feriado
                if (linha.Cell(5).TryGetValue(out string peridiocidade))
                {
                    peridiocidade = peridiocidade.Trim();
                    if (peridiocidade.Contains("anual", StringComparison.OrdinalIgnoreCase))
                    { 
                        feriado.Periodicidade = FeriadoGuids.PERIODIC_ANUAL;
                        Console.Write("ANUAL\n");
                    }
                else
                    {
                        feriado.Periodicidade = FeriadoGuids.PERIODIC_EVENTUAL;
                        Console.Write("EVENTUAL\n");
                    }                   
                }
                lista.Add(feriado);
            }

            return lista;
        }

        /// <summary>
        /// Lê os feriados municipais da planilha e retorna uma lista de objetos FeriadoMunicipal.
        /// </summary>
        public List<FeriadoMunicipal> LerFeriadosMunicipais()
        {
            var lista = new List<FeriadoMunicipal>();
            var planilha = AbrirPlanilha("FERIADOS MUNICIPAIS");

            // Itera as linhas usadas, pulando as duas primeiras (cabeçalhos)
            foreach (var linha in planilha.RowsUsed().Skip(2)) // pula cabeçalhos
            {
                var feriado = new FeriadoMunicipal();

                // Coluna 1: CIDADE
                if (linha.Cell(1).TryGetValue(out string cidade))
                    feriado.Cidade = cidade;
                else
                {
                    Console.WriteLine($"[Aviso] Linha {linha.RowNumber()}: CIDADE inválida ou vazia. Linha ignorada.");
                    continue;
                }
                // Coluna 2: ANIVERSÁRIO DA CIDADE
                if (linha.Cell(2).TryGetValue(out DateTime aniversario))
                    feriado.AniversarioCidade = aniversario;
                    
                else
                {
                    Console.WriteLine($"[Aviso] Linha {linha.RowNumber()}: ANIVERSÁRIO DA CIDADE ausente ou inválido.");
                }

                // Coluna 3: PONTE
                if (linha.Cell(3).TryGetValue(out DateTime ponte))
                    feriado.Ponte = ponte;


                // Coluna 4: PADROEIRO
                var celulaPadroeiro = linha.Cell(4).GetString();

                if (!string.IsNullOrEmpty(celulaPadroeiro))
                {
                    var texto = celulaPadroeiro.Trim();

                    // Divide o texto em partes, considerando que a primeira parte é a data
                    var partes = texto.Split(' ', StringSplitOptions.RemoveEmptyEntries);

                    // Tenta converter a primeira parte em data
                    if (DateTime.TryParse(partes[0], out DateTime dataPadroeiro))
                    {
                        feriado.Padroeiro = dataPadroeiro;
                    }
                }

                // Coluna 5: FERIADO MUNICIPAL (data + nome em mesma célula, com quebra de linha)
                var celulaFeriadoMunicipal = linha.Cell(5).GetString();
                DateTime? dataFeriado = null;
                string descricao = string.Empty;

                if (!string.IsNullOrWhiteSpace(celulaFeriadoMunicipal))
                {
                    var texto = celulaFeriadoMunicipal.Trim();

                    // Divide o texto em partes, considerando que a primeira parte é a data
                    var partes = texto.Split(' ', StringSplitOptions.RemoveEmptyEntries);

                    // Tenta converter a primeira parte em data
                    if (DateTime.TryParse(partes[0], out DateTime data))
                    {
                        dataFeriado = data;
                        descricao = string.Join(' ', partes.Skip(1)).Trim();
                    }
                }

                // Coluna 6: Periodicidade do feriado
                var periodicidade = FeriadoGuids.PERIODIC_ANUAL;
                if (linha.Cell(6).TryGetValue(out string peridiocidade))
                {
                    if (peridiocidade.Contains("eventual", StringComparison.OrdinalIgnoreCase))
                        periodicidade = FeriadoGuids.PERIODIC_EVENTUAL;
                }

                if (dataFeriado.HasValue)
                {
                    feriado.OutrosFeriados.Add((
                        dataFeriado.Value,
                        descricao,
                        periodicidade
                    ));
                }

                lista.Add(feriado);
            }

            return lista;
        }
    }
}