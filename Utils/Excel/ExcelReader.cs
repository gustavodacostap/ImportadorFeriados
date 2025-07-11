using ClosedXML.Excel;
using ImportadorFeriados.Models.Importação;

namespace ImportadorFeriados.Utils.Excel
{
    public class ExcelReader
    {
        private readonly string _caminhoArquivo;

        /// <summary>
        /// Construtor que recebe o caminho do arquivo Excel a ser lido.
        /// </summary>
        public ExcelReader(string caminhoArquivo) =>
            _caminhoArquivo = caminhoArquivo;

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
                string msg = $"!! Erro ao abrir o arquivo Excel. Verifique se o arquivo '{_caminhoArquivo}' está fechado antes de rodar o programa.";
                Console.WriteLine(msg);
                Environment.Exit(1); // encerra o programa com código de erro 1
                return null!; // só para o compilador não reclamar
            }

            if (!workbook.Worksheets.Any(p => p.Name == nomePlanilha))
            {
                string msg = $"!! Planilha '{nomePlanilha}' não encontrada!";
                Console.WriteLine(msg);
                Environment.Exit(1);
                return null!;
            }

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

                    // Verifica se a data é 09/07 (Consciência Negra)
                    if (data.Day == 9 && data.Month == 7)
                    {
                        feriado.Abrangencia = FeriadoGuids.ABRANG_ESTADUAL;
                        feriado.UFCode = 19;
                    }
                    else
                    {
                        feriado.Abrangencia = FeriadoGuids.ABRANG_NACIONAL;
                        feriado.UFCode = 0;
                    }
                }
                else
                {
                    continue;
                }

                // Coluna 3: Descrição do feriado
                if (linha.Cell(3).TryGetValue(out string descricao))
                {
                    feriado.Descricao = descricao.Trim();

                    // Verifica se é uma ponte (ignora maiúsculas/minúsculas)
                    if (descricao.Contains("ponte", StringComparison.OrdinalIgnoreCase))
                    {
                        feriado.Tipo = FeriadoGuids.TIPO_PONTE;
                    }                        
                    else
                    {
                        feriado.Tipo = FeriadoGuids.TIPO_FERIADO;
                    }
                }

                // Coluna 5: Periodicidade do feriado
                if (linha.Cell(4).TryGetValue(out string peridiocidade))
                {
                    peridiocidade = peridiocidade.Trim();
                    if (peridiocidade.Contains("anual", StringComparison.OrdinalIgnoreCase))
                    { 
                        feriado.Periodicidade = FeriadoGuids.PERIODIC_ANUAL;
                    }
                else
                    {
                        feriado.Periodicidade = FeriadoGuids.PERIODIC_EVENTUAL;
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
                    continue;
                }
                // Coluna 2: ANIVERSÁRIO DA CIDADE
                if (linha.Cell(2).TryGetValue(out DateTime aniversario))
                    feriado.AniversarioCidade = aniversario;
                    
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

                if (dataFeriado.HasValue)
                {
                    feriado.OutrosFeriados.Add((
                        dataFeriado.Value,
                        descricao
                    ));
                }

                lista.Add(feriado);
            }

            return lista;
        }
    }
}