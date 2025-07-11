using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using ImportadorFeriados.Models.Exportação;
using ImportadorFeriados.Models.Importação;

namespace ImportadorFeriados.Utils.Excel
{
    public class GeradorExcel(List<FeriadoNE> feriadosNE,
    List<FeriadoMunicipalBruto> feriadosMunicipaisBanco,
    List<FeriadoMunicipal> feriadosMunicipaisExcel,
    string caminhoSalvarArquivo)
    {
        public void GerarExcelDeFeriados()
        {
            var workbook = new XLWorkbook();

            // =========================
            // 1) FERIADOS NACIONAIS/ESTADUAIS
            // =========================
            var wsNE = workbook.Worksheets.Add("FERIADOS NACIONAIS E ESTADUAIS");

            // Cabeçalho: DATA (A e B mescladas), FERIADOS (C), PERIODICIDADE (D)
            wsNE.Cell("A1").Value = "DATA";
            wsNE.Range("A1:B1").Merge().Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            wsNE.Cell("C1").Value = "FERIADOS";
            wsNE.Cell("D1").Value = "PERIODICIDADE";
            wsNE.Row(1).Height = 25;

            int linha = 2;
            foreach (var feriado in feriadosNE.OrderBy(f => f.Data))
            {
                wsNE.Row(linha).Height = 25;
                wsNE.Cell(linha, 1).Value = feriado.Data.ToString("dd/MM/yyyy");
                wsNE.Cell(linha, 2).Value = feriado.Data.ToString("dddd", new System.Globalization.CultureInfo("pt-BR"));
                wsNE.Cell(linha, 3).Value = feriado.Descricao;
                wsNE.Cell(linha, 4).Value = feriado.Periodicidade;
                linha++;
            }

            // =========================
            // 2) FERIADOS MUNICIPAIS
            // =========================
            var wsMun = workbook.Worksheets.Add("FERIADOS MUNICIPAIS");

            // Cabeçalho
            wsMun.Cell("A1").Value = "CIDADE";
            wsMun.Cell("B1").Value = "ANIVERSÁRIO DA CIDADE";
            wsMun.Cell("C1").Value = "PONTE";
            wsMun.Cell("D1").Value = "PADROEIRO";
            wsMun.Cell("E1").Value = "FERIADO MUNICIPAL";

            int linhaMun = 2;

            var excelNormalizado = feriadosMunicipaisExcel
                .GroupBy(f => TextoUtils.RemoverAcentos(f.Cidade.Trim().ToLower()))
                .ToDictionary(g => g.Key, g => g.First());

            var bancoNormalizado = feriadosMunicipaisBanco
                .GroupBy(f => TextoUtils.RemoverAcentos(f.Cidade.Trim().ToLower()))
                .ToDictionary(g => g.Key, g => g.ToList());

            // // Para manter a ordem original do Excel
            foreach (var cidadeExcelOriginal in feriadosMunicipaisExcel.Select(f => f.Cidade.Trim()).Distinct())
            {
                var nomeNormalizado = TextoUtils.RemoverAcentos(cidadeExcelOriginal.ToLower());

                if (!bancoNormalizado.ContainsKey(nomeNormalizado))
                    continue;

                var feriadoExcel = excelNormalizado[nomeNormalizado];
                var feriadosDaCidadeBanco = bancoNormalizado[nomeNormalizado];

                // Linha vazia intercalada
                wsMun.Range($"A{linhaMun}:E{linhaMun}").Style.Fill.BackgroundColor = XLColor.DarkGray;
                linhaMun++;

                // Preenche dados no Excel
                wsMun.Cell(linhaMun, 1).Value = cidadeExcelOriginal;

                // Verifica por data
                DateTime? aniversario = feriadoExcel?.AniversarioCidade;
                DateTime? ponte = feriadoExcel?.Ponte;
                DateTime? padroeiro = feriadoExcel?.Padroeiro;
                DateTime? outro = feriadoExcel?.OutrosFeriados.FirstOrDefault().Data;

                wsMun.Cell(linhaMun, 2).Value = feriadosDaCidadeBanco.FirstOrDefault(f => DatasIguais(f.Data, aniversario))?.Data.ToString("dd/MM/yyyy");
                wsMun.Cell(linhaMun, 3).Value = feriadosDaCidadeBanco.FirstOrDefault(f => DatasIguais(f.Data, ponte))?.Data.ToString("dd/MM/yyyy");
                wsMun.Cell(linhaMun, 4).Value = feriadosDaCidadeBanco.FirstOrDefault(f => DatasIguais(f.Data, padroeiro))?.Data.ToString("dd/MM/yyyy");

                // Pega os feriados que não sejam igual a nenhum dos anteriores
                var datasUsadas = new[] { aniversario, ponte, padroeiro }.Where(d => d.HasValue)
                    .Select(d => d.Value)
                    .ToHashSet();
                var feriadosExtras = feriadosDaCidadeBanco.Where(f => !datasUsadas.Contains(f.Data))
                    .OrderBy(f => f.Data)
                    .ToList();

                // Monta string: data\nDescrição\n...
                var textoFeriadosExtras = string.Join("\n", feriadosExtras
                    .Select(f => $"{f.Data:dd/MM/yyyy}\n{f.Descricao}"));

                var celulaFeriadosExtra = wsMun.Cell(linhaMun, 5);
                celulaFeriadosExtra.Value = textoFeriadosExtras;
                celulaFeriadosExtra.Style.Alignment.WrapText = true;

                linhaMun++;
            }

            // Ajusta largura das colunas
            wsNE.Columns().AdjustToContents();
            wsMun.Columns().AdjustToContents();

            EstilizarPlanilha(wsNE, "A1:D1");
            EstilizarPlanilha(wsMun, "A1:E1");

            // Salva o arquivo
            workbook.SaveAs(caminhoSalvarArquivo);

            Console.WriteLine($"Arquivo Excel salvo em: {caminhoSalvarArquivo}");
        }

        public static void EstilizarPlanilha(IXLWorksheet ws, string cabecalhoRange)
        {
            var header = ws.Range(cabecalhoRange);
            header.Style.Fill.BackgroundColor = XLColor.LightGreen;
            header.Style.Font.Bold = true;

            var rangeUsado = ws.RangeUsed();
            rangeUsado.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            rangeUsado.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            rangeUsado.Style.Border.OutsideBorderColor = XLColor.Black;
            rangeUsado.Style.Border.InsideBorderColor = XLColor.Black;
            rangeUsado.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            rangeUsado.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        }
        public static bool DatasIguais(DateTime data1, DateTime? data2)
        {
            if (data2 == null) return false;
            return data1.Date == data2.Value.Date;
        }
    }  
}
