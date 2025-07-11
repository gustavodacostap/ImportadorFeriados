namespace ImportadorFeriados.Models.Exportação
{
    /// <summary>
    /// Representa um feriado municipal bruto, apenas com os dados
    /// necessários para a exportação em Excel
    /// </summary>
    public class FeriadoMunicipalBruto
    {
        public string Cidade { get; set; } = "";
        public DateTime Data { get; set; }
        public string Descricao { get; set; } = "";
        public string Tipo { get; set; } = "";
    }
}
