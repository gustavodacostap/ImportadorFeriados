namespace ImportadorFeriados.Models.Exportação
{
    /// <summary>
    /// Representa um feriado nacional ou estadual apenas com os dados
    /// necessários para a exportação em Excel
    /// </summary>
    public class FeriadoNE
    {
        public DateTime Data { get; set; }
        public string Descricao { get; set; } = string.Empty;
        public string Periodicidade { get; set; } = string.Empty;
    }
}
