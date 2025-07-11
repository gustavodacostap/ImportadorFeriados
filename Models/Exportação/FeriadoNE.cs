namespace ImportadorFeriados.Models.Exportação
{
    public class FeriadoNE
    {
        public DateTime Data { get; set; }
        public string Descricao { get; set; } = string.Empty;
        public string Periodicidade { get; set; } = string.Empty; // ANUAL ou EVENTUAL
    }
}
