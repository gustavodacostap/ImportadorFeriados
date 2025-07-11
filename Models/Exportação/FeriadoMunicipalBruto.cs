namespace ImportadorFeriados.Models.Exportação
{
    public class FeriadoMunicipalBruto
    {
        public string Cidade { get; set; } = "";
        public DateTime Data { get; set; }
        public string Descricao { get; set; } = "";
        public string Tipo { get; set; } = "";
    }
}
