namespace ImportadorFeriados.Models.Importação
{
    /// <summary>
    /// Representa um feriado nacional ou estadual, com os dados básicos
    /// extraídos do Excel antes de serem convertidos para o formato de banco.
    /// </summary>
    public class FeriadoNacionalEstadual
    {
        public DateTime Data { get; set; }
        public string Descricao { get; set; } = string.Empty;
        public string Tipo { get; set; } = string.Empty;
        public string Abrangencia { get; set; } = string.Empty;
        public string Periodicidade { get; set; } = string.Empty;
        public short UFCode { get; set; }
    }
}
