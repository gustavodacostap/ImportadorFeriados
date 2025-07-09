namespace ImportadorFeriados.Models
{
    /// <summary>
    /// Representa um feriado municipal, com os dados básicos
    /// extraídos do Excel antes de serem convertidos para o formato de banco.
    /// </summary>
    public class FeriadoMunicipal
    {
        public string Cidade { get; set; } = string.Empty;
        public DateTime? AniversarioCidade { get; set; }
        public string? DescAniversario=> AniversarioCidade.HasValue ? "Aniversário da Cidade" : null;

        public DateTime? Ponte { get; set; }
        public string? DescPonte => Ponte.HasValue ? "Ponte (Aniversário da Cidade)" : null;

        public DateTime? Padroeiro { get; set; }
        public string? DescPadroeiro => Padroeiro.HasValue ? "Padroeiro(a) da cidade" : null;

        public List<(DateTime Data, string Descricao, string Periodicidade)> OutrosFeriados { get; set; } = new();
    }
}
