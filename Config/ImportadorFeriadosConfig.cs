namespace ImportadorFeriados.Config
{
    // Classe que representa as configurações necessárias para executar o importador de feriados
    public class ImportadorFeriadosConfig
    {
        public string Usuario { get; set; } = string.Empty;
        public string Senha { get; set; } = string.Empty;
        public string Schema { get; set; } = string.Empty;
        public string NomeDoBanco { get; set; } = string.Empty;
        public string Hostname { get; set; } = string.Empty;
        public string Porta { get; set; } = string.Empty;
        public string CaminhoArquivoExcel { get; set; } = string.Empty;
    }
}
