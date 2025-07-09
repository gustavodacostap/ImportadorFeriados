namespace ImportadorFeriados.Models
{
    /// <summary>
    /// Representa um feriado com dados correspondentes ao formato do banco.
    /// </summary>
    public class FeriadoParaBanco
    {
        public int CD_FERIADO { get; set; } 
        public string DS_FERIADO { get; set; } = string.Empty;  // Ex: "Aniversário da Cidade"
        public short DIA_FERIADO { get; set; }
        public short MES_FERIADO { get; set; }
        public short? ANO_FERIADO { get; set; }                 // Pode ser null para feriado anual

        public string CD_TP_FERIADO { get; set; } = string.Empty;  // GUID do tipo: Feriado ou Ponte
        public string CD_ABRANGENCIA { get; set; } = string.Empty; // GUID: Nacional, Estadual, Municipal
        public string CD_PERIODICIDADE { get; set; } = string.Empty; // GUID: Anual ou Eventual

        public short CD_UF { get; set; }                        // 0 (nacional), ou 19 (São Paulo)
        public char IND_ATIVO { get; set; } = '1';              // Sempre '1'
        public short DELETED { get; set; } = 0;                 // Sempre 0

        public string USU_INCL { get; set; } = string.Empty;    // Usuário que realizou a inclusão
        public short CD_UNI_INCL { get; set; } = 175;           // Sempre 175
        public DateTime DT_HR_INCL { get; set; } = DateTime.Now;

        public string? USU_ALT { get; set; } = null;
        public short? CD_UNI_ALT { get; set; } = null;
        public DateTime? DT_HR_ALT { get; set; } = null;

        public string CD_GUID_REFERENCIA_PAI { get; set; } = Guid.NewGuid().ToString();
        public string? Cidade { get; set; }
    }
}
