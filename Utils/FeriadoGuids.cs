namespace ImportadorFeriados.Utils
{
    /// <summary>
    /// Classe utilitária contendo constantes dos GUIDs usados para categorizar tipos de feriado,
    /// abrangências e periodicidades.
    /// </summary>
    public static class FeriadoGuids
    {
        // Tipos de feriado
        public const string TIPO_FERIADO = "4350D9CB-8094-4B1E-AACC-9FAA1094A2BD";
        public const string TIPO_PONTE = "517C442E-C1AA-49AC-9982-2F321EA6140F";

        // Abrangência do feriado
        public const string ABRANG_NACIONAL = "67B5E01A-BE2B-40C0-92A2-28D0CF3A21C3";
        public const string ABRANG_ESTADUAL = "ACDA7436-9ED0-41F1-8965-EE68BE76F298";
        public const string ABRANG_MUNICIPAL = "B3793B67-2FF3-4E39-95B3-634C2FE29E9C";

        // Periodicidade do feriado
        public const string PERIODIC_ANUAL = "C46BC619-B713-42CB-BB1C-104958B4B787";
        public const string PERIODIC_EVENTUAL = "BA391D41-1DA8-44AE-A9CB-1DA81AAC7F0E";
    }
}
