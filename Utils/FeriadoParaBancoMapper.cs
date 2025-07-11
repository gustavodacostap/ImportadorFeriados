using ImportadorFeriados.Models;
using ImportadorFeriados.Models.Importação;
using ImportadorFeriados.Utils;

namespace ImportadorFeriados.Utils
{
    public static class FeriadoParaBancoMapper
    {
        /// <summary>Mapeia feriados nacionais/estaduais para o model FeriadoParaBanco.</summary>
        public static IEnumerable<FeriadoParaBanco> FromNacionaisEstaduais(
            this IEnumerable<FeriadoNacionalEstadual> feriadosNacionaisEstaduais,
            string usuarioInclusao
        )
        {
            foreach (var f in feriadosNacionaisEstaduais)
            {
                yield return new FeriadoParaBanco
                {
                    DS_FERIADO = f.Descricao,
                    DIA_FERIADO = (short)f.Data.Day,
                    MES_FERIADO = (short)f.Data.Month,
                    ANO_FERIADO = f.Periodicidade == FeriadoGuids.PERIODIC_ANUAL ? null : (short)f.Data.Year,
                    CD_TP_FERIADO = f.Tipo,
                    CD_ABRANGENCIA = f.Abrangencia,
                    CD_PERIODICIDADE = f.Periodicidade,
                    CD_UF = f.UFCode,
                    USU_INCL = usuarioInclusao
                };
            }
        }

        /// <summary>Mapeia feriados municipais para o model FeriadoParaBanco.</summary>
        public static IEnumerable<FeriadoParaBanco> FromMunicipais(
            this IEnumerable<FeriadoMunicipal> feriadoMunicipais,
            string usuarioInclusao
        )
        {
            foreach (var f in feriadoMunicipais)
            {
                // 1) Aniversário de cidade
                if (f.AniversarioCidade.HasValue)
                {
                    yield return new FeriadoParaBanco
                    {
                        DS_FERIADO = f.DescAniversario!,
                        DIA_FERIADO = (short)f.AniversarioCidade.Value.Day,
                        MES_FERIADO = (short)f.AniversarioCidade.Value.Month,
                        ANO_FERIADO = null,
                        CD_TP_FERIADO = FeriadoGuids.TIPO_FERIADO,
                        CD_ABRANGENCIA = FeriadoGuids.ABRANG_MUNICIPAL,
                        CD_PERIODICIDADE = FeriadoGuids.PERIODIC_ANUAL,
                        CD_UF = 19,
                        USU_INCL = usuarioInclusao,
                        Cidade = f.Cidade
                    };
                }

                // 2) Ponte
                if (f.Ponte.HasValue)
                {
                    yield return new FeriadoParaBanco
                    {
                        DS_FERIADO = f.DescPonte!,
                        DIA_FERIADO = (short)f.Ponte.Value.Day,
                        MES_FERIADO = (short)f.Ponte.Value.Month,
                        ANO_FERIADO = (short?)f.Ponte.Value.Year,
                        CD_TP_FERIADO = FeriadoGuids.TIPO_PONTE,
                        CD_ABRANGENCIA = FeriadoGuids.ABRANG_MUNICIPAL,
                        CD_PERIODICIDADE = FeriadoGuids.PERIODIC_EVENTUAL,
                        CD_UF = 19,
                        USU_INCL = usuarioInclusao,
                        Cidade = f.Cidade
                    };
                }

                // 3) Padroeiro
                if (f.Padroeiro.HasValue)
                {
                    yield return new FeriadoParaBanco
                    {
                        DS_FERIADO = f.DescPadroeiro!,
                        DIA_FERIADO = (short)f.Padroeiro.Value.Day,
                        MES_FERIADO = (short)f.Padroeiro.Value.Month,
                        ANO_FERIADO = null,
                        CD_TP_FERIADO = FeriadoGuids.TIPO_FERIADO,
                        CD_ABRANGENCIA = FeriadoGuids.ABRANG_MUNICIPAL,
                        CD_PERIODICIDADE = FeriadoGuids.PERIODIC_ANUAL,
                        CD_UF = 19,
                        USU_INCL = usuarioInclusao,
                        Cidade = f.Cidade
                    };
                }

                // 4) Outros feriados
                foreach (var (data, desc) in f.OutrosFeriados)
                {
                    yield return new FeriadoParaBanco
                    {
                        DS_FERIADO = desc,
                        DIA_FERIADO = (short)data.Day,
                        MES_FERIADO = (short)data.Month,
                        ANO_FERIADO =  null,
                        CD_TP_FERIADO = FeriadoGuids.TIPO_FERIADO,
                        CD_ABRANGENCIA = FeriadoGuids.ABRANG_MUNICIPAL,
                        CD_PERIODICIDADE = FeriadoGuids.PERIODIC_ANUAL,
                        CD_UF = 19,
                        USU_INCL = usuarioInclusao,
                        Cidade = f.Cidade
                    };
                }
            }
        }
    }
}
