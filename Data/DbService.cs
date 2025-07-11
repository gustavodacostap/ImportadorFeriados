using ImportadorFeriados.Models;
using ImportadorFeriados.Models.Exportação;
using ImportadorFeriados.Utils;
using System.Data.Odbc;

namespace ImportadorFeriados.Data
{
    // Classe responsável por toda a comunicação com o banco de dados
    public class DbService
    {
        private readonly string _connectionString;
        private readonly string _schema;

        // Record que representa o resultado de uma tentativa de inserção de feriado
        public record ResultadoInsercaoFeriado(int? CdFeriado, int? LocNu, bool FeriadoJaExistia);

        public DbService(string connectionString, string schema) =>
            (_connectionString, _schema) = (connectionString, schema);

        /// <summary>
        /// Define o schema padrão para a conexão antes de executar qualquer comando no DB2.
        /// </summary>
        private void SetSchema(OdbcConnection conn)
        {
            using var cmd = conn.CreateCommand();
            cmd.CommandText = $"SET SCHEMA {_schema}";
            cmd.ExecuteNonQuery();
        }

        /// <summary>
        /// Insere um feriado no banco de dados, tratando duplicidade e vínculo com a cidade (quando municipal).
        /// </summary>
        public ResultadoInsercaoFeriado InserirFeriado(FeriadoParaBanco feriado)
        {
            using var conn = new OdbcConnection(_connectionString);
            conn.Open();
            SetSchema(conn);

            if (FeriadoJaExiste(conn, feriado, out int cdFeriadoExistente))
            {
                // Trata feriado que já existe no banco (pode ou não estar vinculado à cidade)
                return TratarFeriadoExistente(conn, feriado, cdFeriadoExistente);
            }

            // Insere novo feriado
            int novoCdFeriado = InserirNovoFeriado(conn, feriado);

            // Tenta obter LOC_NU para feriados municipais
            int? locNu = ObterLocNuSeMunicipal(feriado);

            return new ResultadoInsercaoFeriado(novoCdFeriado, locNu, false);
        }

        /// <summary>
        /// Insere vínculo entre um feriado e uma localidade (cidade) na tabela TB_FERIADO_LOCALIDADE.
        /// </summary>
        public void InserirFeriadoLocalidade(int cdFeriado, int locNu, string usuarioInclusao)
        {
            using var conn = new OdbcConnection(_connectionString);
            conn.Open();

            SetSchema(conn);

            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"
                INSERT INTO TB_FERIADO_LOCALIDADE (
                    CD_FERIADO, LOC_NU, DELETED, USU_INCL,
                    CD_UNI_INCL, DT_HR_INCL, USU_ALT, CD_UNI_ALT,
                    DT_HR_ALT
                ) VALUES (
                    ?, ?, NULL, ?, 
                    175, CURRENT TIMESTAMP, NULL, NULL,
                    NULL
                )";

            cmd.Parameters.AddWithValue("@CD_FERIADO", cdFeriado);
            cmd.Parameters.AddWithValue("@LOC_NU", locNu);
            cmd.Parameters.AddWithValue("@USU_INCL", usuarioInclusao);

            cmd.ExecuteNonQuery();
        }

        /// <summary>
        /// Verifica se um feriado com as mesmas características já existe na tabela TB_FERIADO.
        /// </summary>
        /// <returns>True se já existir; retorna também o CD_FERIADO encontrado.</returns>
        private static bool FeriadoJaExiste(OdbcConnection conn, FeriadoParaBanco feriado, out int cdFeriado)
        {
            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"
                SELECT CD_FERIADO FROM TB_FERIADO 
                WHERE DIA_FERIADO = ? AND MES_FERIADO = ? AND 
                      (ANO_FERIADO = ? OR (? IS NULL AND ANO_FERIADO IS NULL)) AND 
                      CD_TP_FERIADO = ? AND CD_ABRANGENCIA = ? AND CD_PERIODICIDADE = ?";

            cmd.Parameters.AddWithValue("@DIA_FERIADO", feriado.DIA_FERIADO);
            cmd.Parameters.AddWithValue("@MES_FERIADO", feriado.MES_FERIADO);
            cmd.Parameters.AddWithValue("@ANO_FERIADO", feriado.ANO_FERIADO ?? (object)DBNull.Value);
            cmd.Parameters.AddWithValue("@ANO_FERIADO_NULL", feriado.ANO_FERIADO ?? (object)DBNull.Value);
            cmd.Parameters.AddWithValue("@CD_TP_FERIADO", feriado.CD_TP_FERIADO);
            cmd.Parameters.AddWithValue("@CD_ABRANGENCIA", feriado.CD_ABRANGENCIA);
            cmd.Parameters.AddWithValue("@CD_PERIODICIDADE", feriado.CD_PERIODICIDADE);

            var result = cmd.ExecuteScalar();
            cdFeriado = result != null ? Convert.ToInt32(result) : 0;
            return result != null;
        }

        /// <summary>
        /// Define o que fazer quando o feriado já existe: ignora ou cria vínculo com a cidade se ainda não houver (caso municipal).
        /// </summary>
        private ResultadoInsercaoFeriado TratarFeriadoExistente(OdbcConnection conn, FeriadoParaBanco feriado, int cdFeriado)
        {
            // Só faz verificação de vínculo se for municipal
            if (feriado.CD_ABRANGENCIA == FeriadoGuids.ABRANG_MUNICIPAL && !string.IsNullOrWhiteSpace(feriado.Cidade))
            {
                int? locNu = ObterLocalidadePorNome(feriado.Cidade);
                if (locNu.HasValue)
                {
                    // Verifica se já está vinculado à cidade
                    if (FeriadoJaVinculado(conn, cdFeriado, locNu.Value))
                    {
                        return new ResultadoInsercaoFeriado(null, null, true);
                    }
                    else
                    {
                        //Feriado existe, mas ainda não vinculado à cidade. Cria o vínculo.
                        return new ResultadoInsercaoFeriado(cdFeriado, locNu, true);
                    }
                }
                else
                {
                    return new ResultadoInsercaoFeriado(null, null, true);
                }
            }

            // Se não for municipal, ignora
            return new ResultadoInsercaoFeriado(null, null, true);
        }

        /// <summary>
        /// Verifica se o feriado já está vinculado à localidade na tabela TB_FERIADO_LOCALIDADE.
        /// </summary>
        private static bool FeriadoJaVinculado(OdbcConnection conn, int cdFeriado, int locNu)
        {
            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"
                SELECT 1 FROM TB_FERIADO_LOCALIDADE
                WHERE CD_FERIADO = ? AND LOC_NU = ?";

            cmd.Parameters.AddWithValue("@CD_FERIADO", cdFeriado);
            cmd.Parameters.AddWithValue("@LOC_NU", locNu);

            return cmd.ExecuteScalar() != null;
        }

        /// <summary>
        /// Insere o novo feriado na tabela TB_FERIADO e retorna o CD_FERIADO gerado.
        /// </summary>
        private static int InserirNovoFeriado(OdbcConnection conn, FeriadoParaBanco feriado)
        {
            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"
                INSERT INTO TB_FERIADO (
                    DS_FERIADO, DIA_FERIADO, MES_FERIADO, ANO_FERIADO,
                    CD_TP_FERIADO, CD_ABRANGENCIA, CD_PERIODICIDADE, CD_UF,
                    IND_ATIVO, DELETED, USU_INCL, CD_UNI_INCL, DT_HR_INCL,
                    USU_ALT, CD_UNI_ALT, DT_HR_ALT, CD_GUID_REFERENCIA_PAI
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)";

            cmd.Parameters.AddWithValue("@DS_FERIADO", feriado.DS_FERIADO);
            cmd.Parameters.AddWithValue("@DIA_FERIADO", feriado.DIA_FERIADO);
            cmd.Parameters.AddWithValue("@MES_FERIADO", feriado.MES_FERIADO);
            cmd.Parameters.AddWithValue("@ANO_FERIADO", feriado.ANO_FERIADO ?? (object)DBNull.Value);
            cmd.Parameters.AddWithValue("@CD_TP_FERIADO", feriado.CD_TP_FERIADO);
            cmd.Parameters.AddWithValue("@CD_ABRANGENCIA", feriado.CD_ABRANGENCIA);
            cmd.Parameters.AddWithValue("@CD_PERIODICIDADE", feriado.CD_PERIODICIDADE);
            cmd.Parameters.AddWithValue("@CD_UF", feriado.CD_UF);
            cmd.Parameters.AddWithValue("@IND_ATIVO", feriado.IND_ATIVO);
            cmd.Parameters.AddWithValue("@DELETED", feriado.DELETED);
            cmd.Parameters.AddWithValue("@USU_INCL", feriado.USU_INCL);
            cmd.Parameters.AddWithValue("@CD_UNI_INCL", feriado.CD_UNI_INCL);
            cmd.Parameters.AddWithValue("@DT_HR_INCL", feriado.DT_HR_INCL);
            cmd.Parameters.AddWithValue("@USU_ALT", DBNull.Value);
            cmd.Parameters.AddWithValue("@CD_UNI_ALT", DBNull.Value);
            cmd.Parameters.AddWithValue("@DT_HR_ALT", DBNull.Value);
            cmd.Parameters.AddWithValue("@CD_GUID_REFERENCIA_PAI", feriado.CD_GUID_REFERENCIA_PAI);

            cmd.ExecuteNonQuery();

            // Recupera o ID gerado para o novo registro
            using var idCmd = conn.CreateCommand();
            idCmd.CommandText = "SELECT IDENTITY_VAL_LOCAL() FROM SYSIBM.SYSDUMMY1";
            return Convert.ToInt32(idCmd.ExecuteScalar());
        }

        /// <summary>
        /// Se o feriado for municipal, retorna o código da localidade correspondente.
        /// </summary>
        private int? ObterLocNuSeMunicipal(FeriadoParaBanco feriado)
        {
            if (feriado.CD_ABRANGENCIA == FeriadoGuids.ABRANG_MUNICIPAL && !string.IsNullOrWhiteSpace(feriado.Cidade))
            {
                var locNu = ObterLocalidadePorNome(feriado.Cidade);
                if (!locNu.HasValue)
                    Console.WriteLine($"!!! Localidade não encontrada ao tentar vincular novo feriado: {feriado.Cidade}");

                return locNu;
            }

            return null;
        }

        /// <summary>
        /// Retorna o código da localidade (LOC_NU) a partir do nome da cidade.
        /// </summary>
        private int? ObterLocalidadePorNome(string nomeCidade)
        {
            using var conn = new OdbcConnection(_connectionString);
            conn.Open();

            SetSchema(conn);

            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"
                SELECT LOC_NU, LOC_NO
                FROM TB_LOCALIDADE WHERE UFE_SG = 'SP'";

            using var reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                if (reader["LOC_NO"] is string locNo)
                {
                    // Remove acentos e compara em minúsculo
                    var nomeSemAcento = TextoUtils.RemoverAcentos(nomeCidade).ToLower();
                    var locNoSemAcento = TextoUtils.RemoverAcentos(locNo).ToLower();

                    if (locNoSemAcento == nomeSemAcento)
                    {
                        return reader.GetInt32(reader.GetOrdinal("LOC_NU"));
                    }
                }
            }
            return null;
        }

        /// <summary>
        /// Busca feriados nacionais e estaduais cadastrados no banco de dados para um determinado ano.
        /// Feriados com ANO_FERIADO nulo são considerados como anuais e assumem o ano informado na busca.
        /// </summary>
        public List<FeriadoNE> BuscarFeriadosNacionaisEstaduais(int ano)
        {
            var feriados = new List<FeriadoNE>();

            using var conn = new OdbcConnection(_connectionString);
            conn.Open();
            SetSchema(conn);

            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"
                SELECT 
                    DIA_FERIADO, MES_FERIADO, 
                    COALESCE(ANO_FERIADO, ?) AS ANO_FERIADO,
                    DS_FERIADO,
                    CD_PERIODICIDADE
                FROM TB_FERIADO
                WHERE CD_ABRANGENCIA IN (?, ?)
                AND (ANO_FERIADO = ? OR ANO_FERIADO IS NULL)
                ORDER BY MES_FERIADO, DIA_FERIADO";

            cmd.Parameters.AddWithValue("@ANO1", ano);
            cmd.Parameters.AddWithValue("@ABR_NACIONAL", FeriadoGuids.ABRANG_NACIONAL);
            cmd.Parameters.AddWithValue("@ABR_ESTADUAL", FeriadoGuids.ABRANG_ESTADUAL);
            cmd.Parameters.AddWithValue("@ANO2", ano);

            using var reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                int dia = Convert.ToInt32(reader["DIA_FERIADO"]);
                int mes = Convert.ToInt32(reader["MES_FERIADO"]);
                int anoFeriado = Convert.ToInt32(reader["ANO_FERIADO"]);

                feriados.Add(new FeriadoNE
                {
                    Data = new DateTime(anoFeriado, mes, dia),
                    Descricao = reader["DS_FERIADO"].ToString() ?? "",
                    Periodicidade = (reader["CD_PERIODICIDADE"]?.ToString() == FeriadoGuids.PERIODIC_ANUAL) ? "ANUAL" : "EVENTUAL"
                });
            }

            return feriados;
        }

        /// <summary>
        /// Busca feriados municipais cadastrados no banco para um determinado ano, vinculados às localidades (cidades).
        /// Feriados com ANO_FERIADO nulo são considerados como anuais e assumem o ano informado.
        /// Feriados relacionados à "Consciência Negra" são ignorados, pois antes de 2024 não eram feriados nacionais, então no banco ainda estão relacionados com algumas cidades.
        /// </summary>
        public List<FeriadoMunicipalBruto> BuscarFeriadosMunicipais(int ano)
        {
            var feriados = new List<FeriadoMunicipalBruto>();

            using var conn = new OdbcConnection(_connectionString);
            conn.Open();
            SetSchema(conn);

            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"
                SELECT 
                    L.LOC_NO AS CIDADE,
                    F.DS_FERIADO,
                    F.DIA_FERIADO,
                    F.MES_FERIADO,
                    COALESCE(F.ANO_FERIADO, ?) AS ANO_FERIADO,
                    F.CD_TP_FERIADO
                FROM TB_FERIADO F
                INNER JOIN TB_FERIADO_LOCALIDADE FL ON FL.CD_FERIADO = F.CD_FERIADO
                INNER JOIN TB_LOCALIDADE L ON L.LOC_NU = FL.LOC_NU
                WHERE F.CD_ABRANGENCIA = ?
                AND (F.ANO_FERIADO = ? OR F.ANO_FERIADO IS NULL)
                ORDER BY L.LOC_NO";

            cmd.Parameters.AddWithValue("@ANO", ano);
            cmd.Parameters.AddWithValue("@ABR_MUNICIPAL", FeriadoGuids.ABRANG_MUNICIPAL);
            cmd.Parameters.AddWithValue("@ANO2", ano);

            using var reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                string descricao = reader["DS_FERIADO"]?.ToString() ?? "";
                string cidade = reader["CIDADE"].ToString() ?? "";
                string tipo = reader["CD_TP_FERIADO"]?.ToString() ?? "";
                int dia = Convert.ToInt32(reader["DIA_FERIADO"]);
                int mes = Convert.ToInt32(reader["MES_FERIADO"]);
                int anoF = Convert.ToInt32(reader["ANO_FERIADO"]);

                // Ignora feriados relacionados à Consciência Negra
                if (descricao.ToLower().Contains("consci") && descricao.ToLower().Contains("negra"))
                    continue;

                feriados.Add(new FeriadoMunicipalBruto
                {
                    Cidade = cidade,
                    Data = new DateTime(anoF, mes, dia),
                    Descricao = descricao,
                    Tipo = tipo
                });
            }

            return feriados;
        }
    }
}