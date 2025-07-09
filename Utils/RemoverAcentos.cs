using System.Globalization;
using System.Text;

namespace ImportadorFeriados.Utils
{
    // Classe utilitária para manipulação de texto
    public static class TextoUtils
    {
        /// <summary> 
        /// Remove acentos de uma string.
        /// </summary>
        public static string RemoverAcentos(string texto)
        {
            if (string.IsNullOrWhiteSpace(texto))
                return texto;

            var normalized = texto.Normalize(NormalizationForm.FormD);
            var sb = new StringBuilder();

            foreach (var c in normalized)
            {
                var unicodeCategory = CharUnicodeInfo.GetUnicodeCategory(c);
                if (unicodeCategory != UnicodeCategory.NonSpacingMark)
                {
                    sb.Append(c);
                }
            }

            return sb.ToString().Normalize(NormalizationForm.FormC);
        }
    }
}