namespace ImportadorFeriados.Utils
{
    public static class LogUtils
    {
        public static void MostrarBarraProgresso(int atual, int total, int largura = 30)
        {
            double progresso = (double)atual / total;
            int hashtags = (int)(progresso * largura);
            string barra = new string('#', hashtags).PadRight(largura, '-');
            Console.Write($"\rImportando feriados: [{barra}] {progresso:P0}");
        }
    }
}
