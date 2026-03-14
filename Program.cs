using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using ClosedXML.Excel;

class Program
{
    // Lê arquivo CSV, XLS ou XLSX e retorna a lista de sorteios (cada sorteio = lista de 6 inteiros)
    static List<List<int>> LerArquivo(string caminho)
    {
        var sorteios = new List<List<int>>();
        string ext = Path.GetExtension(caminho)?.ToLower() ?? "";

        if (ext == ".csv")
        {
            var linhas = File.Exists(caminho) ? File.ReadAllLines(caminho) : Array.Empty<string>();
            if (linhas.Length <= 1) return sorteios;

            string cabecalho = linhas[0] ?? string.Empty;
            char separador = cabecalho.Contains(",") ? ',' : ';';

            foreach (var linha in linhas.Skip(1))
            {
                if (string.IsNullOrWhiteSpace(linha)) continue;

                var colunas = linha.Split(separador);
                if (colunas.Length < 8) continue;

                var numeros = new List<int>();
                for (int i = 2; i <= 7; i++)
                {
                    string bruto = (colunas[i] ?? string.Empty).Trim().Trim('"', '\'');
                    if (int.TryParse(bruto, NumberStyles.Integer, CultureInfo.InvariantCulture, out int numero))
                        numeros.Add(numero);
                }

                if (numeros.Count == 6)
                    sorteios.Add(numeros.OrderBy(x => x).ToList());
            }
        }
        else if (ext == ".xls" || ext == ".xlsx")
        {
            using (var wb = new XLWorkbook(caminho))
            {
                var ws = wb.Worksheets.First();
                // Considera primeira linha como cabeçalho e os 6 números nas colunas C..H (3..8)
                foreach (var row in ws.RowsUsed().Skip(1))
                {
                    var numeros = new List<int>();
                    for (int col = 3; col <= 8; col++)
                    {
                        // Tratamento de possível nulo na célula
                        var cellValue = (row.Cell(col).GetValue<string>() ?? string.Empty).Trim();
                        if (!int.TryParse(cellValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out int numero))
                        {
                            numeros.Clear();
                            break;
                        }
                        numeros.Add(numero);
                    }

                    if (numeros.Count == 6)
                        sorteios.Add(numeros.OrderBy(x => x).ToList());
                }
            }
        }

        return sorteios;
    }

    // Gera combinações de tamanho k a partir de uma lista
    static IEnumerable<IEnumerable<int>> Combinacoes(List<int> lista, int k)
    {
        if (k == 0) { yield return new List<int>(); yield break; }
        for (int i = 0; i < lista.Count; i++)
        {
            var resto = lista.Skip(i + 1).ToList();
            foreach (var comb in Combinacoes(resto, k - 1))
                yield return new List<int> { lista[i] }.Concat(comb);
        }
    }

    // Calcula frequências de combinações de tamanho "tamanho"
    static List<(string chave, int freq, double prob)> CalcularCombinacoes(List<List<int>> sorteios, int tamanho)
    {
        var frequencias = new Dictionary<string, int>();

        foreach (var sorteio in sorteios)
        {
            foreach (var comb in Combinacoes(sorteio, tamanho))
            {
                var chave = string.Join("-", comb.OrderBy(x => x));
                if (!frequencias.ContainsKey(chave)) frequencias[chave] = 0;
                frequencias[chave]++;
            }
        }

        int total = frequencias.Values.Sum();
        if (total == 0) total = 1;

        return frequencias
            .OrderByDescending(x => x.Value)
            .Select(x => (x.Key, x.Value, (double)x.Value / total))
            .ToList();
    }

    // Frequência absoluta de cada número (1..60)
    static List<(int numero, int freq)> FrequenciasNumerosAbsolutas(List<List<int>> sorteios)
    {
        var freq = new Dictionary<int, int>();
        for (int i = 1; i <= 60; i++) freq[i] = 0;

        foreach (var sorteio in sorteios)
            foreach (var numero in sorteio)
                freq[numero]++;

        return freq.OrderByDescending(kv => kv.Value)
                   .Select(kv => (kv.Key, kv.Value)).ToList();
    }

    // Probabilidade por sorteio (freq/total de sorteios)
    static List<(int numero, int freq, double prob)> CalcularNumerosPorSorteio(List<List<int>> sorteios)
    {
        int totalSorteios = sorteios.Count;
        if (totalSorteios == 0) totalSorteios = 1;

        return FrequenciasNumerosAbsolutas(sorteios)
            .Where(x => x.freq > 0)
            .Select(x => (x.numero, x.freq, (double)x.freq / totalSorteios))
            .OrderByDescending(x => x.freq)
            .ToList();
    }

    // Probabilidade por aparições (freq/total de aparições)
    static List<(int numero, int freq, double prob)> CalcularNumerosPorAparicoes(List<List<int>> sorteios)
    {
        int totalAparicoes = sorteios.Count * 6;
        if (totalAparicoes == 0) totalAparicoes = 1;

        return FrequenciasNumerosAbsolutas(sorteios)
            .Where(x => x.freq > 0)
            .Select(x => (x.numero, x.freq, (double)x.freq / totalAparicoes))
            .OrderByDescending(x => x.freq)
            .ToList();
    }

    // Exibe tabela de combinações no console
    static void MostrarTabelaCombinacoes(List<List<int>> sorteios, int tamanho)
    {
        var top = CalcularCombinacoes(sorteios, tamanho)
                    .OrderByDescending(x => x.freq)
                    .Take(6)
                    .ToList();

        Console.WriteLine($"\n=== Top combinações de {tamanho} números ===");
        Console.WriteLine($"{"Combinação",-30} | {"Frequência",10} | {"Probabilidade",12}");
        Console.WriteLine(new string('-', 60));

        if (top.Count == 0)
        {
            Console.WriteLine("Nenhuma combinação encontrada.");
            return;
        }

        foreach (var item in top)
            Console.WriteLine($"{item.chave,-30} | {item.freq,10} | {item.prob,10:P2}");
    }

    // Exibe tabela de números no console
    static void MostrarTabelaNumeros(List<List<int>> sorteios, string tipo)
    {
        var lista = (tipo == "sorteio") ? CalcularNumerosPorSorteio(sorteios) : CalcularNumerosPorAparicoes(sorteios);
        string titulo = tipo == "sorteio" ? "Top números isolados (por sorteio)" : "Top números isolados (por aparições)";

        Console.WriteLine($"\n=== {titulo} ===");
        Console.WriteLine($"{"Número",-10} | {"Frequência",10} | {"Probabilidade",12}");
        Console.WriteLine(new string('-', 40));

        if (lista.Count == 0)
        {
            Console.WriteLine("Nenhum número encontrado.");
            return;
        }

        foreach (var item in lista.Take(6))
            Console.WriteLine($"{item.numero,-10} | {item.freq,10} | {item.prob,10:P2}");
    }

    // Exporta todas as análises para Excel usando ClosedXML (sem gráficos)
    static void ExportarParaExcelSemGraficos(List<List<int>> sorteios, string pastaData)
    {
        Directory.CreateDirectory(pastaData);
        string dataHora = DateTime.Now.ToString("dd-MM-yyyy_HH-mm");
        string nomeArquivo = $"analises_consolidado_{dataHora}.xlsx";
        string caminhoExcel = Path.Combine(pastaData, nomeArquivo);

        var freqList = FrequenciasNumerosAbsolutas(sorteios);
        int totalSorteios = sorteios.Count; if (totalSorteios == 0) totalSorteios = 1;
        int totalAparicoes = sorteios.Count * 6; if (totalAparicoes == 0) totalAparicoes = 1;

        using (var wb = new XLWorkbook())
        {
            // Planilha: Números Comparação
            var wsComp = wb.Worksheets.Add("Números Comparação");
            wsComp.Cell(1, 1).Value = "Número";
            wsComp.Cell(1, 2).Value = "Frequência";
            wsComp.Cell(1, 3).Value = "Probabilidade (por sorteio)";
            wsComp.Cell(1, 4).Value = "Probabilidade (por aparições)";

            int r = 2;
            foreach ((int numero, int freq) in freqList)
            {
                wsComp.Cell(r, 1).Value = numero;
                wsComp.Cell(r, 2).Value = freq;
                wsComp.Cell(r, 3).Value = (double)freq / totalSorteios;
                wsComp.Cell(r, 4).Value = (double)freq / totalAparicoes;
                r++;
            }

            if (r > 2)
            {
                wsComp.Column(3).Style.NumberFormat.Format = "0.00%";
                wsComp.Column(4).Style.NumberFormat.Format = "0.00%";
                wsComp.Range(1, 1, r - 1, 4).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            }
            wsComp.Columns().AdjustToContents();

            // Planilha: Top10 Números
            var top10 = freqList.Take(10).ToList();
            var wsTop = wb.Worksheets.Add("Top10 Números");
            wsTop.Cell(1, 1).Value = "Número";
            wsTop.Cell(1, 2).Value = "Frequência";
            wsTop.Cell(1, 3).Value = "Prob (por sorteio)";

            int row = 2;
            foreach ((int numero, int freq) in top10)
            {
                wsTop.Cell(row, 1).Value = numero;
                wsTop.Cell(row, 2).Value = freq;
                wsTop.Cell(row, 3).Value = (double)freq / totalSorteios;
                row++;
            }
            if (row > 2)
            {
                wsTop.Column(3).Style.NumberFormat.Format = "0.00%";
                wsTop.Range(1, 1, row - 1, 3).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            }
            wsTop.Columns().AdjustToContents();

            // Planilhas: Top2..Top6 (todas as combinações)
            for (int k = 2; k <= 6; k++)
            {
                var combs = CalcularCombinacoes(sorteios, k)
                            .OrderByDescending(x => x.freq)
                            .ToList();

                var ws = wb.Worksheets.Add($"Top{k}");
                ws.Cell(1, 1).Value = "Combinação";
                ws.Cell(1, 2).Value = "Frequência";
                ws.Cell(1, 3).Value = "Probabilidade";

                int rr = 2;
                foreach (var item in combs)
                {
                    ws.Cell(rr, 1).Value = item.chave;
                    ws.Cell(rr, 2).Value = item.freq;
                    ws.Cell(rr, 3).Value = item.prob;
                    rr++;
                }

                if (rr > 2)
                {
                    ws.Column(3).Style.NumberFormat.Format = "0.00%";
                    ws.Range(1, 1, rr - 1, 3).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                }
                ws.Columns().AdjustToContents();
            }

            wb.SaveAs(caminhoExcel);
        }

        Console.WriteLine($"\nArquivo Excel consolidado gerado em: {caminhoExcel}");
    }

    static void Main(string[] args)
    {
        string pastaData = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Data");
        // Tenta localizar automaticamente "sorteio" com CSV, XLSX ou XLS na pasta Data
        string? caminhoArquivo = null;
        var candidatos = new[]
        {
            Path.Combine(pastaData, "sorteio.csv"),
            Path.Combine(pastaData, "sorteio.xlsx"),
            Path.Combine(pastaData, "sorteio.xls"),
        };
        foreach (var c in candidatos)
        {
            if (File.Exists(c)) { caminhoArquivo = c; break; }
        }

        if (caminhoArquivo == null)
        {
            Console.WriteLine("Arquivo de entrada não encontrado em:");
            foreach (var c in candidatos) Console.WriteLine(" - " + c);
            Console.WriteLine("Crie a pasta Data e coloque 'sorteio.csv' ou 'sorteio.xlsx' ou 'sorteio.xls' nela.");
            return;
        }

        var sorteios = LerArquivo(caminhoArquivo);
        Console.WriteLine($"Sorteios válidos carregados: {sorteios.Count}");
        if (sorteios.Count == 0)
        {
            Console.WriteLine("Nenhum sorteio válido foi encontrado.");
            return;
        }

        while (true)
        {
            Console.WriteLine("\n=== Menu Principal ===");
            Console.WriteLine("1 - Números isolados (probabilidade por sorteio)");
            Console.WriteLine("2 - Números isolados (probabilidade por aparições)");
            Console.WriteLine("3 - Duplas");
            Console.WriteLine("4 - Trios");
            Console.WriteLine("5 - Quádruplos");
            Console.WriteLine("6 - Quíntuplos");
            Console.WriteLine("7 - Sextetos");
            Console.WriteLine("13 - Consolidado em Excel (todas as análises, sem gráficos)");
            Console.WriteLine("0 - Sair");
            Console.Write("\nDigite sua escolha: ");

            var entrada = Console.ReadLine();
            if (!int.TryParse(entrada, out int escolha))
            {
                Console.WriteLine("Entrada inválida. Digite apenas números.");
                continue;
            }

            if (escolha == 0)
            {
                Console.WriteLine("Encerrando o programa...");
                break;
            }
            else if (escolha == 1)
            {
                MostrarTabelaNumeros(sorteios, "sorteio");
            }
            else if (escolha == 2)
            {
                MostrarTabelaNumeros(sorteios, "aparicoes");
            }
            else if (escolha >= 3 && escolha <= 7)
            {
                int tamanho = escolha - 1;
                MostrarTabelaCombinacoes(sorteios, tamanho);
            }
            else if (escolha == 13)
            {
                try
                {
                    ExportarParaExcelSemGraficos(sorteios, pastaData);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Falha ao gerar Excel: " + ex.Message);
                }
            }
            else
            {
                Console.WriteLine("Opção inválida. Escolha entre 0, 1..7 ou 13.");
            }
        }
    }
}





//using System;
//using System.Collections.Generic;
//using System.Globalization;
//using System.IO;
//using System.Linq;
//using System.Text;
//using ExcelDataReader;
//using ClosedXML.Excel;

//class Program
//{
//    static List<List<int>> LerArquivo(string caminho)
//    {
//        var sorteios = new List<List<int>>();
//        string ext = Path.GetExtension(caminho)?.ToLower() ?? "";

//        if (ext == ".csv")
//        {
//            var linhas = File.Exists(caminho) ? File.ReadAllLines(caminho) : Array.Empty<string>();
//            if (linhas.Length <= 1) return sorteios;

//            string cabecalho = linhas[0] ?? string.Empty;
//            char separador = cabecalho.Contains(",") ? ',' : ';';

//            foreach (var linha in linhas.Skip(1))
//            {
//                if (string.IsNullOrWhiteSpace(linha)) continue;

//                var colunas = linha.Split(separador);
//                if (colunas == null || colunas.Length < 8) continue;

//                var numeros = new List<int>();
//                for (int i = 2; i <= 7; i++)
//                {
//                    string bruto = (colunas[i] ?? string.Empty).Trim().Trim('"', '\'');
//                    if (int.TryParse(bruto, NumberStyles.Integer, CultureInfo.InvariantCulture, out int numero))
//                        numeros.Add(numero);
//                }

//                if (numeros.Count == 6)
//                    sorteios.Add(numeros.OrderBy(x => x).ToList());
//            }
//        }
//        else if (ext == ".xls" || ext == ".xlsx")
//        {
//            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

//            using (var stream = File.Open(caminho, FileMode.Open, FileAccess.Read))
//            using (var reader = ExcelReaderFactory.CreateReader(stream))
//            {
//                int rowIndex = 0;
//                do
//                {
//                    while (reader.Read())
//                    {
//                        if (rowIndex == 0) { rowIndex++; continue; }

//                        var numeros = new List<int>();
//                        for (int i = 2; i <= 7; i++)
//                        {
//                            var val = reader.GetValue(i);
//                            var s = (val?.ToString() ?? string.Empty).Trim();
//                            if (!int.TryParse(s, out int numero)) { numeros.Clear(); break; }
//                            numeros.Add(numero);
//                        }

//                        if (numeros.Count == 6)
//                            sorteios.Add(numeros.OrderBy(x => x).ToList());

//                        rowIndex++;
//                    }
//                } while (reader.NextResult());
//            }
//        }

//        return sorteios;
//    }

//    static IEnumerable<IEnumerable<int>> Combinacoes(List<int> lista, int k)
//    {
//        if (k == 0) { yield return new List<int>(); yield break; }
//        for (int i = 0; i < lista.Count; i++)
//        {
//            var resto = lista.Skip(i + 1).ToList();
//            foreach (var comb in Combinacoes(resto, k - 1))
//                yield return new List<int> { lista[i] }.Concat(comb);
//        }
//    }

//    static List<(string chave, int freq, double prob)> CalcularCombinacoes(List<List<int>> sorteios, int tamanho)
//    {
//        var frequencias = new Dictionary<string, int>();

//        foreach (var sorteio in sorteios)
//        {
//            foreach (var comb in Combinacoes(sorteio, tamanho))
//            {
//                var chave = string.Join("-", comb.OrderBy(x => x));
//                if (!frequencias.ContainsKey(chave)) frequencias[chave] = 0;
//                frequencias[chave]++;
//            }
//        }

//        int total = frequencias.Values.Sum();
//        if (total == 0) total = 1;

//        return frequencias
//            .OrderByDescending(x => x.Value)
//            .Select(x => (x.Key, x.Value, (double)x.Value / total))
//            .ToList();
//    }

//    static List<(int numero, int freq)> FrequenciasNumerosAbsolutas(List<List<int>> sorteios)
//    {
//        var freq = new Dictionary<int, int>();
//        for (int i = 1; i <= 60; i++) freq[i] = 0;

//        foreach (var sorteio in sorteios)
//            foreach (var numero in sorteio)
//                freq[numero]++;

//        return freq.OrderByDescending(kv => kv.Value)
//                   .Select(kv => (kv.Key, kv.Value)).ToList();
//    }

//    static List<(int numero, int freq, double prob)> CalcularNumerosPorSorteio(List<List<int>> sorteios)
//    {
//        int totalSorteios = sorteios.Count; if (totalSorteios == 0) totalSorteios = 1;

//        return FrequenciasNumerosAbsolutas(sorteios)
//            .Where(x => x.freq > 0)
//            .Select(x => (x.numero, x.freq, (double)x.freq / totalSorteios))
//            .ToList();
//    }

//    static List<(int numero, int freq, double prob)> CalcularNumerosPorAparicoes(List<List<int>> sorteios)
//    {
//        int totalAparicoes = sorteios.Count * 6; if (totalAparicoes == 0) totalAparicoes = 1;

//        return FrequenciasNumerosAbsolutas(sorteios)
//            .Where(x => x.freq > 0)
//            .Select(x => (x.numero, x.freq, (double)x.freq / totalAparicoes))
//            .ToList();
//    }

//    static void MostrarTabelaCombinacoes(List<List<int>> sorteios, int tamanho)
//    {
//        var top = CalcularCombinacoes(sorteios, tamanho)
//                    .OrderByDescending(x => x.freq)
//                    .Take(6)
//                    .ToList();

//        Console.WriteLine($"\n=== Top combinações de {tamanho} números ===");
//        Console.WriteLine($"{"Combinação",-30} | {"Frequência",10} | {"Probabilidade",12}");
//        Console.WriteLine(new string('-', 60));

//        if (top.Count == 0)
//        {
//            Console.WriteLine("Nenhuma combinação encontrada.");
//            return;
//        }

//        foreach (var item in top)
//            Console.WriteLine($"{item.chave,-30} | {item.freq,10} | {item.prob,10:P2}");
//    }

//    static void MostrarTabelaNumeros(List<List<int>> sorteios, string tipo)
//    {
//        var top = (tipo == "sorteio") ? CalcularNumerosPorSorteio(sorteios) : CalcularNumerosPorAparicoes(sorteios);
//        string titulo = tipo == "sorteio" ? "Top números isolados (por sorteio)" : "Top números isolados (por aparições)";

//        Console.WriteLine($"\n=== {titulo} ===");
//        Console.WriteLine($"{"Número",-10} | {"Frequência",10} | {"Probabilidade",12}");
//        Console.WriteLine(new string('-', 40));

//        if (top.Count == 0)
//        {
//            Console.WriteLine("Nenhum número encontrado.");
//            return;
//        }

//        foreach (var item in top.Take(6))
//            Console.WriteLine($"{item.numero,-10} | {item.freq,10} | {item.prob,10:P2}");
//    }

//    // Exporta todas as análises para Excel usando ClosedXML (sem gráficos)
//    static void ExportarParaExcelSemGraficos(List<List<int>> sorteios, string pastaData)
//    {
//        Directory.CreateDirectory(pastaData);
//        string dataHora = DateTime.Now.ToString("dd-MM-yyyy_HH-mm");
//        string nomeArquivo = $"analises_consolidado_{dataHora}.xlsx";
//        string caminhoExcel = Path.Combine(pastaData, nomeArquivo);

//        var freqList = FrequenciasNumerosAbsolutas(sorteios);
//        int totalSorteios = sorteios.Count; if (totalSorteios == 0) totalSorteios = 1;
//        int totalAparicoes = sorteios.Count * 6; if (totalAparicoes == 0) totalAparicoes = 1;

//        using (var wb = new XLWorkbook())
//        {
//            // Planilha: Números Comparação
//            var wsComp = wb.Worksheets.Add("Números Comparação");
//            wsComp.Cell(1, 1).Value = "Número";
//            wsComp.Cell(1, 2).Value = "Frequência";
//            wsComp.Cell(1, 3).Value = "Probabilidade (por sorteio)";
//            wsComp.Cell(1, 4).Value = "Probabilidade (por aparições)";

//            int r = 2;
//            foreach ((int numero, int freq) in freqList)
//            {
//                wsComp.Cell(r, 1).Value = numero;
//                wsComp.Cell(r, 2).Value = freq;
//                wsComp.Cell(r, 3).Value = (double)freq / totalSorteios;
//                wsComp.Cell(r, 4).Value = (double)freq / totalAparicoes;
//                r++;
//            }

//            if (r > 2)
//            {
//                wsComp.Range(1, 1, r - 1, 4).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
//                wsComp.Column(3).Style.NumberFormat.Format = "0.00%";
//                wsComp.Column(4).Style.NumberFormat.Format = "0.00%";
//            }
//            wsComp.Columns().AdjustToContents();

//            // Planilha: Top10 Números
//            var top10 = freqList.Take(10).ToList();
//            var wsTop = wb.Worksheets.Add("Top10 Números");
//            wsTop.Cell(1, 1).Value = "Número";
//            wsTop.Cell(1, 2).Value = "Frequência";
//            wsTop.Cell(1, 3).Value = "Prob (por sorteio)";

//            int row = 2;
//            foreach ((int numero, int freq) in top10)
//            {
//                wsTop.Cell(row, 1).Value = numero;
//                wsTop.Cell(row, 2).Value = freq;
//                wsTop.Cell(row, 3).Value = (double)freq / totalSorteios;
//                row++;
//            }
//            if (row > 2)
//            {
//                wsTop.Range(1, 1, row - 1, 3).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
//                wsTop.Column(3).Style.NumberFormat.Format = "0.00%";
//            }
//            wsTop.Columns().AdjustToContents();

//            // Planilhas: Top2..Top6 (combinações)
//            for (int k = 2; k <= 6; k++)
//            {
//                var combs = CalcularCombinacoes(sorteios, k)
//                            .OrderByDescending(x => x.freq)
//                            .ToList();

//                var ws = wb.Worksheets.Add($"Top{k}");
//                ws.Cell(1, 1).Value = "Combinação";
//                ws.Cell(1, 2).Value = "Frequência";
//                ws.Cell(1, 3).Value = "Probabilidade";

//                int rr = 2;
//                foreach (var item in combs)
//                {
//                    ws.Cell(rr, 1).Value = item.chave;
//                    ws.Cell(rr, 2).Value = item.freq;
//                    ws.Cell(rr, 3).Value = item.prob;
//                    rr++;
//                }

//                if (rr > 2)
//                {
//                    ws.Range(1, 1, rr - 1, 3).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
//                    ws.Column(3).Style.NumberFormat.Format = "0.00%";
//                }
//                ws.Columns().AdjustToContents();
//            }

//            wb.SaveAs(caminhoExcel);
//        }

//        Console.WriteLine($"\nArquivo Excel consolidado gerado em: {caminhoExcel}");
//    }

//    static void Main(string[] args)
//    {
//        string pastaData = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Data");
//        string caminhoArquivo = Path.Combine(pastaData, "sorteio.csv");

//        if (!File.Exists(caminhoArquivo))
//        {
//            Console.WriteLine("Arquivo não encontrado em: " + caminhoArquivo);
//            return;
//        }

//        var sorteios = LerArquivo(caminhoArquivo);
//        Console.WriteLine($"Sorteios válidos carregados: {sorteios.Count}");
//        if (sorteios.Count == 0)
//        {
//            Console.WriteLine("Nenhum sorteio válido foi encontrado.");
//            return;
//        }

//        while (true)
//        {
//            Console.WriteLine("\n=== Menu Principal ===");
//            Console.WriteLine("1 - Números isolados (probabilidade por sorteio)");
//            Console.WriteLine("2 - Números isolados (probabilidade por aparições)");
//            Console.WriteLine("3 - Duplas");
//            Console.WriteLine("4 - Trios");
//            Console.WriteLine("5 - Quádruplos");
//            Console.WriteLine("6 - Quíntuplos");
//            Console.WriteLine("7 - Sextetos");
//            Console.WriteLine("13 - Consolidado em Excel (todas as análises, sem gráficos)");
//            Console.WriteLine("0 - Sair");
//            Console.Write("\nDigite sua escolha: ");

//            var entrada = Console.ReadLine();
//            if (!int.TryParse(entrada, out int escolha))
//            {
//                Console.WriteLine("Entrada inválida. Digite apenas números.");
//                continue;
//            }

//            if (escolha == 0)
//            {
//                Console.WriteLine("Encerrando o programa...");
//                break;
//            }
//            else if (escolha == 1)
//            {
//                MostrarTabelaNumeros(sorteios, "sorteio");
//            }
//            else if (escolha == 2)
//            {
//                MostrarTabelaNumeros(sorteios, "aparicoes");
//            }
//            else if (escolha >= 3 && escolha <= 7)
//            {
//                int tamanho = escolha - 1;
//                MostrarTabelaCombinacoes(sorteios, tamanho);
//            }
//            else if (escolha == 13)
//            {
//                try
//                {
//                    ExportarParaExcelSemGraficos(sorteios, pastaData);
//                }
//                catch (Exception ex)
//                {
//                    Console.WriteLine("Falha ao gerar Excel: " + ex.Message);
//                }
//            }
//            else
//            {
//                Console.WriteLine("Opção inválida. Escolha entre 0, 1..7 ou 13.");
//            }
//        }
//    }
//}




////using System;
////using System.Collections.Generic;
////using System.Globalization;
////using System.IO;
////using System.Linq;
////using System.Reflection;
////using System.Runtime.CompilerServices;
////using System.Text;
////using OfficeOpenXml;
////using OfficeOpenXml.Drawing.Chart;
////using ExcelDataReader;

////class Program
////{
////    // Module initializer: executa antes de Main e antes de qualquer inicializador estático
////    [ModuleInitializer]
////    public static void ModuleInit()
////    {
////        ConfigureEpplusLicenseNonCommercial();
////    }

////    /// <summary>
////    /// Configura a licença do EPPlus de forma robusta:
////    /// - EPPlus 8: ExcelPackage.License = new LicenseConfiguration { LicenseType = LicenseType.NonCommercial }
////    /// - EPPlus <=7: ExcelPackage.LicenseContext = LicenseContext.NonCommercial
////    /// - Fallback: define variável de ambiente EPPlusLicenseContext=NonCommercial
////    /// Também imprime diagnóstico da assembly carregada e do estado da licença.
////    /// </summary>
////    static void ConfigureEpplusLicenseNonCommercial()
////    {
////        try
////        {
////            var excelPackageType = typeof(ExcelPackage);

////            // 1) Tenta EPPlus 8 (propriedade License)
////            var licenseProp = excelPackageType.GetProperty("License", BindingFlags.Public | BindingFlags.Static);
////            if (licenseProp != null && licenseProp.CanWrite)
////            {
////                var licenseConfigType = Type.GetType("OfficeOpenXml.LicenseConfiguration, EPPlus")
////                                         ?? Type.GetType("OfficeOpenXml.LicenseConfiguration, OfficeOpenXml");
////                var licenseTypeEnum = Type.GetType("OfficeOpenXml.LicenseType, EPPlus")
////                                     ?? Type.GetType("OfficeOpenXml.LicenseType, OfficeOpenXml");

////                if (licenseConfigType != null && licenseTypeEnum != null)
////                {
////                    var licenseConfig = Activator.CreateInstance(licenseConfigType);
////                    var ltProp = licenseConfigType.GetProperty("LicenseType", BindingFlags.Public | BindingFlags.Instance);
////                    if (ltProp != null)
////                    {
////                        var nonCommercial = Enum.Parse(licenseTypeEnum, "NonCommercial");
////                        ltProp.SetValue(licenseConfig, nonCommercial);
////                        licenseProp.SetValue(null, licenseConfig);
////                        // sucesso para EPPlus 8
////                    }
////                }
////            }
////        }
////        catch
////        {
////            // ignora e tenta fallback
////        }

////        try
////        {
////            var excelPackageType = typeof(ExcelPackage);

////            // 2) Tenta EPPlus <=7 (propriedade LicenseContext)
////            var licenseContextProp = excelPackageType.GetProperty("LicenseContext", BindingFlags.Public | BindingFlags.Static);
////            if (licenseContextProp != null && licenseContextProp.CanWrite)
////            {
////                var licenseContextType = Type.GetType("OfficeOpenXml.LicenseContext, EPPlus")
////                                      ?? Type.GetType("OfficeOpenXml.LicenseContext, OfficeOpenXml");
////                if (licenseContextType != null)
////                {
////                    var nonCommercial = Enum.Parse(licenseContextType, "NonCommercial");
////                    licenseContextProp.SetValue(null, nonCommercial);
////                }
////            }
////        }
////        catch
////        {
////            // ignora e tenta fallback
////        }

////        try
////        {
////            // 3) Fallback: define variável de ambiente para o processo atual
////            Environment.SetEnvironmentVariable("EPPlusLicenseContext", "NonCommercial", EnvironmentVariableTarget.Process);
////        }
////        catch
////        {
////            // ignore
////        }

////        // 4) Diagnóstico: imprime versão da assembly EPPlus e estado atual da licença (se acessível)
////        try
////        {
////            var asm = typeof(ExcelPackage).Assembly;
////            Console.WriteLine($"EPPlus assembly: {asm.GetName().Name}, version {asm.GetName().Version}");

////            var excelPackageType = typeof(ExcelPackage);
////            var licenseProp = excelPackageType.GetProperty("License", BindingFlags.Public | BindingFlags.Static);
////            if (licenseProp != null)
////            {
////                var val = licenseProp.GetValue(null);
////                Console.WriteLine($"ExcelPackage.License is set: {val != null}");
////            }
////            else
////            {
////                var licenseContextProp = excelPackageType.GetProperty("LicenseContext", BindingFlags.Public | BindingFlags.Static);
////                if (licenseContextProp != null)
////                {
////                    var val = licenseContextProp.GetValue(null);
////                    Console.WriteLine($"ExcelPackage.LicenseContext = {val}");
////                }
////                else
////                {
////                    Console.WriteLine("No License or LicenseContext property found on ExcelPackage type.");
////                }
////            }
////        }
////        catch (Exception ex)
////        {
////            Console.WriteLine("Erro ao verificar licença EPPlus: " + ex.Message);
////        }
////    }

////    static List<List<int>> LerArquivo(string caminho)
////    {
////        var sorteios = new List<List<int>>();
////        string ext = Path.GetExtension(caminho)?.ToLower() ?? "";

////        if (ext == ".csv")
////        {
////            var linhas = File.Exists(caminho) ? File.ReadAllLines(caminho) : Array.Empty<string>();
////            if (linhas.Length <= 1) return sorteios;

////            string cabecalho = linhas[0] ?? string.Empty;
////            char separador = cabecalho.Contains(",") ? ',' : ';';

////            foreach (var linha in linhas.Skip(1))
////            {
////                if (string.IsNullOrWhiteSpace(linha)) continue;

////                var colunas = linha.Split(separador);
////                if (colunas == null || colunas.Length < 8) continue;

////                var numeros = new List<int>();
////                for (int i = 2; i <= 7; i++)
////                {
////                    string bruto = (colunas[i] ?? string.Empty).Trim().Trim('"', '\'');
////                    if (int.TryParse(bruto, NumberStyles.Integer, CultureInfo.InvariantCulture, out int numero))
////                        numeros.Add(numero);
////                }

////                if (numeros.Count == 6)
////                    sorteios.Add(numeros.OrderBy(x => x).ToList());
////            }
////        }
////        else if (ext == ".xls" || ext == ".xlsx")
////        {
////            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

////            using (var stream = File.Open(caminho, FileMode.Open, FileAccess.Read))
////            using (var reader = ExcelReaderFactory.CreateReader(stream))
////            {
////                int rowIndex = 0;
////                do
////                {
////                    while (reader.Read())
////                    {
////                        if (rowIndex == 0) { rowIndex++; continue; }

////                        var numeros = new List<int>();
////                        for (int i = 2; i <= 7; i++)
////                        {
////                            var val = reader.GetValue(i);
////                            var s = (val?.ToString() ?? string.Empty).Trim();
////                            if (!int.TryParse(s, out int numero)) { numeros.Clear(); break; }
////                            numeros.Add(numero);
////                        }

////                        if (numeros.Count == 6)
////                            sorteios.Add(numeros.OrderBy(x => x).ToList());

////                        rowIndex++;
////                    }
////                } while (reader.NextResult());
////            }
////        }

////        return sorteios;
////    }

////    static IEnumerable<IEnumerable<int>> Combinacoes(List<int> lista, int k)
////    {
////        if (k == 0) { yield return new List<int>(); yield break; }
////        for (int i = 0; i < lista.Count; i++)
////        {
////            var resto = lista.Skip(i + 1).ToList();
////            foreach (var comb in Combinacoes(resto, k - 1))
////                yield return new List<int> { lista[i] }.Concat(comb);
////        }
////    }

////    static List<(string chave, int freq, double prob)> CalcularCombinacoes(List<List<int>> sorteios, int tamanho)
////    {
////        var frequencias = new Dictionary<string, int>();

////        foreach (var sorteio in sorteios)
////        {
////            foreach (var comb in Combinacoes(sorteio, tamanho))
////            {
////                var chave = string.Join("-", comb.OrderBy(x => x));
////                if (!frequencias.ContainsKey(chave)) frequencias[chave] = 0;
////                frequencias[chave]++;
////            }
////        }

////        int total = frequencias.Values.Sum();
////        if (total == 0) total = 1;

////        return frequencias
////            .OrderByDescending(x => x.Value)
////            .Take(6)
////            .Select(x => (x.Key, x.Value, (double)x.Value / total))
////            .ToList();
////    }

////    static List<(int numero, int freq)> FrequenciasNumerosAbsolutas(List<List<int>> sorteios)
////    {
////        var freq = new Dictionary<int, int>();
////        for (int i = 1; i <= 60; i++) freq[i] = 0;

////        foreach (var sorteio in sorteios)
////            foreach (var numero in sorteio)
////                freq[numero]++;

////        return freq.OrderByDescending(kv => kv.Value)
////                   .Select(kv => (kv.Key, kv.Value)).ToList();
////    }

////    static List<(int numero, int freq, double prob)> CalcularNumerosPorSorteio(List<List<int>> sorteios)
////    {
////        int totalSorteios = sorteios.Count; if (totalSorteios == 0) totalSorteios = 1;

////        return FrequenciasNumerosAbsolutas(sorteios)
////            .Where(x => x.freq > 0)
////            .Take(6)
////            .Select(x => (x.numero, x.freq, (double)x.freq / totalSorteios))
////            .ToList();
////    }

////    static List<(int numero, int freq, double prob)> CalcularNumerosPorAparicoes(List<List<int>> sorteios)
////    {
////        int totalAparicoes = sorteios.Count * 6; if (totalAparicoes == 0) totalAparicoes = 1;

////        return FrequenciasNumerosAbsolutas(sorteios)
////            .Where(x => x.freq > 0)
////            .Take(6)
////            .Select(x => (x.numero, x.freq, (double)x.freq / totalAparicoes))
////            .ToList();
////    }

////    static void MostrarTabelaCombinacoes(List<List<int>> sorteios, int tamanho)
////    {
////        var top = CalcularCombinacoes(sorteios, tamanho);

////        Console.WriteLine($"\n=== Top combinações de {tamanho} números ===");
////        Console.WriteLine($"{"Combinação",-25} | {"Frequência",10} | {"Probabilidade",12}");
////        Console.WriteLine(new string('-', 55));

////        if (top.Count == 0)
////        {
////            Console.WriteLine("Nenhuma combinação encontrada.");
////            return;
////        }

////        foreach (var item in top)
////            Console.WriteLine($"{item.chave,-25} | {item.freq,10} | {item.prob,10:P2}");
////    }

////    static void MostrarTabelaNumeros(List<List<int>> sorteios, string tipo)
////    {
////        var top = (tipo == "sorteio") ? CalcularNumerosPorSorteio(sorteios) : CalcularNumerosPorAparicoes(sorteios);
////        string titulo = tipo == "sorteio" ? "Top números isolados (por sorteio)" : "Top números isolados (por aparições)";

////        Console.WriteLine($"\n=== {titulo} ===");
////        Console.WriteLine($"{"Número",-10} | {"Frequência",10} | {"Probabilidade",12}");
////        Console.WriteLine(new string('-', 40));

////        if (top.Count == 0)
////        {
////            Console.WriteLine("Nenhum número encontrado.");
////            return;
////        }

////        foreach (var item in top)
////            Console.WriteLine($"{item.numero,-10} | {item.freq,10} | {item.prob,10:P2}");
////    }

////    static void ExportarParaExcelComComparacaoEGraficos(List<List<int>> sorteios, string pastaData)
////    {
////        Directory.CreateDirectory(pastaData);
////        string dataHora = DateTime.Now.ToString("dd-MM-yyyy_HH-mm");
////        string nomeArquivo = $"analises_consolidado_{dataHora}.xlsx";
////        string caminhoExcel = Path.Combine(pastaData, nomeArquivo);

////        using (var package = new ExcelPackage())
////        {
////            var freqList = FrequenciasNumerosAbsolutas(sorteios);
////            int totalSorteios = sorteios.Count; if (totalSorteios == 0) totalSorteios = 1;
////            int totalAparicoes = sorteios.Count * 6; if (totalAparicoes == 0) totalAparicoes = 1;

////            var wsComp = package.Workbook.Worksheets.Add("Números Comparação");
////            wsComp.Cells[1, 1].Value = "Número";
////            wsComp.Cells[1, 2].Value = "Frequência";
////            wsComp.Cells[1, 3].Value = "Probabilidade (por sorteio)";
////            wsComp.Cells[1, 4].Value = "Probabilidade (por aparições)";

////            int r = 2;
////            foreach (var (numero, freq) in freqList)
////            {
////                wsComp.Cells[r, 1].Value = numero;
////                wsComp.Cells[r, 2].Value = freq;
////                wsComp.Cells[r, 3].Value = (double)freq / totalSorteios;
////                wsComp.Cells[r, 4].Value = (double)freq / totalAparicoes;
////                r++;
////            }
////            if (wsComp.Dimension != null)
////            {
////                wsComp.Cells[2, 3, r - 1, 3].Style.Numberformat.Format = "0.00%";
////                wsComp.Cells[2, 4, r - 1, 4].Style.Numberformat.Format = "0.00%";
////                wsComp.Cells[wsComp.Dimension.Address].AutoFitColumns();
////            }

////            var top10 = freqList.Take(10).ToList();
////            var wsChart = package.Workbook.Worksheets.Add("Gráfico Números");
////            wsChart.Cells[1, 1].Value = "Número";
////            wsChart.Cells[1, 2].Value = "Frequência";
////            wsChart.Cells[1, 3].Value = "Prob (por sorteio)";

////            int row = 2;
////            foreach (var item in top10)
////            {
////                wsChart.Cells[row, 1].Value = item.numero;
////                wsChart.Cells[row, 2].Value = item.freq;
////                wsChart.Cells[row, 3].Value = (double)item.freq / totalSorteios;
////                row++;
////            }
////            if (wsChart.Dimension != null)
////            {
////                wsChart.Cells[2, 3, row - 1, 3].Style.Numberformat.Format = "0.00%";
////                wsChart.Cells[wsChart.Dimension.Address].AutoFitColumns();
////            }

////            var chartNums = wsChart.Drawings.AddChart("Top10Numeros", eChartType.ColumnClustered);
////            chartNums.Title.Text = "Top 10 Números Mais Frequentes";
////            chartNums.SetPosition(0, 0, 4, 0);
////            chartNums.SetSize(800, 420);
////            if (row - 1 >= 2)
////                chartNums.Series.Add(wsChart.Cells[2, 2, row - 1, 2], wsChart.Cells[2, 1, row - 1, 1]).Header = "Frequência";

////            for (int k = 2; k <= 6; k++)
////            {
////                var topComb = CalcularCombinacoes(sorteios, k);
////                var ws = package.Workbook.Worksheets.Add($"Top{k}");

////                ws.Cells[1, 1].Value = "Combinação";
////                ws.Cells[1, 2].Value = "Frequência";
////                ws.Cells[1, 3].Value = "Probabilidade";

////                int rr = 2;
////                foreach (var item in topComb)
////                {
////                    ws.Cells[rr, 1].Value = item.chave;
////                    ws.Cells[rr, 2].Value = item.freq;
////                    ws.Cells[rr, 3].Value = item.prob;
////                    rr++;
////                }

////                if (ws.Dimension != null)
////                {
////                    ws.Cells[2, 3, rr - 1, 3].Style.Numberformat.Format = "0.00%";
////                    ws.Cells[ws.Dimension.Address].AutoFitColumns();
////                }

////                var chartComb = ws.Drawings.AddChart($"GraficoTop{k}", eChartType.ColumnClustered);
////                chartComb.Title.Text = $"Top {Math.Min(topComb.Count, 10)} Combinações de {k} números";
////                chartComb.SetPosition(0, 0, 4, 0);
////                chartComb.SetSize(800, 420);
////                if (rr - 1 >= 2)
////                    chartComb.Series.Add(ws.Cells[2, 2, rr - 1, 2], ws.Cells[2, 1, rr - 1, 1]).Header = "Frequência";
////            }

////            var fi = new FileInfo(caminhoExcel);
////            package.SaveAs(fi);
////        }

////        Console.WriteLine($"\nArquivo Excel consolidado gerado em: {caminhoExcel}");
////    }

////    static void Main(string[] args)
////    {
////        // ModuleInit já chamou ConfigureEpplusLicenseNonCommercial antes de Main
////        string pastaData = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Data");
////        string caminhoArquivo = Path.Combine(pastaData, "sorteio.csv");

////        if (!File.Exists(caminhoArquivo))
////        {
////            Console.WriteLine("Arquivo não encontrado em: " + caminhoArquivo);
////            return;
////        }

////        var sorteios = LerArquivo(caminhoArquivo);
////        Console.WriteLine($"Sorteios válidos carregados: {sorteios.Count}");
////        if (sorteios.Count == 0)
////        {
////            Console.WriteLine("Nenhum sorteio válido foi encontrado.");
////            return;
////        }

////        while (true)
////        {
////            Console.WriteLine("\n=== Menu Principal ===");
////            Console.WriteLine("1 - Números isolados (probabilidade por sorteio)");
////            Console.WriteLine("2 - Números isolados (probabilidade por aparições)");
////            Console.WriteLine("3 - Duplas");
////            Console.WriteLine("4 - Trios");
////            Console.WriteLine("5 - Quádruplos");
////            Console.WriteLine("6 - Quíntuplos");
////            Console.WriteLine("7 - Sextetos");
////            Console.WriteLine("13 - Consolidado em Excel (comparação + gráficos)");
////            Console.WriteLine("0 - Sair");
////            Console.Write("\nDigite sua escolha: ");

////            var entrada = Console.ReadLine();
////            if (!int.TryParse(entrada, out int escolha))
////            {
////                Console.WriteLine("Entrada inválida. Digite apenas números.");
////                continue;
////            }

////            if (escolha == 0)
////            {
////                Console.WriteLine("Encerrando o programa...");
////                break;
////            }
////            else if (escolha == 1)
////            {
////                MostrarTabelaNumeros(sorteios, "sorteio");
////            }
////            else if (escolha == 2)
////            {
////                MostrarTabelaNumeros(sorteios, "aparicoes");
////            }
////            else if (escolha >= 3 && escolha <= 7)
////            {
////                int tamanho = escolha - 1;
////                MostrarTabelaCombinacoes(sorteios, tamanho);
////            }
////            else if (escolha == 13)
////            {
////                try
////                {
////                    ExportarParaExcelComComparacaoEGraficos(sorteios, pastaData);
////                }
////                catch (Exception ex)
////                {
////                    Console.WriteLine("Falha ao gerar Excel: " + ex.Message);
////                }
////            }
////            else
////            {
////                Console.WriteLine("Opção inválida. Escolha entre 0, 1..7 ou 13.");
////            }
////        }
////    }
////}