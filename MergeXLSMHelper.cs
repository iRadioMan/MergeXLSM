using ClosedXML.Excel;

namespace MergeXLSM
{
    static class MergeXLSMHelper
    {
        public static void CopyRowWithoutFormulas(IXLRow sourceRow, IXLRow targetRow)
        {
            for (int i = 1; i <= sourceRow.CellCount(); i++) // for every cell in the row
            {
                var sourceCell = sourceRow.Cell(i);
                var targetCell = targetRow.Cell(i);

                targetCell.Value = sourceCell.Value; // copy value only
            }
        }
        public static void ParseArgs(string[] args, ref string templateFile, ref string resultFile,
                                  ref string workdir, ref int firstRow, ref int lastRow, ref string rowTag)
        {
            var argsDict = new Dictionary<string, string>();

            for (int i = 0; i < args.Length - 1; i += 2)
            {
                if (args[i].StartsWith('-') && !argsDict.ContainsKey(args[i]))
                    argsDict[args[i]] = args[i + 1];
            }

            if (argsDict.TryGetValue("-i", out string? tempFile))
                templateFile = tempFile;

            if (argsDict.TryGetValue("-o", out string? resFile))
                resultFile = resFile;

            if (argsDict.TryGetValue("-dir", out string? dir))
                workdir = Path.GetFullPath(dir);

            if (argsDict.TryGetValue("-first", out string? first) && int.TryParse(first, out int fRow))
                firstRow = fRow;
            else if (argsDict.ContainsKey("-first"))
                Console.WriteLine("[Warning] Invalid -first value, using default.");

            if (argsDict.TryGetValue("-last", out string? last) && int.TryParse(last, out int lRow))
                lastRow = lRow;
            else if (argsDict.ContainsKey("-last"))
                Console.WriteLine("[Warning] Invalid -last value, using default.");

            if (argsDict.TryGetValue("-rowTag", out string? t))
                rowTag = t;
        }
        public static void PrintArgs(ref string templatePath, ref string resultPath, ref string workDir, 
            ref int firstRow, ref int lastRow, ref string rowTag)
        {
            Console.WriteLine();
            Console.WriteLine("Arguments:");
            Console.WriteLine($"Template file: {templatePath}");
            Console.WriteLine($"Result file: {resultPath}");
            Console.WriteLine($"Working directory: {workDir}");
            Console.WriteLine($"First row: {firstRow}");
            Console.WriteLine($"Last row: {lastRow}");
            Console.WriteLine($"rowTag: {rowTag}");
            Console.WriteLine();
            Console.WriteLine("Cleaning old results...");
        }
        public static void PrintUsage()
        {
            Console.WriteLine("Usage: MergeXLSM.exe -i <template.xlsm> -o <result.xlsm> -dir <directory> -first <number> -last <number> -tag <text>");
            Console.WriteLine();
            Console.WriteLine("Options:");
            Console.WriteLine("  -i <file>     Template XLSM file (default: Template.xlsm)");
            Console.WriteLine("  -o <file>     Output XLSM file (default: Result.xlsm)");
            Console.WriteLine("  -dir <path>   Directory containing XLSM files (default: XLSM)");
            Console.WriteLine("  -first <num>  First row to process (default: 6)");
            Console.WriteLine("  -last <num>   Last row to process (default: 1000)");
            Console.WriteLine("  -tag <text>   Tag to search for in the first column (default: Fact)");
            Console.WriteLine();
            Console.WriteLine("Example:");
            Console.WriteLine("  MergeXLSM.exe -i Template.xlsm -o Result.xlsm -dir XLSM -first 6 -last 1000 -tag Fact");
        }
    }
}
