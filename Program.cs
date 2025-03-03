using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing;
using System.Text;

Console.OutputEncoding = Encoding.UTF8; // set encoding to utf-8 to support cyrillic symbols

string workdir = Environment.CurrentDirectory + @"\XLSM\";
string templateFile = "Template.xlsm";
string resultFile = "Result.xlsm";
int firstRow = 6;
int lastRow = 1000;
string rowTag = "Fact";

if (args.Contains("-h") || args.Contains("--help"))
{
    PrintUsage();
    return;
}

Console.WriteLine("[ Merge XLSM ]\n");

ParseArgs(args, ref templateFile, ref resultFile, ref workdir, ref firstRow, ref lastRow, ref rowTag);

Console.WriteLine("Arguments:");
Console.WriteLine($"Template file: {templateFile}");
Console.WriteLine($"Result file: {resultFile}");
Console.WriteLine($"Working directory: {workdir}");
Console.WriteLine($"First row: {firstRow}");
Console.WriteLine($"Last row: {lastRow}");
Console.WriteLine($"rowTag: {rowTag}");
Console.WriteLine();

Console.WriteLine("Cleaning results...");

if (File.Exists(workdir + resultFile)) // delete result file if exists
    File.Delete(workdir + resultFile);

if (!File.Exists(workdir + templateFile)) // check if template file exists
{
    Console.WriteLine($"Template file {templateFile} is not found");
    return;
}

Console.WriteLine($"Creating a new worksheet from {templateFile}");

var resultBook = new XLWorkbook(workdir + templateFile); // create new workbook from template that will contain merged results
var ws = resultBook.Worksheet(1); // open first worksheet
var curRow = firstRow; // a current row to copy data

Console.WriteLine("Getting files in a working dir...");

foreach (var file in Directory.GetFiles(workdir)) // get all files in a working dir
{
    if (file.EndsWith(templateFile) || !file.EndsWith(".xlsm")) // skip template and non-xlsm files
        continue;

    Console.WriteLine($"XLSM file found: {file}");

    try
    {
        using (var srcbook = new XLWorkbook(file)) // open source workbook
        {
            var srcsheet = srcbook.Worksheet(1); // open first worksheet in a source workbook
            int cpCnt = 0; // initialize a row counter for source workbook

            for (int i = firstRow; i <= lastRow; ++i)
            {
                if (srcsheet.Cell(i, 1).Value.ToString() == rowTag) // if first cell in a row contains rowTag
                {
                    //srcsheet.Row(i).CopyTo(ws.Row(curRow)); // copy that row to worksheet
                    CopyRowWithoutFormulas(srcsheet.Row(i), ws.Row(curRow));
                    curRow++; // switch to next row
                    cpCnt++; // count copied rows (for only last source workbook to notify user)
                }
            }

            Console.WriteLine($"Rows copied: {cpCnt.ToString()}");
        }
    }
    catch (Exception exc)
    {
        Console.WriteLine($"[Error] Unable to process file: {exc.Message}");
    }
}

Console.WriteLine();
Console.WriteLine($"Total merged rows: {curRow - firstRow}.\nSaving results...");
resultBook.SaveAs(workdir + resultFile);
Console.WriteLine($"Saved to {resultFile}");

static void CopyRowWithoutFormulas(IXLRow sourceRow, IXLRow targetRow)
{
    for (int i = 1; i <= sourceRow.CellCount(); i++) // for every cell in the row
    {
        var sourceCell = sourceRow.Cell(i);
        var targetCell = targetRow.Cell(i);

        targetCell.Value = sourceCell.Value; // copy value only
    }
}

static void ParseArgs(string[] args, ref string templateFile, ref string resultFile,
                          ref string workdir, ref int firstRow, ref int lastRow, ref string rowTag)
{
    var argsDict = new Dictionary<string, string>();

    for (int i = 0; i < args.Length - 1; i += 2)
    {
        if (args[i].StartsWith("-") && !argsDict.ContainsKey(args[i]))
            argsDict[args[i]] = args[i + 1];
    }

    if (argsDict.TryGetValue("-i", out string? tempFile))
        templateFile = tempFile;
    if (argsDict.TryGetValue("-o", out string? resFile))
        resultFile = resFile;
    if (argsDict.TryGetValue("-dir", out string? dir))
        workdir = dir + @"\";
    if (argsDict.TryGetValue("-first", out string? first) && int.TryParse(first, out int fRow))
        firstRow = fRow;
    if (argsDict.TryGetValue("-last", out string? last) && int.TryParse(last, out int lRow))
        lastRow = lRow;
    if (argsDict.TryGetValue("-rowTag", out string? t))
        rowTag = t;
}
static void PrintUsage()
{
    Console.WriteLine("Usage: MergeXLSM.exe -i <template.xlsm> -o <result.xlsm> -dir <directory> -first <number> -last <number> -tag <text>");
    Console.WriteLine();
    Console.WriteLine("Arguments:");
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