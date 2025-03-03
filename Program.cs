using ClosedXML.Excel;
using System.Text;

Console.OutputEncoding = Encoding.UTF8; // set encoding to utf-8 to support cyrillic symbols

string workdir = Path.Combine(Environment.CurrentDirectory, "XLSM");
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

Console.WriteLine("[ Merge XLSM ]");

ParseArgs(args, ref templateFile, ref resultFile, ref workdir, ref firstRow, ref lastRow, ref rowTag);

string templatePath = Path.Combine(workdir, templateFile);
string resultPath = Path.Combine(workdir, resultFile);

Console.WriteLine();
Console.WriteLine("Arguments:");
Console.WriteLine($"Template file: {templatePath}");
Console.WriteLine($"Result file: {resultPath}");
Console.WriteLine($"Working directory: {workdir}");
Console.WriteLine($"First row: {firstRow}");
Console.WriteLine($"Last row: {lastRow}");
Console.WriteLine($"rowTag: {rowTag}");
Console.WriteLine();
Console.WriteLine("Cleaning old results...");

if (File.Exists(resultPath)) // delete result file if exists
{
    try
    {
        File.Delete(resultPath);
    }
    catch (Exception e)
    {
        Console.WriteLine($"[Error] Cannot delete old result file: {e.Message}");
        return;
    }
}

if (!File.Exists(templatePath)) // check if template file exists
{
    Console.WriteLine($"Template file {templateFile} is not found");
    return;
}

Console.WriteLine($"Creating a new worksheet from {templateFile}");

var resultBook = new XLWorkbook(templatePath); // create new workbook from template that will contain merged results
var ws = resultBook.Worksheet(1); // open first worksheet
var curRow = firstRow; // a current row to copy data

Console.WriteLine("Getting files in a working dir...");

foreach (var file in Directory.GetFiles(workdir, "*.xlsm")) // get xlsm files in a working dir
{
    if (Path.GetFileName(file).Equals(templateFile, StringComparison.OrdinalIgnoreCase)) // skip template file
        continue;

    Console.WriteLine($"XLSM file found: {file}");

    try
    {
        using var srcbook = new XLWorkbook(file); // open source workbook
        var srcsheet = srcbook.Worksheet(1); // open first worksheet in a source workbook
        int cpCnt = 0; // initialize a row counter for source workbook

        for (int i = firstRow; i <= lastRow; ++i)
        {
            if (srcsheet.Cell(i, 1).Value.ToString() == rowTag) // if first cell in a row contains rowTag
            {
                CopyRowWithoutFormulas(srcsheet.Row(i), ws.Row(curRow)); // copy that row
                curRow++; // switch to next row
                cpCnt++; // count copied rows (for only last source workbook to notify user)
            }
        }

        Console.WriteLine($"Rows copied: {cpCnt}");
    }
    catch (Exception e)
    {
        Console.WriteLine($"[Error] Unable to process file: {e.Message}");
    }
}

Console.WriteLine();
Console.WriteLine($"Total merged rows: {curRow - firstRow}.\nSaving results...");

try 
{ 
    resultBook.SaveAs(resultPath);
    Console.WriteLine($"Saved to {resultPath}");
}
catch (Exception e)
{
    Console.WriteLine($"[Error] Failed to save result file: {e.Message}");
}

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