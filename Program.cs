using MergeXLSM;
using ClosedXML.Excel;
using System.Text;

Console.OutputEncoding = Encoding.UTF8; // set encoding to utf-8 to support cyrillic symbols

string workDir = Path.Combine(Environment.CurrentDirectory, "XLSM");
string templateFile = "Template.xlsm";
string resultFile = "Result.xlsm";
int firstRow = 6;
int lastRow = 1000;
string rowTag = "Fact";

if (args.Contains("-h") || args.Contains("--help"))
{
    MergeXLSMHelper.PrintUsage();
    return;
}

Console.WriteLine($"[ MergeXLSM - A tool to merge XLSX/XLSM files ]");
Console.WriteLine("Tip: use -h or --help option for help");

MergeXLSMHelper.ParseArgs(args, ref templateFile, ref resultFile, ref workDir, ref firstRow, ref lastRow, ref rowTag);

string templatePath = Path.Combine(workDir, templateFile);
string resultPath = Path.Combine(workDir, resultFile);

MergeXLSMHelper.PrintArgs(ref templatePath, ref resultPath, ref workDir, ref firstRow, ref lastRow, ref rowTag);

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

foreach (var file in Directory.GetFiles(workDir, "*.xls?")) // get xlsx/xlsm files in a working dir
{
    if (Path.GetFileName(file).Equals(templateFile, StringComparison.OrdinalIgnoreCase)) // skip template file
        continue;

    Console.WriteLine($"File found: {file}");

    try
    {
        using var srcbook = new XLWorkbook(file); // open source workbook
        var srcsheet = srcbook.Worksheet(1); // open first worksheet in a source workbook
        int cpCnt = 0; // initialize a row counter for source workbook

        for (int i = firstRow; i <= lastRow; ++i)
        {
            if (srcsheet.Cell(i, 1).Value.ToString() == rowTag) // if first cell in a row contains rowTag
            {
                MergeXLSMHelper.CopyRowWithoutFormulas(srcsheet.Row(i), ws.Row(curRow)); // copy that row
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