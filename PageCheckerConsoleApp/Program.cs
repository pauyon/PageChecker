using PageChecker.Library;
using Spectre.Console;

AnsiConsole.Write(
    new FigletText("Page Checker")
        .Centered()
        .Color(Color.Green));

XmlReaderUtility.MarketExcelPath = AnsiConsole.Ask<string>("File path of [green]market spreadsheet[/]:");
XmlReaderUtility.SalesExcelPath = AnsiConsole.Ask<string>("File path of [green]sales run[/] to compare against:");

Console.WriteLine("Loading files...");

XmlReaderUtility.SetDefaultFilePaths();
XmlReaderUtility.OpenExcelFiles();

var exportPath = "report.xlsx";
XmlReaderUtility.ExportResults(exportPath);

Console.WriteLine("All done!");
