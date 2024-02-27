using PageChecker.Library;
using Spectre.Console;

AnsiConsole.Write(
    new FigletText("Page Checker")
        .Centered()
        .Color(Color.Green));

var rootPath = AnsiConsole.Ask<string>("Folder path that has [green]market spreadsheet[/] and [green]sales run[/]:");
XmlReaderUtility.SetDirectoryPath(rootPath);

var files = XmlReaderUtility.GetDirectoryFiles();

var salesSheet = AnsiConsole.Prompt(
    new SelectionPrompt<string>()
        .Title("Which file is the [green]sales run[/]?")
        .PageSize(10)
        .MoreChoicesText("[grey](Move up and down to reveal more files)[/]")
        .AddChoices(files));

files.Remove(salesSheet);
var marketSheet = files[0];

XmlReaderUtility.OpenSalesSheet(salesSheet);
XmlReaderUtility.OpenMarketSheet(marketSheet);

XmlReaderUtility.ExportResults();

Console.WriteLine("All done!");
