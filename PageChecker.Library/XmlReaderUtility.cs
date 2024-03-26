using ClosedXML.Excel;
using PageChecker.Domain.Models;
using System.Text.RegularExpressions;

namespace PageChecker.Library;

public static class XmlReaderUtility
{
    public static DirectoryInfo RootDirectory { get; private set; }
    public static XLWorkbook MarketWorkbook { get; internal set; } = new XLWorkbook();
    public static XLWorkbook SalesWorkbook { get; internal set; } = new XLWorkbook();

    public static List<string> MarketHeaders { get; internal set; } = new();
    public static List<string> SalesHeaders { get; internal set; } = new();


    public static IXLWorksheet GetMarketWorksheet(int index)
    {
        if (index == 0)
        {
            return new XLWorkbook().Worksheet("blank");
        }

        return MarketWorkbook.Worksheets.Worksheet(index);
    }

    public static IXLWorksheet GetSalesWorksheet(int index)
    {
        if (index == 0)
        {
            return new XLWorkbook().Worksheet("blank");
        }

        return SalesWorkbook.Worksheets.Worksheet(index);
    }

    public static List<Market> GetMarketWorksheetData()
    {
        var marketData = new List<Market>();

        var worksheet = GetMarketWorksheet(1);

        MarketHeaders = GetWorksheetHeaders(worksheet, skip: 1, take: 1);

        var rows = worksheet.RangeUsed().RowsUsed().Skip(2); // Skip header row

        foreach (var row in rows)
        {
            if (row.CellsUsed().Count() < 6)
            {
                break;
            }

            var customer = row.Cell(MarketHeaders.IndexOf("Customer") + 1).Value.ToString();
            var size = row.Cell(MarketHeaders.IndexOf("Size") + 1).Value.ToString();
            var rep = row.Cell(MarketHeaders.IndexOf("Rep") + 1).Value.ToString();
            var categories = row.Cell(MarketHeaders.IndexOf("Categories") + 1).Value.ToString();
            var contractStatus = row.Cell(MarketHeaders.IndexOf("Contract Status") + 1).Value.ToString();
            var artwork = row.Cell(MarketHeaders.IndexOf("Artwork") + 1).Value.ToString();
            var notes = row.Cell(MarketHeaders.IndexOf("Notes") + 1).Value.ToString();
            var placement = row.Cell(MarketHeaders.IndexOf("Placement") + 1).Value.ToString();
            //var accountingNotes = row.Cell(MarketHeaders.IndexOf("ACCOUNTING USE ONLY") + 1).Value.ToString();


            marketData.Add(new Market
            {
                Customer = customer,
                Size = Convert.ToDouble(size),
                Rep = rep,
                Categories = categories,
                ContractStatus = contractStatus,
                Artwork = artwork,
                Notes = notes,
                Placement = placement,
                //AccountingNotes = accountingNotes,
            });
        }

        return marketData;
    }

    public static List<SalesRun> GetSalesWorksheetData()
    {
        var salesData = new List<SalesRun>();

        var worksheet = GetSalesWorksheet(1);
        
        SalesHeaders = GetWorksheetHeaders(worksheet, skip: 1, take: 1);

        var rows = worksheet.RangeUsed().RowsUsed().Skip(1); // Skip header row

        foreach (var row in rows)
        {
            var client = row.Cell(SalesHeaders.IndexOf("Client") + 1).Value.ToString();
            var product = row.Cell(SalesHeaders.IndexOf("Product") + 1).Value.ToString();
            var description = row.Cell(SalesHeaders.IndexOf("Description") + 1).Value.ToString();
            var salesRep = row.Cell(SalesHeaders.IndexOf("Sales Rep") + 1).Value.ToString();
            var net = row.Cell(SalesHeaders.IndexOf("Net") + 1).Value.ToString();
            var barter = row.Cell(SalesHeaders.IndexOf("Barter") + 1).Value.ToString();


            salesData.Add(new SalesRun
            {
                Client = client,
                Product = product,
                Description = description,
                SalesRep = salesRep,
                Net = net,
                Barter  = barter,
            });
        }

        return salesData;
    }

    public static List<Market> CompareSheets(List<Market> marketSheetData, List<SalesRun> salesSheetData)
    {
        foreach(var salesRow in salesSheetData)
        {
            foreach(var marketRow in marketSheetData)
            {
                var salesClient = Regex.Replace(salesRow.Client.ToLower(), @"\([a-zA-Z0-9 .-]+\)", "").Replace(" ", "");
                var marketClient = marketRow.Customer.ToLower().Replace(" ", "");

                var salesPageSize = GetPageSizeNumericValue(salesRow.Description);
                var marketPageSize = marketRow.Size;

                if (salesClient.Contains(marketClient) && 
                    salesPageSize == marketPageSize)
                {
                    marketRow.PassedCheck = true;
                }
            }
        }

        return marketSheetData;
    }

    public static double GetPageSizeNumericValue(string pageDescription)
    {
        if (string.IsNullOrEmpty(pageDescription))
        {
            return 0;
        }

        if (pageDescription.ToLower().Contains("full page"))
        {
            return 1;
        }

        if (pageDescription.ToLower().Contains("1/2 page"))
        {
            return 0.5;
        }

        return 0;
    }

    public static void ExportResults(string folderPath)
    {
        var resultsFilePath = Path.Combine(folderPath, "Results.xlsx");
        
        if (File.Exists(resultsFilePath))
        {
            File.Delete(resultsFilePath);
        }

        var marketSheetData = GetMarketWorksheetData();
        var salesSheetData = GetSalesWorksheetData();

        var checkedMarketData = CompareSheets(marketSheetData, salesSheetData);

        var workbook = new XLWorkbook();
        var ws = workbook.Worksheets.Add("Report");

        ws.Range(1, 1, 1, MarketHeaders.Take(6).Count()).Style.Fill.SetBackgroundColor(XLColor.LightBlue);

        var rowNum = 1;
        var colNum = 1;

        foreach(var header in MarketHeaders.Take(6))
        {
            ws.Cell(rowNum, colNum).Value = header;
            colNum++;
        }

        foreach (var item in checkedMarketData)
        {
            rowNum++;

            ws.Cell(rowNum, 1).Value = item.PassedCheck ? "X" : "";
            ws.Cell(rowNum, 2).Value = item.Customer;
            ws.Cell(rowNum, 3).Value = item.Size;
            ws.Cell(rowNum, 4).Value = item.Rep;
            ws.Cell(rowNum, 5).Value = item.Categories;
            ws.Cell(rowNum, 6).Value = item.ContractStatus;

            ws.Range(rowNum, 1, 1, 1).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);

            if (!item.PassedCheck) 
            {
                ws.Range(rowNum, 1, rowNum, MarketHeaders.Take(6).Count()).Style.Fill.SetBackgroundColor(XLColor.Yellow);
            }

            if (item.ContractStatus.ToLower() == "pay per lead")
            {
                ws.Range(rowNum, 1, rowNum, MarketHeaders.Take(6).Count()).Style.Fill.SetBackgroundColor(XLColor.LightPastelPurple);
            }
        }

        ws.Columns().AdjustToContents();

        workbook.SaveAs(resultsFilePath);
    }

    public static List<string> GetWorksheetHeaders(IXLWorksheet worksheet, int skip, int take)
    {
        var headers = new List<string>();
        var headerRow = worksheet.RangeUsed().RowsUsed().Skip(skip).Take(take);

        foreach (var row in headerRow)
        {
            var cells = row.Cells();

            foreach (var cell in cells)
            {
                headers.Add(cell.Value.ToString());
            }
        }

        return headers;
    }

    public static void SetRootDirectoryPath(string directoryPath)
    {
        RootDirectory = new DirectoryInfo(directoryPath);
    }

    public static List<string> GetRootDirectoryFolders()
    {
        var folders = RootDirectory.GetDirectories().Where(f => !f.Attributes.HasFlag(FileAttributes.Hidden)).Select(x => x.Name).ToList();
        return folders;
    }

    public static void OpenSalesSheet(string filename)
    {
        var filepath = Path.Combine(RootDirectory.FullName, filename);

        SalesWorkbook = new XLWorkbook(filepath);
    }

    public static void OpenMarketSheet(string filename)
    {
        var filepath = Path.Combine(RootDirectory.FullName, filename);

        MarketWorkbook = new XLWorkbook(filepath);
    }
}
