using ClosedXML.Excel;
using PageChecker.Domain.Models;

namespace PageChecker.Library;

public class ExcelReaderUtility : FileReaderBase, IFileReaderUtility
{
    public XLWorkbook MarketWorkbook { get; internal set; } = new XLWorkbook();
    public XLWorkbook SalesRunWorkbook { get; internal set; } = new XLWorkbook();

    public List<MarketClient> GetMarketClientSheetData()
    {
        var marketData = new List<MarketClient>();
        var worksheet = MarketWorkbook.Worksheets.Worksheet(1);

        MarketClientSheetHeaders = GetWorksheetHeaders(worksheet, skip: 1, take: 1);
        var rows = worksheet.RangeUsed().RowsUsed().Skip(2); // Skip header row

        foreach (var row in rows)
        {
            if (row.CellsUsed().Count() < 6)
            {
                break;
            }

            var customer = row.Cell(MarketClientSheetHeaders.IndexOf("Customer") + 1).Value.ToString();
            var size = row.Cell(MarketClientSheetHeaders.IndexOf("Size") + 1).Value.ToString();
            var rep = row.Cell(MarketClientSheetHeaders.IndexOf("Rep") + 1).Value.ToString();
            var categories = row.Cell(MarketClientSheetHeaders.IndexOf("Categories") + 1).Value.ToString();
            var contractStatus = row.Cell(MarketClientSheetHeaders.IndexOf("Contract Status") + 1).Value.ToString();
            var artwork = row.Cell(MarketClientSheetHeaders.IndexOf("Artwork") + 1).Value.ToString();
            var notes = row.Cell(MarketClientSheetHeaders.IndexOf("Notes") + 1).Value.ToString();
            var placement = row.Cell(MarketClientSheetHeaders.IndexOf("Placement") + 1).Value.ToString();

            marketData.Add(new MarketClient
            {
                Customer = customer,
                Size = Convert.ToDouble(size),
                Rep = rep,
                Categories = categories,
                ContractStatus = contractStatus,
                Artwork = artwork,
                Notes = notes,
                Placement = placement,
            });
        }

        return marketData;
    }

    public List<SalesRun> GetSalesSheetData()
    {
        var salesData = new List<SalesRun>();
        var worksheet = SalesRunWorkbook.Worksheets.Worksheet(1);

        SalesSheetHeaders = GetWorksheetHeaders(worksheet, skip: 1, take: 1);

        var rows = worksheet.RangeUsed().RowsUsed().Skip(1); // Skip header row

        foreach (var row in rows)
        {
            var client = row.Cell(SalesSheetHeaders.IndexOf("Client") + 1).Value.ToString();
            var product = row.Cell(SalesSheetHeaders.IndexOf("Product") + 1).Value.ToString();
            var description = row.Cell(SalesSheetHeaders.IndexOf("Description") + 1).Value.ToString();
            var salesRep = row.Cell(SalesSheetHeaders.IndexOf("Sales Rep") + 1).Value.ToString();
            var net = row.Cell(SalesSheetHeaders.IndexOf("Net") + 1).Value.ToString();
            var barter = row.Cell(SalesSheetHeaders.IndexOf("Barter") + 1).Value.ToString();


            salesData.Add(new SalesRun
            {
                Client = client,
                Product = product,
                Description = description,
                SalesRep = salesRep,
                Net = net,
                Barter = barter,
            });
        }

        return salesData;
    }

    public void AnalyzeAndExportResults(string folderPath, string marketClientSheetPath, string salesRunPath)
    {
        MarketWorkbook = new XLWorkbook(marketClientSheetPath);
        SalesRunWorkbook = new XLWorkbook(salesRunPath);

        var resultsExportPath = Path.Combine(folderPath, "Results.xlsx");

        RemoveExistingResultsExcel(resultsExportPath);

        var marketClientSheetData = GetMarketClientSheetData();
        var salesRunSheetData = GetSalesSheetData();
        var checkedMarketData = CompareSheetsData(marketClientSheetData, salesRunSheetData);

        GenerateResultsExcel(checkedMarketData, resultsExportPath);
    }

    private List<string> GetWorksheetHeaders(IXLWorksheet worksheet, int skip, int take)
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
}
