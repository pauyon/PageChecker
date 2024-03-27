using ClosedXML.Excel;
using PageChecker.Domain.Models;
using System.Text.RegularExpressions;

namespace PageChecker.Library;

public class XmlReaderUtility : ReaderBase, IReaderUtility
{
    public XLWorkbook MarketWorkbook { get; internal set; } = new XLWorkbook();
    public XLWorkbook SalesWorkbook { get; internal set; } = new XLWorkbook();

    public List<Market> GetMarketSheetData()
    {
        var worksheet = MarketWorkbook.Worksheets.Worksheet(1);
        var rows = worksheet.RangeUsed().RowsUsed().Skip(2); // Skip header row

        var marketData = new List<Market>();

        foreach (var row in rows)
        {
            if (row.CellsUsed().Count() < 6)
            {
                break;
            }

            var customer = row.Cell(1).Value.ToString();
            var size = row.Cell(2).Value.ToString();
            var rep = row.Cell(3).Value.ToString();
            var categories = row.Cell(4).Value.ToString();
            var contractStatus = row.Cell(5).Value.ToString();
            var artwork = row.Cell(6).Value.ToString();
            var notes = row.Cell(7).Value.ToString();
            var placement = row.Cell(8).Value.ToString();
            var accountingNotes = row.Cell(9).Value.ToString();


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
                AccountingNotes = accountingNotes,
            });
        }

        return marketData;
    }

    public List<SalesRun> GetSalesSheetData()
    {
        var worksheet = SalesWorkbook.Worksheets.Worksheet(1);
        var rows = worksheet.RangeUsed().RowsUsed().Skip(1); // Skip header row

        var salesData = new List<SalesRun>();
        
        foreach (var row in rows)
        {
            var client = row.Cell(1).Value.ToString();
            var product = row.Cell(2).Value.ToString();
            var description = row.Cell(3).Value.ToString();
            var salesRep = row.Cell(4).Value.ToString();
            var net = row.Cell(5).Value.ToString();
            var barter = row.Cell(6).Value.ToString();


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

    public void ExportResults(string folderPath, string marketSheetFilename, string salesSheetFilename)
    {
        MarketWorkbook = new XLWorkbook(Path.Combine(folderPath, marketSheetFilename));
        SalesWorkbook = new XLWorkbook(Path.Combine(folderPath, salesSheetFilename));

        var resultsExportPath = Path.Combine(folderPath, "Results.xlsx");

        RemoveExistingResultsXml(resultsExportPath);

        var marketSheetData = GetMarketSheetData();
        var salesSheetData = GetSalesSheetData();
        var checkedMarketData = CompareSheetsData(marketSheetData, salesSheetData);

        GenerateResultsExcel(checkedMarketData, resultsExportPath);
    }
}
