﻿using ClosedXML.Excel;
using PageChecker.Domain.Models;
using System.Text.RegularExpressions;

namespace PageChecker.Library;

public class XmlReaderUtility : ReaderBase, IReaderUtility
{
    public XLWorkbook MarketWorkbook { get; internal set; } = new XLWorkbook();
    public XLWorkbook SalesWorkbook { get; internal set; } = new XLWorkbook();

    public IXLWorksheet OpenMarketSheet(int index)
    {
        if (index == 0)
        {
            return new XLWorkbook().Worksheet("blank");
        }

        return MarketWorkbook.Worksheets.Worksheet(index);
    }

    public IXLWorksheet OpenSalesWorksheet(int index)
    {
        if (index == 0)
        {
            return new XLWorkbook().Worksheet("blank");
        }

        return SalesWorkbook.Worksheets.Worksheet(index);
    }

    public List<Market> GetMarketSheetData()
    {
        var marketData = new List<Market>();

        var worksheet = OpenMarketSheet(1);

        var rows = worksheet.RangeUsed().RowsUsed().Skip(2); // Skip header row

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
        var salesData = new List<SalesRun>();

        var worksheet = OpenSalesWorksheet(1);

        var rows = worksheet.RangeUsed().RowsUsed().Skip(1); // Skip header row

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

    public List<Market> CompareSheetsData(List<Market> marketSheetData, List<SalesRun> salesSheetData)
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

    public void OpenSalesSheet(string filename)
    {
        var filepath = Path.Combine(RootDirectory.FullName, filename);

        SalesWorkbook = new XLWorkbook(filepath);
    }

    public void OpenMarketSheet(string filename)
    {
        var filepath = Path.Combine(RootDirectory.FullName, filename);

        MarketWorkbook = new XLWorkbook(filepath);
    }

    public void ExportResults(string folderPath)
    {
        var resultsFilePath = Path.Combine(folderPath, "Results.xlsx");

        if (File.Exists(resultsFilePath))
        {
            File.Delete(resultsFilePath);
        }

        var marketSheetData = GetMarketSheetData();
        var salesSheetData = GetSalesSheetData();

        var checkedMarketData = CompareSheetsData(marketSheetData, salesSheetData);

        var workbook = new XLWorkbook();
        var ws = workbook.Worksheets.Add("Report");

        ws.Range(1, 1, 1, 10).Style.Fill.SetBackgroundColor(XLColor.LightBlue);

        var rowNum = 1;

        ws.Cell(rowNum, 1).Value = "PassedCheck";
        ws.Cell(rowNum, 2).Value = "Customer";
        ws.Cell(rowNum, 3).Value = "Size";
        ws.Cell(rowNum, 4).Value = "Rep";
        ws.Cell(rowNum, 5).Value = "Categories";
        ws.Cell(rowNum, 6).Value = "Contract Status";
        /*ws.Cell(rowNum, 7).Value = "Artwork";
        ws.Cell(rowNum, 8).Value = "Notes";
        ws.Cell(rowNum, 9).Value = "Placement";
        ws.Cell(rowNum, 10).Value = "Accounting Notes";*/

        foreach (var item in checkedMarketData)
        {
            rowNum++;

            ws.Cell(rowNum, 1).Value = item.PassedCheck ? "X" : "";
            ws.Cell(rowNum, 2).Value = item.Customer;
            ws.Cell(rowNum, 3).Value = item.Size;
            ws.Cell(rowNum, 4).Value = item.Rep;
            ws.Cell(rowNum, 5).Value = item.Categories;
            ws.Cell(rowNum, 6).Value = item.ContractStatus;
            /*ws.Cell(rowNum, 7).Value = item.Artwork;
            ws.Cell(rowNum, 8).Value = item.Notes;
            ws.Cell(rowNum, 9).Value = item.Placement;
            ws.Cell(rowNum, 10).Value = item.AccountingNotes;*/

            ws.Range(rowNum, 1, 1, 1).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);

            if (!item.PassedCheck)
            {
                ws.Range(rowNum, 1, rowNum, 10).Style.Fill.SetBackgroundColor(XLColor.Yellow);
            }

            if (item.ContractStatus.ToLower() == "pay per lead")
            {
                ws.Range(rowNum, 1, rowNum, 10).Style.Fill.SetBackgroundColor(XLColor.LightPastelPurple);
            }
        }

        ws.Columns().AdjustToContents();

        workbook.SaveAs(resultsFilePath);
    }
}
