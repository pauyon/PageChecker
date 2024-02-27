﻿using ClosedXML.Excel;
using PageChecker.Domain.Models;
using System.Text.RegularExpressions;

namespace PageChecker.Library;

public static class XmlReaderUtility
{
    public static string MarketExcelPath { set; get; } = string.Empty;
    public static string SalesExcelPath { set; get; } = string.Empty;

    public static XLWorkbook MarketWorkbook { get; internal set; } = new XLWorkbook();
    public static XLWorkbook SalesWorkbook { get; internal set; } = new XLWorkbook();

    public static void OpenExcelFiles()
    {
        MarketWorkbook = new XLWorkbook(MarketExcelPath);
        SalesWorkbook = new XLWorkbook(SalesExcelPath);
    }

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

    public static void SetDefaultFilePaths()
    {
        if (MarketExcelPath == ".")
        {
            MarketExcelPath = $"C:\\Users\\pauli\\Desktop\\Test Market Sheet.xlsx";
        }

        if (SalesExcelPath == ".")
        {
            SalesExcelPath = $"C:\\Users\\pauli\\Desktop\\Test Sales Sheet.xlsx";
        }
    }

    public static List<Market> GetMarketWorksheetData()
    {
        var marketData = new List<Market>();

        var worksheet = GetMarketWorksheet(1);

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
            var contractStatus = row.Cell(4).Value.ToString();
            var notes = row.Cell(5).Value.ToString();
            var placement = row.Cell(6).Value.ToString();
            var accountingNotes = row.Cell(7).Value.ToString();


            marketData.Add(new Market
            {
                Customer = customer,
                Size = Convert.ToDouble(size),
                Rep = rep,
                ContractStatus = contractStatus,
                Notes = notes,
                Placement = placement,
                AccountingNotes = accountingNotes,
            });
        }

        return marketData;
    }

    public static List<SalesRun> GetSalesWorksheetData()
    {
        var salesData = new List<SalesRun>();

        var worksheet = GetSalesWorksheet(1);

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
                Net = Convert.ToDecimal(net),
                Barter  = Convert.ToInt32(barter),
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

    public static void ExportResults(string exportPath)
    {
        var marketSheetData = GetMarketWorksheetData();
        var salesSheetData = GetSalesWorksheetData();

        var passedMarketData = CompareSheets(marketSheetData, salesSheetData);

        var workbook = new XLWorkbook();
        var ws = workbook.Worksheets.Add("Report");

        ws.Range(1, 1, 1, 10).Style.Fill.SetBackgroundColor(XLColor.Blush);

        ws.Cell(1, 1).Value = "PassedCheck";
        ws.Cell(1, 2).Value = "Customer";
        ws.Cell(1, 3).Value = "Size";
        ws.Cell(1, 4).Value = "Rep";
        ws.Cell(1, 5).Value = "Categories";
        ws.Cell(1, 6).Value = "Contract Status";
        ws.Cell(1, 7).Value = "Artwork";
        ws.Cell(1, 8).Value = "Notes";
        ws.Cell(1, 9).Value = "Placement";
        ws.Cell(1, 10).Value = "Accounting Notes";

        var rowNum = 2;

        foreach (var item in passedMarketData)
        {
            ws.Range(rowNum, 1, 1, 1).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);

            ws.Cell(rowNum, 1).Value = item.PassedCheck ? "X" : "";
            ws.Cell(rowNum, 2).Value = item.Customer;
            ws.Cell(rowNum, 3).Value = item.Size;
            ws.Cell(rowNum, 4).Value = item.Rep;
            ws.Cell(rowNum, 5).Value = item.Categories;
            ws.Cell(rowNum, 6).Value = item.ContractStatus;
            ws.Cell(rowNum, 7).Value = item.Artwork;
            ws.Cell(rowNum, 8).Value = item.Notes;
            ws.Cell(rowNum, 9).Value = item.Placement;
            ws.Cell(rowNum, 10).Value = item.AccountingNotes;

            rowNum++;
        }

        ws.Columns().AdjustToContents();

        workbook.SaveAs(exportPath);
    }
}