using ClosedXML.Excel;
using PageChecker.Domain.Models;
using System.Text.RegularExpressions;

namespace PageChecker.Library
{
    public class ReaderBase
    {
        public DirectoryInfo RootDirectory { get; private set; } = new DirectoryInfo(".");

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

        public void SetRootDirectoryPath(string directoryPath)
        {
            RootDirectory = new DirectoryInfo(directoryPath);
        }

        public List<string> GetRootDirectoryFolders()
        {
            var folders = RootDirectory.GetDirectories().Select(x => x.Name).ToList();
            return folders;
        }

        public void RemoveExistingResultsXml(string resultsSheetFilePath)
        {
            if (File.Exists(resultsSheetFilePath))
            {
                File.Delete(resultsSheetFilePath);
            }
        }

        public List<Market> CompareSheetsData(List<Market> marketSheetData, List<SalesRun> salesSheetData)
        {
            foreach (var salesRow in salesSheetData)
            {
                foreach (var marketRow in marketSheetData)
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

        public void GenerateResultsExcel(List<Market> checkedMarketData, string resultsExportPath)
        {
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

            workbook.SaveAs(resultsExportPath);
        }
    }
}
