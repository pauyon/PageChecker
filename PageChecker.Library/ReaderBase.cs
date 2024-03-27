using ClosedXML.Excel;
using PageChecker.Domain.Models;
using System.Text.RegularExpressions;

namespace PageChecker.Library
{
    public class ReaderBase : IReaderBase
    {
        public DirectoryInfo WorkspaceDirectory { get; set; } = new DirectoryInfo(".");
        public List<string> MarketClientSheetHeaders { get; set; } = new();
        public List<string> SalesSheetHeaders { get; set; } = new();

        /// <summary>
        /// Gets the numeric page size value of a string page size.
        /// </summary>
        /// <param name="pageDescription">String value of page size.</param>
        /// <returns>Numeric value of page size.</returns>
        public double GetPageSizeNumericValue(string pageDescription)
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

        /// <summary>
        /// Sets path of workspace directory.
        /// </summary>
        /// <param name="directoryPath">Full path of workspace directory.</param>
        public void SetWorkspaceDirectoryPath(string directoryPath)
        {
            WorkspaceDirectory = new DirectoryInfo(directoryPath);
        }

        /// <summary>
        /// Gets all files within the workspace directory.
        /// </summary>
        /// <returns></returns>
        public List<string> GetWorkspaceFolders()
        {
            var folders = WorkspaceDirectory.GetDirectories().Select(x => x.Name).ToList();
            return folders;
        }

        /// <summary>
        /// Deletes existing results excel file.
        /// </summary>
        /// <param name="resultsSheetFilePath">Path of results excel file.</param>
        public void RemoveExistingResultsExcel(string resultsSheetFilePath)
        {
            if (File.Exists(resultsSheetFilePath))
            {
                File.Delete(resultsSheetFilePath);
            }
        }

        /// <summary>
        /// Compare market sheet and sales run sheet data.
        /// </summary>
        /// <param name="marketClientSheetData">List of data within market sheet.</param>
        /// <param name="salesRunSheetData">List of data within sales run sheet.</param>
        /// <returns>List of market data after comparison.</returns>
        public List<MarketClient> CompareSheetsData(List<MarketClient> marketClientSheetData, List<SalesRun> salesRunSheetData)
        {
            foreach (var salesRow in salesRunSheetData)
            {
                foreach (var marketRow in marketClientSheetData)
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

            return marketClientSheetData;
        }

        /// <summary>
        /// Export processed data to excel file.
        /// </summary>
        /// <param name="checkedMarketData">List of market data that has been processed.</param>
        /// <param name="resultsExportPath">Export path of results excel file.</param>
        public void GenerateResultsExcel(List<MarketClient> checkedMarketData, string resultsExportPath)
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
