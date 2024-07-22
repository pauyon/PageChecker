using ClosedXML.Excel;
using Microsoft.Extensions.Logging;
using PageChecker.Domain.Models;
using System.Text.RegularExpressions;

namespace PageChecker.Library
{
    public class FileReaderBase : IFileReaderBase
    {
        private readonly ILogger _logger;

        public FileReaderBase(ILogger logger)
        {
            _logger = logger;
        }

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

            if (pageDescription.ToLower().Contains("two page"))
            {
                return 2;
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
            _logger.LogInformation("Getting workspace folders...");
            var folders = WorkspaceDirectory.GetDirectories().Select(x => x.Name).ToList();
            return folders;
        }

        /// <summary>
        /// Deletes existing results excel file.
        /// </summary>
        /// <param name="resultsSheetFilePath">Path of results excel file.</param>
        public void RemoveExistingResultsExcel(string resultsSheetFilePath)
        {
            _logger.LogInformation("Checking for already existing results sheet");
            if (File.Exists(resultsSheetFilePath))
            {
                _logger.LogInformation("Existing results sheet found. Deleteing...");
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
            try
            {
                _logger.LogInformation("Comparing sales and market sheet data...");
                foreach (var salesRow in salesRunSheetData)
                {
                    var salesRunClientName = Regex.Replace(salesRow.ClientName.ToLower(), @"\([a-zA-Z0-9 .-]+\)", "").Replace(" ", "").Trim();

                    foreach (var marketRow in marketClientSheetData)
                    {
                        var marketClientName = marketRow.AccurateCustomerName.ToLower().Replace(" ", "").Trim();

                        var salesClientPageSize = GetPageSizeNumericValue(salesRow.Description);
                        var marketClientPageSize = marketRow.Size;

                        if (salesRunClientName.Contains(marketClientName) &&
                            salesClientPageSize == marketClientPageSize)
                        {
                            marketRow.PassedCheck = true;
                        }
                    }
                }

                return marketClientSheetData;
            }
            catch(Exception ex)
            {
                _logger.LogError(ex, "An error ocurred comparing sales and market sheet data.");
                throw new Exception("An error ocurred comparing sales and market sheet data.");
            }
        }

        /// <summary>
        /// Export processed data to excel file.
        /// </summary>
        /// <param name="checkedMarketData">List of market data that has been processed.</param>
        /// <param name="resultsExportPath">Export path of results excel file.</param>
        public void GenerateResultsExcel(List<MarketClient> checkedMarketData, string resultsExportPath)
        {
            try
            {
                _logger.LogInformation("Generating results sheet...");

                var workbook = new XLWorkbook();
                var ws = workbook.Worksheets.Add("Results");

                var rowLength = 7;

                ws.Range(1, 1, 1, rowLength).Style.Fill.SetBackgroundColor(XLColor.LightBlue);

                var rowNum = 1;

                ws.Cell(rowNum, 1).Value = "PassedCheck";
                ws.Cell(rowNum, 2).Value = "Customer";
                ws.Cell(rowNum, 3).Value = "Accounting Customer Name";
                ws.Cell(rowNum, 4).Value = "Size";
                ws.Cell(rowNum, 5).Value = "Rep";
                ws.Cell(rowNum, 6).Value = "Categories";
                ws.Cell(rowNum, 7).Value = "Contract Status";

                foreach (var item in checkedMarketData)
                {
                    rowNum++;

                    ws.Cell(rowNum, 1).Value = item.PassedCheck ? "X" : "";
                    ws.Cell(rowNum, 2).Value = item.CustomerName;
                    ws.Cell(rowNum, 3).Value = item.AccountingCustomerName;
                    ws.Cell(rowNum, 4).Value = item.Size;
                    ws.Cell(rowNum, 5).Value = item.Rep;
                    ws.Cell(rowNum, 6).Value = item.Categories;
                    ws.Cell(rowNum, 7).Value = item.ContractStatus;

                    ws.Range(rowNum, 1, 1, 1).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);

                    if (!item.PassedCheck)
                    {
                        ws.Range(rowNum, 1, rowNum, rowLength).Style.Fill.SetBackgroundColor(XLColor.Yellow);
                    }

                    if (item.ContractStatus.ToLower() == "pay per lead")
                    {
                        ws.Range(rowNum, 1, rowNum, rowLength).Style.Fill.SetBackgroundColor(XLColor.LightPastelPurple);
                    }
                }

                ws.Columns().AdjustToContents();

                workbook.SaveAs(resultsExportPath);
            }
            catch(Exception ex)
            {
                _logger.LogError(ex, "An error ocurred generating results sheet.");
                throw new Exception("An error ocurred generating resutls sheet.");
            }
        }

        /// <summary>
        /// Renames the accounting use only column for ease of use.
        /// </summary>
        /// <param name="headers"></param>
        /// <returns></returns>
        public List<string> RenameAccountingColumn(List<string> headers)
        {
            var accountingUseOnlyColumn = headers.First(x => x.ToLower().Contains("accounting use only"));
            var indexOfAccountingUseOnlyColumn = headers.IndexOf(accountingUseOnlyColumn);

            _logger.LogInformation($"Renaming column '{indexOfAccountingUseOnlyColumn}' to AccountingCustomerName");
            headers[indexOfAccountingUseOnlyColumn] = "AccountingCustomerName";

            return headers;
        }
    }
}
