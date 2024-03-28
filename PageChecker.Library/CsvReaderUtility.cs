using Microsoft.VisualBasic.FileIO;
using PageChecker.Domain.Models;

namespace PageChecker.Library
{
    public class CsvReaderUtility : FileReaderBase, IFileReaderUtility
    {
        public FileInfo MarketClientSheet { get; set; } = new FileInfo("./");
        public FileInfo SalesRunSheet { get; set; } = new FileInfo("./");

        public void AnalyzeAndExportResults(string folderPath, string marketClientSheetFilename, string salesSheetFilename)
        {
            MarketClientSheet = new FileInfo(Path.Combine(folderPath, marketClientSheetFilename));
            SalesRunSheet = new FileInfo(Path.Combine(folderPath, salesSheetFilename));

            var marketFolder = folderPath.Split("\\").Last();
            var resultsExportPath = Path.Combine(folderPath, $"{marketFolder}-Results.xlsx");

            RemoveExistingResultsExcel(resultsExportPath);

            var marketClientSheetData = GetMarketClientSheetData();
            var salesRunSheetData = GetSalesSheetData();
            var checkedMarketData = CompareSheetsData(marketClientSheetData, salesRunSheetData);

            GenerateResultsExcel(checkedMarketData, resultsExportPath);
        }

        public List<MarketClient> GetMarketClientSheetData()
        {
            var marketData = new List<MarketClient>();
            var data = ReadCsvFileData(MarketClientSheet.FullName, false);

            MarketClientSheetHeaders = data.Skip(1).First().Split(";").ToList();
            MarketClientSheetHeaders = RenameAccountingColumn(MarketClientSheetHeaders);

            foreach (var line in data.Skip(2).ToList())
            {
                var columns = line.Split(';').ToList();

                var customer = columns[MarketClientSheetHeaders.IndexOf("Customer")];
                var size = columns[MarketClientSheetHeaders.IndexOf("Size")];
                var rep = columns[MarketClientSheetHeaders.IndexOf("Rep")];
                var categories = columns[MarketClientSheetHeaders.IndexOf("Categories")];
                var contractStatus = columns[MarketClientSheetHeaders.IndexOf("Contract Status")];
                var artwork = columns[MarketClientSheetHeaders.IndexOf("Artwork")];
                var notes = columns[MarketClientSheetHeaders.IndexOf("Notes")];
                var placement = columns[MarketClientSheetHeaders.IndexOf("Placement")];
                var accountingCustomerName = columns[MarketClientSheetHeaders.IndexOf("AccountingCustomerName")];

                marketData.Add(new MarketClient
                {
                    CustomerName = customer,
                    Size = Convert.ToDouble(size),
                    Rep = rep,
                    Categories = categories,
                    ContractStatus = contractStatus,
                    Artwork = artwork,
                    Notes = notes,
                    Placement = placement,
                    AccountingCustomerName = accountingCustomerName
                });
            }

            return marketData;
        }

        public List<string> ReadCsvFileData(string filePath, bool readToEndOfFile = true)
        {
            var parser = new TextFieldParser(new StringReader(File.ReadAllText(filePath)))
            {
                HasFieldsEnclosedInQuotes = true,
                Delimiters = new string[] { "," },
                TrimWhiteSpace = true
            };

            var allCsvLines = new List<string>();

            // Reads all fields on the current line of the CSV file and returns as a string array
            // Joins each field together with new delimiter "|"
            while (!parser.EndOfData)
            {
                allCsvLines.Add(string.Join(";", parser.ReadFields()));
            }

            if (readToEndOfFile)
            {
                return allCsvLines;
            }
            else
            {
                var trimmedData = new List<string>();

                foreach( var line in allCsvLines)
                {
                    var columnData = line.Split(";").ToList();

                    // Exit on first empty line detected
                    if (columnData.All(x => string.IsNullOrEmpty(x)))
                    {
                        return trimmedData;
                    }

                    trimmedData.Add(line);
                }

                return trimmedData;
            }
        }

        public List<SalesRun> GetSalesSheetData()
        {
            var salesRunData = new List<SalesRun>();
            var data = ReadCsvFileData(SalesRunSheet.FullName, false);

            SalesSheetHeaders = data.First().Split(";").ToList();

            foreach (var line in data.Skip(1).ToList())
            {
                var columns = line.Split(';').ToList();

                var client = columns[SalesSheetHeaders.IndexOf("Client")];
                var product = columns[SalesSheetHeaders.IndexOf("Product")];
                var description = columns[SalesSheetHeaders.IndexOf("Description")];
                var salesRep = columns[SalesSheetHeaders.IndexOf("Sales Rep")];
                var net = columns[SalesSheetHeaders.IndexOf("Net")];
                var barter = columns[SalesSheetHeaders.IndexOf("Barter")];


                salesRunData.Add(new SalesRun
                {
                    ClientName = client,
                    Product = product,
                    Description = description,
                    SalesRep = salesRep,
                    Net = net,
                    Barter = barter,
                });
            }

            return salesRunData;
        }
    }
}
