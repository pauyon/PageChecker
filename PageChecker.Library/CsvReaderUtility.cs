using PageChecker.Domain.Models;

namespace PageChecker.Library
{
    public class CsvReaderUtility : ReaderBase, IReaderUtility
    {
        public void AnalyzeAndExportResults(string folderPath, string marketSheetFilename, string salesSheetFilename)
        {
            throw new NotImplementedException();
        }

        public List<Market> GetMarketSheetData()
        {
            throw new NotImplementedException();
        }

        public List<SalesRun> GetSalesSheetData()
        {
            throw new NotImplementedException();
        }
    }
}
