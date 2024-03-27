using ClosedXML.Excel;

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
    }
}
