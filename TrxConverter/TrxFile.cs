using System.IO;

namespace TrxConverter
{
    public class TrxFile
    {
        private FileInfo TrxFileInfo { get; set; }

        public TrxFile(FileInfo trxInfo)
        {
            TrxFileInfo = trxInfo;
        }

        public TrxFile(string trxInfo)
        {
            TrxFileInfo = new FileInfo(trxInfo);
        }

        public string GetFilePath()
        {
            return TrxFileInfo.FullName;
        }

    }
}