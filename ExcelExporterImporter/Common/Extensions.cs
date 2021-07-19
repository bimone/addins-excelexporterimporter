using System.IO;

//using OfficeOpenXml.Table;

namespace ExcelExporterImporter.Common
{
    public static class Extensions
    {
        /// <summary>
        ///     Valid if the file is locked
        /// </summary>
        /// <param name="file"></param>
        /// <returns>True or False</returns>
        public static bool IsFileLocked(this FileInfo file)
        {
            FileStream stream = null;
            try
            {
                stream = file.Open(FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            }
            catch (IOException)
            {
                return true;
            }
            finally
            {
                if (stream != null)
                    stream.Close();
            }

            //file is not locked
            return false;
        }
    }
}