using System.Collections;
using System.IO;
using System.Linq;

namespace DataAnalysis
{
    public class DataSaveWrite
    {
        /*
         * This will save the data ArrayList to a file 
         */
        public static void WriteDataToFile(ArrayList dataArr, string fileName)
        {
            string currentDirectory = Directory.GetCurrentDirectory() + "/Data/something.txt";  // Save it in the data Folder where this code is
            File.WriteAllLines(currentDirectory, dataArr.Cast<string>().ToArray());
        }
    }
}