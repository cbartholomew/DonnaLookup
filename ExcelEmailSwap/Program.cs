using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelEmailSwap
{
    class Program
    {
        static void Main(string[] args)
        {
            string sourceDocument = "C:\\Users\\Christopher\\Documents\\Temp\\ACTIVE EMAIL LIST.xlsx";
            string sourceReplacement = "C:\\Users\\Christopher\\Documents\\Temp\\OutputFINAL.xlsx";
            List<string[]> sourceReplacementColumns
            = uwExcel.createFromExcel(sourceReplacement, "A", "D", "Sheet1");

            // 37 - 44
            foreach (string[] sourceReplacementRow in sourceReplacementColumns)
            {
                var sourceReplacementEmail = sourceReplacementRow[1];

                uwExcel.replaceFromExcel(sourceDocument, "A", "CC", sourceReplacementEmail,"","Sheet1");                
            }
        }
    }
}
