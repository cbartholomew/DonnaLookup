using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelNameSwap
{
    class Program
    {

        static void Main(string[] args)
        {
            uwExcel.createFromExcel(Configuration.file_path, "A", "Z", "CSE Alumni");
            uwExcel.createFromExcel(Configuration.file_path, "A", "Z", "CSE Donors");
        }
    }
}
