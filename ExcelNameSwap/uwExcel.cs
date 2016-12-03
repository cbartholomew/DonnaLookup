using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.IO;


namespace ExcelNameSwap
{
    public class uwExcel
    {
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public static Workbook createWorkbook()
        {
            Application xlApp = new Application();
            xlApp.UserControl = false;
            Workbook xlWorkBook = xlApp.Workbooks.Add();

            return xlWorkBook;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public static Workbook getWorkbook(string path)
        {
            Workbook xwWorkBook = null;
            try
            {
                if (String.IsNullOrEmpty(path))
                {
                    throw new ArgumentNullException("No XLS Path Found");
                }
                Application xwApp = new Application();

                // init new appliction
                xwWorkBook = xwApp.Workbooks.Open(path);
            }
            catch (Exception e)
            {
                throw e;
            }

            return xwWorkBook;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="xwWorkBook"></param>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public static Worksheet getWorksheet(Workbook xwWorkBook, string sheetName = "Sheet1")
        {
            Worksheet xwWorksheet = null;
            try
            {
                // create new work sheet
                xwWorksheet = (Worksheet)xwWorkBook.Sheets[sheetName];
            }
            catch (Exception e)
            {
                throw e;
            }

            return xwWorksheet;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="range"></param>
        /// <param name="customerWorksheet"></param>
        /// <returns></returns>
        public static string[] GetRangeValue(string range, Worksheet customerWorksheet)
        {
            Range workingRange = customerWorksheet.get_Range(range);

            System.Array arr = (System.Array)workingRange.Cells.Value2;

            string[] output = ConvertToStringArray(arr);

            return output;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="values"></param>
        /// <returns></returns>
        private static string[] ConvertToStringArray(System.Array values)
        {
            string[] theArray = new string[values.Length];
            for (int i = 1; i <= values.Length; i++)
            {
                if (values.GetValue(1, i) == null)
                    theArray[i - 1] = "";
                else
                    theArray[i - 1] = (string)values.GetValue(1, i).ToString();
            }
            return theArray;
        }

        /// <summary>
        /// This is the default setting.
        /// </summary>
        /// <param name="path"></param>
        public static void createFromExcel(string path, 
            string startColumn, 
            string endColumn, 
            string worksheet = "Sheet1")
        {
            Configuration c = new Configuration();

            Workbook customerWorkbook = uwExcel.getWorkbook(path);
            Worksheet customerWorksheet = uwExcel.getWorksheet(customerWorkbook, worksheet);
            Range xlsRange = customerWorksheet.UsedRange;

            List<string[]> rows = new List<string[]>();
           
            // iterate through the rows to create work items
            foreach (Range row in xlsRange.Rows)
            {
                List<string> emailList = new List<string>();

                int rowNumber = row.Row;
                // skip header
                if (rowNumber == 1)
                    continue;
                // the range here is column A (1) will always be the title and column B (2) will be the description

                string[] range = uwExcel.GetRangeValue(startColumn + rowNumber + ":" + endColumn + rowNumber + "", customerWorksheet);

                emailList.Add(getValue(range, Configuration.COLUMNS.email_address));
                emailList.Add(getValue(range, Configuration.COLUMNS.email_address1));
                emailList.Add(getValue(range, Configuration.COLUMNS.email_address2));
                emailList.Add(getValue(range, Configuration.COLUMNS.email_address3));
                emailList.Add(getValue(range, Configuration.COLUMNS.email_address4));
                emailList.Add(getValue(range, Configuration.COLUMNS.email_address5));
                //emailList.Add(getValue(range, Configuration.COLUMNS.email_address6));   

                // remove all blanks
                emailList.RemoveAll(x=>x == "");

                string domain = GetEmail(emailList);

                string emailToWrite = emailList.Find(p => p.Contains(domain));

                // write value
                setValue(ref customerWorksheet, rowNumber, Convert.ToInt32(Configuration.COLUMNS.email_address) + 1, emailToWrite);

                Console.WriteLine(rowNumber.ToString());
            }
            // save                
            customerWorkbook.Save();
            customerWorkbook.Close();
        }

        public static string GetEmail(List<string> emailList) 
        {
            string emailToReturn = "";
            Dictionary<string,Configuration.EMAIL_WEIGHT> emailWeights 
                = new Dictionary<string,Configuration.EMAIL_WEIGHT>();

            foreach (string email in emailList)
            {


                string hostDomain = email.Split('@')[1];
                string domain = hostDomain.Split('.')[0];

                if (emailWeights.ContainsKey(domain))
                {
                    continue;
                }
                else
                {
                    emailWeights.Add(domain, Configuration.GetWeight(domain.ToUpper()));
                }            
            }

            var myList = emailWeights.ToList();

            myList.Sort((pair2, pair1) => pair1.Value.CompareTo(pair2.Value));

            emailToReturn = myList[0].Key;

            return emailToReturn;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="assignedTo"></param>
        /// <param name="title"></param>
        /// <param name="description"></param>
        /// <param name="iterationPath"></param>
        public static void createTestWorkItemToFile(string assignedTo, string title, string description, string iterationPath = @"CIP\V2 Shim AD Testing")
        {
            StringBuilder output = new StringBuilder();

            output.AppendLine(assignedTo);
            output.AppendLine(iterationPath);
            output.AppendLine(title);
            output.AppendLine(description);

            File.AppendAllText("testOutput.txt", output.ToString());
        }

        public static string getValue(string[] range, Configuration.COLUMNS column)
        { 
            return range[Convert.ToInt32(column)].ToString();
        }

        public static void setValue(ref Worksheet customerWorksheet, 
            int row, 
            int column, 
            string value)
        {
            Range setRange = (Range)customerWorksheet.Cells[row, column];
            setRange.set_Value(RangeValueDataType: Type.GetType(value), value: value);
        }
    }

}


