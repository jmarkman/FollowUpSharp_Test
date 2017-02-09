using System;
using System.Collections.Generic;
using OfficeOpenXml;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace mailMergeBeta
{
    /// <summary>
    /// Excel document creation class that stores the information gathered from the query in
    /// a MS Excel spreadsheet for usage in Mail Merge.</summary>
    class ExcelWrite
    {
        private ExcelPackage qfuExcel;
        private ExcelWorksheet ws;
        /// <summary>
        /// Constructor for the ExcelWrite class. Doesn't accept any arguments as of now.
        /// </summary>
        public ExcelWrite()
        {
            // Instantiation of the EPPlus ExcelPackage class
            qfuExcel = new ExcelPackage();
            // Creation of worksheet in a workbook in the ExcelPackage object
            qfuExcel.Workbook.Worksheets.Add("Quote Follow Ups");
            // Targeting of worksheets works on an index basis, with the sheet index starting at position 1
            ws = qfuExcel.Workbook.Worksheets[1];
            ws.Name = "Records";

            // Value assignment to cells is direct
            ws.Cells["A1"].Value = "Control_Number";
            ws.Cells["B1"].Value = "First_Name";
            ws.Cells["C1"].Value = "Broker_Email";

            // Changing font style is a boolean value, while color changing is more direct
            ws.Cells["A1"].Style.Font.Bold = true;
            ws.Cells["B1"].Style.Font.Bold = true;
            ws.Cells["C1"].Style.Font.Bold = true;
        }

        /// <summary>
        /// This method, using a list provided as an argument, adds the contents of that list to
        /// a specific column. Since this handles the control numbers, this will be column 1.
        /// </summary>
        /// <param name="_ctrlNums">Accepts a string-based List as an argument. This will 100%
        /// be a list of control numbers represented as strings.</param>
        public void addCtrlNums(List<string> _ctrlNums)
        {
            /*
             * for each row "row" where row is less than the length of the list
             *     for each item in the list "_ctrlNums" for the length of the list
             *        ws.Cells[row, 1].Value = _ctrlNums[i]
             */
            for (int row = 2, i = 0; i < _ctrlNums.Count; row++, i++)
            {
                ws.Cells[row, 1].Value = _ctrlNums[i];
            }
            // Excel generally works on an index 1 basis, so column A would be 1
        }

        /// <summary>
        /// This method, using a list provided as an argument, adds the contents of that list to
        /// a specific column. Since this handles the names, this will be column 2.
        /// </summary>
        /// <param name="_names">Accepts a string-based List as an argument. This will 100%
        /// be a list of broker names represented as strings.</param>
        public void addNames(List<string> _names)
        {
            for (int row = 2, i = 0; i < _names.Count; row++, i++)
            {
                ws.Cells[row, 2].Value = _names[i];
            }
        }

        /// <summary>
        /// This method, using a list provided as an argument, adds the contents of that list to
        /// a specific column. Since this handles the emails, this will be column 3.
        /// </summary>
        /// <param name="_emails">Accepts a string-based List as an argument. This will 100%
        /// be a list of emails represented as strings.</param>
        public void addEmails(List<string> _emails)
        {
            for (int row = 2, i = 0; i < _emails.Count; r++, i++)
            {
                ws.Cells[row, 3].Value = _emails[i];
            }
        }

        // TODO: Add methods for adding property location names to sheet
        // TODO: Add methods for adding effective dates to sheet

        /// <summary>
        /// Saves the worksheet to a specified direcetory since EPPlus works with Excel in memory.
        /// </summary>
        public void saveWS()
        {
            DateTime theDate = DateTime.Now; // Instantiate a datetime object to get today's date
            string today = theDate.ToString("MM-dd-yyyy"); // Get today's date as a string in the format month-day-year
            string archivePath = $@"C:\Users\{Environment.UserName}\Documents\FollowUpSharp\Quote Follow Ups Archive\";
            string excelWorkbookName = $"Follow Ups for {today}.xlsx";
            /*
             * Capture the Excel file currently being worked with in memory and store it in a byte array
             * so it can be saved to the user's computer
             */
            Byte[] bin = qfuExcel.GetAsByteArray();
            try
            {
                File.WriteAllBytes($@"C:\Users\{Environment.UserName}\Documents\FollowUpSharp\followups.xlsx", bin);
                File.WriteAllBytes(Path.Combine(archivePath, excelWorkbookName), bin);
            }
            catch (DirectoryNotFoundException)
            {
                Console.WriteLine("Directory not found! Creating...");
                Directory.CreateDirectory($@"C:\Users\{Environment.UserName}\Documents\FollowUpSharp");
                Directory.CreateDirectory($@"C:\Users\{Environment.UserName}\Documents\FollowUpSharp\Quote Follow Ups Archive\");
                File.WriteAllBytes($@"C:\Users\{Environment.UserName}\Documents\FollowUpSharp\followups.xlsx", bin);
                File.WriteAllBytes(Path.Combine(archivePath, excelWorkbookName), bin);
            }
        }

        public void removeDuplicates()
        {
            /*
             * This was an interesting read on using the internal
             * library for accessing, using, and disposing of Excel 
             * COM objects: https://coderwall.com/p/app3ya/read-excel-file-in-c
             * 
             * I don't know why EPPlus doesn't have a built in version of this,
             * although I can also accomplish this with a linear search since
             * Big O notation speed doesn't matter too much for this application
             * right now, since it'll still be faster than writing emails/generating 
             * emails by storing them in a signature and editing them individually ALL DAY LONG.
             * 
             * Getting the items from the actual WKFC database will be sort of its own mess
             * in terms of speed since its a perfect storm of Dan's SQL ability being low and
             * the way to access ANYTHING in the database is a long-ass list of inner joins.
             * 
             * Below, 4 objects are instantiated:
             * excel, excelWB, excelWS, sheetRange
             * All 4 of these should be pretty self explanatory in what their instances represent
             */
            Excel.Application excel = new Excel.Application();
            Excel.Workbook excelWB = excel.Workbooks.Open($@"C:\Users\{Environment.UserName}\Documents\FollowUpSharp\followups.xlsx");
            Excel.Worksheet excelWS = excelWB.Sheets[1];
            Excel.Range sheetRange = excelWS.UsedRange;

            // Call upon Excel's "Remove Duplicate Values" function
            sheetRange.RemoveDuplicates(1);

            /*
             * Now, we need to release the COM objects from memory. Mr. Garland says
             * that "[i]f this is not properly done, then there will be lingering 
             * processes that will hold the file access writes to your Excel workbook."
             * No bueno. 
             */
            Marshal.ReleaseComObject(sheetRange);
            Marshal.ReleaseComObject(excelWS);
            excelWB.Close(true);
            Marshal.ReleaseComObject(excelWB);
            Marshal.ReleaseComObject(excel);
        }      
    }
}
