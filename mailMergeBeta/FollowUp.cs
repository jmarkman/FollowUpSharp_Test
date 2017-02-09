using System;
using System.Collections.Generic;
using System.Diagnostics;
using Word = Microsoft.Office.Interop.Word;

namespace mailMergeBeta
{
    class FollowUp
    {
        static void Main(string[] args)
        {
            /* Process
             * 1. Query SQL server for follow ups
             * 2. Remove duplicates from QFU list
             * 3. Add to Excel file
             * 4. Use Excel file in Mail Merge Word Document
             * 5. Send
             */

            Stopwatch watch = new Stopwatch();
            watch.Start();

            List<string> dbCtrlNums = new List<string>();
            List<string> dbFirstNames = new List<string>();
            List<string> dbEmails = new List<string>();

            Query imsQFU = new Query();

            // .Distinct().ToList();
            Console.WriteLine($"Getting control numbers... Current time: {watch.ElapsedMilliseconds}");
            dbCtrlNums = imsQFU.fetchCtrlNum();
            Console.WriteLine($"Getting broker names... Current time: {watch.ElapsedMilliseconds}");
            dbFirstNames = imsQFU.fetchNames();
            Console.WriteLine($"Getting emails... Current time: {watch.ElapsedMilliseconds}");
            dbEmails = imsQFU.fetchEmails();


            ExcelWrite sheet = new ExcelWrite();
            Console.WriteLine($"Instantiating Excel class... Current time: {watch.ElapsedMilliseconds}");
            sheet.addCtrlNums(dbCtrlNums);
            sheet.addNames(dbFirstNames);
            sheet.addEmails(dbEmails);
            sheet.saveWS();
            sheet.removeDuplicates();
            Console.WriteLine($"Finishing Excel usage... Current time: {watch.ElapsedMilliseconds}");
            // TODO: see how this plays out - if it sucks, implement a search of some kind instead of accessing the Excel COM

            beginMerge();
            Console.WriteLine($"Finishing Mail Merge... Current time: {watch.ElapsedMilliseconds}");
        }

        public static void beginMerge()
        {
            string filePath = $@"C:\Users\{Environment.UserName}\Documents\FollowUpSharp\followups.xlsx";

            // Instantiation of the Interop Objects for the Application and the Document
            Word.Application app = new Word.Application();
            Word.Document doc = app.Documents.Add($@"C:\Users\{Environment.UserName}\Documents\testMerge.docx", Visible: false);
            var mrg = doc.MailMerge; // Easy access to the mail merge object

            app.Visible = false;
            doc.Select();
            // "Connect" to the excel spreadsheet we just made
            mrg.OpenDataSource(filePath, SQLStatement: "SELECT * FROM [Records$]");
 
            mrg.Destination = Word.WdMailMergeDestination.wdSendToEmail;
            mrg.SuppressBlankLines = true;
            mrg.DataSource.FirstRecord = (int)Word.WdMailMergeDefaultRecord.wdDefaultFirstRecord;
            mrg.DataSource.LastRecord = (int)Word.WdMailMergeDefaultRecord.wdDefaultLastRecord;
            mrg.MailAddressFieldName = "Broker_Email";
            mrg.Execute();
            doc.Close(SaveChanges: Word.WdSaveOptions.wdDoNotSaveChanges);
            app.Quit();

            mrg = null;
            doc = null;
            app = null;
        }
    }
}
