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

            Console.ReadKey();
        }

        public static void beginMerge()
        {
            // Instantiation of the Interop Objects for the Application and the Document
            Word.Application app = new Word.Application();
            Word.Document doc = app.Documents.Add($@"C:\Users\{Environment.UserName}\Documents\testMerge.docm");

            // "Connect" to the excel spreadsheet we just made
            doc.MailMerge.OpenDataSource(Name: $@"C:\Users\{Environment.UserName}\Documents\testArchive\Follow Ups.xlsx", ReadOnly: true, Connection: "Quote Follow Up Records");

            doc.MailMerge.Destination = Word.WdMailMergeDestination.wdSendToEmail;
            Console.WriteLine(doc.MailMerge.MailAddressFieldName);
            //doc.MailMerge.Execute();
        }
    }
}
