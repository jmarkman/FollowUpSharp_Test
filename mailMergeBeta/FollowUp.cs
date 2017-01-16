using System.Collections.Generic;

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

            List<string> dbCtrlNums = new List<string>();
            List<string> dbFirstNames = new List<string>();
            List<string> dbEmails = new List<string>();

            Query imsQFU = new Query();

            // .Distinct().ToList();

            dbCtrlNums = imsQFU.fetchCtrlNum();
            dbFirstNames = imsQFU.fetchNames();
            dbEmails = imsQFU.fetchEmails();

            ExcelWrite sheet = new ExcelWrite();
            sheet.addCtrlNums(dbCtrlNums);
            sheet.addNames(dbFirstNames);
            sheet.addEmails(dbEmails);
            sheet.saveWS();
            sheet.removeDuplicates();
            // TODO: see how this plays out - if it sucks, implement a search of some kind instead of accessing the Excel COM

            MailMerge.beginMerge();
            // Don't do this

            //Console.ReadKey();
        }
    }
}
