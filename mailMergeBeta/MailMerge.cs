using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace mailMergeBeta
{
    class MailMerge
    {
        public static void beginMerge()
        {
            Word.Application app = new Word.Application();
            app.Visible = true;
            Word.Document doc = app.Documents.Open($@"C:\Users\{Environment.UserName}\Documents\testMerge.docm");
            
            // "Connect" to the excel spreadsheet we just made
            doc.MailMerge.OpenDataSource(Name: $@"C:\Users\{Environment.UserName}\Documents\testArchive\Follow Ups.xlsx", ReadOnly: true, Connection: "Quote Follow Up Records");
            

                        
        }
    }
}
