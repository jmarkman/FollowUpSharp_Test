using System;
using System.Collections.Generic;
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
            var app = new Word.Application();
            var doc = new Word.Document();
            var mailMerge = doc.MailMerge;

            // "Connect" to the excel spreadsheet we just made
            mailMerge.OpenDataSource(@"D:\Work\Follow Ups.xlsx", ReadOnly: false,  LinkToSource: true, Connection: @"Provider=Microsoft.ACE.OLEDB.12.0;User ID=Admin;Data Source=D:\Work\Follow Ups.xlsx;Mode=Read;Extended Properties=""HDR=YES;IMEX=1;"";Jet OLEDB:System database="""";Jet OLEDB:Registry Path="""";Je", SQLStatement: "SELECT * FROM 'Quote Follow Up Records$'");

        }
    }
}
