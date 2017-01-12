using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.IO;

namespace mailMergeBeta
{
    class ExcelWrite
    {
        private string sheetTitle, sheetCompany;
        private ExcelPackage qfuExcel;
        static void main(String[] args)
        {
            sheetTitle = "Test Title";
            sheetCompany = "My Company";
            using (qfuExcel = new ExcelPackage())
            {
                qfuExcel.Workbook.Properties.Title = sheetTitle;
                qfuExcel.Workbook.Properties.Company = sheetCompany;
                qfuExcel.Workbook.Worksheets.Add("Sheet1");
                ExcelWorksheet ws = qfuExcel.Workbook.Worksheets[1];
                ws.Name = "Sheet1";

                ws.Cells["A1"].Value = "Control_Number";
                ws.Cells["B1"].Value = "First_Name";
                ws.Cells["C1"].Value = "Broker_Email";

                Byte[] bin = qfuExcel.GetAsByteArray();
                File.WriteAllBytes(@"D:\Work\QFU\Follow Ups.xlsx", bin);

            }
        }
    }
}
