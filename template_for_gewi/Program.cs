using ClosedXML.Excel;
using System;
using System.Globalization;
using System.IO;

namespace TemplateForGeWi
{
    class TemplateGenerator
    {
        internal static readonly CultureInfo ParseCulture = CultureInfo.InvariantCulture;

        static void Main(string[] args)
        {
            Console.WriteLine("This is {0}.", Environment.GetCommandLineArgs());
            Console.WriteLine(" current working directory is {0}", Directory.GetCurrentDirectory());
            var path = Utilities.BaseCreatedDirectory;

            // just to see if I'm doing it wright:
            var filePath1 = Utilities.PathCombine(path, "formulae.xlsx");
            CreateSimpleTestFile(filePath1);
            Utilities.UnpackPackage(filePath1);

            // this is gonna be the real thing:
            var filePath2 = Utilities.PathCombine(path, "template-for-gewi.xlsx");
            CreateTemplateForGeWi(filePath2, false);
            Utilities.UnpackPackage(filePath2);

        } // static void Main(string[] args)

        public static void CreateTemplateForGeWi(string fileName, bool withIteration = false)
        {
            string sheetName = withIteration ? "Template With Iteration" : "Template";
            uint nOffsetLines = 10;
            uint nDataLines = 30;
            if (File.Exists(fileName))
            {
                File.Delete(fileName);
            }
            using (var wb = new XLWorkbook())
            {
                var sheet = wb.Worksheets.Add(sheetName);

                // this is the "real row" counter
                uint iRow = 0;

                // insert nOffsetLines empty rows:
                for (uint i = 0; i < nOffsetLines; i++)
                {
                    iRow++;
                    string siRow = iRow.ToString();
                    sheet.Cell((int)iRow, 1).Value = "Row " + siRow;
                }

                // colum O (as in Oliver)
                uint fixed_cetaProd_row = iRow;
                uint fixed_cetaProd_col = 15;
                var fixed_cetaProd = sheet.Cell((int)iRow, (int)fixed_cetaProd_col).Value = 0.50;

                // create the product rows:
                for (uint i = 0; i < nDataLines; i++)
                {
                    iRow++;
                    string siRow = iRow.ToString();

                    sheet.Cell((int)iRow, 1).Value = "Row " + siRow; // column A 
                    sheet.Cell((int)iRow, 2).Value = 2; // column B
                    sheet.Cell((int)iRow, 3).Value = 3; // column C
                    sheet.Cell((int)iRow, 4).Value = 4; // column D
                    sheet.Cell((int)iRow, 5).Value = 5; // column E
                    sheet.Cell((int)iRow, 6).Value = 6; // column F
                    sheet.Cell((int)iRow, 7).Value = 7; // column G
                    sheet.Cell((int)iRow, 8).Value = 8; // column H
                    sheet.Cell((int)iRow, 9).Value = 9; // column I
                    sheet.Cell((int)iRow, 10).Value = 10; // column J
                    sheet.Cell((int)iRow, 11).Value = 11; // column K
                    sheet.Cell((int)iRow, 12).Value = 12; // column L

                    // column M
                    var column_M = sheet.Cell((int)iRow, 13);
                    column_M.FormulaA1 = "=IF( AND( ISNUMBER(N" + siRow + "), N" + siRow + "<>0) , N" + siRow + ", 0.00000001 )";
                    //column_M.Value = 2.345;
                    // column N
                    var column_N = sheet.Cell((int)iRow, 14);
                    var dummy_M = column_M.Value;
                    column_N.FormulaA1 = "=IF( K" + siRow + ">2300.0 , 1.0/(2.0*(LOG(2.51/K" + siRow + "/(M" + siRow + ")^0.5+L11/H" + siRow + "/3.71)))^2 , 64/K" + siRow + " )";
                } // for (uint i = 0; i < nDataLines; i++)

                // insert nOffsetLines empty rows:
                for (uint i = 0; i < nOffsetLines; i++)
                {
                    iRow++;
                    string siRow = iRow.ToString();
                    sheet.Cell((int)iRow, 1).Value = "Row " + siRow;
                }

                // colum O (as in Oliver)
                uint fixed_cetaVap_row = iRow;
                uint fixed_cetaVap_col = 15;
                var fixed_cetaVap = sheet.Cell((int)iRow, (int)fixed_cetaVap_col).Value = 0.50;

                // create vapour rows:
                for (uint i = 0; i < nDataLines; i++)
                {
                    iRow++;
                    string siRow = iRow.ToString();
                    sheet.Cell((int)iRow, 1).Value = "Row " + siRow; // column A

                    var column_P = sheet.Cell((int)iRow, 16);
                    column_P.FormulaA1 = "=O" + siRow + "*$O$" + fixed_cetaVap_row + " + N" + siRow + "* I" + siRow + " / H" + siRow;

                }

                // enable and configure iteration:
                wb.Iterate = true; // Excel's default is false (isn't it?)
                wb.IterateCount = 50; // Excel's default is 100
                wb.IterateDelta = 0.01; // Excel's default is 0.001
                // save the new workbook
                Console.WriteLine("saving template as \"{0}\"", fileName);
                wb.SaveAs(fileName);
            } // using (var wb = new XLWorkbook())



        } // public static void CreateTemplateForGeWi(string fileName, bool withIteration = false)

        // This is just for testing if I'm doing it correctly:
        public static void CreateSimpleTestFile(string fileName)
        {
            if (File.Exists(fileName))
            {
                File.Delete(fileName);
            }
            using (var wb = new XLWorkbook())
            {
                var sheet = wb.Worksheets.Add("Formulae");
                var a1 = sheet.Cell("A1");
                var b1 = sheet.Cell("B1");
                var d1 = sheet.Cell("D1");
                a1.Value = 1.2345;
                b1.Value = 2.3456;
                d1.FormulaA1 = "=A1+B1";
                var d4 = sheet.Cell("D4");
                d4.FormulaA1 = "=IF( D1 > 3.5, 10, 20 )";
                // correct: 
                sheet.Cell("D5").FormulaA1 = "= IF( D1 < 3.5, 10, 20 )";
                // correct:
                sheet.Cell("E5").FormulaA1 = "= IF( ISNUMBER(C1), \"C1 is a1 number\", \"C1 is not a1 number\" )";
                // correct:
                sheet.Cell("D6").FormulaA1 = "= IF( AND(ISNUMBER(C1), C1>1.1), C1, 0.001 )";
                // correct:
                sheet.Cell("D7").FormulaA1 = "= IF( AND(ISNUMBER(D1), D1<>0), D1, 1.001 )";

                wb.SaveAs(fileName);
            }
        } // public static void CreateSimpleTestFile(string fileName)

    } // class TemplateGenerator

} // namespace TemplateForGeWi