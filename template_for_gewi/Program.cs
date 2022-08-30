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
            CreateTemplateForGeWi(filePath2);
            Utilities.UnpackPackage(filePath2);

        } // static void Main(string[] args)

        public static void CreateTemplateForGeWi(string fileName)
        {
            string sheetName = "Leitungsliste";
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

                // colum O (not 0 [Zero] but O as in Oliver) holds zetaProd
                uint fixed_zetaProd_row = iRow;
                uint fixed_zetaProd_col = 15;
                var fixed_zetaProd = sheet.Cell((int)iRow, (int)fixed_zetaProd_col).Value = 0.50;
                string fixed_zetaProdAddress = XLHelper.GetColumnLetterFromNumber((int)fixed_zetaProd_col);
                fixed_zetaProdAddress = "$" + fixed_zetaProdAddress + "$" + fixed_zetaProd_row;

                // create the product rows:
                for (uint i = 0; i < nDataLines; i++)
                {
                    iRow++;
                    string siRow = iRow.ToString();
                    string sStreamNr = (i + 1).ToString();

                    // column A 
                    sheet.Cell((int)iRow, 1).Value = "f(prod_name_" + sStreamNr + ")";
                    // column B
                    sheet.Cell((int)iRow, 2).Value = "f(prod_massflow_" + sStreamNr + ")";
                    // column C
                    sheet.Cell((int)iRow, 3).Value = "f(prod_temperature_" + sStreamNr + ")";
                    // column D
                    sheet.Cell((int)iRow, 4).Value = "f(prod_pressure_" + sStreamNr + ")";
                    // column E
                    sheet.Cell((int)iRow, 5).Value = "f(prod_density_" + sStreamNr + ")";
                    // column F
                    sheet.Cell((int)iRow, 6).Value = "f(prod_viscosity_" + sStreamNr + ")";
                    // column G
                    var column_G = sheet.Cell((int)iRow, 7);
                    column_G.FormulaA1 = "=IF(  AND( ISNUMBER(E" + siRow + "), E" + siRow + "<>0 ), B" + siRow + "/E" + siRow + ", -1 )";
                    // column H
                    sheet.Cell((int)iRow, 8).Value = "f(prod_nominal_diameter_" + sStreamNr + ")";
                    // column I
                    sheet.Cell((int)iRow, 9).Value = "f(prod_length_" + sStreamNr + ")";
                    // column J
                    var column_J = sheet.Cell((int)iRow, 10);
                    column_J.FormulaA1 = "=IF( AND( ISNUMBER(H" + siRow + "), H" + siRow + "<>0), G" + siRow + "/3600/(3.14/4*(H" + siRow + "/1000)^2), -1 )"; // column J
                    // column K
                    var column_K = sheet.Cell((int)iRow, 11);
                    column_K.FormulaA1 = "=IF( AND( ISNUMBER(F" + siRow + "), F" + siRow + "<>0), J" + siRow + "*H" + siRow + "/1000*E" + siRow + "/(F" + siRow + "/1000), -1)"; // column K
                    // column L                                                                                                                                                                                 // column L
                    sheet.Cell((int)iRow, 12).Value = 0.30;
                    // column M
                    var column_M = sheet.Cell((int)iRow, 13);
                    column_M.FormulaA1 = "=IF( AND( ISNUMBER(N" + siRow + "), N" + siRow + "<>0) , N" + siRow + ", 0.00000001 )";
                    // column N
                    var column_N = sheet.Cell((int)iRow, 14);
                    var dummy_M = column_M.Value;
                    column_N.FormulaA1 = "=IF( K" + siRow + ">2300.0 , 1.0/( 2.0*( LOG( 2.51/K" + siRow + "/( M" + siRow + " )^0.5+L11/H" + siRow + "/3.71 ) ) )^2 , 64/K" + siRow + " )";
                    // colum O (not 0 [Zero] but O as in Oliver)
                    sheet.Cell((int)iRow, 15).Value = "f(prod_elbows_" + sStreamNr + ")";
                    // column P
                    var column_P = sheet.Cell((int)iRow, 16);
                    column_P.FormulaA1 = "=IF( AND( ISNUMBER(H" + siRow + "),H" + siRow + "<>0), O" + siRow + "*" + fixed_zetaProdAddress + "+N" + siRow + "*I" + siRow + "/H" + siRow + ", -1)";
                    // column Q
                    var column_Q = sheet.Cell((int)iRow, 17);
                    column_Q.FormulaA1 = "=IF( AND( ISNUMBER(J" + siRow + "),J" + siRow + "<>0), P" + siRow + "*E" + siRow + "/2*J" + siRow + "^2/100, -1)";
                    // column R
                    var column_R = sheet.Cell((int)iRow, 18);
                    column_R.FormulaA1 = "=IF( AND( ISNUMBER(E" + siRow + "),E" + siRow + "<>0), Q" + siRow + "/E" + siRow + "/9.81*100, -1)";
                } // for (uint i = 0; i < nDataLines; i++)

                // insert nOffsetLines empty rows:
                for (uint i = 0; i < nOffsetLines; i++)
                {
                    iRow++;
                    string siRow = iRow.ToString();
                    sheet.Cell((int)iRow, 1).Value = "Row " + siRow;
                }

                // colum O (not 0 [Zero] but O as in Oliver) holds zetaVap
                uint fixed_zetaVap_row = iRow;
                uint fixed_zetaVap_col = 15;
                var fixed_zetaVap = sheet.Cell((int)iRow, (int)fixed_zetaVap_col).Value = 0.50;
                string fixed_zetaVapAddress = XLHelper.GetColumnLetterFromNumber((int)fixed_zetaVap_col);
                fixed_zetaVapAddress = "$" + fixed_zetaVapAddress + "$" + fixed_zetaVap_row;

                // create vapour rows:
                for (uint i = 0; i < nDataLines; i++)
                {
                    iRow++;
                    string siRow = iRow.ToString();
                    string sStreamNr = (i + 1).ToString();

                    // column A
                    sheet.Cell((int)iRow, 1).Value = "f(vap_name_" + sStreamNr + ")";
                    // column B
                    sheet.Cell((int)iRow, 2).Value = "f(vap_massflow_" + sStreamNr + ")";
                    // column C
                    sheet.Cell((int)iRow, 3).Value = "f(vap_temperature_" + sStreamNr + ")";
                    // column D
                    sheet.Cell((int)iRow, 4).FormulaA1 = "=IF( ISNUMBER($C" + siRow + ") , EXP( 19.06597-4098.23/($C" + siRow + "+237.46532) ), -1 )";
                    // column E
                    sheet.Cell((int)iRow, 5).FormulaA1 = "=IF( ISNUMBER($C" + siRow + "), (0.217*$D" + siRow + "/($C" + siRow + "+273.15)), -1)";
                    // column F
                    sheet.Cell((int)iRow, 6).Value = "f(vap_viscosity_" + sStreamNr + ")";
                    // column G
                    var column_G = sheet.Cell((int)iRow, 7);
                    column_G.FormulaA1 = "=IF(  AND( ISNUMBER(E" + siRow + "), E" + siRow + "<>0 ), B" + siRow + "/E" + siRow + ", -1 )";
                    // column H
                    sheet.Cell((int)iRow, 8).Value = "f(vap_nominal_diamater_" + sStreamNr + ")";
                    // column I
                    sheet.Cell((int)iRow, 9).Value = "f(vap_length_" + sStreamNr + ")";
                    // column J
                    var column_J = sheet.Cell((int)iRow, 10);
                    column_J.FormulaA1 = "=IF( AND( ISNUMBER(H" + siRow + "), H" + siRow + "<>0), G" + siRow + "/3600/(3.14/4*(H" + siRow + "/1000)^2), -1 )"; // column J
                    // column K
                    var column_K = sheet.Cell((int)iRow, 11);
                    column_K.FormulaA1 = "=IF( AND( ISNUMBER(F" + siRow + "), F" + siRow + "<>0), J" + siRow + "*H" + siRow + "/1000*E" + siRow + "/(F" + siRow + "/1000), -1)"; // column K
                    // column L
                    sheet.Cell((int)iRow, 12).Value = 0.30;
                    // column M
                    var column_M = sheet.Cell((int)iRow, 13);
                    column_M.FormulaA1 = "=IF( AND( ISNUMBER(N" + siRow + "), N" + siRow + "<>0) , N" + siRow + ", 0.00000001 )";
                    // column N
                    sheet.Cell((int)iRow, 14).Value = 0.03;
                    // colum O (not 0 [Zero] but O as in Oliver)
                    sheet.Cell((int)iRow, 15).Value = "f(vap_elbows_" + sStreamNr + ")";
                    // column P
                    var column_P = sheet.Cell((int)iRow, 16);
                    column_P.FormulaA1 = "=IF( AND( ISNUMBER(H" + siRow + "),H" + siRow + "<>0), O" + siRow + "*" + fixed_zetaProdAddress + "+N" + siRow + "*I" + siRow + "/H" + siRow + ", -1)";
                    // column Q
                    var column_Q = sheet.Cell((int)iRow, 17);
                    column_Q.FormulaA1 = "=IF( AND( ISNUMBER(J" + siRow + "),J" + siRow + "<>0), P" + siRow + "*E" + siRow + "/2*J" + siRow + "^2/100, -1)";
                }

                // enable and configure iteration:
                wb.Iterate = true; // Excel's default is false (isn't it?)
                wb.IterateCount = 100; // Excel's default is 100
                wb.IterateDelta = 0.001; // Excel's default is 0.001
                var saveOptions = new SaveOptions { EvaluateFormulasBeforeSaving = true };
                // save the new workbook
                Console.WriteLine("saving template as \"{0}\"", fileName);
                wb.SaveAs(fileName, saveOptions);
            } // using (var wb = new XLWorkbook())
        } // public static void CreateTemplateForGeWi(string fileName)

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