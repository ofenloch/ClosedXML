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

            // this is an iteration demo:
            var filePath3 = Utilities.PathCombine(path, "drag-coefficient.xlsx");
            CreateDragCoefficientXLSX(filePath3);
            Utilities.UnpackPackage(filePath3);

            // this is a simple file for comparing:
            var filePath4 = Utilities.PathCombine(path, "basic-table.xlsx");
            CreateBasicTable(filePath4);
            Utilities.UnpackPackage(filePath4);

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

                sheet.Cell("A8").Value = "Bezeichnung";
                sheet.Cell("A9").Value = "";

                sheet.Cell("B8").Value = "Massenstrom";
                sheet.Cell("B9").Value = "[kg/h]";

                sheet.Cell("C8").Value = "Temperatur";
                sheet.Cell("C9").Value = "[°C]";

                sheet.Cell("D8").Value = "Druck";
                sheet.Cell("D9").Value = "[mbar]";

                sheet.Cell("E8").Value = "Dichte";
                sheet.Cell("E9").Value = "[kg/m³]";

                sheet.Cell("F8").Value = "Viskosität";
                sheet.Cell("F9").Value = "[cP]";

                sheet.Cell("G8").Value = "Volumenstrom";
                sheet.Cell("G9").Value = "[m³/h]";

                sheet.Cell("H8").Value = "Nennweite";
                sheet.Cell("H9").Value = "[mm]";

                sheet.Cell("I8").Value = "ca. Länge";
                sheet.Cell("I9").Value = "[mm]";

                sheet.Cell("J8").Value = "Geschw.";
                sheet.Cell("J9").Value = "[m/s]";

                sheet.Cell("K8").Value = "Reynolds-Zahl";
                sheet.Cell("K9").Value = "[1]";

                sheet.Cell("L8").Value = "Rauhigkeit";
                sheet.Cell("L9").Value = "[mm]";

                sheet.Cell("M8").Value = "lambda_0";
                sheet.Cell("M9").Value = "[1]";

                sheet.Cell("N8").Value = "lambda";
                sheet.Cell("N9").Value = "[1]";

                sheet.Cell("O8").Value = "Anz. Krümmer";
                sheet.Cell("O9").Value = "[1]";

                sheet.Cell("P8").Value = "Verl.beiwert";
                sheet.Cell("P9").Value = "tot.";

                sheet.Cell("Q8").Value = "Druckverlust";
                sheet.Cell("Q9").Value = "mbar";

                sheet.Cell("R8").Value = "entspr.";
                sheet.Cell("R9").Value = "m Flsg.säule";

                // adjust all columns in one shot to row 8
                sheet.Columns().AdjustToContents(8);


                // colum O (not 0 [Zero] but O as in Oliver) holds zetaProd
                uint fixed_zetaProd_row = iRow;
                uint fixed_zetaProd_col = 22;
                var fixed_zetaProd = sheet.Cell((int)iRow, (int)fixed_zetaProd_col);
                fixed_zetaProd.Value = 0.50;
                fixed_zetaProd.Style.Fill.BackgroundColor = XLColor.LightYellow;
                string fixed_zetaProdAddress = XLHelper.GetColumnLetterFromNumber((int)fixed_zetaProd_col);
                fixed_zetaProdAddress = "$" + fixed_zetaProdAddress + "$" + fixed_zetaProd_row;
                var fixed_zetaProdLabel = sheet.Cell((int)iRow, (int)fixed_zetaProd_col - 1);
                fixed_zetaProdLabel.Value = "zeta=";
                fixed_zetaProdLabel.Style.Fill.BackgroundColor = XLColor.LightYellow;

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
                    //         =B11/E11
                    var column_G = sheet.Cell((int)iRow, 7);
                    column_G.FormulaA1 = "=IF(  AND( ISNUMBER(E" + siRow + "), E" + siRow + "<>0 ), B" + siRow + "/E" + siRow + ", -1 )";
                    // column H
                    sheet.Cell((int)iRow, 8).Value = "f(prod_inner_tubediameter_" + sStreamNr + ")";
                    // column I
                    sheet.Cell((int)iRow, 9).Value = "f(prod_tubelength_" + sStreamNr + ")";
                    // column J
                    //         =G11/3600/(3,14/4*(H11/1000)^2)
                    var column_J = sheet.Cell((int)iRow, 10);
                    column_J.FormulaA1 = "=IF( AND( ISNUMBER(H" + siRow + "), H" + siRow + "<>0), G" + siRow + "/3600/(3.14/4*(H" + siRow + "/1000)^2), -1 )"; // column J
                    // column K
                    //         =J11*H11/1000*E11/(F11/1000)
                    var column_K = sheet.Cell((int)iRow, 11);
                    column_K.FormulaA1 = "=IF( AND( ISNUMBER(F" + siRow + "), F" + siRow + "<>0), J" + siRow + "*H" + siRow + "/1000*E" + siRow + "/(F" + siRow + "/1000), -1)"; // column K
                    // column L                                                                                                                                                                                 // column L
                    sheet.Cell((int)iRow, 12).Value = 0.30;
                    // column M (initial value for iteration zeta_0)
                    //         =WENN(N11=0;0,00000001;N11)
                    // We have to initialize the iteration properly to make it work. This means, cell M MUST ALWAYS 
                    // have a numeric value in the beginnimg. To assure this is always the case, we check if cell N 
                    // is a numeric value and not zero. If these conditions are not met, we set cell M to 0.00000001.
                    var column_M = sheet.Cell((int)iRow, 13);
                    column_M.FormulaA1 = "=IF( AND( ISNUMBER(N" + siRow + "), N" + siRow + "<>0 ) , N" + siRow + ", 0.00000001 )";
                    var dummy_M = column_M.Value;
                    // column N (iterative calculation of zeta)
                    //         =WENN(K11>2300;1/(2*(LOG(2,51/K11/(M11)^0,5+L11/H11/3,71)))^2;64/K11)
                    var column_N = sheet.Cell((int)iRow, 14);
                    column_N.FormulaA1 = "=IF( K" + siRow + ">2300.0 , 1.0/( 2.0*( LOG( 2.51/K" + siRow + "/( M" + siRow + " )^0.5+L" + siRow + "/H" + siRow + "/3.71 ) ) )^2 , 64/K" + siRow + " )";
                    // colum O (not 0 [Zero] but O as in Oliver)
                    sheet.Cell((int)iRow, 15).Value = "f(prod_elbows_" + sStreamNr + ")";
                    // column P
                    //         =O11*$O$10+N11*I11/H11
                    var column_P = sheet.Cell((int)iRow, 16);
                    column_P.FormulaA1 = "=IF( AND( ISNUMBER(H" + siRow + "),H" + siRow + "<>0), O" + siRow + "*" + fixed_zetaProdAddress + "+N" + siRow + "*I" + siRow + "/H" + siRow + ", -1)";
                    // column Q
                    //         =P11*E11/2*J11^2/100
                    var column_Q = sheet.Cell((int)iRow, 17);
                    column_Q.FormulaA1 = "=IF( AND( ISNUMBER(J" + siRow + "),J" + siRow + "<>0), P" + siRow + "*E" + siRow + "/2*J" + siRow + "^2/100, -1)";
                    // column R
                    //         =Q11/E11/9,81*100
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
                uint fixed_zetaVap_col = 22;
                var fixed_zetaVap = sheet.Cell((int)iRow, (int)fixed_zetaVap_col);
                fixed_zetaVap.Value = 0.50;
                fixed_zetaVap.Style.Fill.BackgroundColor = XLColor.LightYellow;
                string fixed_zetaVapAddress = XLHelper.GetColumnLetterFromNumber((int)fixed_zetaVap_col);
                fixed_zetaVapAddress = "$" + fixed_zetaVapAddress + "$" + fixed_zetaVap_row;
                var fixed_zetaVapLabel = sheet.Cell((int)iRow, (int)fixed_zetaVap_col - 1);
                fixed_zetaVapLabel.Value = "zeta=";
                fixed_zetaVapLabel.Style.Fill.BackgroundColor = XLColor.LightYellow;

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
                    //         =EXP(19,06597-4098,23/($C33+237,46532))
                    sheet.Cell((int)iRow, 4).FormulaA1 = "=IF( ISNUMBER($C" + siRow + ") , EXP( 19.06597-4098.23/($C" + siRow + "+237.46532) ), -1 )";
                    // column E
                    //         =(0,217*$D33/($C33+273,15))
                    sheet.Cell((int)iRow, 5).FormulaA1 = "=IF( ISNUMBER($C" + siRow + "), (0.217*$D" + siRow + "/($C" + siRow + "+273.15)), -1)";
                    // column F
                    sheet.Cell((int)iRow, 6).Value = "f(vap_viscosity_" + sStreamNr + ")";
                    // column G
                    //         =B33/E33
                    var column_G = sheet.Cell((int)iRow, 7);
                    column_G.FormulaA1 = "=IF(  AND( ISNUMBER(E" + siRow + "), E" + siRow + "<>0 ), B" + siRow + "/E" + siRow + ", -1 )";
                    // column H
                    sheet.Cell((int)iRow, 8).Value = "f(vap_inner_tubediameter_" + sStreamNr + ")";
                    // column I
                    sheet.Cell((int)iRow, 9).Value = "f(vap_tubelength_" + sStreamNr + ")";
                    // column J
                    //         =G33/3600/(3,14/4*(H33/1000)^2)
                    var column_J = sheet.Cell((int)iRow, 10);
                    column_J.FormulaA1 = "=G" + siRow + "/3600/(3.14/4*(H" + siRow + "/1000)^2)"; // column J
                    // column K
                    var column_K = sheet.Cell((int)iRow, 11);
                    column_K.Value = "";
                    // column L
                    sheet.Cell((int)iRow, 12).Value = 0.30;
                    // column M
                    var column_M = sheet.Cell((int)iRow, 13);
                    column_M.Value = "";
                    // column N
                    sheet.Cell((int)iRow, 14).Value = 0.03;
                    // colum O (not 0 [Zero] but O as in Oliver)
                    sheet.Cell((int)iRow, 15).Value = "f(vap_elbows_" + sStreamNr + ")";
                    // column P
                    //         =O33*$O$32+N33*I33/H33
                    var column_P = sheet.Cell((int)iRow, 16);
                    column_P.FormulaA1 = "=IF( AND( ISNUMBER(H" + siRow + "),H" + siRow + "<>0), O" + siRow + "*" + fixed_zetaVapAddress + "+N" + siRow + "*I" + siRow + "/H" + siRow + ", -1)";
                    // column Q
                    //         =P33*E33/2*J33^2/100
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
                var dummy = sheet.Cell("D5").Value;
                // correct:
                sheet.Cell("E5").FormulaA1 = "= IF( ISNUMBER(C1), \"C1 is a1 number\", \"C1 is not a1 number\" )";
                // correct:
                sheet.Cell("D6").FormulaA1 = "= IF( AND(ISNUMBER(C1), C1>1.1), C1, 0.001 )";
                dummy = sheet.Cell("D6").Value;
                // correct:
                sheet.Cell("D7").FormulaA1 = "= IF( AND(ISNUMBER(D1), D1<>0), D1, 1.001 )";
                dummy = sheet.Cell("D7").Value;

                // enable and configure iteration:
                wb.Iterate = true; // Excel's default is false (isn't it?)
                wb.IterateCount = 100; // Excel's default is 100
                wb.IterateDelta = 0.001; // Excel's default is 0.001
                var saveOptions = new SaveOptions { EvaluateFormulasBeforeSaving = true };
                // save the new workbook
                Console.WriteLine("saving template as \"{0}\"", fileName);
                wb.SaveAs(fileName, saveOptions);
            } // using (var wb = new XLWorkbook())

            // re-open the workbook and see what's in there
            using (var wb = new XLWorkbook(fileName))
            {
                var sheet = wb.Worksheets.Worksheet("Formulae");
                var dummy = sheet.Cell("D5").Value;
                Console.WriteLine("Cell D5 has (cached) value \"{0}\"", dummy);
                dummy = sheet.Cell("D6").Value;
                Console.WriteLine("Cell D6 has (cached) value \"{0}\"", dummy);
                dummy = sheet.Cell("D7").Value;
                Console.WriteLine("Cell D7 has (cached) value \"{0}\"", dummy);
            } // using (var wb = new XLWorkbook(fileName))

        } // public static void CreateSimpleTestFile(string fileName)


        public static void CreateDragCoefficientXLSX(string fileName)
        {
            if (File.Exists(fileName))
            {
                File.Delete(fileName);
            }
            var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Drag Coefficient");

            // <math>\frac{1}{\sqrt{\zeta}}\ =\ -2\ log \left( \frac{2.51}{Re \sqrt{\zeta}} + \frac{K / d_i}{3.71} \right) </math>

            worksheet.Cell("A1").Value = "inner diameter in mm";
            worksheet.Cell("B1").Value = 60.0; // inner diameter in mm
            worksheet.Cell("A2").Value = "velocity in m/s";
            worksheet.Cell("B2").Value = 20.0; // velocity in m/s
            worksheet.Cell("A3").Value = "Reynolds Number";
            worksheet.Cell("B3").Value = 2331.0; // Reynolds Number

            worksheet.Cell("A5").Value = "Zeta_0:";
            // We have to initialize the iteration properly to make it work.
            // This means, cell B5 MUST have a numeric value in the beginnimg. 
            // To assure this is always the case, we check i B6 is a numeric value and not zero.
            worksheet.Cell("B5").FormulaA1 = "=IF( AND( ISNUMBER(B6), B6<>0 ) , B6, 0.00000001 )";
            worksheet.Cell("A6").Value = "Zeta_N:";
            worksheet.Cell("B6").FormulaA1 = "=IF( B3>2300.0 , 1/( 2*(LOG(2.51/B3/(B5)^0.5+B2/B1/3.71)) )^2 , 64/B3)";

            workbook.Iterate = true;
            workbook.IterateCount = 100;
            workbook.IterateDelta = 0.00001;

            workbook.SaveAs(fileName);
        } // public static void CreateDragCoefficientXLSX(string fileName)

        public static void CreateBasicTable(string fileName)
        {
            if (File.Exists(fileName))
            {
                File.Delete(fileName);
            }
            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("Basic Table");
                // Title
                ws.Cell("B2").Value = "Contacts";

                // First Names
                ws.Cell("B3").Value = "FName";
                ws.Cell("B4").Value = "John";
                ws.Cell("B5").Value = "Hank";
                ws.Cell("B6").Value = "Dagny";

                // Last Names
                ws.Cell("C3").Value = "LName";
                ws.Cell("C4").Value = "Galt";
                ws.Cell("C5").Value = "Rearden";
                ws.Cell("C6").Value = "Taggart";

                // Id
                ws.Cell("D3").Value = "ID";
                ws.Cell("D4").Value = 3.141592653589;
                ws.Cell("D5").Value = 2.718281828459;
                ws.Cell("D6").Value = 1.414213562373;

                // Formula
                ws.Cell("E3").Value = "Formula";
                ws.Cell("E4").FormulaA1 = "=D6*D5";
                ws.Cell("E5").FormulaA1 = "=D6*D4";
                ws.Cell("E6").FormulaA1 = "=D6*D3";

                workbook.SaveAs(fileName);
            }
        } // public static void CreateBasicTable(string fileName)

    } // class TemplateGenerator

} // namespace TemplateForGeWi