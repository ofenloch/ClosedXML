using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

namespace ClosedXML.Examples.Misc
{
    public class Iteration : IXLExample
    {
        public void Create(String filePath)
        {
            using (var wb = new XLWorkbook())
            {
                var worksheet = wb.Worksheets.Add("Iteration");

                // solve equation
                //
                //      zeta = 1 / ( 2 * log( ( 2.51 / Re / sqrt(zeta) + (K/di/3.71) ) )^2
                //
                // iterativley for zeta (drag coefficient)

                worksheet.Cell("A1").Value = "inner tube diameter in mm:";
                worksheet.Cell("B1").Value = 60.0; // inner diameter in mm
                worksheet.Cell("A2").Value = "velocity in m/s:";
                worksheet.Cell("B2").Value = 20.0; // velocity in m/s
                worksheet.Cell("A3").Value = "Reynolds Number:";
                worksheet.Cell("B3").Value = 2331.0; // Reynolds Number

                worksheet.Cell("A5").Value = "Zeta_0:";
                // We have to initialize the iteration properly to make it work.
                // This means, cell B5 MUST have a numeric value in the beginnimg. 
                // To assure this is always the case, we check i B6 is a numeric value and not zero.
                worksheet.Cell("B5").FormulaA1 = "=IF( AND( ISNUMBER(B6), B6<>0 ) , B6, 0.00000001 )";
                worksheet.Cell("A6").Value = "Zeta_N:";
                worksheet.Cell("B6").FormulaA1 = "=IF( B3>2300.0 , 1/( 2*(LOG(2.51/B3/(B5)^0.5+B2/B1/3.71)) )^2 , 64/B3)";

                wb.Iterate = true;
                wb.IterateCount = 100;
                wb.IterateDelta = 0.00001;

                wb.SaveAs(filePath);
            }
        } // public void Create(String filePath)
    } // public class Iteration : IXLExample
} // namespace ClosedXML.Examples.Misc