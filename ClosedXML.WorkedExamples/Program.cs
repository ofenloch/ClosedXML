using ClosedXML.Excel;
using System;
using System.IO;



namespace ClosedXML.WorkedExamples
{
    public class Program
    {
        public static string BaseCreatedDirectory
        {
            get
            {
                var path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Created");
                if (!Directory.Exists(path)) Directory.CreateDirectory(path);
                return path;
            }
        }

        public static string BaseModifiedDirectory
        {
            get
            {
                var path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Modified");
                if (!Directory.Exists(path)) Directory.CreateDirectory(path);
                return path;
            }
        }

        static void Main(string[] args)
        {
            var path = Program.BaseCreatedDirectory;
            var filePath = Path.Combine(path, "FormulaeWithIteration.xlsx");
            if (File.Exists(filePath))
            {
                File.Delete(filePath);
            }

            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Formulae With Iteration");
            ws.Cell("A1").Value = "Value 1";
            ws.Cell("B1").Value = "Value 2";
            ws.Cell("C1").Value = "Product";
            ws.Cell("A3").Value = 1.2345;
            ws.Cell("B3").Value = 2.3456;
            ws.Cell("C3").FormulaA1 = "=A3*B3";
            ws.Cell("D3").FormulaA1 = "=IF(E3=0, 0.001, E3)";
            ws.Cell("E3").FormulaA1 = "=(1.0+D3)*C3)";
            ws.Cell("D5").FormulaA1 = "=IF(E5=0, 0.001, E5)";
            ws.Cell("E5").FormulaA1 = "=1.0+D5";

            wb.CalculateMode = XLCalculateMode.Auto;
            wb.FullCalculationOnLoad = true;

            Console.WriteLine("saving XLWorkbook as {0}", filePath);
            wb.SaveAs(filePath);

        } // static void Main(string[] args)


    } // public class Program

} // namespace ClosedXML.WorkedExamples