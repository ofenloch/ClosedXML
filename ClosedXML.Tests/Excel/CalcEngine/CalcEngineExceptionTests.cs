using ClosedXML.Excel;
using ClosedXML.Excel.CalcEngine.Exceptions;
using ClosedXML.Tests.Utils;
using NUnit.Framework;
using System;
using System.Globalization;
using System.Threading;

namespace ClosedXML.Tests.Excel.CalcEngine
{
    [TestFixture]
    public class CalcEngineExceptionTests
    {
        [OneTimeSetUp]
        public void SetCultureInfo()
        {
            Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture("en-US");
        }

        [Test]
        public void InvalidCharNumber()
        {
            Assert.Throws<CellValueException>(() => XLWorkbook.EvaluateExpr("CHAR(-2)"));
            Assert.Throws<CellValueException>(() => XLWorkbook.EvaluateExpr("CHAR(270)"));
        }

        [Test]
        public void DivisionByZero()
        {
            Assert.Throws<DivisionByZeroException>(() => XLWorkbook.EvaluateExpr("0/0"));
            Assert.Throws<DivisionByZeroException>(() => new XLWorkbook().AddWorksheet().Evaluate("0/0"));
        }

        [Test]
        public void InvalidFunction()
        {
            Exception ex;
            ex = Assert.Throws<NameNotRecognizedException>(() => XLWorkbook.EvaluateExpr("XXX(A1:A2)"));
            Assert.That(ex.Message, Is.EqualTo("The identifier `XXX` was not recognised."));

            var ws = new XLWorkbook().AddWorksheet();
            ex = Assert.Throws<NameNotRecognizedException>(() => ws.Evaluate("XXX(A1:A2)"));
            Assert.That(ex.Message, Is.EqualTo("The identifier `XXX` was not recognised."));
        }

        [Test]
        public void NestedNameNotRecognizedException()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").SetFormulaA1("=XXX");
            ws.Cell("A2").SetFormulaA1(@"=IFERROR(A1, ""Success"")");

            Assert.AreEqual("Success", ws.Cell("A2").Value);
        }

        [Test]
        [TestCase(true, (uint)1, 0.001)]
        [TestCase(true, (uint)2, 0.001)]
        [TestCase(true, (uint)3, 0.001)]
        [TestCase(true, (uint)50, 0.01)]
        [TestCase(false, (uint)50, 0.01)]
        public void Iteration(bool iterate, uint iterateCount, double iterateDelta)
        {
            using (var tmpFile = new TemporaryFile())
            {
                var saveOptions = new SaveOptions { EvaluateFormulasBeforeSaving = true };

                using (var wb = new XLWorkbook())
                {
                    wb.Iterate = iterate;
                    wb.IterateCount = iterateCount;
                    wb.IterateDelta = iterateDelta;

                    var ws = wb.AddWorksheet("Iteration");

                    ws.Cell("A1").Value = "Fibonacci";

                    var b2 = ws.Cell("B2");
                    var b3 = ws.Cell("B3");
                    var b4 = ws.Cell("B4");

                    b2.FormulaA1 = "=IF( ISNUMBER(B3), B3, 1 )";
                    var dummy_b2 = b2.Value;
                    b3.FormulaA1 = "=IF( ISNUMBER(B4), B4, 1 )";
                    var dummy_b3 = b3.Value;
                    b4.FormulaA1 = "=IF( AND( ISNUMBER(B3), ISNUMBER(B2)), B2+B3, 1 )";

                    wb.SaveAs(tmpFile.Path, saveOptions);
                }

                using (var wb = new XLWorkbook(tmpFile.Path))
                {
                    Assert.AreEqual(wb.Iterate, iterate);
                    if (wb.Iterate == true)
                    {
                        Assert.AreEqual(wb.IterateCount, iterateCount);
                        Assert.AreEqual(wb.IterateDelta, iterateDelta);
                    }
                    else
                    {
                        Assert.AreEqual(wb.IterateCount, 0);
                        Assert.AreEqual(wb.IterateDelta, 0);
                    }
                    var ws = wb.Worksheet(1);

                    if (wb.Iterate == true)
                    {
                        ws.RecalculateAllFormulas();
                        var val_b2 = ws.Cell("B2").Value;
                        var val_b3 = ws.Cell("B3").Value;
                        var val_b4 = ws.Cell("B4").Value;
                        Assert.AreEqual(val_b2, 2);
                        Assert.AreEqual(val_b3, 2);
                        Assert.AreEqual(val_b4, 4);
                    }
                    else
                    {
                        var getValueB2 = new TestDelegate(() => { var v = ws.Cell("B2").Value; });
                        var getValueB3 = new TestDelegate(() => { var v = ws.Cell("B3").Value; });
                        var getValueB4 = new TestDelegate(() => { var v = ws.Cell("B4").Value; });
                        Exception ex;
                        ex = Assert.Throws<System.InvalidOperationException>( getValueB2 );
                        ex = Assert.Throws<System.InvalidOperationException>( getValueB3 );
                        ex = Assert.Throws<System.InvalidOperationException>( getValueB4 );
                    }
                }
            } // using (var tmpFile = new TemporaryFile())
        } // public void Iteration(bool iterate, uint iterateCount, double iterateDelta)
    }
}
