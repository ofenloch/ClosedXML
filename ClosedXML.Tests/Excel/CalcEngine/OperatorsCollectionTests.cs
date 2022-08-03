﻿using ClosedXML.Excel.CalcEngine;
using NUnit.Framework;
using System.Collections.Generic;
using System.Globalization;
using ClosedXML.Excel;

namespace ClosedXML.Tests.Excel.CalcEngine
{
    // TODO: Once array/dynamic array formulas are supported, remove internal API and replace with workbook formulas.
    /// <summary>
    /// Tests for arrays and reference operators.
    /// </summary>
    [TestFixture]
    public class OperatorsCollectionTests
    {
        private static readonly CalcContext Ctx = new(null, CultureInfo.InvariantCulture, null, null, null);

        [Test]
        public void ArrayOperandSameSizeArray_ElementsAreCalculatedAsScalarValues()
        {
            var typesPerColumn = new ConstArray(new ScalarValue[5, 5]
            {
                { true, 1, "1", "one", Error.CellReference },
                { true, 1, "1", "one", Error.CellReference },
                { true, 1, "1", "one", Error.CellReference },
                { true, 1, "1", "one", Error.CellReference },
                { true, 1, "1", "one", Error.CellReference }
            });
            var typesPerRow = new ConstArray(new ScalarValue[5, 5]
            {
                { true, true, true, true, true },
                { 2, 2, 2, 2, 2 },
                { "2", "2", "2", "2", "2"},
                { "two", "two", "two", "two", "two"},
                { Error.NumberInvalid, Error.NumberInvalid, Error.NumberInvalid, Error.NumberInvalid, Error.NumberInvalid }
            });
            var result = ToArray(AnyValue.Concat(typesPerColumn, typesPerRow, Ctx));

            for (var row = 0; row < 5; ++row)
            {
                for (var col = 0; col < 5; ++col)
                {
                    var lhs = typesPerColumn[row, col].ToAnyValue();
                    var rhs = typesPerRow[row, col].ToAnyValue();
                    AnyValue.Concat(lhs, rhs, Ctx).TryPickScalar(out var expectedResult, out var _);
                    var actualValue = result[row, col];
                    Assert.AreEqual(expectedResult, actualValue);
                }
            }
        }

        [Test]
        public void ArrayOperandDifferentSizedArray_ResizeAndUseNAForMissingValues()
        {
            AnyValue lhs = new ConstArray(new ScalarValue[2, 1] { { 1 }, { 2 } });
            AnyValue rhs = new ConstArray(new ScalarValue[1, 2] { { 3, 4 } });

            var result = ToArray(AnyValue.BinaryPlus(lhs, rhs, Ctx));

            Assert.AreEqual(result.Width, 2);
            Assert.AreEqual(result.Height, 2);
            Assert.AreEqual(result[0, 0], ScalarValue.FromT1(4));
            Assert.AreEqual(result[0, 1], ScalarValue.FromT3(Error.NoValueAvailable));
            Assert.AreEqual(result[1, 0], ScalarValue.FromT3(Error.NoValueAvailable));
            Assert.AreEqual(result[1, 1], ScalarValue.FromT3(Error.NoValueAvailable));
        }

        [Test]
        public void ArrayOperandScalar_ScalarUpscaledToArray()
        {
            AnyValue array = new ConstArray(new ScalarValue[1, 2] { { 1, 2 } });
            AnyValue scalar = ScalarValue.FromT0(true).ToAnyValue();

            var arrayPlusScalarResult = ToArray(AnyValue.BinaryPlus(array, scalar, Ctx));
            Assert.AreEqual(arrayPlusScalarResult[0, 0], ScalarValue.FromT1(2));
            Assert.AreEqual(arrayPlusScalarResult[0, 1], ScalarValue.FromT1(3));

            var scalarPlusArrayResult = ToArray(AnyValue.BinaryPlus(scalar, array, Ctx));
            Assert.AreEqual(scalarPlusArrayResult[0, 0], ScalarValue.FromT1(2));
            Assert.AreEqual(scalarPlusArrayResult[0, 1], ScalarValue.FromT1(3));
        }

        [Test]
        public void ArrayOperandSingleCellReference_ReferencedCellValueUpscaledToArray()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet() as XLWorksheet;
            ws.Cell("A1").Value = "5";
            AnyValue array = new ConstArray(new ScalarValue[1, 2] { { 10, 5 } });
            AnyValue singleCellReference = new Reference(new XLRangeAddress(XLAddress.Create("A1"), XLAddress.Create("A1")));
            var ctx = new CalcContext(null, CultureInfo.InvariantCulture, wb, ws, null);

            var arrayDividedByReference = ToArray(AnyValue.BinaryDiv(array, singleCellReference, ctx));
            Assert.AreEqual(2, arrayDividedByReference.Width);
            Assert.AreEqual(1, arrayDividedByReference.Height);
            Assert.AreEqual(arrayDividedByReference[0, 0], ScalarValue.FromT1(2));
            Assert.AreEqual(arrayDividedByReference[0, 1], ScalarValue.FromT1(1));

            var referenceDividedByArray = ToArray(AnyValue.BinaryDiv(singleCellReference, array, ctx));
            Assert.AreEqual(2, referenceDividedByArray.Width);
            Assert.AreEqual(1, referenceDividedByArray.Height);
            Assert.AreEqual(referenceDividedByArray[0, 0], ScalarValue.FromT1(0.5));
            Assert.AreEqual(referenceDividedByArray[0, 1], ScalarValue.FromT1(1));
        }

        [Test]
        public void ArrayOperandAreaReference_ReferenceBehavesAsArray()
        {
            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet() as XLWorksheet;
            ws.Cell("A1").Value = "5";
            ws.Cell("B1").Value = 1;
            ws.Cell("C1").Value = 2;
            AnyValue array = new ConstArray(new ScalarValue[1, 2] { { 10, 5 } });
            AnyValue areaReference = new Reference(new XLRangeAddress(XLAddress.Create("A1"), XLAddress.Create("C1")));

            var arrayMultArea = ToArray(AnyValue.BinaryMult(array, areaReference, new CalcContext(null, CultureInfo.InvariantCulture, wb, ws, null)));

            Assert.AreEqual(3, arrayMultArea.Width);
            Assert.AreEqual(1, arrayMultArea.Height);
            Assert.AreEqual((ScalarValue)50, arrayMultArea[0, 0]);
            Assert.AreEqual((ScalarValue)5, arrayMultArea[0, 1]);
            Assert.AreEqual((ScalarValue)Error.NoValueAvailable, arrayMultArea[0, 2]);

            var areaMultArray = ToArray(AnyValue.BinaryMult(areaReference, array, new CalcContext(null, CultureInfo.InvariantCulture, wb, ws, null)));

            Assert.AreEqual(3, areaMultArray.Width);
            Assert.AreEqual(1, areaMultArray.Height);
            Assert.AreEqual((ScalarValue)50, areaMultArray[0, 0]);
            Assert.AreEqual((ScalarValue)5, areaMultArray[0, 1]);
            Assert.AreEqual((ScalarValue)Error.NoValueAvailable, areaMultArray[0, 2]);
        }

        [Test]
        public void ArrayOperandReferenceWithMultipleAreas_ReferenceBehavesAsArrayFullOfValueErrors()
        {
            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet() as XLWorksheet;
            ws.Cell("A1").Value = "5";
            ws.Cell("B1").Value = 1;
            ws.Cell("C1").Value = 2;
            AnyValue array = new ConstArray(new ScalarValue[1, 3] { { Error.DivisionByZero, 10, 5 } });
            AnyValue multiAreaReference = new Reference(new List<XLRangeAddress>() { new XLRangeAddress(XLAddress.Create("A1"), XLAddress.Create("A1")), new XLRangeAddress(XLAddress.Create("B1"), XLAddress.Create("C1")) });

            var arrayMultReference = ToArray(AnyValue.BinaryMult(array, multiAreaReference, new CalcContext(null, CultureInfo.InvariantCulture, wb, ws, null)));
            Assert.AreEqual(3, arrayMultReference.Width);
            Assert.AreEqual(1, arrayMultReference.Height);
            Assert.AreEqual((ScalarValue)Error.DivisionByZero, arrayMultReference[0, 0]);
            Assert.AreEqual((ScalarValue)Error.CellValue, arrayMultReference[0, 1]);
            Assert.AreEqual((ScalarValue)Error.CellValue, arrayMultReference[0, 2]);

            var referenceMultArray = ToArray(AnyValue.BinaryMult(multiAreaReference, array, new CalcContext(null, CultureInfo.InvariantCulture, wb, ws, null)));
            Assert.AreEqual(3, referenceMultArray.Width);
            Assert.AreEqual(1, referenceMultArray.Height);
            Assert.AreEqual((ScalarValue)Error.CellValue, referenceMultArray[0, 0]);
            Assert.AreEqual((ScalarValue)Error.CellValue, referenceMultArray[0, 1]);
            Assert.AreEqual((ScalarValue)Error.CellValue, referenceMultArray[0, 2]);
        }

        [TestCase("A1:A1*B2:B2")]
        [TestCase("A1:A1*2")]
        [TestCase("10*B2")]
        public void SingleCellReferenceOperandSingleCellReference_UsesScalarsInCells(string formula)
        {
            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet() as XLWorksheet;
            ws.Cell("A1").Value = 10;
            ws.Cell("B2").Value = 2;
            var result = ws.Evaluate(formula);
            Assert.AreEqual(20, result);
        }

        [Test]
        public void AreaReferenceOperandAreaReference_BehavesAsArrays()
        {
            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet() as XLWorksheet;
            ws.Cell("A1").Value = 1;
            ws.Cell("B1").Value = 2;
            ws.Cell("C5").Value = 10;
            ws.Cell("E5").Value = 30;

            AnyValue leftReference = new Reference(new XLRangeAddress(XLAddress.Create("A1"), XLAddress.Create("B1")));
            AnyValue rightReference = new Reference(new XLRangeAddress(XLAddress.Create("C5"), XLAddress.Create("E5")));
            var result = ToArray(AnyValue.BinaryPlus(leftReference, rightReference, new CalcContext(null, CultureInfo.InvariantCulture, wb, ws, null)));
            Assert.AreEqual(3, result.Width);
            Assert.AreEqual(1, result.Height);
            Assert.AreEqual((ScalarValue)11, result[0, 0]);
            Assert.AreEqual((ScalarValue)2, result[0, 1]);
            Assert.AreEqual((ScalarValue)Error.NoValueAvailable, result[0, 2]);
        }

        [Test]
        public void BothAreasMultiAreaReferences_TurnsIntoSingleErrorValue()
        {
            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet() as XLWorksheet;
            AnyValue multiAreaReference = new Reference(new List<XLRangeAddress> { new XLRangeAddress(XLAddress.Create("A1"), XLAddress.Create("B1")), new XLRangeAddress(XLAddress.Create("C1"), XLAddress.Create("D1")) });
            var result = AnyValue.BinaryPlus(multiAreaReference, multiAreaReference, new CalcContext(null, CultureInfo.InvariantCulture, wb, ws, null));

            Assert.AreEqual((AnyValue)Error.CellValue, result);
        }

        [Test]
        public void AreaReferenceOperandMultiAreaReferences_TurnsIntoArrayOfErrors()
        {
            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet() as XLWorksheet;
            AnyValue multiAreaReference = new Reference(new List<XLRangeAddress> { new XLRangeAddress(XLAddress.Create("A1"), XLAddress.Create("B1")), new XLRangeAddress(XLAddress.Create("C1"), XLAddress.Create("D1")) });
            AnyValue singleAreaReference = new Reference(new XLRangeAddress(XLAddress.Create("A1"), XLAddress.Create("E2")));

            var multiAreaOperandSingleArea = ToArray(AnyValue.BinaryPlus(multiAreaReference, singleAreaReference, new CalcContext(null, CultureInfo.InvariantCulture, wb, ws, null)));

            Assert.AreEqual(5, multiAreaOperandSingleArea.Width);
            Assert.AreEqual(2, multiAreaOperandSingleArea.Height);
            multiAreaOperandSingleArea.ForEach(x => Assert.AreEqual(x, (ScalarValue)Error.CellValue));

            var singleAreaOperandMultiArea = ToArray(AnyValue.BinaryPlus(singleAreaReference, multiAreaReference, new CalcContext(null, CultureInfo.InvariantCulture, wb, ws, null)));

            Assert.AreEqual(5, singleAreaOperandMultiArea.Width);
            Assert.AreEqual(2, singleAreaOperandMultiArea.Height);
            singleAreaOperandMultiArea.ForEach(x => Assert.AreEqual(x, (ScalarValue)Error.CellValue));
        }

        [Test]
        public void UnaryOperatorOnArray()
        {
            AnyValue allTypes = new ConstArray(new ScalarValue[2, 2] { { true, -5 }, { "2", "one" } });
            var result = ToArray(allTypes.UnaryMinus(new CalcContext(null, CultureInfo.InvariantCulture, null, null, null)));
            Assert.AreEqual(2, result.Width);
            Assert.AreEqual(2, result.Height);
            Assert.AreEqual((ScalarValue)(-1), result[0, 0]);
            Assert.AreEqual((ScalarValue)5, result[0, 1]);
            Assert.AreEqual((ScalarValue)(-2), result[1, 0]);
            Assert.AreEqual((ScalarValue)Error.CellValue, result[1, 1]);
        }

        [Test]
        public void UnaryOperatorOnSingleCellReference()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("B3").Value = 4;
            var result = ws.Evaluate("-B3:B3");
            Assert.AreEqual(-4, result);
        }

        [Test]
        public void UnaryOperatorOnAreaReference()
        {
            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet() as XLWorksheet;
            ws.Cells("B3:D4").Value = 100;
            AnyValue areaReference = new Reference(new XLRangeAddress(XLAddress.Create("B3"), XLAddress.Create("D4")));

            var result = ToArray(areaReference.UnaryPercent(new CalcContext(null, CultureInfo.InvariantCulture, wb, ws, null)));
            Assert.AreEqual(3, result.Width);
            Assert.AreEqual(2, result.Height);
            result.ForEach(value => Assert.AreEqual((ScalarValue)1, value));
        }

        [Test]
        public void UnaryOperatorOnMultiAreaReference_TurnsIntoSingleErrorValue()
        {
            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet() as XLWorksheet;
            AnyValue reference = new Reference(new List<XLRangeAddress> { new XLRangeAddress(XLAddress.Create("A1"), XLAddress.Create("B1")), new XLRangeAddress(XLAddress.Create("C1"), XLAddress.Create("D1")) });

            var result = reference.UnaryPercent(new CalcContext(null, CultureInfo.InvariantCulture, wb, ws, null));
            Assert.AreEqual((AnyValue)Error.CellValue, result);
        }

        private static Array ToArray(AnyValue value)
        {
            return value.Match(
                logical => throw new System.InvalidOperationException(),
                number => throw new System.InvalidOperationException(),
                text => throw new System.InvalidOperationException(),
                error => throw new System.InvalidOperationException(),
                array => array,
                reference => throw new System.InvalidOperationException());
        }
    }
}
