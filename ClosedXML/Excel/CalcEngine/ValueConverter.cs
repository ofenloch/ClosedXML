using System;
using System.Globalization;

namespace ClosedXML.Excel.CalcEngine
{
    internal class ValueConverter
    {
        private readonly CultureInfo _culture;
        private readonly CalcContext _ctx;

        public ValueConverter(CultureInfo culture, CalcContext ctx)
        {
            _culture = culture;
            _ctx = ctx;
        }

        public CultureInfo Culture => _culture;

        internal OneOf<double, Error> ToNumber(in AnyValue? value)
        {
            if (!value.HasValue)
                return Error.CellValue;

            if (value.Value.TryPickScalar(out var scalar, out var collection))
                return scalar.ToNumber(_culture);

            return collection.Match(
                    array => throw new NotImplementedException("Not sure what to do with it."),
                    reference =>
                    {
                        if (reference.TryGetSingleCellValue(out var scalarValue, _ctx))
                            return scalarValue.ToNumber(_culture);

                        throw new NotImplementedException("Not sure what to do with it.");
                    });
        }

        internal OneOf<string, Error> ToText(AnyValue value)
        {
            if (value.TryPickScalar(out var scalar, out var collection))
                return scalar.ToText(_culture);

            if (collection.TryPickT0(out var array, out var _))
                return array[0, 0].ToText(_culture);

            throw new NotImplementedException("Conversion from reference to text is not implemented yet.");
        }
    }
}
