using System;
using System.Collections.Generic;
using System.Globalization;

namespace ClosedXML.Excel.CalcEngine.Functions
{
    /// <summary>
    /// A collection of adapter functions from a more a generic formula function to more specific ones.
    /// </summary>
    internal static class SignatureAdapter
    {
        public static CalcEngineFunction Adapt(Func<CalcContext, string, ScalarValue?, AnyValue> f)
        {
            return (ctx, args) =>
            {
                var arg0Input = args[0] ?? AnyValue.From(0.0);
                if (!ToText(arg0Input, ctx.Culture).TryPickT0(out var arg0, out var error))
                    return error;

                var arg1 = args.Length > 1 && args[1].HasValue
                    ? ToScalarValue(args[1], ctx)
                    : default(ScalarValue?);

                return f(ctx, arg0, arg1);
            };
        }

        public static CalcEngineFunction Adapt(Func<double, AnyValue> f)
        {
            return (ctx, args) => ToNumber(in args[0], ctx).Match(
                    number => f(number),
                    error => error);
        }

        public static CalcEngineFunction Adapt(Func<CalcContext, double, List<Reference>, AnyValue> f)
        {
            return (ctx, args) =>
            {
                if (!ToNumber(args[0] ?? 0, ctx).TryPickT0(out var number, out var error))
                    return error;

                var references = new List<Reference>();
                for (var i = 1; i < args.Length; ++i)
                {
                    if (!(args[i] ?? 0).TryPickReference(out var reference))
                        return Error.CellValue;

                    references.Add(reference);
                }

                return f(ctx, number, references);
            };
        }

        /// <summary>
        /// Convert input value to a scalar.
        /// </summary>
        private static ScalarValue ToScalarValue(in AnyValue? value, CalcContext ctx)
        {
            if (!value.HasValue)
                return 0;

            if (value.Value.TryPickScalar(out var scalar, out var collection))
                return scalar;

            if (collection.TryPickT0(out var array, out var reference))
                return array[0, 0];

            if (reference.TryGetSingleCellValue(out var referenceScalar, ctx))
                return referenceScalar;

            return Error.CellValue;
        }

        private static OneOf<double, Error> ToNumber(in AnyValue? value, CalcContext ctx)
        {
            if (!value.HasValue)
                return Error.CellValue;

            if (value.Value.TryPickScalar(out var scalar, out var collection))
                return scalar.ToNumber(ctx.Culture);

            return collection.Match(
                array => throw new NotImplementedException("Not sure what to do with it."),
                reference =>
                {
                    if (reference.TryGetSingleCellValue(out var scalarValue, ctx))
                        return scalarValue.ToNumber(ctx.Culture);

                    throw new NotImplementedException("Not sure what to do with it.");
                });
        }

        private static OneOf<string, Error> ToText(in AnyValue value, CultureInfo culture)
        {
            if (value.TryPickScalar(out var scalar, out var collection))
                return scalar.ToText(culture);

            if (collection.TryPickT0(out var array, out var _))
                return array[0, 0].ToText(culture);

            throw new NotImplementedException("Conversion from reference to text is not implemented yet.");
        }
    }
}
