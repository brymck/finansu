/* 
 * FinAnSu
 * http://code.google.com/p/finansu/
 * 
 * Copyright 2011, Bryan McKelvey
 * Licensed under the MIT license
 * http://www.opensource.org/licenses/mit-license.php
 */

using System;
using ExcelDna.Integration;

namespace FinAnSu
{
    public static class InterestRates
    {
        #region Forward Rate Agreements
        [ExcelFunction("Returns the theoretical forward rate between tenor A and tenor B.")]
        public static double FRA(
            [ExcelArgument("is the rate through point A.", Name = "rate_a")] double rate1,
            [ExcelArgument("is time in days to point A.", Name = "ttm_a")] double ttm1,
            [ExcelArgument("is the rate through point B.", Name = "rate_b")] double rate2,
            [ExcelArgument("is time in days to point B.", Name = "ttm_b")] double ttm2,
            [ExcelArgument("is the basis in days (360, 365, etc.).", Name = "basis")] double basis)
        {
            return ((1 + rate2 * ttm2 / basis) / (1 + rate1 * ttm1 / basis) - 1) * basis / (ttm2 - ttm1);
        }

        [ExcelFunction("For short-term contracts with a maturity of less than one year from now, " +
            "returns the theoretical long forward rate given a currency forward and FX rates.")]
        public static double FRAFromFXLong(
            [ExcelArgument("is the foreign exchange spot rate.",Name="fx_spot")] double fxSpot,
            [ExcelArgument("is the long foreign exchange swap.", Name = "fx_swap_long")] double fxSwapLong,
            [ExcelArgument("is the short foreign exchange swap.", Name = "fx_swap_short")] double fxSwapShort,
            [ExcelArgument("is the foreign forward rate agreement.", Name = "foreign_fra")] double fgnFra,
            [ExcelArgument("is the time in days to the start of the FRA.", Name = "start_days")] double startDays,
            [ExcelArgument("is the time in days to the end of the FRA.", Name = "end_days")] double endDays,
            [ExcelArgument("is the domestic basis in days (360, 365, etc.).", Name = "domestic_basis")] double domBasis,
            [ExcelArgument("is the foreign basis in days (360, 365, etc.).", Name = "foreign_basis")] double fgnBasis)
        {
            return ((fxSpot + fxSwapLong) / (fxSpot + fxSwapShort) * (1 + fgnFra * (endDays - startDays) / fgnBasis) - 1) *
                domBasis / (endDays - startDays);
        }

        [ExcelFunction("For short-term contracts with a maturity of less than one year from now, " +
            "returns the theoretical short forward rate given a currency forward and FX rates.")]
        public static double FRAFromFXShort(
            [ExcelArgument("is the foreign exchange spot rate.", Name = "fx_spot")] double fxSpot,
            [ExcelArgument("is the long foreign exchange swap.", Name = "fx_swap_long")] double fxSwapLong,
            [ExcelArgument("is the short foreign exchange swap.", Name = "fx_swap_short")] double fxSwapShort,
            [ExcelArgument("is the foreign forward rate agreement.", Name = "foreign_fra")] double fgnFra,
            [ExcelArgument("is the time in days to the start of the FRA.", Name = "start_days")] double startDays,
            [ExcelArgument("is the time in days to the end of the FRA.", Name = "end_days")] double endDays,
            [ExcelArgument("is the domestic basis in days (360, 365, etc.).", Name = "domestic_basis")] double domBasis,
            [ExcelArgument("is the foreign basis in days (360, 365, etc.).", Name = "foreign_basis")] double fgnBasis)
        {
            return ((fxSpot + fxSwapShort) / (fxSpot + fxSwapLong) * (1 + fgnFra * (endDays - startDays) / fgnBasis) - 1) *
                domBasis / (endDays - startDays);
        }
        #endregion

        [ExcelFunction("Returns the Black-76 European payer/receiver swaption valuation.")]
        public static double Swaption(
            [ExcelArgument("is whether the instrument is a payer (p) or a receiver (r).",
                Name = "pay_rec_flag")] string payRecFlag,
            [ExcelArgument("is the tenor of the swap in years.", Name = "tenor")] double t1,
            [ExcelArgument("is the number of compoundings per year", Name = "periods")] double m,
            [ExcelArgument("is the current underlying swap rate.", Name = "swap_rate")] double F,
            [ExcelArgument("is the option's strike rate.", Name = "strike_rate")] double X,
            [ExcelArgument("is the time to maturity in years.", Name = "time_to_expiry")] double T,
            [ExcelArgument("is the risk-free rate through expiry.", Name = "risk-free_rate")] double r,
                        [ExcelArgument("is the implied volatility at expiry.", Name = "volatility")] double v)
        {
            double d1 = (Math.Log(F / X) + v * v / 2 * T) / (v * Math.Sqrt(T));
            double d2 = d1 - v * Math.Sqrt(T);

            if (IsPayer(payRecFlag))
            {
                return ((1 - 1 / Math.Pow(1 + F / m, t1 * m)) / F) * Math.Exp(-r * T) * (F * Statistics.CND(d1) - X * Statistics.CND(d2));
            }
            else
            {
                return ((1 - 1 / Math.Pow(1 + F / m, t1 * m)) / F) * Math.Exp(-r * T) * (X * Statistics.CND(-d2) - F * Statistics.CND(-d1));
            }
        }

        #region Helper Functions
        private static bool IsPayer(string payRecFlag)
        {
            switch (payRecFlag.ToLower())
            {
                case "p":
                case "pay":
                case "payer":
                    return true;
                case "r":
                case "rec":
                case "receive":
                case "receiver":
                    return false;
                default:
                    return true;
            }
        }
        #endregion
    }
}
