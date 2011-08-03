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
    public static class Conversion
    {
        [ExcelFunction("Returns the dollar duration of a cash flow, given a start and end date")]
        public static double DollarDur(
            [ExcelArgument("is the cash flow to occur in the future.", Name = "cash_flow")] double cashFlow,
            [ExcelArgument("is the starting or valuation date.", Name = "valuation_date")] double valDate,
            [ExcelArgument("is the ending or payment date.", Name = "payment_date")] double pmtDate,
            double df, string dc)
        {
            return cashFlow * (FwdToDf(dfToFwd(df, valDate, pmtDate, dc) + .0001,
                valDate, pmtDate, dc) - df);
        }

        [ExcelFunction("Converts a discount factor to a forward rate.")]
        public static double dfToFwd(double df, double valDate, double pmtDate,
            string dc)
        {
            switch (dc.ToUpper())
            {
                case "ACT/360":
                    return (Math.Pow(df, 91.31 / (valDate - pmtDate))
                        - 1) * 360 / 91.31;
                case "30/360":
                case "30E/360":
                    return (Math.Pow(df, 182.62 / (valDate - pmtDate))
                        - 1) * 2;
                default:
                    return 0;
            }
        }

        [ExcelFunction("Converts a forward rate to a discount factor.")]
        public static double FwdToDf(double fwd, double valDate, double pmtDate,
            string dc)
        {
            switch (dc.ToUpper())
            {
                case "ACT/360":
                    return Math.Pow(1 + fwd * 91.31 / 360, (valDate - pmtDate) / 91.31);
                case "30E/360":
                case "30/360":
                    return Math.Pow(1 + fwd / 2, (valDate - pmtDate) / 182.62);
                default:
                    return 0;
            }
        }

        [ExcelFunction("Adds a spread to a rate, allowing for various daycounts.")]
        public static double AddSpread(double fwd, double spd, string fwdDc, string spdDc)
        {
            return FwdToFwd(FwdToFwd(fwd, fwdDc, spdDc) + spd, spdDc, fwdDc);
        }

        [ExcelFunction("Converts a forward rate to a forward rate with a different daycount.")]
        public static double FwdToFwd(double fwd, string fromDc, string toDc)
        {
            return dfToFwd(FwdToDf(fwd, 0, 365.24, fromDc),
                0, 365.24, toDc);
        }
    }
}