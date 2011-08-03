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
    public static class YieldCurve
    {
        [ExcelFunction("Returns the discount factor rate given Eurodollar futures and swap rates.")]
        public static double Interpolate(
            [ExcelArgument("is a range of Eurodollar expiration dates.", Name = "e$_dates")] double[] edd,
            [ExcelArgument("is a range of Eurodollar prices.", Name = "e$_prices")] double[] edp)
        {
            return 2;
        }

        [ExcelFunction("Calculates the spot stub.")]
        public static double Stub(
            [ExcelArgument("is the target valuation date.", Name = "valuation_date")] double valuationDate,
            [ExcelArgument("is the overnight value date (for LIBOR, usually 1 business day after trade date).", Name = "overnight_start")] double overnightStart,
            [ExcelArgument("is the overnight value.", Name = "overnight_value")] double overnightValue,
            [ExcelArgument("is the term valuation date (for LIBOR, usually 2 business days after trade date).", Name = "term_start")] double termStart,
            [ExcelArgument("is either 1 or 2 term LIBOR expiration dates.", Name = "term_ends")] double[] termEnds,
            [ExcelArgument("is either 1 or 2 term LIBOR values.", Name = "term_values")] double[] termValues)
        {
            double termValue;

            if (termEnds.Length == 2 && termValues.Length == 2)
            {
                double pct = (valuationDate - termEnds[0]) / (termEnds[1] - termEnds[0]);
                termValue = pct * termValues[0] + (1 - pct) * termValues[1];
            }
            else
            {
                termValue = termValues[0];
            }

            return ((1 + overnightValue * (termStart - overnightStart) / 360) * (1 + termValue * (valuationDate - termStart) / 360) - 1) *
                360 / (valuationDate - overnightStart);
        }

        [ExcelFunction("Interpolates a rate based on a cubic spline.")]
        public static double[,] CubicSpline(
            [ExcelArgument("is a range of unknown dates for which to find the interpolated rates.", Name = "interp_dates")] double[] idates,
            [ExcelArgument("is a range of known dates or term points.", Name = "known_dates")] double[] dates,
            [ExcelArgument("is a range of known rates.", Name = "known_rates")] double[] rates
            )
        {
            int count = dates.Length;
            double[] h = new double[count];
            double[] alpha = new double[count];
            double[] l = new double[count];
            double[] mu = new double[count];
            double[] z = new double[count];
            double[] b = new double[count];
            double[] c = new double[count];
            double[] d = new double[count];
            double[,] results = new double[idates.Length, 1];

            if (count != rates.Length)
            {
                return new double[,] { { 0 } };
            }
            else
            {
                h[0] = dates[1] - dates[0];
                h[count - 1] = 0;
                alpha[0] = 0;
                alpha[count - 1] = 0;
                l[0] = 1;
                l[count - 1] = 0;
                mu[0] = 0;
                mu[count - 1] = 0;
                z[0] = 0;
                z[count - 1] = 0;

                for (int i = 1; i < count - 1; i++)
                {
                    h[i] = dates[i + 1] - dates[i];
                    alpha[i] = 3 / h[i] * (rates[i + 1] - rates[i]) - 3 / h[i - 1] * (rates[i] - rates[i - 1]);
                    l[i] = 2 * (dates[i + 1] - dates[i - 1]) - h[i - 1] * mu[i - 1];
                    mu[i] = h[i] / l[i];
                    z[i] = (alpha[i] - h[i - 1] * z[i - 1]) / l[i];
                }

                for (int i = count - 2; i >= 0; i--)
                {
                    c[i] = z[i] - mu[i] * c[i + 1];
                    b[i] = (rates[i + 1] - rates[i]) / h[i] - h[i] * (c[i + 1] + 2 * c[i]) / 3;
                    d[i] = (c[i + 1] - c[i]) / h[i] / 3;
                }

                for (int i = 0; i < idates.Length; i++)
                {
                    if (idates[i] < dates[0] || idates[i] > dates[count - 1])
                    {
                        results[i, 0] = 0;
                    }
                    else
                    {
                        double x;
                        int j = 0;

                        while (idates[i] > dates[j + 1])
                        {
                            j++;
                        }
                        x = idates[i] - dates[j];
                        results[i, 0] = d[j] * Math.Pow(x, 3) + c[j] * Math.Pow(x, 2) + b[j] * x + rates[j];
                    }
                }

                return results;
            }
        }

        [ExcelFunction("Provides days between a start and end date based on daycount convention.")]
        public static double Days(
            [ExcelArgument("is the starting date.", Name = "start_date")] double startDate,
            [ExcelArgument("is the ending date.", Name = "end_date")] double endDate,
            [ExcelArgument("is the daycount convention (e.g. \"Act/360\", \"30/360\").")] string convention)
        {
            DateTime start = DateTime.FromOADate(startDate);
            DateTime end = DateTime.FromOADate(endDate);
            int startDay = start.Day;
            int startMonth = start.Month;
            int startYear = start.Year;
            int endDay = end.Day;
            int endMonth = end.Month;
            int endYear = end.Year;

            switch (convention.Substring(0, convention.IndexOf("/")).ToLower())
            {
                case "act":
                    return endDate - startDate;
                case "30":
                    if (startDay == 31 || (startMonth == 2 && startDay >= 28))
                    {
                        startDay = 30;
                    }

                    if (endDay == 31 || (endMonth == 2 && endDay >= 28))
                    {
                        if (startDay == 30)
                        {
                            endDay = 30;
                        }
                        else
                        {
                            endDay = 1;
                            if (endMonth == 12)
                            {
                                endMonth = 1;
                                endYear++;
                            }
                            else
                            {
                                endMonth++;
                            }
                        }
                    }

                    return (endDay - startDay) + (endMonth - startMonth) * 30 + (endYear - startYear) * 360;
                default:
                    return 0;
            }
        }

        [ExcelFunction("Provides days between a start and end date based on daycount convention.")]
        public static double Years(
            [ExcelArgument("is the starting date.", Name = "start_date")] double startDate,
            [ExcelArgument("is the ending date.", Name = "end_date")] double endDate,
            [ExcelArgument("is the daycount convention (e.g. \"Act/360\", \"30/360\").")] string convention)
        {
            double days = Days(startDate, endDate, convention);

            switch (convention.Substring(convention.IndexOf("/") + 1).ToLower())
            {
                case "360":
                    return days / 360;
                case "365":
                    return days / 365;
                case "act/act":
                    return 0;
                default:
                    return 0;
            }
        }

        public static double[,] dfFromForwards(double[] dates, double[] rates, double startDay, string convention)
        {
            double[,] dfs = new double[dates.Length, 1];
            double prevDf = 1;
            double prevDay = startDay;

            if (dates.Length != rates.Length)
            {
                return new double[,] { { 0 } };
            }
            else
            {
                for (int i = 0; i < dates.Length; i++)
                {
                    dfs[i, 0] = prevDf / (1 + rates[i] * Years(prevDay, dates[i], convention));
                    prevDay = dates[i];
                }
                return dfs;
            }
        }
    }
}