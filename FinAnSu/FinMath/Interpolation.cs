/* 
 * FinAnSu
 * http://code.google.com/p/finansu/
 * 
 * Copyright 2011, Bryan McKelvey
 * Licensed under the MIT license
 * http://www.opensource.org/licenses/mit-license.php
 */

using System;
using System.Collections.Generic;
using System.Linq;
using ExcelDna.Integration;

namespace FinAnSu
{
    public static class Interpolation
    {
        [ExcelFunction("Returns the interpolated rate given two sets of continuously compounded " +
            "rates and tenors.")]
        public static double ContinuousInterpolation(
            [ExcelArgument("is a range of tenors in years .", Name = "tenors")] double[] tenors,
            [ExcelArgument("is a range of continuously compounded rates.", Name = "rates")] double[] rates,
            [ExcelArgument("is the tenor for which you want an interpolated rate.", Name = "tenor_interp")] double ti)
        {
            Dictionary<double, double> pairs = new Dictionary<double, double>();
            for (int i = 0; i < tenors.Length; i++)
            {
                pairs.Add(tenors[i], rates[i]);
            }
            IEnumerable<double> lower = from pair in pairs
                                        where pair.Key <= ti
                                        select pair.Key;
            IEnumerable<double> higher = from pair in pairs
                                         where pair.Key >= ti
                                         select pair.Key;
            double t1 = lower.Max();
            double t2 = higher.Min();
            double r1 = pairs[t1];

            if (t1 == t2)
            {
                return r1;
            }

            double r2 = pairs[t2];
            double df1 = Math.Exp(-r1 * t1);
            double df2 = Math.Exp(-r2 * t2);

            double r1_2 = -Math.Log(df2 / df1) / (t2 - t1); // continuously compounded rate between points 1 and 2
            double dfi = df1 * Math.Exp(-r1_2 * (ti - t1)); // df for interpolated tenor
            return -Math.Log(dfi) / ti;
        }
    }
}
