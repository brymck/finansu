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
using System.Text;
using ExcelDna.Integration;

namespace FinAnSu
{
    public static partial class Statistics
    {
        [ExcelFunction(IsHidden = true)]
        public static double Var(double[] x)
        {
            double xAvg = x.Average();
            int n = x.Length;
            double v = 0;

            for (int i = 0; i < n; i++)
            {
                v += Math.Pow(x[i] - xAvg, 2);
            }
            return v / (n - 1);
        }

        [ExcelFunction(IsHidden = true)]
        public static double Covar(double[] x, double[] y)
        {
            double xAvg = x.Average();
            double yAvg = y.Average();
            int n = x.Length;

            if (n == y.Length)
            {
                double c = 0;
                for (int i = 0; i < n; i++)
                {
                    c += (x[i] - xAvg) * (y[i] - yAvg);
                }
                return c / n;
            }
            else
            {
                return 0;
            }
        }
    }
}
