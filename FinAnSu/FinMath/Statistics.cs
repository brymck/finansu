/* 
 * FinAnSu
 * http://code.google.com/p/finansu/
 * 
 * Copyright 2011, Bryan McKelvey
 * Licensed under the MIT license
 * http://www.opensource.org/licenses/mit-license.php
 */

using System;
using System.Linq;
using ExcelDna.Integration;

namespace FinAnSu
{
    public static partial class Statistics
    {
        #region Cumulative Normal Distribution
        [ExcelFunction("Returns the standard normal cumulative distribution (has a mean of zero and" +
            "a standard deviation of one).")]
        public static double CND(
            [ExcelArgument("is the value for which you want the distribution.", Name = "z")] double z)
        {
            double L;
            double K;
            double dCND;
            const double b = 0.2316419;
            const double a1 = 0.31938153;
            const double a2 = 0.356563782;
            const double a3 = 1.781477937;
            const double a4 = 1.821255978;
            const double a5 = 1.330274429;

            L = Math.Abs(z);
            K = 1 / (1 + b * L);
            dCND = 1 - 1 / Math.Sqrt(2 * Math.PI) *
                Math.Exp(-L * L / 2) *
                (a1 * K -
                a2 * K * K +
                a3 * K * K * K -
                a4 * K * K * K * K +
                a5 * K * K * K * K * K);

            if (z < 0) { return 1 - dCND; }
            else { return dCND; }
        }
        #endregion

        #region Probability Density Function
        [ExcelFunction("Returns the probability density function.")]
        public static double PDF(
            [ExcelArgument("is the value for which you want the distribution.", Name = "z")] double z)
        {
            return Math.Exp(-(z * z) / 2) / Math.Sqrt(2 * Math.PI);
        }
        #endregion

        #region Normal Distribution Function
        [ExcelFunction("Returns the normal distribution function.")]
        public static double ND(
            [ExcelArgument("is the value for which you want the distribution.", Name = "x")] double x)
        {
            return Math.Exp(-x * x / 2) / Math.Sqrt(2 * Math.PI);
        }
        #endregion

        #region Bivariate Normal Distribution Function
        [ExcelFunction("Returns the bivariate normal distribution function.")]
        public static double BND(double x, double y, double rho)
        {
            return Math.Exp(-1 / (2 * (1 - rho * rho)) * (x * x + y * y - 2 * x * y * rho)) /
                (2 * Math.PI * Math.Sqrt(1 - rho * rho));
        }
        #endregion

        #region Inverse Cumulative Normal Distribution Function
        [ExcelFunction("Returns the inverse cumulative normal distribution function.")]
        public static double CNDEV(
                        [ExcelArgument("is the value for which you want the distribution.", Name = "U")] double U)
        {
            double x;
            double r;
            double[] A = new double[] { 2.50662823884, -18.61500062529, 41.39119773534, -25.44106049637 };
            double[] b = new double[] { -8.4735109309, 23.08336743743, -21.06224101826, 3.13082909833 };
            double[] c = new double[]{0.337475482272615, 0.976169019091719, 0.160797971491821, 2.76438810333863E-02,
                3.8405729373609E-03, 3.951896511919E-04, 3.21767881767818E-05, 2.888167364E-07, 3.960315187E-07};

            x = U - .5;
            if (Math.Abs(x) < .92)
            {
                r = x * x;
                r = x * (((A[3] * r + A[2]) * r + A[1]) * r + A[0]) /
                    ((((b[3] * r + b[2]) * r + b[1]) * r + b[0]) * r + 1);
                return r;
            }
            else
            {
                if (x >= 0)
                {
                    r = 1 - U;
                }
                else
                {
                    r = U;
                }
                r = Math.Log(-Math.Log(r));
                r = c[0] + r * (c[1] + r * (c[2] + r * (c[3] + r + (c[4] +
                    r * (c[5] + r * (c[6] + r * (c[7] + r * c[8])))))));
                if (x < 0)
                {
                    return -r;
                }
                else
                {
                    return r;
                }
            }
        }
        #endregion
    }
}
