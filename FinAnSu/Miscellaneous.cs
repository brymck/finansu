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
    public static class Miscellaneous
    {
        [ExcelFunction("Returns the value of an individual's human capital at a given age.")]
        public static double HumanCapital(
            [ExcelArgument("is the age in years at which to perform the valuation.",Name="age")] double age,
            [ExcelArgument("is the annual earnings at the given age.", Name = "initial_earnings")] double initEarnings,
            [ExcelArgument("is the real earnings growth, adjusted for inflation, Social Security and pension payments.",
                Name = "earnings_growth")] double earningsGrowth,
            [ExcelArgument("is the life expectancy in years.", Name = "life_expectancy")] double lifeExpectancy,
            [ExcelArgument("is the inflation-adjusted risk-free rate", Name = "risk-free_rate")] double rfr,
            [ExcelArgument("is the individual's risk-adjusted discount rate.", Name = "discount_rate")] double disc)
        {
            double temp = 0;
            for (int t = (int)age + 1; t <= lifeExpectancy; t++)
            {
                temp += initEarnings * Math.Pow(1 + earningsGrowth, t - age) / Math.Pow(1 + rfr + disc, t - age);
            }
            return temp;
        }
    }
}
