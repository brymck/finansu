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
    public static class PortfolioManagement
    {
        #region "Execution"
        [ExcelFunction("Returns the quoted spread from the buyer's perspective")]
        public static decimal QuotedSpread(
            [ExcelArgument("is the market bid price at execution time.", Name = "bid")] decimal bid,
            [ExcelArgument("is the market ask price at execution time.", Name = "ask")] decimal ask,
            [ExcelArgument("is the actual execution price.", Name = "execution")] decimal execution)
        {
            return ask - bid;
        }

        [ExcelFunction("Returns the effective spread from the buyer's perspective")]
        public static decimal EffectiveSpread(
            [ExcelArgument("is the market bid price at execution time.", Name = "bid")] decimal bid,
            [ExcelArgument("is the market ask price at execution time.", Name = "ask")] decimal ask,
            [ExcelArgument("is the actual execution price.", Name = "execution")] decimal execution)
        {
            return ask + bid - execution * 2;
        }

        // add VWAP, opening and closing price implicit transaction costs
        #endregion

        #region "Rate of Return"
        [ExcelFunction("Returns the chain-linked time-weighted rate of return (TWR).")]
        public static decimal TWR(
            [ExcelArgument("is a range of days in which cash flows occurred.", Name = "days")] decimal[] days,
            [ExcelArgument("is a range of ending balances on days in which cash flows occurred.",
                Name = "balances")] decimal[] ebals,
            [ExcelArgument("is a range of cash flow amounts (positive = contribution, negative = withdrawal.",
                Name = "cash_flows")] decimal[] cfs)
        {
            if (days.Length != ebals.Length || ebals.Length != cfs.Length)
            {
                return 0;
            }
            else
            {
                decimal ret = 1;
                for (int i = 1; i < days.Length; i++)
                {
                    ret *= (ebals[i] - cfs[i]) / ebals[i - 1];
                }
                return ret - 1;
            }
        }

        // add MWR VI.131
        #endregion

        #region "CAPM"
        [ExcelFunction("Returns the expected return on an asset according to the capital asset pricing model")]
        public static double CAPM(
            [ExcelArgument("is the risk-free rate of interest.", Name = "risk_free_rate")] double rfr,
            [ExcelArgument("is the sensitivity of expected excess asset returns to the expected " +
                "excess market returns. See =Beta().", Name = "beta")] double b,
            [ExcelArgument("is the expected return of the market portfolio.",
                Name = "expected_market_return")] double er_m)
        {
            return rfr + b * (er_m - rfr);
        }

        [ExcelFunction("Returns the relation of an asset's returns with market returns. (1 = perfect correlation, " +
            "0 = no correlation, -1 = perfect negative correlation.)")]
        public static double Beta(
            [ExcelArgument("is an array representing the asset's returns.", Name = "asset_returns")] double[] r_a,
            [ExcelArgument("is an array representing market returns.",Name="market_returns")] double[] r_m
            )
        {
            return Statistics.Covar(r_a, r_m) / Statistics.Var(r_m);
        }
        #endregion
    }
}
