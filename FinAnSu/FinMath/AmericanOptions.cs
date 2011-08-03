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
using ExcelDna.Integration;

namespace FinAnSu
{
    public static partial class Options
    {
        [ExcelFunction("Returns the Barone-Adesi-Whaley approximation for an American option.")]
        public static double American(
            [ExcelArgument("is whether the instrument is a call (c) or a put (p).",
                Name = "call_put_flag")] string callPutFlag,
            [ExcelArgument("is the current value of the underlying stock.", Name = "stock_price")] double S,
            [ExcelArgument("is the option's strike price.", Name = "strike_price")] double X,
            [ExcelArgument("is the time to maturity in years.", Name = "time_to_expiry")] double T,
            [ExcelArgument("is the risk-free rate through expiry.", Name = "risk-free_rate")] double r,
            [ExcelArgument("is the annual dividend yield.", Name = "dividend_yield")] double b,
            [ExcelArgument("is the implied volatility at expiry.", Name = "volatility")] double v)
        {
            double s;
            if (IsCall(callPutFlag))
            {
                s = 1;
            }
            else
            {
                s = -1;
            }

            if ((s == 1) && (b >= r))
            {
                return GBlackScholes(callPutFlag, S, X, T, r, b, v);
            }
            else
            {
                double Sk = CritComdtyPrice(callPutFlag, X, T, r, b, v);
                double N = 2 * b / (v * v);
                double k = 2 * r / (v * v * (1 - Math.Exp(-r * T)));
                double d1 = (Math.Log(Sk / X) + (b + v * v / 2) * T) / (v * Math.Sqrt(T));
                double Q = (-(N - 1) + s * Math.Sqrt(Math.Pow(N - 1, 2) + 4 * k)) / 2;
                double a = s * (Sk / Q) * (1 - Math.Exp((b - r) * T) * Statistics.CND(s * d1));

                if (s * S < s * Sk)
                {
                    return GBlackScholes(callPutFlag, S, X, T, r, b, v) + a * Math.Pow(S / Sk, Q);
                }
                else
                {
                    return s * (S - X);
                }
            }
        }

        [ExcelFunction("Returns the binomial valuation for an American option")]
        public static double AmericanBinomial(
            [ExcelArgument("is whether the instrument is a call (c) or a put (p).",
                Name = "call_put_flag")] string callPutFlag,
            [ExcelArgument("is the current value of the underlying stock.", Name = "stock_price")] double S,
            [ExcelArgument("is the option's strike price.", Name = "strike_price")] double K,
            [ExcelArgument("is the time to maturity in years.", Name = "time_to_expiry")] double T,
            [ExcelArgument("is the risk-free rate through expiry.", Name = "risk-free_rate")] double r,
            [ExcelArgument("is the annual dividend yield.", Name = "dividend_yield")] double q,
            [ExcelArgument("is the implied volatility at expiry.", Name = "volatility")] double v,
            [ExcelArgument("is the number of calculations performed to increase precision. The default is 500.",
                Name = "iterations")] int iter)
        {
            if (iter == 0) iter = 500;
            double delta_t = T / iter;
            double R = Math.Exp(r * delta_t);
            double Rinv = 1.0 / R;
            double u = Math.Exp(v * Math.Sqrt(delta_t));
            double uu = u * u;
            double d = 1.0 / u;
            double p_up = (Math.Exp((r - q) * (delta_t)) - d) / (u - d);
            double p_down = 1.0 - p_up;
            double[] prices = new double[iter + 1];
            double[] values = new double[iter + 1];
            prices[0] = S * Math.Pow(d, iter);  // fill in the endnodes.
            for (int i = 1; i <= iter; ++i) prices[i] = uu * prices[i - 1];
            if (callPutFlag == "p")
            {
                for (int i = 0; i <= iter; ++i) values[i] = Math.Max(0, (K - prices[i]));
            }
            else
            {
                for (int i = 0; i <= iter; ++i) values[i] = Math.Max(0, (prices[i] - K));
            }
            for (int step = iter - 1; step >= 0; --step)
            {
                for (int i = 0; i <= step; ++i)
                {
                    values[i] = (p_up * values[i + 1] + p_down * values[i]) * Rinv;
                    prices[i] = d * prices[i + 1];
                    if (callPutFlag == "p")
                    {
                        values[i] = Math.Max(values[i], K - prices[i]);
                    }
                    else
                    {
                        values[i] = Math.Max(values[i], prices[i] - K);
                    }
                }
            }
            return values[0];
        }

        private static double CritComdtyPrice(string callPutFlag, double X, double T, double r, double b, double v)
        {
            // Newton Raphson algorithm to solve for the critical commodity price for a call/put
            double s; // toggles positive for call and negative for put
            if (IsCall(callPutFlag))
            {
                s = 1;
            }
            else
            {
                s = -1;
            }
            double N = 2 * b / (v * v);
            double m = 2 * r / (v * v);
            double qu = (-(N - 1) + s * Math.Sqrt(Math.Pow(N - 1, 2) + 4 * m)) / 2;
            double su = X / (1 - 1 / qu);
            double h = -s * (b * T + s * 2 * v * Math.Sqrt(T)) * X / (s * (su - X));
            double Si = su + X * Math.Exp(h) - su * Math.Exp(h);

            double k = 2 * r / (v * v * (1 - Math.Exp(-r * T)));
            double d1 = (Math.Log(Si / X) + (b + v * v / 2) * T) / (v * Math.Sqrt(T));
            double Q = (-(N - 1) + s * Math.Sqrt(Math.Pow(N - 1, 2) + 4 * k)) / 2;
            double LHS = s * (Si - X);
            double RHS = GBlackScholes(callPutFlag, Si, X, T, r, b, v) + s *
                (1 - Math.Exp((b - r) * T) * Statistics.ND(s * d1)) * Si / Q;

            double bi = s * Math.Exp((b - r) * T) * Statistics.CND(s * d1) * (1 - 1 / Q) +
                s * (1 - s * Math.Exp((b - r) * T) * Statistics.CND(s * d1) / (v * Math.Sqrt(T))) / Q;
            double E = 0.000001;

            // Newton-Raphson algorithm for finding critical price Si
            while (Math.Abs(LHS - RHS) / X > E)
            {
                Si = (X + s * (RHS - bi * Si)) / (1 - s * bi);
                d1 = (Math.Log(Si / X) + (b + v * v / 2) * T) / (v * Math.Sqrt(T));
                LHS = s * (Si - X);
                RHS = GBlackScholes(callPutFlag, Si, X, T, r, b, v) + s *
                    (1 - Math.Exp((b - r) * T) * Statistics.CND(s * d1)) * Si / Q;
                bi = s * Math.Exp((b - r) * T) * Statistics.CND(s * d1) * (1 - 1 / Q) +
                    s * (1 - s * Math.Exp((b - r) * T) * Statistics.CND(s * d1) / (v * Math.Sqrt(T))) / Q;
            }

            return Si;
        }
    }
}
