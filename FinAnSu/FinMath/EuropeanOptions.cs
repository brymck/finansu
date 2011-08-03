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
        #region Black-Scholes
        [ExcelFunction("Returns the Black-76 valuation for options on futures and forwards.")]
        public static double Black76(
            [ExcelArgument("is whether the instrument is a call (c) or a put (p).",
                Name = "call_put_flag")] string callPutFlag,
            [ExcelArgument("is the current forward value.", Name = "forward")] double F,
            [ExcelArgument("is the option's strike price.", Name = "strike_price")] double X,
            [ExcelArgument("is the time to maturity in years.", Name = "time_to_expiry")] double T,
            [ExcelArgument("is the risk-free rate through expiry.", Name = "risk-free_rate")] double r,
            [ExcelArgument("is the implied volatility at expiry.", Name = "volatility")] double v)
        {
            double d1 = (Math.Log(F / X) + (v * v / 2) * T) / (v * Math.Sqrt(T));
            double d2 = d1 - v * Math.Sqrt(T);

            if (IsCall(callPutFlag))
            {
                return Math.Exp(-r * T) * (F * Statistics.CND(d1) - X * Statistics.CND(d2));
            }
            else
            {
                return Math.Exp(-r * T) * (X * Statistics.CND(-d2) - F * Statistics.CND(-d1));
            }
        }

        [ExcelFunction("Returns the Black-Scholes European call/put valuation.")]
        public static double BlackScholes(
            [ExcelArgument("is whether the instrument is a call (c) or a put (p).",
                Name = "call_put_flag")] string callPutFlag,
            [ExcelArgument("is the current value of the underlying stock.", Name = "stock_price")] double S,
            [ExcelArgument("is the option's strike price.", Name = "strike_price")] double K,
            [ExcelArgument("is the time to maturity in years.", Name = "time_to_expiry")] double T,
            [ExcelArgument("is the risk-free rate through expiry.", Name = "risk-free_rate")] double r,
            [ExcelArgument("is the annual dividend yield.", Name = "dividend_yield")] double q,
            [ExcelArgument("is the implied volatility at expiry.", Name = "volatility")] double v)
        {
            double d1 = (Math.Log(S / K) + (r - q + v * v / 2) * T) / (v * Math.Sqrt(T));
            double d2 = d1 - v * Math.Sqrt(T);

            if (IsCall(callPutFlag))
            {
                return S * Math.Exp(-q * T) * Statistics.CND(d1) - K * Math.Exp(-r * T) * Statistics.CND(d2);
            }
            else
            {
                return K * Math.Exp(-r * T) * Statistics.CND(-d2) - S * Math.Exp(-q * T) * Statistics.CND(-d1);
            }
        }

        [ExcelFunction("Returns the generalized Black-Scholes European call/put valuation.")]
        public static double GBlackScholes(
            [ExcelArgument("is whether the instrument is a call (c) or a put (p).",
                Name = "call_put_flag")] string callPutFlag,
            [ExcelArgument("is the current value of the underlying stock.", Name = "stock_price")] double S,
            [ExcelArgument("is the option's strike price.", Name = "strike_price")] double X,
            [ExcelArgument("is the time to maturity in years.", Name = "time_to_expiry")] double T,
            [ExcelArgument("is the risk-free rate through expiry.", Name = "risk-free_rate")] double r,
            [ExcelArgument("is the annualized cost of carry.", Name = "cost_of_carry")] double b,
            [ExcelArgument("is the implied volatility at expiry.", Name = "volatility")] double v)
        {
            double d1 = (Math.Log(S / X) + (b + v * v / 2) * T) / (v * Math.Sqrt(T));
            double d2 = d1 - v * Math.Sqrt(T);

            if (IsCall(callPutFlag))
            {
                return S * Math.Exp((b - r) * T) * Statistics.CND(d1) - X * Math.Exp(-r * T) * Statistics.CND(d2);
            }
            else
            {
                return X * Math.Exp(-r * T) * Statistics.CND(-d2) - S * Math.Exp((b - r) * T) * Statistics.CND(-d1);
            }
        }

        [ExcelFunction("Returns the Black-Scholes implied volatility using the Newton-Raphson method.")]
        public static double ImpliedVolatility(
            [ExcelArgument("is whether the instrument is a call (c) or a put (p).",
                Name = "call_put_flag")] string callPutFlag,
            [ExcelArgument("is the current value of the underlying stock.", Name = "stock_price")] double S,
            [ExcelArgument("is the option's strike price.", Name = "strike_price")] double K,
            [ExcelArgument("is the time to maturity in years.", Name = "time_to_expiry")] double T,
            [ExcelArgument("is the risk-free rate through expiry.", Name = "risk-free_rate")] double r,
            [ExcelArgument("is the annual dividend yield.", Name = "dividend_yield")] double q,
            [ExcelArgument("is the Black-Scholes European put/call valuation", Name = "price")] double cm)
        {
            const double epsilon = 0.00000001;
            double vi;
            double ci;
            double vegai;
            double minDiff;

            // Manaster and Koehler seed value
            vi = Math.Sqrt(Math.Abs(Math.Log(S / K) + r * T) * 2 / T);
            ci = BlackScholes(callPutFlag, S, K, T, r, q, vi);
            vegai = Vega(callPutFlag, S, K, T, r, q, vi);
            minDiff = Math.Abs(cm - ci);

            while (minDiff >= epsilon)
            {
                vi -= (ci - cm) / vegai;
                ci = BlackScholes(callPutFlag, S, K, T, r, q, vi);
                vegai = Vega(callPutFlag, S, K, T, r, q, vi);
                minDiff = Math.Abs(cm - ci);
            }

            if (minDiff < epsilon)
            {
                return vi;
            }
            else
            {
                return 0;
            }
        }
        #endregion

        #region First-Order Greeks
        [ExcelFunction("Returns the delta options Greek (sensitivity to changes in the underlying's price).")]
        public static double BSDelta(
            [ExcelArgument("is whether the instrument is a call (c) or a put (p).",
                Name = "call_put_flag")] string callPutFlag,
            [ExcelArgument("is the current value of the underlying stock.", Name = "stock_price")] double S,
            [ExcelArgument("is the option's strike price.", Name = "strike_price")] double K,
            [ExcelArgument("is the time to maturity in years.", Name = "time_to_expiry")] double T,
            [ExcelArgument("is the risk-free rate through expiry.", Name = "risk-free_rate")] double r,
            [ExcelArgument("is the annual dividend yield.", Name = "dividend_yield")] double q,
            [ExcelArgument("is the implied volatility at expiry.", Name = "volatility")] double v)
        {
            double d1 = (Math.Log(S / K) + (r - q + v * v / 2) * T) / (v * Math.Sqrt(T));

            if (IsCall(callPutFlag))
            {
                return Math.Exp(-q * T) * Statistics.CND(d1);
            }
            else
            {
                return -Math.Exp(-q * T) * Statistics.CND(-d1);
            }
        }

        [ExcelFunction("Returns the vega options Greek (sensitivity to changes in volatility).")]
        public static double Vega(
            [ExcelArgument("is whether the instrument is a call (c) or a put (p).",
                Name = "call_put_flag")] string callPutFlag,
            [ExcelArgument("is the current value of the underlying stock.", Name = "stock_price")] double S,
            [ExcelArgument("is the option's strike price.", Name = "strike_price")] double K,
            [ExcelArgument("is the time to maturity in years.", Name = "time_to_expiry")] double T,
            [ExcelArgument("is the risk-free rate through expiry.", Name = "risk-free_rate")] double r,
            [ExcelArgument("is the annual dividend yield.", Name = "dividend_yield")] double q,
            [ExcelArgument("is the implied volatility at expiry.", Name = "volatility")] double v)
        {
            double d1 = (Math.Log(S / K) + (r - q + v * v / 2) * T) / (v * Math.Sqrt(T));

            return S * Math.Exp(-q * T) * Statistics.PDF(d1) * Math.Sqrt(T);
        }

        [ExcelFunction("Returns the theta options Greek (sensitivity to the passage of time; time decay).")]
        public static double Theta(
            [ExcelArgument("is whether the instrument is a call (c) or a put (p).",
                Name = "call_put_flag")] string callPutFlag,
            [ExcelArgument("is the current value of the underlying stock.", Name = "stock_price")] double S,
            [ExcelArgument("is the option's strike price.", Name = "strike_price")] double K,
            [ExcelArgument("is the time to maturity in years.", Name = "time_to_expiry")] double T,
            [ExcelArgument("is the risk-free rate through expiry.", Name = "risk-free_rate")] double r,
            [ExcelArgument("is the annual dividend yield.", Name = "dividend_yield")] double q,
            [ExcelArgument("is the implied volatility at expiry.", Name = "volatility")] double v)
        {
            double d1 = (Math.Log(S / K) + (r - q + v * v / 2) * T) / (v * Math.Sqrt(T));
            double d2 = d1 - v * Math.Sqrt(T);

            if (IsCall(callPutFlag))
            {
                return -Math.Exp(-q * T) * (S * Statistics.PDF(d1) * v) / (2 * Math.Sqrt(T)) -
                    r * K * Math.Exp(-r * T) * Statistics.CND(d2) + q * S * Math.Exp(-q * T) * Statistics.CND(d1);
            }
            else
            {
                return -Math.Exp(-q * T) * (S * Statistics.PDF(d1) * v) / (2 * Math.Sqrt(T)) +
                    r * K * Math.Exp(-r * T) * Statistics.CND(-d2) - q * S * Math.Exp(-q * T) * Statistics.CND(-d1);
            }
        }

        [ExcelFunction("Returns the rho options Greek (sensitivity to the risk-free rate).")]
        public static double Rho(
            [ExcelArgument("is whether the instrument is a call (c) or a put (p).",
                Name = "call_put_flag")] string callPutFlag,
            [ExcelArgument("is the current value of the underlying stock.", Name = "stock_price")] double S,
            [ExcelArgument("is the option's strike price.", Name = "strike_price")] double K,
            [ExcelArgument("is the time to maturity in years.", Name = "time_to_expiry")] double T,
            [ExcelArgument("is the risk-free rate through expiry.", Name = "risk-free_rate")] double r,
            [ExcelArgument("is the annual dividend yield.", Name = "dividend_yield")] double q,
            [ExcelArgument("is the implied volatility at expiry.", Name = "volatility")] double v)
        {
            double d1 = (Math.Log(S / K) + (r - q + v * v / 2) * T) / (v * Math.Sqrt(T));
            double d2 = d1 - v * Math.Sqrt(T);

            if (IsCall(callPutFlag))
            {
                return K * T * Math.Exp(-r * T) * Statistics.CND(d2);
            }
            else
            {
                return -K * T * Math.Exp(-r * T) * Statistics.CND(-d2);
            }
        }
        #endregion

        #region Higher-Order Greeks
        [ExcelFunction("Returns the gamma options Greek (sensitivity to changes in delta; convexity).")]
        public static double Gamma(
            [ExcelArgument("is whether the instrument is a call (c) or a put (p).",
                Name = "call_put_flag")] string callPutFlag,
            [ExcelArgument("is the current value of the underlying stock.", Name = "stock_price")] double S,
            [ExcelArgument("is the option's strike price.", Name = "strike_price")] double K,
            [ExcelArgument("is the time to maturity in years.", Name = "time_to_expiry")] double T,
            [ExcelArgument("is the risk-free rate through expiry.", Name = "risk-free_rate")] double r,
            [ExcelArgument("is the annual dividend yield.", Name = "dividend_yield")] double q,
            [ExcelArgument("is the implied volatility at expiry.", Name = "volatility")] double v)
        {
            double d1 = (Math.Log(S / K) + (r - q + v * v / 2) * T) / (v * Math.Sqrt(T));

            return Math.Exp(-q * T) * Statistics.PDF(d1) / (S * v * Math.Sqrt(T));
        }

        [ExcelFunction("Returns the vanna options Greek (sensitivity of delta to changes in volatility).")]
        public static double Vanna(
            [ExcelArgument("is whether the instrument is a call (c) or a put (p).",
                Name = "call_put_flag")] string callPutFlag,
            [ExcelArgument("is the current value of the underlying stock.", Name = "stock_price")] double S,
            [ExcelArgument("is the option's strike price.", Name = "strike_price")] double K,
            [ExcelArgument("is the time to maturity in years.", Name = "time_to_expiry")] double T,
            [ExcelArgument("is the risk-free rate through expiry.", Name = "risk-free_rate")] double r,
            [ExcelArgument("is the annual dividend yield.", Name = "dividend_yield")] double q,
            [ExcelArgument("is the implied volatility at expiry.", Name = "volatility")] double v)
        {
            double d1 = (Math.Log(S / K) + (r - q + v * v / 2) * T) / (v * Math.Sqrt(T));
            double d2 = d1 - v * Math.Sqrt(T);

            return -Math.Exp(-q * T) * Statistics.PDF(d1) * d2 / v;
        }

        [ExcelFunction("Returns the charm options Greek (sensitivity of delta to the passage of time).")]
        public static double Charm(
            [ExcelArgument("is whether the instrument is a call (c) or a put (p).",
                Name = "call_put_flag")] string callPutFlag,
            [ExcelArgument("is the current value of the underlying stock.", Name = "stock_price")] double S,
            [ExcelArgument("is the option's strike price.", Name = "strike_price")] double K,
            [ExcelArgument("is the time to maturity in years.", Name = "time_to_expiry")] double T,
            [ExcelArgument("is the risk-free rate through expiry.", Name = "risk-free_rate")] double r,
            [ExcelArgument("is the annual dividend yield.", Name = "dividend_yield")] double q,
            [ExcelArgument("is the implied volatility at expiry.", Name = "volatility")] double v)
        {
            double d1 = (Math.Log(S / K) + (r - q + v * v / 2) * T) / (v * Math.Sqrt(T));
            double d2 = d1 - v * Math.Sqrt(T);

            if (IsCall(callPutFlag))
            {
                return -q * Math.Exp(-q * T) * Statistics.CND(d1) + Math.Exp(-q * T) * Statistics.PDF(d1) *
                    (2 * (r - q) * T - d2 * v * Math.Sqrt(T)) / (2 * T * v * Math.Sqrt(T));
            }
            else
            {
                return q * Math.Exp(-q * T) * Statistics.CND(-d1) + Math.Exp(-q * T) * Statistics.PDF(d1) *
                    (2 * (r - q) * T - d2 * v * Math.Sqrt(T)) / (2 * T * v * Math.Sqrt(T));
            }
        }

        [ExcelFunction("Returns the speed options Greek (sensitivity of gamma to changes in the underlying's price).")]
        public static double Speed(
            [ExcelArgument("is whether the instrument is a call (c) or a put (p).",
                Name = "call_put_flag")] string callPutFlag,
            [ExcelArgument("is the current value of the underlying stock.", Name = "stock_price")] double S,
            [ExcelArgument("is the option's strike price.", Name = "strike_price")] double K,
            [ExcelArgument("is the time to maturity in years.", Name = "time_to_expiry")] double T,
            [ExcelArgument("is the risk-free rate through expiry.", Name = "risk-free_rate")] double r,
            [ExcelArgument("is the annual dividend yield.", Name = "dividend_yield")] double q,
            [ExcelArgument("is the implied volatility at expiry.", Name = "volatility")] double v)
        {
            double d1 = (Math.Log(S / K) + (r - q + v * v / 2) * T) / (v * Math.Sqrt(T));

            return -Math.Exp(-q * T) * Statistics.PDF(d1) / (S * S * v * Math.Sqrt(T)) *
                (d1 / (v * Math.Sqrt(T)) + 1);
        }

        [ExcelFunction("Returns the zomma options Greek (sensitivity of gamma to changes in volatility).")]
        public static double Zomma(
            [ExcelArgument("is whether the instrument is a call (c) or a put (p).",
                Name = "call_put_flag")] string callPutFlag,
            [ExcelArgument("is the current value of the underlying stock.", Name = "stock_price")] double S,
            [ExcelArgument("is the option's strike price.", Name = "strike_price")] double K,
            [ExcelArgument("is the time to maturity in years.", Name = "time_to_expiry")] double T,
            [ExcelArgument("is the risk-free rate through expiry.", Name = "risk-free_rate")] double r,
            [ExcelArgument("is the annual dividend yield.", Name = "dividend_yield")] double q,
            [ExcelArgument("is the implied volatility at expiry.", Name = "volatility")] double v)
        {
            double d1 = (Math.Log(S / K) + (r - q + v * v / 2) * T) / (v * Math.Sqrt(T));
            double d2 = d1 - v * Math.Sqrt(T);

            return -Math.Exp(-q * T) * Statistics.PDF(d1) * (d1 * d2 - 1) / (S * v * v * Math.Sqrt(T));
        }

        [ExcelFunction("Returns the color options Greek (sensitivity of gamma to the passage of time).")]
        public static double Color(
            [ExcelArgument("is whether the instrument is a call (c) or a put (p).",
                Name = "call_put_flag")] string callPutFlag,
            [ExcelArgument("is the current value of the underlying stock.", Name = "stock_price")] double S,
            [ExcelArgument("is the option's strike price.", Name = "strike_price")] double K,
            [ExcelArgument("is the time to maturity in years.", Name = "time_to_expiry")] double T,
            [ExcelArgument("is the risk-free rate through expiry.", Name = "risk-free_rate")] double r,
            [ExcelArgument("is the annual dividend yield.", Name = "dividend_yield")] double q,
            [ExcelArgument("is the implied volatility at expiry.", Name = "volatility")] double v)
        {
            double d1 = (Math.Log(S / K) + (r - q + v * v / 2) * T) / (v * Math.Sqrt(T));
            double d2 = d1 - v * Math.Sqrt(T);

            return -Math.Exp(-q * T) * Statistics.PDF(d1) / (2 * S * T * v * Math.Sqrt(T)) *
                (2 * q * T + 1 + (2 * (r - q) * T - d2 * v * Math.Sqrt(T)) / (2 * T * v * Math.Sqrt(T)) * d1);
        }

        [ExcelFunction("Returns the DvegaDtime options Greek (sensitivity of vega to the passage of time).")]
        public static double DvegaDtime(
            [ExcelArgument("is whether the instrument is a call (c) or a put (p).",
                Name = "call_put_flag")] string callPutFlag,
            [ExcelArgument("is the current value of the underlying stock.", Name = "stock_price")] double S,
            [ExcelArgument("is the option's strike price.", Name = "strike_price")] double K,
            [ExcelArgument("is the time to maturity in years.", Name = "time_to_expiry")] double T,
            [ExcelArgument("is the risk-free rate through expiry.", Name = "risk-free_rate")] double r,
            [ExcelArgument("is the annual dividend yield.", Name = "dividend_yield")] double q,
            [ExcelArgument("is the implied volatility at expiry.", Name = "volatility")] double v)
        {
            double d1 = (Math.Log(S / K) + (r - q + v * v / 2) * T) / (v * Math.Sqrt(T));
            double d2 = d1 - v * Math.Sqrt(T);

            return S * Math.Exp(-q * T) * Statistics.PDF(d1) * Math.Sqrt(T) *
                (q + ((r - q) * d1) / (v * Math.Sqrt(T)) - (1 + d1 * d2) / (2 * T));
        }

        [ExcelFunction("Returns the vomma options Greek (sensitivity of vega to changes in volatility).")]
        public static double Vomma(
            [ExcelArgument("is whether the instrument is a call (c) or a put (p).",
                Name = "call_put_flag")] string callPutFlag,
            [ExcelArgument("is the current value of the underlying stock.", Name = "stock_price")] double S,
            [ExcelArgument("is the option's strike price.", Name = "strike_price")] double K,
            [ExcelArgument("is the time to maturity in years.", Name = "time_to_expiry")] double T,
            [ExcelArgument("is the risk-free rate through expiry.", Name = "risk-free_rate")] double r,
            [ExcelArgument("is the annual dividend yield.", Name = "dividend_yield")] double q,
            [ExcelArgument("is the implied volatility at expiry.", Name = "volatility")] double v)
        {
            double d1 = (Math.Log(S / K) + (r - q + v * v / 2) * T) / (v * Math.Sqrt(T));
            double d2 = d1 - v * Math.Sqrt(T);

            return S * Math.Exp(-q * T) * Statistics.PDF(d1) * Math.Sqrt(T) * d1 * d2 / v;
        }

        [ExcelFunction("Returns the dual delta options Greek (probability of finishing in-the-money).")]
        public static double DualDelta(
            [ExcelArgument("is whether the instrument is a call (c) or a put (p).",
                Name = "call_put_flag")] string callPutFlag,
            [ExcelArgument("is the current value of the underlying stock.", Name = "stock_price")] double S,
            [ExcelArgument("is the option's strike price.", Name = "strike_price")] double K,
            [ExcelArgument("is the time to maturity in years.", Name = "time_to_expiry")] double T,
            [ExcelArgument("is the risk-free rate through expiry.", Name = "risk-free_rate")] double r,
            [ExcelArgument("is the annual dividend yield.", Name = "dividend_yield")] double q,
            [ExcelArgument("is the implied volatility at expiry.", Name = "volatility")] double v)
        {
            double d1 = (Math.Log(S / K) + (r - q + v * v / 2) * T) / (v * Math.Sqrt(T));
            double d2 = d1 - v * Math.Sqrt(T);

            if (IsCall(callPutFlag))
            {
                return -Math.Exp(-q * T) * Statistics.CND(d2);
            }
            else
            {
                return Math.Exp(-q * T) * Statistics.CND(-d2);
            }
        }

        [ExcelFunction("Returns the dual gamma options Greek.")]
        public static double DualGamma(
            [ExcelArgument("is whether the instrument is a call (c) or a put (p).",
                Name = "call_put_flag")] string callPutFlag,
            [ExcelArgument("is the current value of the underlying stock.", Name = "stock_price")] double S,
            [ExcelArgument("is the option's strike price.", Name = "strike_price")] double K,
            [ExcelArgument("is the time to maturity in years.", Name = "time_to_expiry")] double T,
            [ExcelArgument("is the risk-free rate through expiry.", Name = "risk-free_rate")] double r,
            [ExcelArgument("is the annual dividend yield.", Name = "dividend_yield")] double q,
            [ExcelArgument("is the implied volatility at expiry.", Name = "volatility")] double v)
        {
            double d1 = (Math.Log(S / K) + (r - q + v * v / 2) * T) / (v * Math.Sqrt(T));
            double d2 = d1 - v * Math.Sqrt(T);

            return Math.Exp(-r * T) * Statistics.PDF(d2) / (K * v * Math.Sqrt(T));
        }
        #endregion

        #region Bermudan
        [ExcelFunction("Returns the binomial valuation for a Bermudan option")]
        public static double BermudanBinomial(
            [ExcelArgument("is whether the instrument is a call (c) or a put (p).",
                Name = "call_put_flag")] string callPutFlag,
            [ExcelArgument("is the current value of the underlying stock.", Name = "stock_price")] double S,
            [ExcelArgument("is the option's strike price.", Name = "strike_price")] double K,
            [ExcelArgument("is the time to maturity in years.", Name = "time_to_expiry")] double T,
            [ExcelArgument("is the risk-free rate through expiry.", Name = "risk-free_rate")] double r,
            [ExcelArgument("is the annual dividend yield.", Name = "dividend_yield")] double q,
            [ExcelArgument("is the implied volatility at expiry.", Name = "volatility")] double v,
            [ExcelArgument("is a range of potential exercise times in years.",
                Name = "potential_exercise_times")] double[] potential_exercise_times,
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
            List<int> potential_exercise_steps = new List<int>();
            for (int i = 0; i < potential_exercise_times.Length; ++i)
            {
                double t = potential_exercise_times[i];
                if (t > 0 && t < T)
                {
                    potential_exercise_steps.Add((int)(t / delta_t));
                }
            }
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
                bool check_exercise_this_step = false;
                for (int j = 0; j < potential_exercise_steps.Count; ++j)
                {
                    if (step == potential_exercise_steps[j]) { check_exercise_this_step = true; };
                }
                for (int i = 0; i <= step; ++i)
                {
                    values[i] = (p_up * values[i + 1] + p_down * values[i]) * Rinv;
                    prices[i] = d * prices[i + 1];
                    if (callPutFlag == "p")
                    {
                        if (check_exercise_this_step) values[i] = Math.Max(values[i], K - prices[i]);
                    }
                    else
                    {
                        if (check_exercise_this_step) values[i] = Math.Max(values[i], prices[i] - K);
                    }
                }
            }
            return values[0];
        }
        #endregion

        #region Helper functions
        private static bool IsCall(string callPutFlag)
        {
            switch (callPutFlag.ToLower())
            {
                case "c":
                case "call":
                    return true;
                case "p":
                case "put":
                    return false;
                default:
                    return true;
            }
        }
        #endregion
    }
}
