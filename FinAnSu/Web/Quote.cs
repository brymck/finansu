/* 
 * FinAnSu
 * http://code.google.com/p/finansu/
 * 
 * Copyright 2011, Bryan McKelvey
 * Licensed under the MIT license
 * http://www.opensource.org/licenses/mit-license.php
 */

using System;
using System.Text.RegularExpressions;
using ExcelDna.Integration;

namespace FinAnSu
{
    public static partial class Web
    {
        private const int ONE_RESULT_PER_PARAM = 1;

        #region Quote
        /// <summary>
        /// Returns the current quote for a security ID from Bloomberg, Google or Yahoo!.
        /// </summary>
        /// <param name="secId">The security ID from the quote service.</param>
        /// <param name="source">The name or abbreviation of the quote service (b, Bloomberg, g, Google, y, Yahoo, etc.).</param>
        /// <param name="names"></param>
        /// <param name="liveUpdating">Whether you want this function to return continuously stream live quotes to the cell.</param>
        /// <param name="freq">The number of seconds between update requests (if live_updating is true).</param>
        /// <returns></returns>
        [ExcelFunction("Returns the current quote for a security ID from Bloomberg, Google or Yahoo!.")]
        public static object[,] Quote(
            [ExcelArgument("is the security ID from the quote service.", Name = "security_id")] string secId,
            [ExcelArgument("is the name or abbreviation of the quote service (b, Bloomberg, g, Google, y, Yahoo, etc.). Defaults to Bloomberg.",
                Name = "source")] string source,
            [ExcelArgument("is a list of which values to return. Accepts any combination of \"px%dtbahlv\" (price, change, % change, date, time, bid, ask, high, low, volume). " +
                "Bid/ask are not available through Google. Use =QuoteParams() for help if necessary. Defaults to price.", Name = "params")] string names,
            [ExcelArgument("is whether you want this function to return continuously stream live quotes to the cell. Defaults to false.",
                Name = "live_updating")] bool liveUpdating,
            [ExcelArgument("is the number of seconds between update requests (if live_updating is true). " +
                "Defaults to 15 seconds.", Name = "frequency")] double freq,
            [ExcelArgument("is whether to display the headers for each column. Defaults to false.", Name = "show_headers")] bool showHeaders)
        {
            if (names == "") names = "p";
            QuoteSource src = QuoteSourceCollection.GetSourceByName(source);
            string url = src.FullURL(secId);
            string[] patterns = src.GetQuoteParamsByChars(names);
            string[] headers = (showHeaders ? QuoteSourceCollection.GetQuoteNamesByChars(names) : new string[] { });

            return (liveUpdating ? Import(url, patterns, ONE_RESULT_PER_PARAM, true, freq, headers) :
                                   GetWebData(url, patterns, ONE_RESULT_PER_PARAM, headers));
        }

        [ExcelFunction("Returns live quotes for a security ID from Bloomberg, Google or Yahoo!. Same as =Quote() " +
            "with the live_updating parameter set to true.")]
        public static object[,] LiveQuote(
            [ExcelArgument("is the security ID from the quote service.", Name = "security_id")] string secId,
            [ExcelArgument("is the name or abbreviation of the quote service (b, Bloomberg, g, Google, y, Yahoo, etc.). Defaults to Bloomberg.",
                Name = "source")] string source,
            [ExcelArgument("is a list of which values to return. Accepts any combination of \"px%dtbahlv\" (price, change, % change, date, time, bid, ask, high, low, volume). " +
                "Bid/ask are not available through Google. Use =QuoteParams() for help if necessary. Defaults to price.", Name = "params")] string names,
            [ExcelArgument("is the number of seconds between update requests (if live_updating is TRUE). " +
                "Defaults to 15 seconds.", Name = "frequency")] double freq,
            [ExcelArgument("is whether to display the headers for each column. Defaults to false.", Name = "show_headers")] bool showHeaders)

        {
            return Quote(secId, source, names, true, freq, showHeaders);
        }

        [ExcelFunction("Returns live quotes for a security ID from Bloomberg. Same as =Quote() " +
            "with the source parameter set to \"Bloomberg\".")]
        public static object[,] BloombergQuote(
            [ExcelArgument("is the security ID from Bloomberg.", Name = "security_id")] string secId,
            [ExcelArgument("is a list of which values to return. Accepts any combination of \"px%dtbahlv\" (price, change, % change, date, time, bid, ask, high, low, volume). " +
                "Bid/ask are not available through Google. Use =QuoteParams() for help if necessary. Defaults to price.", Name = "params")] string names,
            [ExcelArgument("is whether you want this function to return continuously stream live quotes to the cell. Defaults to false.",
                Name = "live_updating")] bool liveUpdating,
            [ExcelArgument("is the number of seconds between update requests (if live_updating is TRUE). " +
                "Defaults to 15 seconds.", Name = "frequency")] double freq,
            [ExcelArgument("is whether to display the headers for each column. Defaults to false.", Name = "show_headers")] bool showHeaders)
        {
            return Quote(secId, "b", names, liveUpdating, freq, showHeaders);
        }

        [ExcelFunction("Returns live quotes for a security ID from Google. Same as =Quote() " +
            "with the source parameter set to \"Google\".")]
        public static object[,] GoogleQuote(
            [ExcelArgument("is the security ID from Google.", Name = "security_id")] string secId,
            [ExcelArgument("is a list of which values to return. Accepts any combination of \"px%dtbahlv\" (price, change, % change, date, time, bid, ask, high, low, volume). " +
                "Bid/ask are not available through Google. Use =QuoteParams() for help if necessary. Defaults to price.", Name = "params")] string names,
            [ExcelArgument("is whether you want this function to return continuously stream live quotes to the cell. Defaults to false.",
                Name = "live_updating")] bool liveUpdating,
            [ExcelArgument("is the number of seconds between update requests (if live_updating is TRUE). " +
                "Defaults to 15 seconds.", Name = "frequency")] double freq,
            [ExcelArgument("is whether to display the headers for each column. Defaults to false.", Name = "show_headers")] bool showHeaders)
        {
            return Quote(secId, "g", names, liveUpdating, freq, showHeaders);
        }

        [ExcelFunction("Returns live quotes for a security ID from Yahoo!. Same as =Quote() " +
            "with the source parameter set to \"Yahoo\".")]
        public static object[,] YahooQuote(
            [ExcelArgument("is the security ID from Yahoo!.", Name = "security_id")] string secId,
            [ExcelArgument("is a list of which values to return. Accepts any combination of \"px%dtbahlv\" (price, change, % change, date, time, bid, ask, high, low, volume). " +
                "Bid/ask are not available through Google. Use =QuoteParams() for help if necessary. Defaults to price.", Name = "params")] string names,
            [ExcelArgument("is whether you want this function to return continuously stream live quotes to the cell. Defaults to false.",
                Name = "live_updating")] bool liveUpdating,
            [ExcelArgument("is the number of seconds between update requests (if live_updating is TRUE). " +
                "Defaults to 15 seconds.", Name = "frequency")] double freq,
            [ExcelArgument("is whether to display the headers for each column. Defaults to false.", Name = "show_headers")] bool showHeaders)
        {
            return Quote(secId, "y", names, liveUpdating, freq, showHeaders);
        }
        #endregion

        #region Quote Helper Functions
        [ExcelFunction("Builds a text string for use in =Quote() designating which values you would like returned.")]
        public static string QuoteParams(
            [ExcelArgument("is whether you wish to return the current price or value.")] bool price,
            [ExcelArgument("is whether you wish to return the day's change.")] bool change,
            [ExcelArgument("is whether you wish to return the day's percentage change.", Name = "pct_change")] bool pctChange,
            [ExcelArgument("is whether you wish to return the latest trade date.")] bool date,
            [ExcelArgument("is whether you wish to return the latest trade time.", Name = "close")] bool time,
            [ExcelArgument("is whether you wish to return the current bid price.")] bool bid,
            [ExcelArgument("is whether you wish to return the current ask price.")] bool ask,
            [ExcelArgument("is whether you wish to return the day's opening price.", Name = "open")] bool openPrice,
            [ExcelArgument("is whether you wish to return the day's high price.")] bool high,
            [ExcelArgument("is whether you wish to return the day's low price.")] bool low,
            [ExcelArgument("is whether you wish to return the day's closing price.")] bool volume,
            [ExcelArgument("is whether you wish to return the market capitalization.", Name = "market_cap")] bool marketCap,
            [ExcelArgument("is whether you wish to return the price/earnings ratio", Name = "p/e")] bool pe)
        {
            return QuoteSourceCollection.QuoteParams(new bool[] { price, change, pctChange, date, time,
                bid, ask, openPrice, high, low, volume,
                marketCap, pe
            });
        }

        [ExcelFunction("Returns FinAnSu's interpretation of an abbreviated security ID.")]
        public static string FullTicker(
            [ExcelArgument("is the security ID from the quote service.", Name = "security_id")] string ticker,
            [ExcelArgument("is the name or abbreviation of the quote service (b, Bloomberg, g, Google, y, Yahoo, etc.). Defaults to Bloomberg.",
                Name = "source")] string source,
            [ExcelArgument("forces FinAnSu to guess at a suffix if none exists if set to true (may result in errors).",
                Name = "force_interpret")] bool forceInterpret)
        {
            return QuoteSourceCollection.GetSourceByName(source).FullTicker(ticker, forceInterpret);
        }

        [ExcelFunction("Returns FinAnSu's intepretation of an abbreviated source name.")]
        public static string ShortenSource(
            [ExcelArgument("is the name or abbreviation of the quote service.", Name = "source")] string source)
        {
            return QuoteSourceCollection.GetSourceByName(source).Abbreviation.ToString();
        }
        #endregion
    }
}