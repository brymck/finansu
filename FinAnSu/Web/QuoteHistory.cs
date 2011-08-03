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
using System.Text;
using ExcelDna.Integration;

namespace FinAnSu
{
    public static partial class Web
    {
        private const int MAX_GLITCH_CHECKS = 5;
        private static double yahooGlitchDate = new DateTime(2011, 1, 28).ToOADate();
        private static Dictionary<char, string> quoteHistoryParams = new Dictionary<char, string>() {
            {'d', "Date"},
            {'o', "Open"},
            {'h', "High"},
            {'l', "Low"},
            {'c', "Close"},
            {'v', "Volume"},
            {'a', "Adj Close"}
        };

        #region Google Finance quote history
        [ExcelFunction("Returns the historical date, open, high, low, close and volume for a security ID from Google Finance.")]
        public static object[,] GoogleHistory(
            [ExcelArgument("is the security ID from Google Finance.", Name = "security_id")] string secId,
            [ExcelArgument("is the date from which to start retrieving history. Defaults to the most recent close.", Name = "start_date")] double dblStartDate,
            [ExcelArgument("is the date at which to stop retrieving history. Defaults to one year ago.", Name = "end_date")] double dblEndDate,
            [ExcelArgument("is a text flag representing whether you want daily (\"d\") or weekly (\"w\") quotes. " +
                "Defaults to daily.")] string period,
            [ExcelArgument("is a list of which values to return. Accepts any combination of \"dohlcva\" (date, open, high, low, close, volume, adj price). " +
                "Use =QuoteHistoryParams() for help if necessary. Defaults to all.", Name = "params")] string names,
            [ExcelArgument("is whether to display the headers for each column. Defaults to false.", Name = "show_headers")] bool showHeaders,
            [ExcelArgument("is whether to sort dates in ascending chronological order. Defaults to false.", Name = "date_order")] bool dateOrder)
        {
            DateTime startDate = (dblStartDate == 0) ? DateTime.Today.AddYears(-1) : DateTime.FromOADate(dblStartDate);
            DateTime endDate = (dblEndDate == 0) ? DateTime.Today : DateTime.FromOADate(dblEndDate);
            names = names.Replace('a', 'c');
            switch (ShortenDate(period))
            {
                case "w":
                    period = "weekly";
                    break;
                case "d":
                default:
                    period = "daily";
                    break;
            }
            string url = string.Format("http://www.google.com/finance/historical?q={0}&startdate={1}&enddate={2}&histperiod={3}&output=csv",
                secId, startDate.ToString("MMM+d,+yyyy"), endDate.ToString("MMM+d,+yyyy"), period);
            return QuoteHistoryParse(url, "d-MMM-yy", names, showHeaders, dateOrder, false);
        }
        #endregion

        #region Yahoo! Finance quote history
        [ExcelFunction("Returns the historical date, open, high, low, close, volume and adjusted price for a security ID from Yahoo! Finance.")]
        public static object[,] YahooHistory(
            [ExcelArgument("is the security ID from Yahoo! Finance.", Name = "security_id")] string secId,
            [ExcelArgument("is the date from which to start retrieving history. Defaults to the most recent close.", Name = "start_date")] double dblStartDate,
            [ExcelArgument("is the date at which to stop retrieving history. Defaults to one year ago.", Name = "end_date")] double dblEndDate,
            [ExcelArgument("is a text flag representing whether you want daily (\"d\"), weekly (\"w\"), monthly (\"m\") or yearly (\"y\") quotes. " +
                "Defaults to daily.")] string period,
            [ExcelArgument("is a list of which values to return. Accepts any combination of \"dohlcva\" (date, open, high, low, close, volume, adj price). " +
                "Use =QuoteHistoryParams() for help if necessary. Defaults to all.", Name = "params")] string names,
            [ExcelArgument("is whether to display the headers for each column. Defaults to false.", Name = "show_headers")] bool showHeaders,
            [ExcelArgument("is whether to sort dates in ascending chronological order. Defaults to false.", Name = "date_order")] bool dateOrder)
        {
            DateTime startDate = (dblStartDate == 0) ? DateTime.Today.AddYears(-1) : DateTime.FromOADate(dblStartDate);
            DateTime endDate = (dblEndDate == 0) ? DateTime.Today : DateTime.FromOADate(dblEndDate);
            switch (ShortenDate(period))
            {
                case "m":
                    period = "m";
                    break;
                case "w":
                    period = "w";
                    break;
                case "y":
                    period = "y";
                    break;
                case "d":
                default:
                    period = "d";
                    break;
            }
            string url = string.Format("http://ichart.finance.yahoo.com/table.csv?s={0}&d={1}&e={2}&f={3}&g={4}&a={5}&b={6}&c={7}&ignore=.csv",
                secId, endDate.Month - 1, endDate.Day, endDate.Year,
                (period == "") ? "d" : period, startDate.Month - 1, startDate.Day, startDate.Year);
            return QuoteHistoryParse(url, "yyyy-MM-dd", names, showHeaders, dateOrder, !endDate.Equals(yahooGlitchDate));
        }
        #endregion

        #region Generic quote history
        [ExcelFunction("Returns the historical date, open, high, low, close, volume and adjusted price for a security ID from the selected quotes provider.")]
        public static object[,] QuoteHistory(
            [ExcelArgument("is the security ID from the quote service.", Name = "security_id")] string secId,
            [ExcelArgument("is the name or abbreviation of the quote service (g, Google, y, Yahoo, etc.). Defaults to Yahoo!.")] string source,
            [ExcelArgument("is the date from which to start retrieving history. Defaults to the most recent close.", Name = "start_date")] double dblStartDate,
            [ExcelArgument("is the date at which to stop retrieving history. Defaults to one year ago.", Name = "end_date")] double dblEndDate,
            [ExcelArgument("is a text flag representing whether you want daily (\"d\"), weekly (\"w\"), monthly (\"m\") or yearly (\"y\") quotes. " +
                "Monthly and yearly quotes are available only through Yahoo!. Defaults to daily.")] string period,
            [ExcelArgument("is a list of which values to return. Accepts any combination of \"dohlcva\" (date, open, high, low, close, volume, adj price). " +
                "Adj price is available only through Yahoo!. Use =QuoteHistoryParams() for help if necessary. Defaults to all.", Name = "params")] string names,
            [ExcelArgument("is whether to display the headers for each column. Defaults to false.", Name = "show_headers")] bool showHeaders,
            [ExcelArgument("is whether to sort dates in ascending chronological order. Defaults to false.", Name = "date_order")] bool dateOrder)
        {
            switch (ShortenSource(source))
            {
                case "g":
                    return GoogleHistory(secId, dblStartDate, dblEndDate, period, names, showHeaders, dateOrder);
                case "y":
                default:
                    return YahooHistory(secId, dblStartDate, dblEndDate, period, names, showHeaders, dateOrder);
            }
        }
        #endregion

        #region Quote history helper functions
        [ExcelFunction("Builds a text string for use in =QuoteHistory() designating which values you would like returned.")]
        public static string QuoteHistoryParams(
            [ExcelArgument("is whether you wish to return the date.")] bool date,
            [ExcelArgument("is whether you wish to return the day's opening price.", Name = "open")] bool openPrice,
            [ExcelArgument("is whether you wish to return the day's high price.")] bool high,
            [ExcelArgument("is whether you wish to return the day's lowest price.")] bool low,
            [ExcelArgument("is whether you wish to return the day's closing price.", Name = "close")] bool closePrice,
            [ExcelArgument("is whether you wish to return the day's volume.")] bool volume,
            [ExcelArgument("is whether you wish to return the day's closing price.")] bool adjClose)
        {
            return QuoteSourceCollection.HistoryParams(new bool[] { date, openPrice, high, low, closePrice, volume, adjClose });
        }

        private static object[,] QuoteHistoryParse(string url, string dateFormat, string names, bool showHeaders, bool dateOrder, bool checkGlitch)
        {
            // Used for determining whether to start on the second row of the CSV file when parsing
            int headerOffset = showHeaders ? 1 : 0;

            if (names != "")
            {
                object[,] csvResults = ImportCSV(url, 0, dateOrder, new string[] { dateFormat, "double" }, true);
                int counter = 0;

                // Fix the super-mega-weird Yahoo! glitch that randomly fails to return data after January 28, 2011
                // by requesting the same CSV again
                if (checkGlitch)
                {
                    while (counter++ < MAX_GLITCH_CHECKS && csvResults[headerOffset, 0].Equals(yahooGlitchDate))
                    {
                        csvResults = ImportCSV(url, 0, dateOrder, new string[] { dateFormat, "double" }, true);
                    }
                }

                // Fill out a list of headers so that we can easily find the text we're looking for and get the appropriate index
                Dictionary<string, int> headers = new Dictionary<string, int>();

                // Get height and width of CSV file
                int rowCount = csvResults.GetLength(0);
                int columnCount = csvResults.Length / rowCount;

                // Add all headers to a string array for storage
                for (int i = 0; i < columnCount; i++)
                {
                    headers.Add(csvResults[0, i].ToString(), i);
                }

                // Convert parameter names to an array of single characters
                char[] nameChars = names.ToLower().ToCharArray();
                object[,] results = new object[rowCount - (1 - headerOffset), nameChars.Length];
                int currentColumn = 0;
                foreach (char nameChar in nameChars)
                {
                    if (quoteHistoryParams.ContainsKey(nameChar) && headers.ContainsKey(quoteHistoryParams[nameChar]))
                    {
                        int matchColumn = headers[quoteHistoryParams[nameChar]];
                        for (int i = 1 - headerOffset; i < rowCount; i++)
                        {
                            results[i - (1 - headerOffset), currentColumn] = csvResults[i, matchColumn];
                        }
                    }
                    else
                    {
                        for (int i = 1 - headerOffset; i < rowCount; i++)
                        {
                            results[i - (1 - headerOffset), currentColumn] = 0;
                        }
                    }
                    currentColumn++;
                }

                return results;
            }

            return ImportCSV(url, 1 - headerOffset, dateOrder, new string[] { dateFormat, "double" }, showHeaders);
        }

        private static string ShortenDate(string dateText)
        {
            switch (dateText.ToLower().Replace(" ", ""))
            {
                case "d":
                case "day":
                case "dly":
                case "daily":
                    return "d";
                case "ww":
                case "wwed":
                case "wkwed":
                case "weeklywednesday":
                case "weekly(wednesday)":
                    return "ww";
                case "wt":
                case "wthurs":
                case "wkthurs":
                case "weeklythursday":
                case "weekly(thursday)":
                    return "wt";
                case "wf":
                case "wfri":
                case "wkfri":
                case "weeklyfriday":
                case "weekly(friday)":
                    return "wf";
                case "w":
                case "wk":
                case "wkly":
                case "week":
                case "weekly":
                    return "w";
                case "bw":
                case "bwk":
                case "bweek":
                case "biweekly":
                case "bi-weekly":
                    return "bw";
                case "m":
                case "mth":
                case "month":
                case "monthly":
                    return "m";
                case "y":
                case "yr":
                case "year":
                case "yearly":
                case "a":
                case "ann":
                case "annual":
                    return "y";
                default:
                    return "";
            }
        }
        #endregion
    }
}
