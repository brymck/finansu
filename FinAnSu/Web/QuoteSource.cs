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
using System.Text.RegularExpressions;
using System.Text;
using System.Windows.Forms;

namespace FinAnSu
{
    public static partial class Web
    {
        /// <summary>Contains information on an individual quote source.</summary>
        private class QuoteSource
        {
            private char abbreviation;
            private string url;
            private string forcePrefix;
            private string forceSuffix;
            private string matchPrefix;
            private string matchSuffix;
            private readonly Regex matchRegex = new Regex("[0-9]");
            private Dictionary<char, string> quoteParams = new Dictionary<char, string>();
            private List<char> historyParams = new List<char>();

            public QuoteSource(char abbreviation, string url,
                string forcePrefix, string forceSuffix, string matchPrefix, string matchSuffix,
                Dictionary<char, string> quoteParams, List<char> historyParams)
            {
                this.abbreviation = abbreviation;
                this.url = url;
                this.forcePrefix = forcePrefix;
                this.forceSuffix = forceSuffix;
                this.matchPrefix = matchPrefix;
                this.matchSuffix = matchSuffix;
                this.quoteParams = quoteParams;
                this.historyParams=historyParams;
            }

            public string[] GetQuoteParamsByChars(string paramAbbreviations)
            {
                List<string> query = new List<string>();
                char[] paramChars = paramAbbreviations.ToCharArray();

                foreach (char paramChar in paramChars)
                {
                    if (quoteParams.ContainsKey(paramChar))
                    {
                        query.Add(quoteParams[paramChar]);
                    }
                    else
                    {
                        query.Add("(0)");
                    }
                }

                return query.ToArray();
            }

            public string FullTicker(string ticker)
            {
                return FullTicker(ticker, false);
            }
            public string FullTicker(string ticker, bool forceInterpret)
            {
                bool hasMatch = (matchRegex.Match(ticker).Length > 0);
                if (ticker != "" && ticker.IndexOf(':') == -1 && ticker.IndexOf('.') == -1)
                {
                    return (hasMatch ? matchPrefix : (forceInterpret ? forcePrefix : "")) + ticker +
                           (hasMatch ? matchSuffix : (forceInterpret ? forceSuffix : ""));
                }
                else
                {
                    return ticker;
                }
            }

            public string FullURL(string ticker)
            {
                return FullURL(ticker, false);
            }
            public string FullURL(string ticker, bool forceInterpret)
            {
                return url + FullTicker(ticker, forceInterpret);
            }

            #region Properties
            public string URL
            {
                get { return url; }
                set { url = value; }
            }

            public char Abbreviation
            {
                get { return abbreviation; }
                set { abbreviation = value; }
            }

            public string ForcePrefix
            {
                get { return forcePrefix; }
                set { forcePrefix = value; }
            }
            public string ForceSuffix
            {
                get { return forceSuffix; }
                set { forceSuffix = value; }
            }
            public string MatchPrefix
            {
                get { return matchPrefix; }
                set { matchPrefix = value; }
            }
            public string MatchSuffix
            {
                get { return matchSuffix; }
                set { matchSuffix = value; }
            }
            #endregion
        }

        /// <summary>
        /// Contains information on all sources available for FinAnSu's quote package
        /// </summary>
        private static class QuoteSourceCollection
        {
            /// <summary>The names of parameters available for use in =Quote()</summary>
            private static Dictionary<char, string> quoteParamNames = new Dictionary<char, string>() {
                {'p', "Price"},
                {'x', "Change"},
                {'%', "% Change"},
                {'d', "Date"},
                {'t', "Time"},
                {'b', "Bid"},
                {'a', "Ask"},
                {'o', "Open"},
                {'h', "High"},
                {'l', "Low"},
                {'v', "Volume"},
                {'M', "Market Cap"},
                {'P', "P/E"}
            };

            /// <summary>The names of parameters available for use in =QuoteHistory()</summary>
            private static Dictionary<char, string> quoteHistoryParamNames = new Dictionary<char, string>() {
                {'d', "Date"},
                {'o', "Open"},
                {'h', "High"},
                {'l', "Low"},
                {'c', "Close"},
                {'v', "Volume"},
                {'a', "Adj Close"}
            };

            /// <summary>Information unique to Bloomberg as a quote source.</summary>
            private static QuoteSource bloomberg = new QuoteSource(
                'b',
                "http://www.bloomberg.com/apps/quote?ticker=",
                "", ":US", "", ":IND",
                new Dictionary<char, string> {
                    {'p', "(?:PRICE|VALUE)\\: <span class=\"amount\">([0-9.,NA-]{1,})"},
                    {'x', "Change</td>\n<td class=\"value[^>]+>([0-9.,NA-]{1,})"},
                    {'%', "Change</td>\n<td class=\"value[^>]+>[0-9.,NA-]{1,} \\(([0-9.,\\-NA]{1,})\\%"},
                    {'d', "\"date\">(.*?)<"},
                    {'t', "\"time\">(.*?)<"},
                    {'b', "Bid</td>\\n<td class=\"value[^>]+>([0-9.,NA-]{1,})"},
                    {'a', "Ask</td>\\n<td class=\"value[^>]+>([0-9.,NA-]{1,})"},
                    {'o', "Open</td>\\n<td class=\"value[^>]+>([0-9.,NA-]{1,})"},
                    {'h', "High</td>\\n<td class=\"value[^>]+>([0-9.,NA-]{1,})"},
                    {'l', "Low</td>\\n<td class=\"value[^>]+>([0-9.,NA-]{1,})"},
                    {'v', "Volume</td>\\n<td class=\"value[^>]+>([0-9.,NA-]{1,})"},
                    {'M', "Market Cap[^<]+</td>\\n<td class=\"value\">([0-9.,NA-]{1,})"},     // Market capitalization
                    {'P', "Price/Earnings[^<]+</td>\\n<td class=\"value\">([0-9.,NA-]{1,})"}, // P/E
                },
                null
            );

            /// <summary>Information unique to Google Finance as a quote source.</summary>
            private static QuoteSource google = new QuoteSource(
                'g',
                "http://www.google.com/ig/api?stock=",
                "NYSE:", "", "", "",
                new Dictionary<char, string> {
                    {'p', "<last data\\=\"([0-9.-]*)\""},
                    {'x', "<change data\\=\"([0-9.-]*)\""},
                    {'%', "<perc_change data\\=\"([0-9.-]*)\""},
                    {'d', "<trade_date_utc data\\=\"([0-9.-]*)\""},
                    {'t', "<trade_date_utc data\\=\"([0-9.-]*)\""},
                    {'o', "<open data\\=\"([0-9.-]*)\""},
                    {'h', "<high data\\=\"([0-9.-]*)\""},
                    {'l', "<low data\\=\"([0-9.-]*)\""},
                    {'v', "<volume data\\=\"([0-9.-]*)\""},
                    {'M', "<market_cap data\\=\"([0-9.-]*)\""},
                    {'P', "<volume data\\=\"([0-9.-]*)\""},
                    {'D', "<volume data\\=\"([0-9.-]*)\""},
                    {'Y', "<volume data\\=\"([0-9.-]*)\""},
                    {'E', "<volume data\\=\"([0-9.-]*)\""},
                    {'S', "<volume data\\=\"([0-9.-]*)\""},
                    {'B', "<volume data\\=\"([0-9.-]*)\""}
                },
                new List<char> { 'd', 'o', 'h', 'l', 'c', 'v'}
            );

            /// <summary>Information unique to Yahoo! Finance as a quote source.</summary>
            private static QuoteSource yahoo = new QuoteSource(
                'y',
                "http://finance.yahoo.com/q?s=",
                "", "", "", ".CME",
                new Dictionary<char, string> {
                    {'p', "(?:</a> |\"yfnc_tabledata1\"><big><b>)<span id=\"yfs_l[19]0_[^>]+>([0-9,.-]{1,})"},
                    {'x', "><span id=\"yfs_c[16]0_.*?(?:\n[\\s\\w\\-]*\n[\\s\\\">\\(]*?)?([0-9.,-]{1,})"},
                    {'%', "yfs_p[24]0_.*?(?:\n[\\s\\w\\-]*\n[\\s\\\">\\(]*?)?([0-9.,-]{1,})%\\)(?:</span>|</b></span></td>)"},
                    {'d', "<span id=\"yfs_market_time\">.*?, (.*?20[0-9]{1,2})"},
                    {'t', "(?:\"time\"|\"yfnc_tabledata1\")><span id=\"yfs_t[51]0_[^>]+>(.*?)<"},
                    {'b', "yfs_b00_[^>]+>([0-9,.-]{1,})"},
                    {'a', "yfs_a00_[^>]+>([0-9,.-]{1,})"},
                    {'o', "Open:</th><td class=\"yfnc_tabledata1\">([0-9,.-]{1,})"},
                    {'h', "yfs_h00_[^>]+>([0-9,.-]{1,})"},
                    {'l', "yfs_g00_[^>]+>([0-9,.-]{1,})"},
                    {'v', "yfs_v00_[^>]+>([0-9,.-]{1,})"},
                    {'M', "yfs_j10_[^>]+>([0-9,.-]{1,}[KMBT]?)"} // Market capitalization
                },
                new List<char> { 'd', 'o', 'h', 'l', 'c', 'v', 'a' }
            );

            /// <summary>Names and abbreviations for each source.</summary>
            private static Dictionary<string, char> abbreviations = new Dictionary<string, char>()
            {
                {"",          'b'},
                {"bloomberg", 'b'},
                {"bberg",     'b'},
                {"bb",        'b'},
                {"b",         'b'},
                {"google",    'g'},
                {"goog",      'g'},
                {"g",         'g'},
                {"yahoo!",    'y'},
                {"yahoo",     'y'},
                {"yhoo",      'y'},
                {"y!",        'y'},
                {"y",         'y'},
            };

            /// <summary>
            /// Returns a quote source based on the name provided.
            /// </summary>
            /// <param name="name">The name or abbreviation associated with a quote source.</param>
            /// <returns>The chosen source, defaulting to Bloomberg.</returns>
            public static QuoteSource GetSourceByName(string name)
            {
                name = name.Trim().ToLower();

                if (abbreviations.ContainsKey(name))
                {
                    switch (abbreviations[name])
                    {
                        case 'b':
                            return Bloomberg;
                        case 'g':
                            return Google;
                        case 'y':
                            return yahoo;
                    }
                }

                return bloomberg;
            }

            public static string QuoteParams(bool[] bools)
            {
                return ParseCharDict(ref quoteParamNames, bools);
            }
            public static string HistoryParams(bool[] bools)
            {
                return ParseCharDict(ref quoteHistoryParamNames, bools);
            }

            public static string[] GetQuoteNamesByChars(string paramAbbreviations)
            {
                List<string> query = new List<string>();
                char[] paramChars = paramAbbreviations.ToCharArray();

                foreach (char paramChar in paramChars)
                {
                    if (quoteParamNames.ContainsKey(paramChar))
                    {
                        query.Add(quoteParamNames[paramChar]);
                    }
                    else
                    {
                        query.Add("(0)");
                    }
                }

                return query.ToArray();
            }

            #region Helper functions
            private static string ParseCharDict(ref Dictionary<char, string> dict, bool[] bools)
            {
                int counter = 0;
                bool foundTrue = false;
                StringBuilder builder = new StringBuilder();

                foreach (char key in dict.Keys)
                {
                    if (foundTrue)
                    {
                        if (bools[counter++])
                        {
                            builder.Append(key);
                        }
                    }
                    else
                    {
                        if (bools[counter++])
                        {
                            builder.Clear();
                            foundTrue = true;
                        }
                        builder.Append(key);
                    }
                }
                return builder.ToString();
            }
            #endregion

            #region Properties
            /// <summary>Information unique to Bloomberg as a quote source.</summary>
            public static QuoteSource Bloomberg
            {
                get { return bloomberg; }
            }

            /// <summary>Information unique to Google Finance as a quote source.</summary>
            public static QuoteSource Google
            {
                get { return google; }
            }

            /// <summary>Information unique to Yahoo! Finance as a quote source.</summary>
            public static QuoteSource Yahoo
            {
                get { return yahoo; }
            }

            /// <summary>An array containing all sources.</summary>
            public static QuoteSource[] Sources
            {
                get { return new QuoteSource[] { bloomberg, google, yahoo }; }
            }
            #endregion
        }
    }
}
