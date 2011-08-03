/* 
 * FinAnSu
 * http://code.google.com/p/finansu/
 * 
 * Copyright 2011, Bryan McKelvey
 * Licensed under the MIT license
 * http://www.opensource.org/licenses/mit-license.php
 */

using System;
using System.IO;
using System.Net;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using ExcelDna.Integration;

namespace FinAnSu
{
    public static partial class Web
    {
        private const string REQUEST_METHOD = "GET";
        public const char ROW_SEPARATOR = '|';
        public const char COLUMN_SEPARATOR = '¦';

        #region GetWebData
        /// <summary>
        /// Returns an array of values based on a URL and regular expression.
        /// </summary>
        /// <param name="url">The full URL of the target website.</param>
        /// <param name="patterns">Regular expression patterns where the first backreference (in parentheses) is the value you wish to retrieve.</param>
        /// <param name="maxLength">The maximum length of the results array.</param>
        /// <returns>An array of values based on a URL and regular expression.</returns>
        [ExcelFunction("Returns an array of values based on a URL and regular expression.")]
        public static object[,] GetWebData(
            [ExcelArgument("is the full URL of the target website.", Name = "url")] string url,
            [ExcelArgument("is a regular expression pattern or patterns where the first backreference (in parentheses)" +
                "is the value you wish to retrieve", Name = "regex_patterns")] object[] patterns,
            [ExcelArgument("is the maximum length of the results array.", Name = "max_length")] int maxLength,
            [ExcelArgument("is a range or array of headers you wish to display.")] object[] headers)
        {
            if (maxLength == 0) { maxLength = 1; }
            if (headers == null) headers = new object[] { };
            HttpWebRequest request;
            bool hasHeaders = (headers.Length > 0);
            int height = maxLength + (hasHeaders ? 1 : 0);
            int width = patterns.Length;
            object[,] result = new object[height, width];

            // Get full source code from target URL
            request = (HttpWebRequest)WebRequest.Create(url);
            request.Method = REQUEST_METHOD;
            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            string fullText = new StreamReader(response.GetResponseStream()).ReadToEnd();

            for (int col = 0; col < width; col++)
            {
                // Execute regular expression
                Regex regExp = new Regex((string)patterns[col], RegexOptions.Multiline | RegexOptions.IgnoreCase);
                MatchCollection matches;
                matches = regExp.Matches(fullText);

                for (int row = 0; row < height; row++)
                {
                    if (row == 0 && hasHeaders)
                    {
                        result[row, col] = (col < headers.Length ? headers[col] : "");
                    }
                    else if (row < matches.Count + (hasHeaders ? 1 : 0))
                    {
                        decimal temp;
                        if (decimal.TryParse(matches[row - (hasHeaders ? 1 : 0)].Groups[1].Value, out temp))
                        {
                            result[row, col] = temp;
                        }
                        else
                        {
                            result[row, col] = matches[row].Groups[1].Value.ToString();
                        }
                    }
                    else
                    {
                        result[row, col] = "";
                    }
                }
            }

            return result;
        }
        #endregion

        /// <summary>
        /// Parses a string returned from an RTD server into a rectangular array
        /// </summary>
        /// <param name="rtdText">The raw string from the RTD server</param>
        /// <returns></returns>
        private static object[,] ParseRtd(string rtdText)
        {
            string[] rows = rtdText.Split(ROW_SEPARATOR);
            int height = rows.Length;
            int width = rows[0].Split(COLUMN_SEPARATOR).Length;

            object[,] values = new object[height, width];
            for (int row = 0; row < height; row++)
            {
                string[] cols = rows[row].Split(COLUMN_SEPARATOR);
                for (int col = 0; col < width; col++)
                {
                    double tempValue;
                    if (double.TryParse(cols[col].ToString(), out tempValue))
                    {
                        values[row, col] = tempValue;
                    }
                    else
                    {
                        values[row, col] = cols[col];
                    }
                }
            }
            return values;
        }
    }
}