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
using System.IO;
using System.Linq;
using System.Net;
using System.Windows.Forms;
using ExcelDna.Integration;

namespace FinAnSu
{
    public static partial class Web
    {
        [ExcelFunction("Returns an array of values from a CSV.")]
        public static object[,] ImportCSV(
            [ExcelArgument("is the URL of the target CSV file.",Name="url")] string url,
            [ExcelArgument("is the first line of the CSV to begin parsing (starting with 0).",
                Name = "start_line")] int startLine,
            [ExcelArgument("is whether to reverse the results.", Name = "reverse")] bool reverse,
            [ExcelArgument("is an array of formats: use \"double\", \"string\" or a date format.",
                Name = "formats")] object[] formats,
            [ExcelArgument("is whether the CSV file contains headers for each column.")] bool hasHeaders)
        {
            CsvParseFormat formatter = new CsvParseFormat();
            WebRequest request;
            HttpWebResponse response;
            List<string[]> sorted = new List<string[]>();
            object[,] parsed;
            int counter = 0;

            foreach (object format in formats)
            {
                formatter.AddFormat(format.ToString());
            }

            request = (HttpWebRequest)WebRequest.Create(url);
            request.Method = "GET";
            response = (HttpWebResponse)request.GetResponse();

            try
            {
                using (StreamReader reader = new StreamReader(response.GetResponseStream()))
                {
                    string line;
                    string[] row;

                    while ((line = reader.ReadLine()) != null)
                    {
                        if (counter >= startLine)
                        {
                            row = line.Split(',');
                            sorted.Add(row);
                        }
                        counter++;
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }

            if (reverse)
            {
                sorted.Reverse(hasHeaders ? 1 : 0, sorted.Count - (hasHeaders ? 1 : 0));
            }

            parsed = new object[sorted.Count, sorted[0].Length];

            for (int i = 0; i < sorted.Count; i++)
            {
                for (int j = 0; j < sorted[i].Length; j++)
                {
                    // Don't bother parsing headers
                    if (hasHeaders && i == 0)
                    {
                        parsed[i, j] = sorted[i][j].ToString();
                    }
                    else
                    {
                        try
                        {
                            parsed[i, j] = formatter.Parse(j, sorted[i][j]);
                        }
                        catch (Exception)
                        {
                            // parsed[i, j] = "";
                        }

                    }

                }
            }

            return parsed;
        }

        private class CsvParseFormat
        {
            private List<string> _formats = new List<string>();
            public void AddFormat(string format)
            {
                _formats.Add(format);
            }
            public object Parse(int key, string unparsed)
            {
                string format;

                if (key < _formats.Count)
                {
                    format = _formats[key];
                }
                else
                {
                    format = _formats.Last();
                }

                switch (format)
                {
                    case "string":
                        return unparsed;
                    case "double":
                        double tempVal;
                        double.TryParse(unparsed, out tempVal);
                        return tempVal;
                    default:
                        DateTime date = DateTime.ParseExact(unparsed, format, null);
                        return date.ToOADate();
                }
            }
        }
    }
}