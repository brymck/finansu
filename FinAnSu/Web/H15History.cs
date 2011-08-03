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
    public static partial class Web
    {
        [ExcelFunction("Returns information from the Fed's H.15 Statistical Release.")]
        public static object[,] H15History(
            [ExcelArgument("is the instrument ID. Go to http://www.federalreserve.gov/releases/h15/data.htm " +
                "and click on one of the data links. The ID is in the URL in the form H15_[id].txt.",
                Name = "instrument_id")] string instrument_id,
            [ExcelArgument("is business day (b), daily (d), weekly Wednesday (ww), weekly Thursday (wt), " +
                "weekly Friday (wf), bi-weekly (bw), monthly (m) or annual (a). Not all frequencies are available " +
                "for all instruments. Default is business day.", Name = "frequency")] string frequency)
        {
            string dateFormat;
            string url;

            frequency = FedFreq(frequency);
            url = string.Format("http://www.federalreserve.gov/releases/h15/data/{0}/H15_{1}.txt",
                FedFreq(frequency), instrument_id.ToUpper());
            switch (frequency)
            {
                case "Monthly":
                    dateFormat = "MM/yyyy";
                    break;
                case "Annual":
                    dateFormat = "yyyy";
                    break;
                default:
                    dateFormat = "MM/dd/yyyy";
                    break;
            }

            return ImportCSV(url, 11, true, new string[] { dateFormat, "double" }, false);
        }

        private static string FedFreq(string unparsed)
        {
            switch (ShortenDate(unparsed))
            {
                case "d":
                    return "Daily";
                case "ww":
                    return "Weekly_Wednesday_";
                case "wt":
                    return "Weekly_Thursday_";
                case "wf":
                    return "Weekly_Friday_";
                case "bw":
                    return "Bi-Weekly_AWednesday_";
                case "m":
                    return "Monthly";
                case "y":
                    return "Annual";
                default:
                    return "Business_day";
            }
        }
    }
}
