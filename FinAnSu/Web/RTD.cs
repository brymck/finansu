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
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.Rtd;

namespace FinAnSu
{
    public static partial class Web
    {
        private const string RTD_NAME = "RTDServers.ImportServer";
        private const string RTD_SERVER = null;
        
        [ExcelFunction("Returns a horizontal array of values based on a URL and regular expression.")]
        public static object[,] Import(
            [ExcelArgument("is the full URL of the target website.", Name = "url")] string url,
            [ExcelArgument("is a regular expression pattern where the first backreference (in parentheses)" +
                "is the value you wish to retrieve", Name = "regex_pattern")] object[] patterns,
            [ExcelArgument("is the maximum length of the results array.", Name = "max_length")] int maxLength,
            [ExcelArgument("is whether you want this function to return continuously stream live quotes to the cell.",
                Name = "live_updating")] bool liveUpdating,
            [ExcelArgument("is the number of seconds between update requests (if live_updating is TRUE). " +
                "Defaults to 15 seconds.", Name = "frequency")] double freq,
            [ExcelArgument("is a range or array of headers you wish to display.")] object[] headers)
        {
            if (headers == null || (headers.Length == 1 && (headers[0] is ExcelMissing || headers[0] is ExcelEmpty)))
            {
                headers = new object[] { };
            }
            if (liveUpdating)
            {
                object objFreq = (object)freq;
                if (freq <= 0 || objFreq is ExcelMissing || objFreq is ExcelEmpty) freq = 15;

                // Build list of params to pass to RTD server, and XlCall will only accept strings
                List<string> rtdParams = new List<string>() { url, maxLength.ToString(), freq.ToString(), patterns.Length.ToString() };

                // Somewhat long-winded, but ExcelDna doesn't expose functions containing string[] arguments,
                // so we have to iterate through an object[]-typed list of patterns
                foreach (string pattern in patterns)
                {
                    rtdParams.Add(pattern);
                }
                foreach (string header in headers)
                {
                    rtdParams.Add(header);
                }

                return ParseRtd(XlCall.RTD(RTD_NAME, RTD_SERVER, rtdParams.ToArray()).ToString());
            }
            else
            {
                return GetWebData(url, patterns, maxLength, headers);
            }
        }
    }
}

namespace RTDServers
{
    public class ImportServer : IRtdServer
    {
        private IRTDUpdateEvent _callback;
        private Timer _timer;
        private Dictionary<int, Query> _queries = new Dictionary<int, Query>();

        public int ServerStart(IRTDUpdateEvent CallbackObject)
        {
            // Called when the first RTD topic is requested
            // Sets server to poll for updates every second
            _callback = CallbackObject;
            _timer = new Timer();
            _timer.Tick += Callback;
            _timer.Interval = 1000;
            return 1;
        }
        public void ServerTerminate()
        {
            // Kills timer and queries so they don't run after workbook closes
            _timer.Dispose();
            _queries = null;
        }
        public object ConnectData(int topicId, ref Array Strings, ref bool GetNewValues)
        {
            // Parses 4 parameters to add to query dictionary
            string url = (string)Strings.GetValue(0);
            int maxLength = int.Parse((string)Strings.GetValue(1));
            double freq = double.Parse((string)Strings.GetValue(2));
            int patternsLength = int.Parse((string)Strings.GetValue(3));
            List<string> patterns = new List<string>();
            List<string> headers = new List<string>();
            for (int i = 4; i < Strings.Length; i++)
            {
                if (i > 3 + patternsLength)
                {
                    headers.Add((string)Strings.GetValue(i));
                }
                else
                {
                    patterns.Add((string)Strings.GetValue(i));
                }
            }

            // Adds query, starts timer and returns #WAIT
            _queries.Add(topicId, new Query(url, patterns.ToArray(), maxLength, freq, headers.ToArray()));
            _timer.Start();
            return _queries[topicId].ToString();
        }

        public void DisconnectData(int topicId)
        {
            // Removes query on disconnect
            _queries.Remove(topicId);
        }

        public int Heartbeat()
        {
            // Called by Excel if a given interval has elapsed (returns true)
            return 1;
        }
        public Array RefreshData(ref int topicCount)
        {
            // Called when Excel requests refresh
            // Returns topic count and an array of IDs and values
            object[,] results = new object[2, _queries.Count];
            topicCount = 0;

            // Prevent overwriting by any Query delegates
            lock (_queries)
            {
                foreach (int topicID in _queries.Keys)
                {
                    // Only return results if they've been updated
                    if (_queries[topicID].Updated)
                    {
                        _queries[topicID].Updated = false;
                        results[0, topicCount] = topicID;
                        results[1, topicCount] = _queries[topicID].ToString();
                        topicCount++;
                    }
                }
            }

            _timer.Start();
            return results;
        }

        private void Callback(object sender, EventArgs e)
        {
            // Stops timer and tells all queries to run their async delegates
            _timer.Stop();
            lock (_queries)
            {
                foreach (KeyValuePair<int, Query> q in _queries)
                {
                    q.Value.AsyncImport();
                }
            }
            _callback.UpdateNotify();
        }
    }

    public class Query
    {
        private const string WAIT_TEXT = "#WAIT";
        private delegate void AsyncImporter();
        private AsyncImporter _importer;
        private string _url;
        private string[] _patterns;
        private string[] _headers;
        private int _maxLength;
        private double _freq;
        private DateTime _nextUpdate = DateTime.Now.AddDays(-1);
        private bool _updated;
        private object[,] _results;

        public Query(string url, string[] patterns, int maxLength, double freq, string[] headers)
        {
            _url = url;
            _patterns = patterns;
            _maxLength = maxLength;
            _freq = freq;
            _headers = headers;
            AsyncImport();
        }

        #region Properties
        public string URL
        {
            get { return _url; }
            set { _url = value; }
        }
        public string[] Patterns
        {
            get { return _patterns; }
            set { _patterns = value; }
        }
        public string[] Headers
        {
            get { return _headers; }
            set { _headers = value; }
        }
        public int MaxLength
        {
            get { return _maxLength; }
            set { _maxLength = value; }
        }
        public bool Updated
        {
            get { return _updated; }
            set { _updated = value; }
        }
        public DateTime NextUpdate
        {
            get { return _nextUpdate; }
            set { _nextUpdate = value; }
        }
        #endregion

        public object[,] Results
        {
            get { return _results; }
            set { _results = value; }
        }
        public void AsyncImport()
        {
            // Only runs async based on query frequency
            if (DateTime.Compare(DateTime.Now, _nextUpdate) > 0)
            {
                _nextUpdate = DateTime.Now.AddSeconds(_freq);
                _importer = new AsyncImporter(RunImportDelegate);
                _importer.BeginInvoke(null, null);
            }
        }
        private void RunImportDelegate()
        {
            _results = FinAnSu.Web.GetWebData(_url, _patterns, _maxLength, _headers);
            _updated = true;
        }
        public override string ToString()
        {
            if (_results == null)
            {
                return WAIT_TEXT;
            }
            else
            {
                int height = _results.GetLength(0);
                int width = _results.GetLength(1);
                StringBuilder parsed = new StringBuilder();

                bool firstRow = true;
                for (int row = 0; row < height; row++)
                {
                    bool firstCol = true;

                    // Only append separate row with semicolon if previous row exist (i.e. not first row)
                    if (firstRow)
                    {
                        firstRow = false;
                    }
                    else
                    {
                        parsed.Append(FinAnSu.Web.ROW_SEPARATOR);
                    }
                    for (int col = 0; col < width; col++)
                    {
                        // Only append separate row with comma if previous row exist (i.e. not first row)
                        if (firstCol)
                        {
                            firstCol = false;
                        }
                        else
                        {
                            parsed.Append(FinAnSu.Web.COLUMN_SEPARATOR);
                        }
                        parsed.Append(_results[row, col].ToString());
                    }
                }
                return parsed.ToString();
            }
        }
    }
}