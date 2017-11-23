using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net;
using System.IO;
using System.Text.RegularExpressions;
using Newtonsoft.Json;

namespace EliteQuant
{
    public sealed class Broker
    {
        public static object[] GetRealTimeQuoteYahoo(string source, string sym)
        {
            string fullurl = _yahooLiveUrl + sym;

            object[] result = new object[_yahooLivePattern.Count];
            HttpWebRequest request;

            request = (HttpWebRequest)WebRequest.Create(fullurl);
            request.Method = "GET";
            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            string fullText = new StreamReader(response.GetResponseStream()).ReadToEnd();

            int i = 0;
            foreach (var kvp in _yahooLivePattern)
            {
                Regex regExp = new Regex((string)kvp.Value, RegexOptions.Multiline | RegexOptions.IgnoreCase);
                MatchCollection matches;
                matches = regExp.Matches(fullText);

                string direction = "><span id=\"yfs_c(?:6[034]|10)[^>]+><img.+alt=\"(\\w(?<!\\d)[\\w'-]*)\">";
                string dirtxt = "";
                bool addsign = false;
                if ((string.Compare(kvp.Key, "Change") == 0) || (string.Compare(kvp.Key, "Change%") == 0))
                    addsign = true;

                if (addsign)
                {
                    Regex regExp2 = new Regex(direction, RegexOptions.Multiline | RegexOptions.IgnoreCase);
                    MatchCollection matches2 = regExp2.Matches(fullText);

                    if (string.Compare(matches2[0].Groups[1].Value.ToString().ToUpper(), "DOWN") == 0)
                    {
                        dirtxt = "-";
                    }
                }

                if (matches.Count > 0)
                {
                    decimal temp;
                    if (decimal.TryParse(matches[0].Groups[1].Value, out temp))
                    {
                        result[i] = (addsign ? dirtxt : "") + temp;
                    }
                    else
                    {
                        result[i] = (addsign ? dirtxt : "") + matches[0].Groups[1].Value.ToString();
                    }
                }
                else
                {
                    result[i] = "";
                }

                i++;
            }

            return result;
        }

        public static object[] GetRealTimeQuoteMarketWatch(string source, string sym)
        {
            string fullurl = _marketWatchLiveUrl + sym;

            object[] result = new object[_yahooLivePattern.Count];
            HttpWebRequest request;

            request = (HttpWebRequest)WebRequest.Create(fullurl);
            request.Method = "GET";
            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            string fullText = new StreamReader(response.GetResponseStream()).ReadToEnd();

            int idx = 0, idx1 = 0, idx2 = 0;
            int i = 0;
            int idx_day_open = -1;
            foreach (var kvp in _marketWatchLivePattern)
            {
                switch (i)
                {
                    case 0:
                    case 1:
                        idx = fullText.IndexOf(kvp.Value);
                        if (idx != -1)
                        {
                            idx1 = fullText.IndexOf("=", idx);
                            idx1 = fullText.IndexOf("=", idx1 + 1);         // second =
                            idx2 = fullText.IndexOf(">", idx1);
                            string v = fullText.Substring(idx1 + 2, idx2 - idx1 - 3);
                            result[i] = double.Parse(v);
                        }
                        else
                        {
                            result[i] = "";
                        }
                        break;
                    case 2:
                        idx = fullText.IndexOf(kvp.Value);
                        if (idx != -1)
                        {
                            idx1 = fullText.IndexOf("=", idx);
                            idx1 = fullText.IndexOf("=", idx1 + 1);         // second =
                            idx2 = fullText.IndexOf(">", idx1);
                            string v = fullText.Substring(idx1 + 2, idx2 - idx1 - 4);
                            result[i] = double.Parse(v) / 100.0;
                        }
                        else
                        {
                            result[i] = "";
                        }
                        break;
                    case 3:
                    case 4:
                        result[i] = result[0];
                        break;
                    case 5:
                        idx = fullText.IndexOf(kvp.Value);
                        if (idx != -1)
                        {
                            idx1 = fullText.IndexOf("=", idx);
                            idx1 = fullText.IndexOf("=", idx1 + 1);         // second =
                            idx2 = fullText.IndexOf(" ", idx1);
                            idx_day_open = idx2;
                            string v = fullText.Substring(idx1 + 2, idx2 - idx1 - 3);
                            result[i] = double.Parse(v);
                        }
                        else
                        {
                            result[i] = "";
                        }
                        break;
                    case 6:
                        if (idx_day_open != -1)
                        {
                            idx1 = fullText.IndexOf("range-high=", idx_day_open);
                            idx2 = fullText.IndexOf(" ", idx1);
                            string v = fullText.Substring(idx1 + 12, idx2 - idx1 - 13);
                            result[i] = double.Parse(v);
                        }
                        else
                        {
                            result[i] = "";
                        }
                        break;
                    case 7:
                        if (idx_day_open != -1)
                        {
                            idx1 = fullText.IndexOf("range-low=", idx_day_open);
                            idx2 = fullText.IndexOf(" ", idx1);
                            string v = fullText.Substring(idx1 + 11, idx2 - idx1 - 12);
                            result[i] = double.Parse(v);
                        }
                        else
                        {
                            result[i] = "";
                        }
                        break;
                    case 8:
                        idx = fullText.IndexOf(kvp.Value);
                        if (idx != -1)
                        {
                            idx1 = fullText.IndexOf(">", idx);
                            idx2 = fullText.IndexOf("</", idx1);
                            string v = fullText.Substring(idx1 + 2, idx2 - idx1 - 3);
                            v = v.Replace("\r", string.Empty);
                            v = v.Replace("\n", string.Empty);
                            v = v.Replace(" ", string.Empty);
                            result[i] = v;
                        }
                        else
                        {
                            result[i] = "";
                        }
                        break;
                    default:
                        result[i] = "";
                        break;
                }

                i++;
            }

            return result;
        }

        public static object[] GetRealTimeQuote(string source, string sym)
        {
            object[] result = new object[_yahooLivePattern.Count];

            string fullurl = _iexLiveUrl + sym + "/quote";
            // Create a request for the URL. 
            HttpWebRequest grequest = (HttpWebRequest)WebRequest.Create(fullurl);
            // If required by the server, set the credentials.
            grequest.Credentials = CredentialCache.DefaultCredentials;
            // Get the response.
            HttpWebResponse gresponse = (HttpWebResponse)grequest.GetResponse();
            // Display the status.
            // Console.WriteLine(((HttpWebResponse)gresponse).StatusDescription);
            // Get the stream containing content returned by the server.
            Stream gdatastream = gresponse.GetResponseStream();
            // Open the stream using a StreamReader for easy access.
            StreamReader greader = new StreamReader(gdatastream);

            // Read the content.
            string quotestr = greader.ReadToEnd();

            // Display the content.
            // Console.WriteLine(quotestr);
            quotestr = quotestr.Replace("//", "");
            quotestr = quotestr.Replace("null", "0");       // pe_ratio = 0 for etf
            // Clean up the streams and the response.
            greader.Close();
            gdatastream.Close();
            gresponse.Close();

            RealTimeDataIEX data = JsonConvert.DeserializeObject<RealTimeDataIEX>(quotestr);

            result[0] = data.latestPrice.ToString();
            result[1] = data.change.ToString();
            result[2] = data.changePercent.ToString();
            result[3] = result[0];
            result[4] = result[0];
            result[5] = data.open.ToString();
            result[6] = result[0];
            result[7] = result[0];
            result[8] = data.latestVolume.ToString();

            return result;
        }

        /// <summary>
        /// https://www.quandl.com/api/v3/datasets/wiki/aapl.csv?start_date=2003-01-01&end_date=2003-03-06&order=asc&api_key=ay68s2CUzKbVuy8GAqxj
        /// </summary>
        /// <param name="source"></param>
        /// <param name="sym"></param>
        /// <param name="startDate"></param>
        /// <param name="endDate"></param>
        /// <param name="period">m, w, y, or d/param>
        /// <returns></returns>
        public static object[,] GetHistoricalQuotes(string source, string sym, DateTime startDate, DateTime endDate, string period = "", bool descending = true, bool useQuandl=true)
        {
            string url;

            if (useQuandl)
            {
                url = string.Format("https://www.quandl.com/api/v3/datasets/wiki/{0}.csv?start_date={1}&end_date={2}&order=des&api_key=ay68s2CUzKbVuy8GAqxj",
                    sym, startDate.ToString("yyyy-MM-dd"), endDate.ToString("yyyy-MM-dd"));
            }
            else
            {
                url = string.Format("http://ichart.finance.yahoo.com/table.csv?s={0}&d={1}&e={2}&f={3}&g={4}&a={5}&b={6}&c={7}&ignore=.csv",
                    sym, endDate.Month - 1, endDate.Day, endDate.Year, (period == "") ? "d" : period, startDate.Month - 1, startDate.Day, startDate.Year);
            }

            // read csv
            WebRequest request;
            HttpWebResponse response;
            
            request = (HttpWebRequest)WebRequest.Create(url);
            request.Method = "GET";
            response = (HttpWebResponse)request.GetResponse();

            List<string[]> sorted = new List<string[]>();
            object[,] ret;
            int counter = 0;

            try
            {
                using (StreamReader reader = new StreamReader(response.GetResponseStream()))
                {
                    string line;
                    string[] row;

                    while ((line = reader.ReadLine()) != null)
                    {

                        row = line.Split(',');
                        sorted.Add(row);

                        counter++;
                    }
                }
            }
            catch (Exception e)
            {
                return (new object[,]{ {  e.Message } });
            }

            bool hasHeaders = true;
            if (!descending)
            {
                sorted.Reverse(hasHeaders ? 1 : 0, sorted.Count - (hasHeaders ? 1 : 0));
            }

            ret = new object[sorted.Count, sorted[0].Length];

            for (int i = 0; i < sorted.Count; i++)
            {
                for (int j = 0; j < sorted[i].Length; j++)
                {
                    // copy header
                    if (hasHeaders && i == 0)
                    {
                        ret[i, j] = sorted[i][j].ToString();
                    }
                    else
                    {
                        try
                        {
                            if (j == 0)  // first column is date
                            {
                                DateTime date = DateTime.ParseExact(sorted[i][j], "yyyy-MM-dd", null);
                                ret[i, j] = date.ToOADate();
                            }
                            else
                            {
                                double tempVal;
                                double.TryParse(sorted[i][j], out tempVal);
                                ret[i, j] = tempVal;
                            }
                        }
                        catch (Exception)
                        {
                            // parsed[i, j] = "";
                        }
                    }
                }
            }  // end for

            return ret;
        }

        #region Yahoo
        private static string _yahooLiveUrl = "http://finance.yahoo.com/q?s=";
        private static Dictionary<string, string> _yahooLivePattern = new Dictionary<string, string>
        {
            { "Price", "yfs_l(?:10|84)_[^>]+>([0-9,.-]{1,})"},
            { "Change", "><span id=\"yfs_c(?:6[034]|10)[^>]+><img[^>]+>[\\s]+([0-9,.-]{1,})"},
            { "Change%", "yfs_p(?:4[034]|20)_[^>]+>\\(?([0-9,.-]{1,}%)\\)?"},
            //{ "Date", "<span id=\"yfs_market_time\">.*?, (.*?20[0-9]{2})"},
            //{ "Time", "<span id=\"yfs_t54_[^>]+>(.*?)<"},
            { "Bid", "yfs_b00_[^>]+>([0-9,.-]{1,})"},
            { "Ask", "yfs_a00_[^>]+>([0-9,.-]{1,})"},
            { "Open", "Open:<\\/th><td[^>]+>([0-9,.-]{1,})"},
            { "High", "yfs_h53_[^>]+>([0-9,.-]{1,})"},
            { "Low", "yfs_g53_[^>]+>([0-9,.-]{1,})"},
            { "Volume", "yfs_v53_[^>]+>([0-9,.-]{1,})"},
            //{ "MarketCap", "yfs_j10_[^>]+>([0-9,.-]{1,}[KMBT]?)"},
            //{ "Sign", "><span id=\"yfs_c(?:6[034]|10)[^>]+><img.+alt=\"(.*?\\w+)\">"}
        };

        private static string _marketWatchLiveUrl = "https://www.marketwatch.com/investing/stock/";
        private static Dictionary<string, string> _marketWatchLivePattern = new Dictionary<string, string>
            {
                { "Price", "<meta name=\"price\" content="},
                { "Change", "<meta name=\"priceChange\" content="},
                { "Change%", "<meta name=\"priceChangePercent\" content="},
                //{ "Date", "<span id=\"yfs_market_time\">.*?, (.*?20[0-9]{2})"},
                //{ "Time", "<span id=\"yfs_t54_[^>]+>(.*?)<"},
                { "Bid", "yfs_b00_[^>]+>([0-9,.-]{1,})"},
                { "Ask", "yfs_a00_[^>]+>([0-9,.-]{1,})"},
                { "Open", "<mw-rangeBar precision=\"2\" day-open="},
                { "High", "yfs_h53_[^>]+>([0-9,.-]{1,})"},
                { "Low", "yfs_g53_[^>]+>([0-9,.-]{1,})"},
                { "Volume", "<span class=\"volume last-value\">"},
                //{ "MarketCap", "yfs_j10_[^>]+>([0-9,.-]{1,}[KMBT]?)"},
                //{ "Sign", "><span id=\"yfs_c(?:6[034]|10)[^>]+><img.+alt=\"(.*?\\w+)\">"}
            };
        #endregion

        // query = @"https://finance.google.com/finance/info?q=" + syms[0];
        // var quote = EliteQuant.Utils.GetRealTimeData(query);
        #region Google finance
        public static List<RealTimeData> GetRealGOOGTimeData(string queries)
        {
            // Create a request for the URL. 
            HttpWebRequest grequest = (HttpWebRequest)WebRequest.Create(queries);
            // If required by the server, set the credentials.
            grequest.Credentials = CredentialCache.DefaultCredentials;
            // Get the response.
            HttpWebResponse gresponse = (HttpWebResponse)grequest.GetResponse();
            // Display the status.
            // Console.WriteLine(((HttpWebResponse)gresponse).StatusDescription);
            // Get the stream containing content returned by the server.
            Stream gdatastream = gresponse.GetResponseStream();
            // Open the stream using a StreamReader for easy access.
            StreamReader greader = new StreamReader(gdatastream);

            // Read the content.
            string quotestr = greader.ReadToEnd();

            // Display the content.
            // Console.WriteLine(quotestr);
            quotestr = quotestr.Replace("//", "");
            // Clean up the streams and the response.
            greader.Close();
            gdatastream.Close();
            gresponse.Close();

            return JsonConvert.DeserializeObject<List<RealTimeData>>(quotestr);
        }

        public class RealTimeData
        {
            public string id { get; set; }
            public string t { get; set; }
            public string e { get; set; }
            public string l { get; set; }
            public string l_fix { get; set; }
            public string l_cur { get; set; }
            public string s { get; set; }
            public string ltt { get; set; }
            public string lt { get; set; }
            public string lt_dts { get; set; }
            public string c { get; set; }
            public string c_fix { get; set; }
            public string cp { get; set; }
            public string cp_fix { get; set; }
            public string ccol { get; set; }
            public string pcls_fix { get; set; }
            public string el { get; set; }
            public string el_fix { get; set; }
            public string el_cur { get; set; }
            public string elt { get; set; }
            public string ec { get; set; }
            public string ec_fix { get; set; }
            public string ecp { get; set; }
            public string ecp_fix { get; set; }
            public string eccol { get; set; }
            public string div { get; set; }
            public string yld { get; set; }
        };

        private static string _iexLiveUrl = "https://api.iextrading.com/1.0/stock/";
        /// <summary>
        /// http://json2csharp.com/
        /// </summary>
        public class RealTimeDataIEX
        {
            public string symbol { get; set; }
            public string companyName { get; set; }
            public string primaryExchange { get; set; }
            public string sector { get; set; }
            public string calculationPrice { get; set; }
            public string open { get; set; }
            public string openTime { get; set; }
            public string close { get; set; }
            public string closeTime { get; set; }
            public string latestPrice { get; set; }
            public string latestSource { get; set; }
            public string latestTime { get; set; }
            public string latestUpdate { get; set; }
            public string latestVolume { get; set; }
            public string iexRealtimePrice { get; set; }
            public string iexRealtimeSize { get; set; }
            public string iexLastUpdated { get; set; }
            public string delayedPrice { get; set; }
            public string delayedPriceTime { get; set; }
            public string previousClose { get; set; }
            public string change { get; set; }
            public string changePercent { get; set; }
            public string iexMarketPercent { get; set; }
            public string iexVolume { get; set; }
            public string avgTotalVolume { get; set; }
            public string iexBidPrice { get; set; }
            public string iexBidSize { get; set; }
            public string iexAskPrice { get; set; }
            public string iexAskSize { get; set; }
            public string marketCap { get; set; }
            public string peRatio { get; set; }
            public string week52High { get; set; }
            public string week52Low { get; set; }
            public string ytdChange { get; set; }
        }
        #endregion
    }
}
