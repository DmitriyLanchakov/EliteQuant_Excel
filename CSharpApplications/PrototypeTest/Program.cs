using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.IO;
using System.Text.RegularExpressions;
using EliteQuant;

namespace PrototypeTest
{
    class Program
    {
        static void Main(string[] args)
        {
            object[] result = Broker.GetRealTimeQuote("", "aapl");
            Console.ReadKey();
        }

        static void Main2(string[] args)
        {
            Console.WriteLine("Hello World");

            string fullurl = "http://www.marketwatch.com/investing/stock/spy";

            Dictionary<string, string> _marketWatchLivePattern = new Dictionary<string, string>
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

            object[] result = new object[_marketWatchLivePattern.Count];

            HttpWebRequest request;
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            //ServicePoint mySp = ServicePointManager.FindServicePoint(fullurl);

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
                            result[i] = double.Parse(v)/100.0;
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
            Console.ReadKey();
        }
    }
}
