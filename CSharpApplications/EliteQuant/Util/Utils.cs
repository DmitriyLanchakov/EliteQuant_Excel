using System;
using System.Collections.Generic;
using System.Reflection;
using System.Linq;
using System.Text;

namespace EliteQuant
{
    public class Utils
    {
        /// <summary>
        /// get path
        /// </summary>
        static public string AssemblyLoadDirectory
        {
            get
            {
                string codeBase = Assembly.GetCallingAssembly().CodeBase;
                UriBuilder uri = new UriBuilder(codeBase);
                string path = Uri.UnescapeDataString(uri.Path);
                return System.IO.Path.GetDirectoryName(path);
            }
        }

        public static string ListEnvironmentVariables()
        {
            string machine = Environment.MachineName;
            string user = Environment.UserDomainName + "\\" + Environment.UserName;

            StringBuilder ret = new StringBuilder();
            foreach (string s in Environment.GetEnvironmentVariables().Keys)
            {
                ret.AppendLine(s + "=" + Environment.GetEnvironmentVariable(s).ToString());
            }
            return ret.ToString();
        }

        public static string GetFuturesSymbol(DateTime dt)
        {
            StringBuilder sb = new StringBuilder();
            switch (dt.Month)
            {
                case 1:
                    sb.Append("F");
                    break;
                case 2:
                    sb.Append("G");
                    break;
                case 3:
                    sb.Append("H");
                    break;
                case 4:
                    sb.Append("J");
                    break;
                case 5:
                    sb.Append("K");
                    break;
                case 6:
                    sb.Append("M");
                    break;
                case 7:
                    sb.Append("N");
                    break;
                case 8:
                    sb.Append("Q");
                    break;
                case 9:
                    sb.Append("U");
                    break;
                case 10:
                    sb.Append("V");
                    break;
                case 11:
                    sb.Append("X");
                    break;
                case 12:
                    sb.Append("Z");
                    break;
                default:
                    break;
            }

            return sb.Append(dt.ToString("yy")).ToString();
        }

        // assuming e.g. H15 expires at 03/01/2015
        public static DateTime DateFromFuturesSymbol(string xyy)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("01");
            switch (xyy.Substring(0,1))
            {
                case "F":
                    sb.Append("Jan");
                    break;
                case "G":
                    sb.Append("Feb");
                    break;
                case "H":
                    sb.Append("Mar");
                    break;
                case "J":
                    sb.Append("Apr");
                    break;
                case "K":
                    sb.Append("May");
                    break;
                case "M":
                    sb.Append("Jun");
                    break;
                case "N":
                    sb.Append("Jul");
                    break;
                case "Q":
                    sb.Append("Aug");
                    break;
                case "U":
                    sb.Append("Sep");
                    break;
                case "V":
                    sb.Append("Oct");
                    break;
                case "X":
                    sb.Append("Nov");
                    break;
                case "Z":
                    sb.Append("Dec");
                    break;
                default:
                    break;
            }

            return DateTime.ParseExact(sb.Append(xyy.Substring(1,2)).ToString(), "ddMMMyy", null);
        }
    }
}
