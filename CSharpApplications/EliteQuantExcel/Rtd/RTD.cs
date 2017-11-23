using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.Rtd;

namespace EliteQuantExcel
{
    public class Rtd
    {
        [ExcelFunction(Description = "RTD time", Category = "EliteQuantExcel - Time")]
        public static Object eqTimeNow(
            [ExcelArgument(Description = "(INT)interval in seconds ")]double interval)
        {
            string[] param = { interval.ToString(), "NOW" };
            object ret = XlCall.RTD("EliteQuantExcel.RTDSimpleTimerServer", null, param);
            return new object[,] { { ret } };
        }

        [ExcelFunction(Description = "RTD time", Category = "EliteQuantExcel - Time")]
        public static Object eqTimeNow2(
            [ExcelArgument(Description = "(INT)interval in seconds ")]double interval)
        {
            string[] param = { interval.ToString(), "NOW" };
            object ret = XlCall.RTD("EliteQuantExcel.RTDTimerServer", null, param);
            return new object[,] { { ret } };
        }

        [ExcelFunction(Description ="Live quote from Yahoo or Google.", Category ="EliteQuantExcel - Data")]
        public static object[,] eqDataRTDQuote(
            [ExcelArgument(Description ="Security/Ticker ID.", Name = "security_id")] string secId,
            [ExcelArgument(Description ="Source (GOOG or YHOO", Name = "source")] string source,
            [ExcelArgument(Description ="Refresh frequency in seconds. Defaults to 5 seconds.", Name = "frequency")] double freq)
        {
            object objFreq = (object)freq;
            if (freq <= 0 || objFreq is ExcelMissing || objFreq is ExcelEmpty) freq = 15;


            List<string> rtdparam = new List<string>() { freq.ToString(), "RealTimeQuote", source, secId};

            object ret = XlCall.RTD("EliteQuantExcel.RTDSimpleTimerServer", null, rtdparam.ToArray());
            string retstr = (string)ret;

            if (retstr != null)
            {
                string[] retarray = retstr.Split(',');

                object[,] result = new object[1, retarray.Length];

                for (int i = 0; i < retarray.Length; i++)
                {
                    result[0, i] = retarray[i];
                }

                return result;
            }
            else
            {
                return null;
            }
        }
    }
}
