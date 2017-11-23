using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelDna.Integration;
using ExcelDna.Integration.Rtd;
using Xl = Microsoft.Office.Interop.Excel;
using EliteQuant;

namespace EliteQuantExcel
{
    public class FreeMarketData
    {
        [ExcelFunction(Description = "Historical quotes from Yahoo or Google.", Category = "EliteQuantExcel - Data")]
        public static object[,] eqDataHistoricalQuotes(
            [ExcelArgument(Description = "Security/Ticker ID.", Name = "security_id")] string secId,
            [ExcelArgument("Start date, defaults to one year ago.", Name = "start_date")] double dblStartDate,
            [ExcelArgument("End date, defaults to today.", Name = "end_date")] double dblEndDate,
            [ExcelArgument("d, w, m, y. Defaults to d = daily.")] string period,
            [ExcelArgument("sort dates in ascending chronological order? Defaults to true.")] bool isDecending
            )
        {
            try
            {
                DateTime startDate = (dblStartDate == 0) ? DateTime.Today.AddYears(-1) : DateTime.FromOADate(dblStartDate);
                DateTime endDate = (dblEndDate == 0) ? DateTime.Today : DateTime.FromOADate(dblEndDate);

                return EliteQuant.Broker.GetHistoricalQuotes("YAHOO", secId, startDate, endDate, period, isDecending);
            }
            catch (Exception e)
            {
                return new object[,] { { e.Message} };
            }
            
        }
    }
}
