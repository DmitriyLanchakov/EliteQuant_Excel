using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EliteQuant.Curves
{
    public class IndexMapping
    {
        static public string ExtIndexName2EQName(string name)
        {
            string ret;
            switch (name)
            {
                case "IDX@USDOIS":
                    ret = "USDOISON Actual/360";
                    break;
                case "IDX@USDLIB1M":
                    ret = "USDLibor1M Actual/360";
                    break;
                case "IDX@USDLIB3M":
                    ret = "USDLibor3M Actual/360";
                    break;
                case "IDX@USDLIB6M":
                    ret = "USDLibor6M Actual/360";
                    break;
                case "IDX@USDLIB12M":
                    ret = "USDLibor12M Actual/360";
                    break;
                default:
                    ret = "";
                    break;
            }

            return ret;
        }

        static public string EQIndexName2ExtName(string name)
        {
            string ret;
            switch (name)
            {
                case "USDOISON Actual/360":
                    ret = "IDX@USDOIS";
                    break;
                case "USDLibor1M Actual/360":
                    ret = "IDX@USDLIB1M";
                    break;
                case "USDLibor3M Actual/360":
                    ret = "IDX@USDLIB3M";
                    break;
                case "USDLibor6M Actual/360":
                    ret = "IDX@USDLIB6M";
                    break;
                case "USDLibor12M Actual/360":
                    ret = "IDX@USDLIB12M";
                    break;
                default:
                    ret = "";
                    break;
            }

            return ret;
        }
    }
}
