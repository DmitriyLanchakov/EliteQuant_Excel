// Reserved names:(CRV@ or IDX@) USDOIS, USDLIB3M, USDLIB1M, USDLIB6M, _yyyyMMdd or _yyyyMMdd
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
    public class Curves
    {
        #region index fixing
        [ExcelFunction(Description = "set historical fixings", Category = "EliteQuantExcel - Indexes")]
        public static string eqIndexesSetHistory(
            [ExcelArgument(Description = "index name (USDOIS, USDLIB3M) ")] string name,
            [ExcelArgument(Description = "historical dates ")] object[] dates,
            [ExcelArgument(Description = "historical fixings ")]double[] quotes,
            [ExcelArgument(Description = "trigger ")]object trigger)
        {
            if (ExcelUtil.CallFromWizard())
                return "";

            string callerAddress = ExcelUtil.getActiveCellAddress();

            try
            {
                if (dates.Length != quotes.Length)
                {
                    return "size mismatch";
                }

                DateVector dt = new DateVector(dates.Length);
                DoubleVector qs = new DoubleVector(dates.Length);

                for (int i = 0; i < dates.Length; i++)
                {
                    dt.Add(new Date(Convert.ToInt32(dates[i])));
                    qs.Add(quotes[i]);
                }

                if (!name.Contains('@'))
                    name = "IDX@" + name;

                string name2 = EliteQuant.Curves.IndexMapping.ExtIndexName2EQName(name);

                RealTimeSeries fixings = new RealTimeSeries(dt, qs);
                IndexManager.instance().setHistory(name2, fixings);

                return name;
            }
            catch (Exception e)
            {
                ExcelUtil.logError(callerAddress, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), e.Message);
                return "#EQ_ERR!";
            }
        }

        [ExcelFunction(Description = "Display historical fixings", Category = "EliteQuantExcel - Indexes")]
        public static object eqIndexeGetHistory(
            [ExcelArgument(Description = "index name (USDOIS, USDLIB3M) ")] String name,
            [ExcelArgument(Description = "trigger ")]object trigger)
        {
            if (ExcelUtil.CallFromWizard())
                return "";

            string callerAddress = ExcelUtil.getActiveCellAddress();

            try
            {
                if (!name.Contains('@'))
                    name = "IDX@" + name;

                string name2;
                name2 = EliteQuant.Curves.IndexMapping.ExtIndexName2EQName(name);
                RealTimeSeries fixings = IndexManager.instance().getHistory(name2);

                double[,] ret = new double[fixings.size(), 2];

                for (int i = 0; i < ret.GetLength(0); i++)
                {
                    ret[i, 0] = fixings.dates()[i].serialNumber();
                    ret[i, 1] = fixings.values()[i];
                }

                return ret;
            }
            catch (Exception e)
            {
                ExcelUtil.logError(callerAddress, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), e.Message);
                return "#EQ_ERR!";
            }
        }

        [ExcelFunction(Description = "clear historical fixings", Category = "EliteQuantExcel - Indexes")]
        public static string eqIndexesClearHistory(
            [ExcelArgument(Description = "index name ")] String name,
            [ExcelArgument(Description = "trigger ")]object trigger)
        {
            if (ExcelUtil.CallFromWizard())
                return "";

            string callerAddress = ExcelUtil.getActiveCellAddress();

            try
            {
                if (!name.Contains('@'))
                    name = "IDX@" + name;

                string name2 = EliteQuant.Curves.IndexMapping.ExtIndexName2EQName(name);
                IndexManager.instance().clearHistory(name2);

                return name;
            }
            catch (Exception e)
            {
                ExcelUtil.logError(callerAddress, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), e.Message);
                return "#EQ_ERR!";
            }
        }

        [ExcelFunction(Description = "Load fixings from file ", Category = "EliteQuantExcel - Indexes")]
        public static object eqIndexesLoadFixingFromFile(
            [ExcelArgument(Description = "trigger ")]object trigger)
        {
            if (ExcelUtil.CallFromWizard())
                return "";

            string callerAddress = ExcelUtil.getActiveCellAddress();

            try
            {
                Xl.Range rng = ExcelUtil.getActiveCellRange();
                DateTime asofdate = DateTime.Today;

                string[] curves = new string[] { "USDOIS", "USDLIB3M", "USDLIB1M", "USDLIB6M", "USDLIB12M" };
                string path = EliteQuant.ConfigManager.Instance.IRRootDir + @"Curves\";
                string name;
                foreach (var s in curves)
                {
                    string[] data = System.IO.File.ReadAllLines(path + s + "_Fixing.csv");
                    DateVector dt = new DateVector(data.Length);
                    DoubleVector qs = new DoubleVector(data.Length);

                    foreach (var fixing in data)
                    {
                        if (fixing == "")
                            continue;

                        string[] ff = fixing.Split(',');
                        dt.Add(EliteQuant.EQConverter.DateTimeToDate(EliteQuant.EQConverter.StringToDateTime(ff[0])));
                        qs.Add(Convert.ToDouble(ff[1]));
                    }

                    name = s;
                    if (!s.Contains("@"))
                        name = "IDX@" + s;

                    string name2 = EliteQuant.Curves.IndexMapping.ExtIndexName2EQName(name);
                    RealTimeSeries fixings = new RealTimeSeries(dt, qs);
                    IndexManager.instance().setHistory(name2, fixings);
                }

                return asofdate;
            }
            catch (Exception e)
            {
                ExcelUtil.logError(callerAddress, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), e.Message);
                return "#EQ_ERR!";
            }
        }
        #endregion

        #region IR curves
        [ExcelFunction(Description = "Construct Flat foward IR curve", Category = "EliteQuantExcel - Curves")]
        public static string eqCurveIRFlatForward(
            [ExcelArgument(Description = "curve id ")] string ObjectId,
            [ExcelArgument(Description = "flat rate ")] double forward,
            [ExcelArgument(Description = "settlement days ")] int settlementdays,
            [ExcelArgument(Description = "calendar ")] string calendar,
            [ExcelArgument(Description = "day counter ")] string daycounter,
            [ExcelArgument(Description = "trigger ")]object trigger)
        {
            if (ExcelUtil.CallFromWizard())
                return "";

            string callerAddress = ExcelUtil.getActiveCellAddress();

            try
            {
                settlementdays = 0;
                calendar = "NULL";
                daycounter = "Actual365Fixed";

                YieldTermStructureHandle handle = new YieldTermStructureHandle(
                    new FlatForward(0, new NullCalendar(), forward, new Actual365Fixed()));

                // Store the option and return its id
                string id = "CRV@" + ObjectId;
                OHRepository.Instance.storeObject(id, handle, callerAddress);
                id += "#" + (String)DateTime.Now.ToString(@"HH:mm:ss");
                return id;
            }
            catch (Exception e)
            {
                ExcelUtil.logError(callerAddress, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), e.Message);
                return "#EQ_ERR!";
            }
        }

        [ExcelFunction(Description = "Construct IR discount zero curve", Category = "EliteQuantExcel - Curves")]
        public static string eqCurveIRDiscountCurve(
            [ExcelArgument(Description = "curve id ")] string ObjectId,
            [ExcelArgument(Description = "tenors ")] object[] dates,
            [ExcelArgument(Description = "discounts ")] double[] discounts,
            [ExcelArgument(Description = "calendar ")] string calendar,
            [ExcelArgument(Description = "day counter ")] string daycounter,
            [ExcelArgument(Description = "Interpolation Method (Linear, LogLinear) ")] string interp,
            [ExcelArgument(Description = "trigger ")]object trigger)
        {
            if (ExcelUtil.CallFromWizard())
                return "";

            string callerAddress = ExcelUtil.getActiveCellAddress();

            try
            {
                calendar = "NULL";
                daycounter = "Actual365Fixed";

                if (dates.Length != discounts.Length)
                {
                    return "size mismatch";
                }

                string interpmethod;
                if (ExcelUtil.isNull(interp))
                {
                    interpmethod = "LOGLINEAR";
                }
                else
                {
                    interpmethod = interp.ToUpper();
                }

                DateVector datesvector = new DateVector(dates.Length);
                DoubleVector discountsvector = new DoubleVector(dates.Length);

                for (int i = 0; i < dates.Length; i++)
                {
                    //datesvector.Add(Conversion.ConvertObject<Date>((DateTime)dates[i], "NA"));
                    datesvector.Add(new Date(Convert.ToInt32(dates[i])));
                    discountsvector.Add(discounts[i]);
                }

                YieldTermStructureHandle handle = null;
                
                if (interpmethod == "LINEAR")
                {
                    handle = new YieldTermStructureHandle(
                        new LinearDiscountCurve(datesvector, discountsvector, new Actual365Fixed(), new NullCalendar()));
                }
                else
                {
                    handle = new YieldTermStructureHandle(
                        new DiscountCurve(datesvector, discountsvector, new Actual365Fixed(), new NullCalendar()));
                } 

                // Store the option and return its id
                string id = "CRV@" + ObjectId;
                OHRepository.Instance.storeObject(id, handle, callerAddress);
                id += "#" + (String)DateTime.Now.ToString(@"HH:mm:ss");
                return id;
            }
            catch (Exception e)
            {
                ExcelUtil.logError(callerAddress, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), e.Message);
                return "#EQ_ERR!";
            }
        }

        // load discount curve
        // save discount curve
        // display discount curve

        [ExcelFunction(Description = "IR Index from IR curve ", Category = "EliteQuantExcel - Curves")]
        public static string eqCurveIRIndex(
            [ExcelArgument(Description = "index id (USDOIS, USDLIB3M, USDLIB1M ")] string ObjectId,
            [ExcelArgument(Description = "currency (USD, GBP, CAD, EUR, JPY ) ")] string Curncy,
            [ExcelArgument(Description = "id or name of discount curve ")] string discountId,
            [ExcelArgument(Description = "trigger ")]object trigger)
        {

            if (ExcelUtil.CallFromWizard())
                return "";

            string callerAddress = ExcelUtil.getActiveCellAddress();

            try
            {
                if (!discountId.Contains('@'))
                    discountId = "CRV@" + discountId;
                EliteQuant.YieldTermStructure curve = OHRepository.Instance.getObject<EliteQuant.YieldTermStructure>(discountId);
                YieldTermStructureHandle h = new YieldTermStructureHandle(curve);

                // market defaults
                IborIndex idx_default;
                if (ObjectId == "USDOIS")       //
                {
                    // Eonia and ois shares defaults
                    //idx_default = new Eonia(h);
                    //idx_default = new OvernightIndex("USDOIS", 0, new USDCurrency(), new TARGET(), new Actual360(), h);
                    idx_default = new FedFunds(h);
                }
                else
                {
                    Period tenor = null;
                    switch (ObjectId)
                    {
                        case "USDLIB3M":
                            tenor = new Period(3, TimeUnit.Months);
                            break;
                        case "USDLIB1M":
                            tenor = new Period(1, TimeUnit.Months);
                            break;
                        case "USDLIB6M":
                            tenor = new Period(6, TimeUnit.Months);
                            break;
                        case "USDLIB12M":
                            tenor = new Period(12, TimeUnit.Months);
                            break;
                    }
                    idx_default = new USDLibor(tenor, h);
                }

                // Store the option and return its id
                string id = "IDX@" + ObjectId;
                OHRepository.Instance.storeObject(id, idx_default, callerAddress);
                id += "#" + (String)DateTime.Now.ToString(@"HH:mm:ss");
                return id;
            }
            catch (Exception e)
            {
                ExcelUtil.logError(callerAddress, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), e.Message);
                return "#EQ_ERR!";
            }
        }

        // TO FINISH
        [ExcelFunction(Description = "IR Index ", Category = "EliteQuantExcel - Curves")]
        public static string eqCurveIRIndex2(
            [ExcelArgument(Description = "index id (USDOIS, USDLIB3M, USDLIB1M ")] string ObjectId,
            [ExcelArgument(Description = "currency (USD, GBP, CAD, EUR, JPY ) ")] string Curncy,
            [ExcelArgument(Description = "tenors ")] object[] dates,
            [ExcelArgument(Description = "discounts ")] double[] discounts,
            [ExcelArgument(Description = "calendar ")] string calendar,
            [ExcelArgument(Description = "day counter ")] string daycounter,
            [ExcelArgument(Description = "id of discount curve ")] string discountId,
            [ExcelArgument(Description = "Interpolation Method (Linear, LogLinear) ")] string interp,
            [ExcelArgument(Description = "trigger ")]object trigger)
        {
            if (ExcelUtil.CallFromWizard())
                return "";

            string callerAddress = ExcelUtil.getActiveCellAddress();

            try
            {
                calendar = "NULL";
                daycounter = "Actual365Fixed";

                if (dates.Length != discounts.Length)
                {
                    return "size mismatch";
                }

                string interpmethod;
                if (ExcelUtil.isNull(interp))
                {
                    interpmethod = "LOGLINEAR";
                }
                else
                {
                    interpmethod = interp.ToUpper();
                }

                DateVector datesvector = new DateVector(dates.Length);
                DoubleVector discountsvector = new DoubleVector(dates.Length);

                for (int i = 0; i < dates.Length; i++)
                {
                    //datesvector.Add(Conversion.ConvertObject<Date>((DateTime)dates[i], "NA"));
                    datesvector.Add(new Date(Convert.ToInt32(dates[i])));
                    discountsvector.Add(discounts[i]);
                }

                YieldTermStructureHandle handle = null;
                if (interpmethod == "LINEAR")
                {
                    handle = new YieldTermStructureHandle(
                        new LinearDiscountCurve(datesvector, discountsvector, new Actual365Fixed(), new NullCalendar()));
                }
                else
                {
                    handle = new YieldTermStructureHandle(
                        new DiscountCurve(datesvector, discountsvector, new Actual365Fixed(), new NullCalendar()));
                } 

                // Store the option and return its id
                string id = "CRV@" + ObjectId;
                OHRepository.Instance.storeObject(id, handle, callerAddress);
                id += "#" + (String)DateTime.Now.ToString(@"HH:mm:ss");
                return id;
            }
            catch (Exception e)
            {
                ExcelUtil.logError(callerAddress, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), e.Message);
                return "#EQ_ERR!";
            }
        }

        [ExcelFunction(Description = "Get IR discount ", Category = "EliteQuantExcel - Curves")]
        public static object eqCurveIRGetDiscountFactor(
            [ExcelArgument(Description = "curve id ")] string ObjectId,
            [ExcelArgument(Description = "tenors ")] DateTime dt,
            [ExcelArgument(Description = "trigger ")]object trigger)
        {
            if (ExcelUtil.CallFromWizard())
                return "";

            string callerAddress = ExcelUtil.getActiveCellAddress();

            try
            {
                Xl.Range rng = ExcelUtil.getActiveCellRange();

                if (!ObjectId.Contains('@'))
                    ObjectId = "CRV@" + ObjectId;

                YieldTermStructureHandle curve = OHRepository.Instance.getObject<YieldTermStructureHandle>(ObjectId);
                Date refDate = curve.referenceDate();

                Date dt2 = EliteQuant.EQConverter.ConvertObject<Date>(dt);
                double discount = curve.discount(dt2);

                return discount;
            }
            catch (Exception e)
            {
                ExcelUtil.logError(callerAddress, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), e.Message);
                return "#EQ_ERR!";
            }
        }

        [ExcelFunction(Description = "Display discount curve", Category = "EliteQuantExcel - Rates")]
        public static object eqIRCurveDisplayDiscountCurve(
            [ExcelArgument(Description = "id of discount curve ")] string ObjectId,
            [ExcelArgument(Description = "trigger ")]object trigger)
        {
            if (ExcelUtil.CallFromWizard())
                return "";

            string callerAddress = "";
            callerAddress = ExcelUtil.getActiveCellAddress();

            try
            {
                Xl.Range rng = ExcelUtil.getActiveCellRange();
                if (!ObjectId.Contains('@'))
                    ObjectId = "CRV@" + ObjectId;

                EliteQuant.YieldTermStructure curve = OHRepository.Instance.getObject<EliteQuant.YieldTermStructure>(ObjectId);
                if (curve.GetType() == typeof(EliteQuant.DiscountCurve))
                {
                    object[,] ret = new object[(curve as DiscountCurve).dates().Count, 2];

                    for (int i = 0; i < (curve as DiscountCurve).dates().Count; i++)
                    {
                        ret[i, 0] = (curve as DiscountCurve).dates()[i].serialNumber();
                        ret[i, 1] = (curve as DiscountCurve).discounts()[i].ToString();
                    }

                    return ret;
                }
                else if (curve.GetType() == typeof(EliteQuant.LinearDiscountCurve))
                {
                    object[,] ret = new object[(curve as LinearDiscountCurve).dates().Count, 2];

                    for (int i = 0; i < (curve as LinearDiscountCurve).dates().Count; i++)
                    {
                        ret[i, 0] = (curve as LinearDiscountCurve).dates()[i].serialNumber();
                        ret[i, 1] = (curve as LinearDiscountCurve).discounts()[i].ToString();
                    }

                    return ret;
                }
                else
                    return 0;
            }
            catch (Exception e)
            {
                ExcelUtil.logError(callerAddress, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), e.Message);
                return e.Message;
            }
        }

        [ExcelFunction(Description = "Save discount curve to disk ", Category = "EliteQuantExcel - Rates")]
        public static object eqIRCurveSaveCurveToDisk(
            [ExcelArgument(Description = "asofdate ")]DateTime asofdate,
            [ExcelArgument(Description = "trigger ")]object trigger)
        {
            if (ExcelUtil.CallFromWizard())
                return "";

            string callerAddress = "";
            callerAddress = ExcelUtil.getActiveCellAddress();

            try
            {
                Xl.Range rng = ExcelUtil.getActiveCellRange();
                if (ExcelUtil.isNull(asofdate))
                    asofdate = DateTime.Today;

                string[] curves = new string[] { "USDOIS", "USDLIB3M", "USDLIB1M", "USDLIB6M", "USDLIB12M" };
                string path;
                foreach (var s in curves)
                {
                    string targetcurve;
                    if (s.Contains("@"))
                    {
                        targetcurve = s;
                    }
                    else
                    {
                        targetcurve = "CRV@" + s;
                    }

                    EliteQuant.YieldTermStructure curve = OHRepository.Instance.getObject<EliteQuant.YieldTermStructure>(targetcurve);
                    if (curve.GetType() == typeof(EliteQuant.DiscountCurve))
                    {
                        asofdate = EliteQuant.EQConverter.DateToDateTime((curve as DiscountCurve).dates()[0]);
                        string[] ret = new string[(curve as DiscountCurve).dates().Count];
                        for (int i = 0; i < (curve as DiscountCurve).dates().Count; i++)
                        {
                            ret[i] = (curve as DiscountCurve).dates()[i].serialNumber() + "," + (curve as DiscountCurve).discounts()[i].ToString();
                        }

                        path = EliteQuant.ConfigManager.Instance.IRRootDir + @"Curves\"
                            + EliteQuant.EQConverter.DateTimeToString(EliteQuant.EQConverter.DateToDateTime((curve as DiscountCurve).dates()[0])) + "\\";
                        if (!System.IO.Directory.Exists(path))
                            System.IO.Directory.CreateDirectory(path);

                        System.IO.File.WriteAllLines(path + s + ".csv", ret);
                    }
                    else if (curve.GetType() == typeof(EliteQuant.LinearDiscountCurve))
                    {
                        asofdate = EliteQuant.EQConverter.DateToDateTime((curve as LinearDiscountCurve).dates()[0]);
                        string[] ret = new string[(curve as LinearDiscountCurve).dates().Count];
                        for (int i = 0; i < (curve as LinearDiscountCurve).dates().Count; i++)
                        {
                            ret[i] = (curve as LinearDiscountCurve).dates()[i].serialNumber() + "," + (curve as LinearDiscountCurve).discounts()[i].ToString();
                        }

                        path = EliteQuant.ConfigManager.Instance.IRRootDir + @"Curves\"
                            + EliteQuant.EQConverter.DateTimeToString(EliteQuant.EQConverter.DateToDateTime((curve as LinearDiscountCurve).dates()[0])) + "\\";
                        if (!System.IO.Directory.Exists(path))
                            System.IO.Directory.CreateDirectory(path);

                        System.IO.File.WriteAllLines(path + s + ".csv", ret);
                    }
                }

                return asofdate;
            }
            catch (Exception e)
            {
                ExcelUtil.logError(callerAddress, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), e.Message);
                return e.Message;
            }
        }

        [ExcelFunction(Description = "Load discount curve from disk ", Category = "EliteQuantExcel - Rates")]
        public static object eqIRCurveLoadCurveFromDisk(
            [ExcelArgument(Description = "asofdate ")]DateTime asofdate,
            [ExcelArgument(Description = "Interpolation Method (Linear, LogLinear) ")] string interp,
            [ExcelArgument(Description = "trigger ")]object trigger)
        {
            if (ExcelUtil.CallFromWizard())
                return "";

            string callerAddress = "";
            callerAddress = ExcelUtil.getActiveCellAddress();

            try
            {
                Xl.Range rng = ExcelUtil.getActiveCellRange();
                if (ExcelUtil.isNull(asofdate))
                    asofdate = DateTime.Today;
                
                string[] curves = new string[] { "USDOIS", "USDLIB3M", "USDLIB1M", "USDLIB6M", "USDLIB12M" };
                string path;
                EliteQuant.IborIndex idx = new EliteQuant.USDLibor(new EliteQuant.Period(3, EliteQuant.TimeUnit.Months));

                foreach (var s in curves)
                {
                    string targetcurve;
                    if (s.Contains("@"))
                    {
                        targetcurve = s;
                    }
                    else
                    {
                        targetcurve = "CRV@" + s;
                    }

                    string interpmethod;
                    if (ExcelUtil.isNull(interp))
                    {
                        interpmethod = "LOGLINEAR";
                    }
                    else
                    {
                        interpmethod = interp.ToUpper();
                    }

                    path = EliteQuant.ConfigManager.Instance.IRRootDir + @"Curves\"
                        + EliteQuant.EQConverter.DateTimeToString(asofdate) + "\\";

                    string[] crv = System.IO.File.ReadAllLines(path + s + ".csv");

                    EliteQuant.DateVector dtv = new EliteQuant.DateVector();
                    EliteQuant.DoubleVector discv = new EliteQuant.DoubleVector();
                    foreach (var c in crv)
                    {
                        string[] cc = c.Split(',');
                        Date dt = new Date(Convert.ToInt32(cc[0]));
                        dtv.Add(dt);
                        discv.Add(Convert.ToDouble(cc[1]));
                    }

                    YieldTermStructure yth = null;

                    if (interpmethod == "LINEAR")
                    {
                        yth = new EliteQuant.LinearDiscountCurve(dtv, discv, idx.dayCounter(), idx.fixingCalendar());
                    }
                    else
                    {
                        yth = new EliteQuant.DiscountCurve(dtv, discv, idx.dayCounter(), idx.fixingCalendar());
                    } 

                    string targetcrvid = "CRV@" + s;
                    OHRepository.Instance.storeObject(targetcrvid, yth, callerAddress);     // callerAddress or null ?
                }

                return asofdate;
            }
            catch (Exception e)
            {
                ExcelUtil.logError(callerAddress, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), e.Message);
                return e.Message;
            }
        }

        #endregion

        #region commodity curve
        [ExcelFunction(Description = "Commod forwards curve", Category = "EliteQuantExcel - Curves")]
        public static string eqCurveCommodForwardsCurve(
            [ExcelArgument(Description = "curve id ")] string ObjectId,
            [ExcelArgument(Description = "curve name (eg. commod ng exchange) ")] string curvename,
            [ExcelArgument(Description = "tenors ")] object[] dates,
            [ExcelArgument(Description = "quotes ")] double[] quotes,
            [ExcelArgument(Description = "calendar ")] string calendar,
            [ExcelArgument(Description = "day counter ")] string daycounter,
            [ExcelArgument(Description = "trigger ")]object trigger)
        {
            if (ExcelUtil.CallFromWizard())
                return "";

            string callerAddress = ExcelUtil.getActiveCellAddress();

            try
            {
                calendar = "NULL";
                daycounter = "Actual365Fixed";

                if (dates.Length != quotes.Length)
                {
                    return "size mismatch";
                }

                DateVector datesvector = new DateVector(dates.Length);
                DoubleVector quotesvector = new DoubleVector(dates.Length);

                for (int i = 0; i < dates.Length; i++)
                {
                    if ((ExcelUtil.isNull(dates[i])) || (quotes[i] == 0))
                        continue;

                    //datesvector.Add(Conversion.ConvertObject<Date>((DateTime)dates[i], "NA"));
                    datesvector.Add(new Date(Convert.ToInt32(dates[i])));
                    quotesvector.Add(quotes[i]);
                }

                CommodityCurveExt curve =
                    new CommodityCurveExt(curvename, datesvector, quotesvector, new NullCalendar(), new Actual365Fixed());
                
                // Store the curve and return its id
                string id = "CRV@" + ObjectId;
                OHRepository.Instance.storeObject(id, curve, callerAddress);
                id += "#" + (String)DateTime.Now.ToString(@"HH:mm:ss");
                return id;
            }
            catch (Exception e)
            {
                ExcelUtil.logError(callerAddress, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), e.Message);
                return "#EQ_ERR!";
            }
        }

        [ExcelFunction(Description = "Commod foward curve display", Category = "EliteQuantExcel - Curves")]
        public static object eqCurveCommodDisplayForwardsCurve(
            [ExcelArgument(Description = "curve id ")] string ObjectId,
            [ExcelArgument(Description = "trigger ")]object trigger)
        {
            if (ExcelUtil.CallFromWizard())
                return "";

            string callerAddress = ExcelUtil.getActiveCellAddress();

            try
            {
                if (!ObjectId.Contains('@'))
                    ObjectId = "CRV@" + ObjectId;

                CommodityCurveExt curve = OHRepository.Instance.getObject<CommodityCurveExt>(ObjectId);
                DateVector dts = curve.dates();
                DoubleVector pts = curve.prices();

                double[,] ret = new double[dts.Count, 2];
                for (int i = 0; i < dts.Count; i++ )
                {
                    ret[i, 0] = dts[i].serialNumber();
                    ret[i, 1] = pts[i];
                }

                return ret;
            }
            catch (Exception e)
            {
                ExcelUtil.logError(callerAddress, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), e.Message);
                return "#EQ_ERR!";
            }
        }

        [ExcelFunction(Description = "Commodity index ", Category = "EliteQuantExcel - Curves")]
        public static string eqCurveCommodIndex(
            [ExcelArgument(Description = "index id ")] string ObjectId,
            [ExcelArgument(Description = "index name ")] string indexname,
            [ExcelArgument(Description = "curve id ")] string curveId,
            [ExcelArgument(Description = "calendar")] string calendar,
            [ExcelArgument(Description = "trigger ")]object trigger)
        {
            if (ExcelUtil.CallFromWizard())
                return "";

            string callerAddress = ExcelUtil.getActiveCellAddress();

            try
            {
                calendar = "NULL";

                CommodityCurveExt curve = OHRepository.Instance.getObject<CommodityCurveExt>(curveId);

                CommodityIndexExt idx = new CommodityIndexExt(indexname, curve, new NullCalendar());

                // Store the index and return its id
                string id = "IDX@" + ObjectId;
                OHRepository.Instance.storeObject(id, idx, callerAddress);
                id += "#" + (String)DateTime.Now.ToString(@"HH:mm:ss");
                return id;
                
            }
            catch (Exception e)
            {
                ExcelUtil.logError(callerAddress, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), e.Message);
                return "#EQ_ERR!";
            }
        }

        [ExcelFunction(Description = "Commodity forward price from index ", Category = "EliteQuantExcel - Curves")]
        public static object eqCurveCommodForwardPrice(
            [ExcelArgument(Description = "index id ")] string ObjectId,
            [ExcelArgument(Description = "forward date ")] DateTime date,
            [ExcelArgument(Description = "trigger ")]object trigger)
        {
            if (ExcelUtil.CallFromWizard())
                return "";

            string callerAddress = ExcelUtil.getActiveCellAddress();

            try
            {
                if (!ObjectId.Contains('@'))
                    ObjectId = "IDX@" + ObjectId;

                CommodityIndexExt idx = OHRepository.Instance.getObject<CommodityIndexExt>(ObjectId);
                Date dt = EliteQuant.EQConverter.ConvertObject<Date>(date);

                // double price = idx.price(dt);
                double price = idx.forwardPrice(dt);

                return price;
            }
            catch (Exception e)
            {
                ExcelUtil.logError(callerAddress, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), e.Message);
                return "#EQ_ERR!";
            }
        }
        #endregion
    }
}
