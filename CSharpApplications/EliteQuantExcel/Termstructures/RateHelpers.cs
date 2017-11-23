using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelDna.Integration;
using ExcelDna.Integration.Rtd;
using Xl = Microsoft.Office.Interop.Excel;

namespace EliteQuantExcel
{
    
    public class RateHelpers
    {
        static EliteQuant.Calendar cal_gbp = new EliteQuant.UnitedKingdom(EliteQuant.UnitedKingdom.Market.Exchange);
        static EliteQuant.Calendar cal_usd = new EliteQuant.UnitedStates(EliteQuant.UnitedStates.Market.Settlement);
        static EliteQuant.JointCalendar cal_usd_gbp = new EliteQuant.JointCalendar(cal_gbp, cal_usd, EliteQuant.JointCalendarRule.JoinHolidays);
        static EliteQuant.DayCounter dc_act_360 = new EliteQuant.Actual360();
        static EliteQuant.DayCounter dc_30_360 = new EliteQuant.Thirty360();
        static EliteQuant.BusinessDayConvention bdc_usd = EliteQuant.BusinessDayConvention.ModifiedFollowing;
        static bool eom_usd = true;
        static int fixingDays_usd = 2;

        #region LIB3M
        [ExcelFunction(Description = "create deposit rate helper", Category = "EliteQuantExcel - Rates")]
        public static string eqIRCurveDepositRateHelper(
            [ExcelArgument(Description = "(String) id of rate helper object ")] String ObjectId,
            [ExcelArgument(Description = "(double) quote of deposit rate ")] double Quote,
            [ExcelArgument(Description = "(String) forward start month, e.g. 7D, 3M ")] String Tenor,
            [ExcelArgument(Description = "int fixingDays ")] int fixingDays,
            [ExcelArgument(Description = "trigger ")]object trigger)
        {
            if (ExcelUtil.CallFromWizard())
                return "";

            string callerAddress = ExcelUtil.getActiveCellAddress();

            try
            {
                // use default value
                // // "london stock exchange"; "Actual/360"; "fixingDays = 2", "MF", "eom = true"
                EliteQuant.IborIndex idx_usdlibor = new EliteQuant.USDLibor(new EliteQuant.Period(3, EliteQuant.TimeUnit.Months));
                if (ExcelUtil.isNull(fixingDays))
                {
                    fixingDays = (int)idx_usdlibor.fixingDays();
                }

                EliteQuant.QuoteHandle quote_ = new EliteQuant.QuoteHandle(new EliteQuant.SimpleQuote(Quote));
                EliteQuant.Period tenor_ = EliteQuant.EQConverter.ConvertObject<EliteQuant.Period>(Tenor);

                EliteQuant.RateHelper rh = new EliteQuant.DepositRateHelper(quote_, tenor_, (uint)fixingDays, cal_usd,
                    bdc_usd, eom_usd, dc_act_360);

                string id = "RHDEP@" + ObjectId;
                OHRepository.Instance.storeObject(id, rh, callerAddress);
                return id + "#" + DateTime.Now.ToString("HH:mm:ss");
            }
            catch (Exception e)
            {
                ExcelUtil.logError(callerAddress, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), e.Message);
                return "#EQ_ERR!";
            }
        }

        /// <summary>
        /// March contract fixes/expires at 3rd Wednesday, for the next 90days
        /// first four non main-cycle months
        /// </summary>
        [ExcelFunction(Description = "create future rate helper", Category = "EliteQuantExcel - Rates")]
        public static string eqIRCurveFuturesRateHelper(
            [ExcelArgument(Description = "(String) id of rate helper object ")] String ObjectId,
            [ExcelArgument(Description = "(double) quote of ED futures e.g. 99.5 ")] double price,
            [ExcelArgument(Description = "(double) convexity adjustment default 0 ")] double convadj,
            [ExcelArgument(Description = "order of ED futures start from 1 ")] int order,
            [ExcelArgument(Description = "trigger ")]object trigger)
        {
            if (ExcelUtil.CallFromWizard())
                return "";

            string callerAddress = ExcelUtil.getActiveCellAddress();

            try
            {
                // use default value
                EliteQuant.IborIndex idx = new EliteQuant.USDLibor(new EliteQuant.Period(3, EliteQuant.TimeUnit.Months));

                EliteQuant.Date today = EliteQuant.Settings.instance().getEvaluationDate();
                EliteQuant.Date settlementdate = idx.fixingCalendar().advance(today, (int)idx.fixingDays(), EliteQuant.TimeUnit.Days);

                EliteQuant.QuoteHandle quote_ = new EliteQuant.QuoteHandle(new EliteQuant.SimpleQuote(price));
                EliteQuant.QuoteHandle conv_ = new EliteQuant.QuoteHandle(new EliteQuant.SimpleQuote(convadj));

                EliteQuant.Date imm_startdate = EliteQuant.IMM.nextDate(settlementdate, false);
                for (int i = 0; i < order-1; i++)
                {
                    imm_startdate = EliteQuant.IMM.nextDate(cal_usd_gbp.advance(imm_startdate, 1, EliteQuant.TimeUnit.Days), false);
                }

                EliteQuant.Date enddate = imm_startdate + 90;

                EliteQuant.RateHelper rh = new EliteQuant.FuturesRateHelper(quote_, imm_startdate, enddate, dc_act_360, conv_);

                string id = "RHED@" + ObjectId;
                OHRepository.Instance.storeObject(id, rh, callerAddress);
                return id + "#" + DateTime.Now.ToString("HH:mm:ss");
            }
            catch (Exception e)
            {
                ExcelUtil.logError(callerAddress, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), e.Message);
                return "#EQ_ERR!";
            }
        }

        [ExcelFunction(Description = "create swap rate helper", Category = "EliteQuantExcel - Rates")]
        public static string eqIRCurveSwapRateHelper(
            [ExcelArgument(Description = "(String) id of rate helper object ")] String ObjectId,
            [ExcelArgument(Description = "(double) quote of swap rate ")] double quote,
            [ExcelArgument(Description = "(String) forward start month, e.g. 7D, 3M ")] String Tenor,
            [ExcelArgument(Description = " spread ")] double spread,
            [ExcelArgument(Description = " name of swap curve(USDLIB3M, USDLIB1M, USDLIB6M, USDLIB12M) ")] string idx,
            [ExcelArgument(Description = " id of discount curve (USDOIS) ")] string discount,
            [ExcelArgument(Description = "trigger ")]object trigger)
        {
            if (ExcelUtil.CallFromWizard())
                return "";

            string callerAddress = ExcelUtil.getActiveCellAddress();

            try
            {
                // use default value
                EliteQuant.IborIndex idx_usdlibor = null;

                EliteQuant.QuoteHandle rate_ = new EliteQuant.QuoteHandle(new EliteQuant.SimpleQuote(quote));
                EliteQuant.Period tenor_ = EliteQuant.EQConverter.ConvertObject<EliteQuant.Period>(Tenor);

                if (ExcelUtil.isNull(spread))
                {
                    spread = 0.0;
                }
                EliteQuant.QuoteHandle spread_ = new EliteQuant.QuoteHandle(new EliteQuant.SimpleQuote(spread));

                EliteQuant.RateHelper rh = null;

                if (ExcelUtil.isNull(idx))
                {
                    idx_usdlibor = new EliteQuant.USDLibor(new EliteQuant.Period(3, EliteQuant.TimeUnit.Months));
                }
                else
                {
                    switch (idx)
                    {
                        case "USDLIB1M":
                            idx_usdlibor = new EliteQuant.USDLibor(new EliteQuant.Period(1, EliteQuant.TimeUnit.Months));
                            break;
                        case "USDLIB6M":
                            idx_usdlibor = new EliteQuant.USDLibor(new EliteQuant.Period(6, EliteQuant.TimeUnit.Months));
                            break;
                        case "USDLIB12M":
                            idx_usdlibor = new EliteQuant.USDLibor(new EliteQuant.Period(12, EliteQuant.TimeUnit.Months));
                            break;
                        default:
                            idx_usdlibor = new EliteQuant.USDLibor(new EliteQuant.Period(3, EliteQuant.TimeUnit.Months));
                            break;
                    }
                }

                if (ExcelUtil.isNull(discount))
                {
                    rh = new EliteQuant.SwapRateHelper(rate_, tenor_,
                        cal_usd_gbp, EliteQuant.Frequency.Semiannual, bdc_usd, dc_30_360, idx_usdlibor);
                }
                else
                {
                    if (!discount.Contains('@'))
                        discount = "CRV@" + discount;

                    EliteQuant.YieldTermStructure curve = OHRepository.Instance.getObject<EliteQuant.YieldTermStructure>(discount);
                    EliteQuant.YieldTermStructureHandle yth = new EliteQuant.YieldTermStructureHandle(curve);
                    rh = new EliteQuant.SwapRateHelper(rate_, tenor_,
                        cal_usd_gbp, EliteQuant.Frequency.Semiannual, bdc_usd, dc_30_360,
                        idx_usdlibor, spread_, new EliteQuant.Period(0, EliteQuant.TimeUnit.Days), yth);
                }

                string id = "RHSWP@" + ObjectId;
                OHRepository.Instance.storeObject(id, rh, callerAddress);
                return id + "#" + DateTime.Now.ToString("HH:mm:ss");
            }
            catch (Exception e)
            {
                ExcelUtil.logError(callerAddress, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), e.Message);
                return "#EQ_ERR!";
            }
        }
        #endregion

        #region OIS
        // 0. deposit rate
        [ExcelFunction(Description = "create ois rate helper", Category = "EliteQuantExcel - Rates")]
        public static string eqIRCurveOISRateHelper(
            [ExcelArgument(Description = "(String) id of rate helper object ")] String ObjectId,
            [ExcelArgument(Description = "(double) quote of swap rate ")] double quote,
            [ExcelArgument(Description = "(String) forward start month, e.g. 7D, 3M ")] String Tenor,
            [ExcelArgument(Description = "trigger ")]object trigger)
        {
            if (ExcelUtil.CallFromWizard())
                return "";

            string callerAddress = ExcelUtil.getActiveCellAddress();

            try
            {
                // "us settlement"; "Actual/360"; "fixingDays = 0", "F", "eom = true"
                EliteQuant.FedFunds idx_ff = new EliteQuant.FedFunds();

                EliteQuant.QuoteHandle rate_ = new EliteQuant.QuoteHandle(new EliteQuant.SimpleQuote(quote));
                EliteQuant.Period tenor_ = EliteQuant.EQConverter.ConvertObject<EliteQuant.Period>(Tenor);

                // USSO
                EliteQuant.RateHelper rh = new EliteQuant.OISRateHelper((uint)fixingDays_usd, tenor_,  rate_, idx_ff);

                string id = "RHOIS@" + ObjectId;
                OHRepository.Instance.storeObject(id, rh, callerAddress);
                return id + "#" + DateTime.Now.ToString("HH:mm:ss");
            }
            catch (Exception e)
            {
                ExcelUtil.logError(callerAddress, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), e.Message);
                return "#EQ_ERR!";
            }
        }

        [ExcelFunction(Description = "create Fed fund basis swap helper", Category = "EliteQuantExcel - Rates")]
        public static string eqIRCurveOISFFBasisSwapHelper(
            [ExcelArgument(Description = "(String) id of rate helper object ")] String ObjectId,
            [ExcelArgument(Description = "(double) quote of swap rate ")] double quote,
            [ExcelArgument(Description = "(double) basis spread ")] double basisspread,
            [ExcelArgument(Description = "(String) forward start month, e.g. 7D, 3M ")] String Tenor,
            [ExcelArgument(Description = "trigger ")]object trigger)
        {
            if (ExcelUtil.CallFromWizard())
                return "";

            string callerAddress = ExcelUtil.getActiveCellAddress();

            try
            {
                // use default value. Eonia and ois has same convention
                EliteQuant.OvernightIndex idx = new EliteQuant.Eonia();
                
                EliteQuant.QuoteHandle rate_ = new EliteQuant.QuoteHandle(new EliteQuant.SimpleQuote(quote));
                EliteQuant.QuoteHandle spread_ = new EliteQuant.QuoteHandle(new EliteQuant.SimpleQuote(basisspread));
                EliteQuant.Period tenor_ = EliteQuant.EQConverter.ConvertObject<EliteQuant.Period>(Tenor);
                EliteQuant.DayCounter dc = new EliteQuant.Actual360();

                // arithmetic average, not compounded. USBG
                EliteQuant.RateHelper rh = new EliteQuant.FixedOISBasisRateHelper(2, tenor_, spread_, rate_, 
                    EliteQuant.Frequency.Quarterly, EliteQuant.BusinessDayConvention.ModifiedFollowing, 
                    dc, idx, EliteQuant.Frequency.Quarterly);

                string id = "RHFFB@" + ObjectId;
                OHRepository.Instance.storeObject(id, rh, callerAddress);
                return id + "#" + DateTime.Now.ToString("HH:mm:ss");
            }
            catch (Exception e)
            {
                ExcelUtil.logError(callerAddress, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), e.Message);
                return "#EQ_ERR!";
            }
        }
        #endregion

        #region LIB Basis
        // http://papers.ssrn.com/sol3/papers.cfm?abstract_id=2219548
        // 3M - (1M+basis) = (3M-fixed) + fixed - (1M+basis) = fixed - (1M+basis)
        [ExcelFunction(Description = "create libor basis swap helper", Category = "EliteQuantExcel - Rates")]
        public static string eqIRCurveLiborBasisSwapHelper(
            [ExcelArgument(Description = "(String) id of rate helper object ")] String ObjectId,
            [ExcelArgument(Description = "(String) base leg (usually USDLIB3M) ")] String baseLeg,
            [ExcelArgument(Description = "(String) basis leg (USDLIB1M, USDLIB6M, etc) ")] String basisLeg,
            [ExcelArgument(Description = "(double) basis spread ")] double basis,
            [ExcelArgument(Description = "(String) basis swap tenor (1Y, 2Y, etc) ")] String tenor,
            [ExcelArgument(Description = "Discount Curve (USDLIB3M or USDOIS) ")] String discount,
            [ExcelArgument(Description = "trigger ")]object trigger)
        {
            if (ExcelUtil.CallFromWizard())
                return "";

            string callerAddress = ExcelUtil.getActiveCellAddress();

            try
            {
                // use default value. Eonia and ois has same convention
                if (!baseLeg.Contains('@'))
                    baseLeg = "IDX" + baseLeg;
                EliteQuant.IborIndex baseidx = OHRepository.Instance.getObject<EliteQuant.IborIndex>(baseLeg);
                EliteQuant.IborIndex basisidx = null;
                switch (basisLeg.ToUpper())
                {
                    case "USDLIB1M":
                        basisidx = new EliteQuant.USDLibor(new EliteQuant.Period(1, EliteQuant.TimeUnit.Months));
                        break;
                    case "USDLIB6M":
                        basisidx = new EliteQuant.USDLibor(new EliteQuant.Period(6, EliteQuant.TimeUnit.Months));
                        break;
                    case "USDLIB12M":
                        basisidx = new EliteQuant.USDLibor(new EliteQuant.Period(12, EliteQuant.TimeUnit.Months));
                        break;
                    default:
                        break;
                }

                EliteQuant.YieldTermStructure curve = null;
                EliteQuant.YieldTermStructureHandle yth = null;
                
                if (!discount.Contains('@'))
                {
                    discount = "CRV@"+discount;
                }
                if (!ExcelUtil.isNull(discount))
                {
                    curve = OHRepository.Instance.getObject<EliteQuant.YieldTermStructure>(discount);
                    yth = new EliteQuant.YieldTermStructureHandle(curve);
                }

                EliteQuant.QuoteHandle basis_ = new EliteQuant.QuoteHandle(new EliteQuant.SimpleQuote(basis));
                EliteQuant.Period tenor_ = EliteQuant.EQConverter.ConvertObject<EliteQuant.Period>(tenor);

                // arithmetic average, not compounded. USBG
                EliteQuant.RateHelper rh = new EliteQuant.IBORBasisRateHelper(2, tenor_, basis_, baseidx, basisidx, yth);

                string id = "RHBAS@" + ObjectId;
                OHRepository.Instance.storeObject(id, rh, callerAddress);
                return id + "#" + DateTime.Now.ToString("HH:mm:ss");
            }
            catch (Exception e)
            {
                ExcelUtil.logError(callerAddress, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), e.Message);
                return "#EQ_ERR!";
            }
        }
        #endregion

        #region boostrapping
        [ExcelFunction(Description = "create curve ", Category = "EliteQuantExcel - Rates")]
        public static string eqIRCurveLinearZero(
            [ExcelArgument(Description = "(String) id of curve (USDOIS, USDLIB3M) ")] string ObjectId,
            [ExcelArgument(Description = "array of rate helpers ")] object[] ratehelpers,
            [ExcelArgument(Description = "Interpolation Method (default LogLinear) ")] string interp,
            [ExcelArgument(Description = "trigger ")]object trigger)
        {
            if (ExcelUtil.CallFromWizard())
                return "";

            string callerAddress = ExcelUtil.getActiveCellAddress();
            Xl.Range rng = ExcelUtil.getActiveCellRange();

            try
            {
                string interpmethod;
                if (ExcelUtil.isNull(interp))
                {
                    interpmethod = "LOGLINEAR";
                }
                else
                {
                    interpmethod = interp.ToUpper();
                }

                EliteQuant.RateHelperVector rhv = new EliteQuant.RateHelperVector();

                EliteQuant.Date today = EliteQuant.Settings.instance().getEvaluationDate();
                List<EliteQuant.Date> dates = new List<EliteQuant.Date>();
                dates.Add(today);       // today has discount 1

                foreach(var rid in ratehelpers)
                {
                    if (ExcelUtil.isNull(rid))
                        continue;

                    try
                    {
                        EliteQuant.RateHelper rh = OHRepository.Instance.getObject<EliteQuant.RateHelper>((string)rid);
                        rhv.Add(rh);
                        dates.Add(rh.latestDate());
                    }
                    catch (Exception)   
                    {
                        // skip null instruments
                    }
                }

                // set reference date to today. or discount to 1
                EliteQuant.YieldTermStructure yth = new EliteQuant.PiecewiseLogLinearDiscount(today, rhv, dc_act_360);

                EliteQuant.DateVector dtv = new EliteQuant.DateVector();
                EliteQuant.DoubleVector discv = new EliteQuant.DoubleVector();
                foreach(var dt in dates)
                {
                    double disc = yth.discount(dt);
                    dtv.Add(dt);
                    discv.Add(disc);
                }

                // reconstruct the discount curve
                // note that discount curve is LogLinear
                EliteQuant.YieldTermStructure yth2 = null;
                yth2 = new EliteQuant.DiscountCurve(dtv, discv, dc_act_360, ObjectId.Contains("OIS") ? cal_usd : cal_usd_gbp);
           
                if (!ObjectId.Contains('@'))
                    ObjectId = "CRV@" + ObjectId;

                //string id = "IRCRV@" + ObjectId;
                string id = ObjectId;
                OHRepository.Instance.storeObject(id, yth2, callerAddress);
                return id + "#" + DateTime.Now.ToString("HH:mm:ss");
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
