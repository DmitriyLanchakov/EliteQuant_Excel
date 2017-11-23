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
    public class IRSwap
    {
        [ExcelFunction(Description = "Interest Rate vanilla swap", Category = "EliteQuantExcel - Instruments")]
        public static object eqInstIRVanillaSwap(
            [ExcelArgument(Description = "trade id ")] string tradeid,
            [ExcelArgument(Description = "payer/receiver (1/0) ")] bool ispayer,
            [ExcelArgument(Description = "notional ")] double notional,
            [ExcelArgument(Description = "fixed rate ")] double fixedRate,
            [ExcelArgument(Description = "start date ")] DateTime startdate,
            [ExcelArgument(Description = " (String) forward start month, e.g. 7D, 3M, 7Y ")] string Tenor,
            [ExcelArgument(Description = "id of libor index ")] string indexid,
            [ExcelArgument(Description = "floating leg spread ")] double spread,
            [ExcelArgument(Description = "id of discount curve ")] string discountId,
            [ExcelArgument(Description = "trigger ")]object trigger)
        {
            if (ExcelUtil.CallFromWizard())
                return "";

            string callerAddress = "";
            callerAddress = ExcelUtil.getActiveCellAddress();

            try
            {
                Xl.Range rng = ExcelUtil.getActiveCellRange();

                // by default
                bool end_of_month = true;
                EliteQuant.DayCounter fixeddc = new EliteQuant.Thirty360();

                EliteQuant.Period tenor_ = EliteQuant.EQConverter.ConvertObject<EliteQuant.Period>(Tenor);

                if (!indexid.Contains('@'))
                    indexid = "IDX@" + indexid;
                IborIndex idx = OHRepository.Instance.getObject<IborIndex>(indexid);
                if (!discountId.Contains('@'))
                    discountId = "CRV@" + discountId;
                YieldTermStructure discountcurve = OHRepository.Instance.getObject<YieldTermStructure>(discountId);
                YieldTermStructureHandle dch = new YieldTermStructureHandle(discountcurve);

                EliteQuant.Date sdate = EliteQuant.EQConverter.ConvertObject<EliteQuant.Date>(startdate);
                EliteQuant.Date fdate = idx.fixingDate(sdate);
                EliteQuant.Date tdate = idx.fixingCalendar().advance(sdate, tenor_);

                Schedule fixedsch = new Schedule(sdate, tdate, new Period(6, TimeUnit.Months),
                    idx.fixingCalendar(), idx.businessDayConvention(), idx.businessDayConvention(),
                    DateGeneration.Rule.Backward, end_of_month);
                Schedule floatingsch = new Schedule(sdate, tdate, idx.tenor(), idx.fixingCalendar(),
                    idx.businessDayConvention(), idx.businessDayConvention(),
                    DateGeneration.Rule.Backward, end_of_month);

                VanillaSwap swap = new VanillaSwap(ispayer ? _VanillaSwap.Type.Payer : _VanillaSwap.Type.Receiver,
                    notional, fixedsch, fixedRate, fixeddc, floatingsch, idx, spread, idx.dayCounter());
                DiscountingSwapEngine engine = new DiscountingSwapEngine(dch);
                swap.setPricingEngine(engine);
                
                Date refDate = discountcurve.referenceDate();

                // Store the futures and return its id
                string id = "SWP@" + tradeid;
                OHRepository.Instance.storeObject(id, swap, callerAddress);
                id += "#" + (String)DateTime.Now.ToString(@"HH:mm:ss");
                return id;
            }
            catch (Exception e)
            {
                ExcelUtil.logError(callerAddress, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), e.Message);
                return "#EQ_ERR!";
            }
        }

        [ExcelFunction(Description = "Interest Rate vanilla OIS swap", Category = "EliteQuantExcel - Instruments")]
        public static object eqInstIROISSwap(
            [ExcelArgument(Description = "trade id ")] string tradeid,
            [ExcelArgument(Description = "payer/receiver (1/0) ")] bool ispayer,
            [ExcelArgument(Description = "notional ")] double notional,
            [ExcelArgument(Description = "fixed rate ")] double fixedRate,
            [ExcelArgument(Description = "start date ")] DateTime startdate,
            [ExcelArgument(Description = " (String) forward start month, e.g. 7D, 3M, 7Y ")] string Tenor,
            [ExcelArgument(Description = "id of overnight index ")] string indexid,
            [ExcelArgument(Description = "floating leg spread ")] double spread,
            [ExcelArgument(Description = "id of discount curve ")] string discountId,
            [ExcelArgument(Description = "trigger ")]object trigger)
        {
            if (ExcelUtil.CallFromWizard())
                return "";

            string callerAddress = "";
            callerAddress = ExcelUtil.getActiveCellAddress();

            try
            {
                Xl.Range rng = ExcelUtil.getActiveCellRange();

                // by default
                // endOfMonth_(1*Months<=swapTenor && swapTenor<=2*Years ? true : false),
                bool end_of_month = true;
                EliteQuant.DayCounter fixeddc = new EliteQuant.Actual360();

                if (!indexid.Contains('@'))
                    indexid = "IDX@" + indexid;
                OvernightIndex idx = OHRepository.Instance.getObject<OvernightIndex>(indexid);
                if (!discountId.Contains('@'))
                    discountId = "CRV@" + discountId;
                YieldTermStructure discountcurve = OHRepository.Instance.getObject<YieldTermStructure>(discountId);
                YieldTermStructureHandle dch = new YieldTermStructureHandle(discountcurve);

                EliteQuant.Period tenor_ = EliteQuant.EQConverter.ConvertObject<EliteQuant.Period>(Tenor);
                EliteQuant.Date sdate = EliteQuant.EQConverter.ConvertObject<EliteQuant.Date>(startdate);
                EliteQuant.Date fdate = idx.fixingDate(sdate);
                EliteQuant.Date tdate = idx.fixingCalendar().advance(sdate, tenor_);

                // fixed leg 1 yr. Forward?
                Schedule fixedsch = new Schedule(sdate, tdate, new Period(1, TimeUnit.Years),
                    idx.fixingCalendar(), idx.businessDayConvention(), idx.businessDayConvention(),
                    DateGeneration.Rule.Forward, end_of_month);

                OvernightIndexedSwap swap = new OvernightIndexedSwap(ispayer ? _OvernightIndexedSwap.Type.Payer : _OvernightIndexedSwap.Type.Receiver,
                    notional, fixedsch, fixedRate, fixeddc, idx, spread);

                DiscountingSwapEngine engine = new DiscountingSwapEngine(dch);
                swap.setPricingEngine(engine);

                Date refDate = discountcurve.referenceDate();

                // Store the futures and return its id
                string id = "SWP@" + tradeid;
                OHRepository.Instance.storeObject(id, swap, callerAddress);
                id += "#" + (String)DateTime.Now.ToString(@"HH:mm:ss");
                return id;
            }
            catch (Exception e)
            {
                ExcelUtil.logError(callerAddress, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), e.Message);
                return "#EQ_ERR!";
            }
        }

        [ExcelFunction(Description = "Interest Rate vanilla basis swap", Category = "EliteQuantExcel - Instruments")]
        public static object eqInstIRBasisSwap(
            [ExcelArgument(Description = "trade id ")] string tradeid,
            [ExcelArgument(Description = "payer/receiver (1/0) ")] bool ispayer,
            [ExcelArgument(Description = "notional ")] double notional,
            [ExcelArgument(Description = "start date ")] DateTime startdate,
            [ExcelArgument(Description = " (String) forward start month, e.g. 7D, 3M, 7Y ")] string Tenor,
            [ExcelArgument(Description = "id of base index ")] string baseindexid,
            [ExcelArgument(Description = "id of basis index ")] string basisindexid,
            [ExcelArgument(Description = "basis leg spread ")] double spread,
            [ExcelArgument(Description = "id of discount curve ")] string discountId,
            [ExcelArgument(Description = "trigger ")]object trigger)
        {
            if (ExcelUtil.CallFromWizard())
                return "";

            string callerAddress = "";
            callerAddress = ExcelUtil.getActiveCellAddress();

            try
            {
                Xl.Range rng = ExcelUtil.getActiveCellRange();

                // by default
                // endOfMonth_(1*Months<=swapTenor && swapTenor<=2*Years ? true : false),
                bool end_of_month = true;
                EliteQuant.DayCounter fixeddc = new EliteQuant.Actual360();

                if (!baseindexid.Contains('@'))
                    baseindexid = "IDX@" + baseindexid;
                IborIndex baseidx = OHRepository.Instance.getObject<IborIndex>(baseindexid);

                if (!basisindexid.Contains('@'))
                    basisindexid = "IDX@" + basisindexid;
                IborIndex basisidx = OHRepository.Instance.getObject<IborIndex>(basisindexid);

                if (!discountId.Contains('@'))
                    discountId = "CRV@" + discountId;
                YieldTermStructure discountcurve = OHRepository.Instance.getObject<YieldTermStructure>(discountId);
                YieldTermStructureHandle dch = new YieldTermStructureHandle(discountcurve);

                EliteQuant.Period tenor_ = EliteQuant.EQConverter.ConvertObject<EliteQuant.Period>(Tenor);
                EliteQuant.Date sdate = EliteQuant.EQConverter.ConvertObject<EliteQuant.Date>(startdate);
                EliteQuant.Date fdate = baseidx.fixingDate(sdate);
                EliteQuant.Date tdate = baseidx.fixingCalendar().advance(sdate, tenor_);

                // fixed leg 1 yr. Forward?
                Schedule basesch = new Schedule(sdate, tdate, baseidx.tenor(),
                    baseidx.fixingCalendar(), baseidx.businessDayConvention(), baseidx.businessDayConvention(),
                    DateGeneration.Rule.Backward, end_of_month);

                Schedule basissch = new Schedule(sdate, tdate, basisidx.tenor(),
                    basisidx.fixingCalendar(), basisidx.businessDayConvention(), basisidx.businessDayConvention(),
                    DateGeneration.Rule.Backward, end_of_month);

                //GenericSwap swap = new GenericSwap((ispayer ? _GenericSwap.Type.Payer : _GenericSwap.Type.Receiver), notional,
                //    basesch, baseidx, baseidx.dayCounter(), basissch, basisidx, basisidx.dayCounter(), spread);

                //DiscountingSwapEngine engine = new DiscountingSwapEngine(dch);
                //swap.setPricingEngine(engine);

                //Date refDate = discountcurve.referenceDate();

                // Store the futures and return its id
                //string id = "SWP@" + tradeid;
                //OHRepository.Instance.storeObject(id, swap, callerAddress);
                //id += "#" + (String)DateTime.Now.ToString(@"HH:mm:ss");
                //return id;
                return 0;
            }
            catch (Exception e)
            {
                ExcelUtil.logError(callerAddress, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), e.Message);
                return "#EQ_ERR!";
            }
        }

        // schedule has 0:n; notionals has 0:n-1 elements
        // in cashflowvectors.hpp, it gets i or the last one if runs out
        [ExcelFunction(Description = "Interest Rate generic swap ", Category = "EliteQuantExcel - Instruments")]
        public static object eqInstIRGenericSwap(
            // trade info
            /*[ExcelArgument(Description = "trade id ")] string tradeid,
            [ExcelArgument(Description = "Entity ")] string entity,
            [ExcelArgument(Description = "Entity ID ")] string entityid,
            [ExcelArgument(Description = "Counterparty ")] string counterparty,
            [ExcelArgument(Description = "Counterparty ID ")] string counterpartyid,
            [ExcelArgument(Description = "swap type ")] string swaptype,
            [ExcelArgument(Description = "Fixing Days ")] object fixingdays,        // use object to catch missing
            [ExcelArgument(Description = "Trade Date ")] object tradedate,
            [ExcelArgument(Description = "Start date ")] object startdate,
            [ExcelArgument(Description = "Maturity date ")] object maturitydate,
            [ExcelArgument(Description = "Tenor ")] string Tenor,
            [ExcelArgument(Description = "is notional schedule given ")] bool isschedulegiven,*/
            [ExcelArgument(Description = "swap type ")] object[] tradeinfo,
            // first leg
            /*[ExcelArgument(Description = "id of first leg index ")] string firstlegindex,
            [ExcelArgument(Description = "first leg frequency ")] string firstlegfreq,
            [ExcelArgument(Description = "first leg convention ")] string firstlegconv,
            [ExcelArgument(Description = "first leg calendar ")] string firstlegcalendar,
            [ExcelArgument(Description = "first leg day counter ")] string firstlegdc,
            [ExcelArgument(Description = "first leg date generation rule ")] string firstlegdgrule,
            [ExcelArgument(Description = "first leg end of month ")] bool firstlegeom,
            [ExcelArgument(Description = "first leg fixed rate ")] double firstlegrate,*/
            [ExcelArgument(Description = "first leg info ")] object[] firstleginfo,       
            [ExcelArgument(Description = "first leg notional(s) ")] object[] firstlegnotionals,     // only object[] works
            [ExcelArgument(Description = "first leg schedule(s) ")] object[,] firstlegschedule,
            // second leg
            /*[ExcelArgument(Description = "id of second leg index ")] string secondlegindex,
            [ExcelArgument(Description = "second leg frequency  ")] string secondlegfreq,
            [ExcelArgument(Description = "second leg convention ")] string secondlegconv,
            [ExcelArgument(Description = "second leg calendar ")] string secondlegcalendar,
            [ExcelArgument(Description = "second leg day counter ")] string secondlegdc,
            [ExcelArgument(Description = "second leg date generation rule ")] string secondlegdgrule,
            [ExcelArgument(Description = "second leg end of month ")] bool secondlegeom,
            [ExcelArgument(Description = "second leg spread ")] double secondlegspread,*/
            [ExcelArgument(Description = "second leg info ")] object[] secondleginfo,
            [ExcelArgument(Description = "second leg notional(s) ")] object[] secondlegnotionals,
            [ExcelArgument(Description = "second leg schedule(s) ")] object[,] secondlegschedule,
            [ExcelArgument(Description = "id of discount curve ")] string discountId
            )
        {
            if (ExcelUtil.CallFromWizard())
                return "";

            string callerAddress = "";
            callerAddress = ExcelUtil.getActiveCellAddress();

            try
            {
                Xl.Range rng = ExcelUtil.getActiveCellRange();

                EliteQuant.Instruments.InterestRateGenericSwap genswap = new EliteQuant.Instruments.InterestRateGenericSwap();

                #region parameters
                if (ExcelUtil.isNull(tradeinfo[0]))
                    return "#EQ_ERR!";
                else
                    genswap.ContractId = (string)tradeinfo[0];

                if (ExcelUtil.isNull(tradeinfo[1]))
                    genswap.Entity = "NA";
                else
                    genswap.Entity = (string)tradeinfo[1];

                if (ExcelUtil.isNull(tradeinfo[2]))
                    genswap.EntityID = "NA";
                else
                    genswap.EntityID = (string)tradeinfo[2];

                if (ExcelUtil.isNull(tradeinfo[3]))
                    genswap.Counterparty = "NA";
                else
                    genswap.Counterparty = (string)tradeinfo[3];

                if (ExcelUtil.isNull(tradeinfo[4]))
                    genswap.CounterpartyID = "NA";
                else
                    genswap.CounterpartyID = (string)tradeinfo[4];

                if (ExcelUtil.isNull(tradeinfo[5]))
                    genswap.SwapType = "Payer";
                else
                    genswap.SwapType = (string)tradeinfo[5];

                if (ExcelUtil.isNull(tradeinfo[6]))
                    genswap.FixingDays = 2;
                else
                    genswap.FixingDays = (int)(double)tradeinfo[6];

                if (ExcelUtil.isNull(tradeinfo[7]))
                    genswap.TradeDate = EliteQuant.EQConverter.DateTimeToString(DateTime.Today);
                else
                    genswap.TradeDate = EliteQuant.EQConverter.DateTimeToString(DateTime.FromOADate((double)tradeinfo[7]));

                // set it temporarily to ""
                if (ExcelUtil.isNull(tradeinfo[8]))
                    genswap.SettlementDate = string.Empty;
                else
                    genswap.SettlementDate = EliteQuant.EQConverter.DateTimeToString(DateTime.FromOADate((double)tradeinfo[8]));

                // set it temporarily to today
                if (ExcelUtil.isNull(tradeinfo[9]))
                    genswap.MaturityDate = string.Empty;
                else
                    genswap.MaturityDate = EliteQuant.EQConverter.DateTimeToString(DateTime.FromOADate((double)tradeinfo[9]));

                // set it temporarily to blank
                if (ExcelUtil.isNull(tradeinfo[10]))
                    genswap.Tenor = string.Empty;
                else
                    genswap.Tenor = (string)tradeinfo[10];

                if (ExcelUtil.isNull(tradeinfo[11]))
                    genswap.IsScheduleGiven = false;
                else
                    //genswap.IsScheduleGiven = Convert.ToBoolean((string)tradeinfo[11]);
                    genswap.IsScheduleGiven = (bool)tradeinfo[11];
                genswap.IsScheduleGiven = false;        // set to false always, amortization currently not supported

                //***************  First Leg *************************//
                if (ExcelUtil.isNull(firstleginfo[0]))
                    genswap.FirstLegIndex = "FIXED";
                else
                    genswap.FirstLegIndex = (string)firstleginfo[0];

                if (ExcelUtil.isNull(firstleginfo[1]))
                    genswap.FirstLegFrequency = "SEMIANNUAL";
                else
                    genswap.FirstLegFrequency = (string)firstleginfo[1];

                if (ExcelUtil.isNull(firstleginfo[2]))
                    genswap.FirstLegConvention = "MODIFIEDFOLLOWING";
                else
                    genswap.FirstLegConvention = (string)firstleginfo[2];

                if (ExcelUtil.isNull(firstleginfo[3]))
                    genswap.FirstLegCalendar = "NYC|LON";
                else
                    genswap.FirstLegCalendar = (string)firstleginfo[3];

                if (ExcelUtil.isNull(firstleginfo[4]))
                    genswap.FirstLegDayCounter = "ACTUAL360";
                else
                    genswap.FirstLegDayCounter = (string)firstleginfo[4];

                if (ExcelUtil.isNull(firstleginfo[5]))
                    genswap.FirstLegDateGenerationRule = "BACKWARD";
                else
                    genswap.FirstLegDateGenerationRule = (string)firstleginfo[5];

                if (ExcelUtil.isNull(firstleginfo[6]))
                    genswap.FirstLegEOM = true;
                else
                    genswap.FirstLegEOM = (bool)firstleginfo[6];

                if (ExcelUtil.isNull(firstleginfo[7]))
                    genswap.FirstLegSpread = 0.0;
                else
                    genswap.FirstLegSpread = (double)firstleginfo[7];

                if (ExcelUtil.isNull(firstlegnotionals))
                {
                    genswap.FirstLegNotionals.Clear();
                    genswap.FirstLegNotionals.Add(0);       // size = 1
                }
                else
                {
                    genswap.FirstLegNotionals.Clear();
                    foreach (var nl in firstlegnotionals)
                    {
                        if (ExcelUtil.isNull(nl))
                            continue;

                        genswap.FirstLegNotionals.Add((double)nl);
                    }
                }

                if (ExcelUtil.isNull(firstlegschedule) || (!genswap.IsScheduleGiven))
                {
                    genswap.FirstLegSchedules.Clear();
                    genswap.FirstLegSchedules.Add(genswap.SettlementDate);
                    genswap.FirstLegSchedules.Add(genswap.MaturityDate);
                }
                else
                {
                    genswap.FirstLegSchedules.Clear();
                    for (int a = 0; a < firstlegschedule.GetLength(0);a++ )
                    {
                        DateTime d;
                        if (ExcelUtil.isNull(firstlegschedule[a, 0]))
                        {
                            // add one more date
                            d = DateTime.FromOADate((double)firstlegschedule[a-1, 1]);
                            genswap.FirstLegSchedules.Add(EliteQuant.EQConverter.DateTimeToString(d));
                            break;
                        }
                            

                        d = DateTime.FromOADate((double)firstlegschedule[a,0]);
                        genswap.FirstLegSchedules.Add(EliteQuant.EQConverter.DateTimeToString(d));
                    }
                }

                //***************  Second Leg *************************//

                if (ExcelUtil.isNull(secondleginfo[0]))
                    genswap.SecondLegIndex = "USDLIB3M";
                else
                    genswap.SecondLegIndex = (string)secondleginfo[0];

                if (ExcelUtil.isNull(secondleginfo[1]))
                    genswap.SecondLegFrequency = "QUARTERLY";
                else
                    genswap.SecondLegFrequency = (string)secondleginfo[1];

                if (ExcelUtil.isNull(secondleginfo[2]))
                    genswap.SecondLegConvention = "MODIFIEDFOLLOWING";
                else
                    genswap.SecondLegConvention = (string)secondleginfo[2];

                if (ExcelUtil.isNull(secondleginfo[3]))
                    genswap.SecondLegCalendar = "NYC|LON";           // nor NYC|LON
                else
                    genswap.SecondLegCalendar = (string)secondleginfo[3];

                if (ExcelUtil.isNull(secondleginfo[4]))
                    genswap.SecondLegDayCounter = "ACTUAL360";
                else
                    genswap.SecondLegDayCounter = (string)secondleginfo[4];

                if (ExcelUtil.isNull(secondleginfo[5]))
                    genswap.SecondLegDateGenerationRule = "BACKWARD";
                else
                    genswap.SecondLegDateGenerationRule = (string)secondleginfo[5];

                if (ExcelUtil.isNull(secondleginfo[6]))
                    genswap.SecondLegEOM = true;
                else
                    genswap.SecondLegEOM = (bool)secondleginfo[6];

                if (ExcelUtil.isNull(secondleginfo[7]))
                    genswap.SecondLegSpread = 0.0;
                else
                    genswap.SecondLegSpread = (double)secondleginfo[7];

                if (ExcelUtil.isNull(secondlegnotionals))
                {
                    genswap.SecondLegNotionals.Clear();
                    genswap.SecondLegNotionals.Add(0);
                }
                else
                {
                    genswap.SecondLegNotionals.Clear();
                    foreach (var nl in secondlegnotionals)
                    {
                        if (ExcelUtil.isNull(nl))
                            continue;

                        genswap.SecondLegNotionals.Add((double)nl);
                    }
                }

                if (ExcelUtil.isNull(secondlegschedule) || (!genswap.IsScheduleGiven))
                {
                    genswap.SecondLegSchedules.Clear();
                    genswap.SecondLegSchedules.Add(genswap.SettlementDate);
                    genswap.SecondLegSchedules.Add(genswap.MaturityDate);
                }
                else
                {
                    genswap.SecondLegSchedules.Clear();
                    for (int a = 0; a < secondlegschedule.GetLength(0); a++)
                    {
                        DateTime d;
                        if (ExcelUtil.isNull(secondlegschedule[a, 0]))
                        {
                            // add one more date
                            d = DateTime.FromOADate((double)secondlegschedule[a - 1, 1]);
                            genswap.SecondLegSchedules.Add(EliteQuant.EQConverter.DateTimeToString(d));
                            break;
                        }


                        d = DateTime.FromOADate((double)secondlegschedule[a, 0]);
                        genswap.SecondLegSchedules.Add(EliteQuant.EQConverter.DateTimeToString(d));
                    }
                }

                #endregion

                #region convert Interest rate generic swap to swap obj
                string firstidx = genswap.FirstLegIndex, secondidx = genswap.SecondLegIndex;
                string firstidx_id = "IDX@" + firstidx;
                string secondidx_id = "IDX@" + secondidx;
                EliteQuant.IborIndex firstidx_obj = null;
                if (!firstidx.Contains("FIXED"))
                    firstidx_obj = OHRepository.Instance.getObject<EliteQuant.IborIndex>(firstidx_id);
                
                EliteQuant.IborIndex secondidx_obj = OHRepository.Instance.getObject<EliteQuant.IborIndex>(secondidx_id);

                genswap.ConstructSwap(firstidx_obj, secondidx_obj);

                if (!discountId.Contains('@'))
                    discountId = "CRV@" + discountId;
                YieldTermStructure discountcurve = OHRepository.Instance.getObject<YieldTermStructure>(discountId);
                YieldTermStructureHandle dch = new YieldTermStructureHandle(discountcurve);

                DiscountingSwapEngine engine = new DiscountingSwapEngine(dch);
                genswap.eqswap_.setPricingEngine(engine);
                #endregion

                string id = "SWP@" + genswap.ContractId;
                OHRepository.Instance.storeObject(id, genswap, callerAddress);
                id += "#" + (String)DateTime.Now.ToString(@"HH:mm:ss");
                return id;
            }
            catch (Exception e)
            {
                ExcelUtil.logError(callerAddress, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), e.Message);
                return "#EQ_ERR!";
            }
        }

        [ExcelFunction(Description = "Display interest rate swap pay/rec schedule", Category = "EliteQuantExcel - Instruments")]
        public static object eqInstDisplayIRSwap(
            [ExcelArgument(Description = "id of IR Swap ")] string tradeid,
            [ExcelArgument(Description = "id of discount curve ")] string discountId,
            [ExcelArgument(Description = "trigger ")]object trigger)
        {
            if (ExcelUtil.CallFromWizard())
                return "";

            string callerAddress = "";
            callerAddress = ExcelUtil.getActiveCellAddress();

            object[,] ret;
            try
            {
                Xl.Range rng = ExcelUtil.getActiveCellRange();


                if (!tradeid.Contains('@'))
                    tradeid = "SWP@" + tradeid;

                if (!discountId.Contains('@'))
                    discountId = "CRV@" + discountId;
                YieldTermStructure discountcurve = OHRepository.Instance.getObject<YieldTermStructure>(discountId);
                Date asofdate = Settings.instance().getEvaluationDate();

                GenericSwap inst = OHRepository.Instance.getObject<GenericSwap>(tradeid);
                
                int rows = Math.Max(inst.firstLegInfo().Count, inst.secondLegInfo().Count);
                ret = new object[rows, 20];     // 10 cols each leg
                for (int i = 0; i < rows; i++)
                {
                    for (int j = 0; j < 20; j++)
                    {
                        ret[i, j] = "";         // initialization. null will be posted as 0; so explicitly set it to ""
                    }
                }

                // first leg
                string[] s;
                DateTime startdate, enddate, paymentdate, resetdate;
                double balance = 0, rate = 0, spread = 0, payment = 0, discount = 0, pv = 0;
                for (int i = 0; i < inst.firstLegInfo().Count; i++ )
                {
                    s = inst.firstLegInfo()[i].Split(',');
                    startdate = EliteQuant.EQConverter.DateToDateTime(new Date(Convert.ToInt32(s[0])));
                    enddate = EliteQuant.EQConverter.DateToDateTime(new Date(Convert.ToInt32(s[1])));
                    paymentdate = EliteQuant.EQConverter.DateToDateTime(new Date(Convert.ToInt32(s[2])));
                    resetdate = (s[3]=="") ? DateTime.MinValue : EliteQuant.EQConverter.DateToDateTime(new Date(Convert.ToInt32(s[3])));
                    balance = Convert.ToDouble(s[4]);
                    rate = Convert.ToDouble(s[5]);
                    spread = Convert.ToDouble(s[6]);
                    payment = Convert.ToDouble(s[7]);
                    
                    // today's cashflow is not included
                    if (EliteQuant.EQConverter.DateTimeToDate(paymentdate).serialNumber() <= asofdate.serialNumber())
                    {
                        discount = 0.0;
                    }
                    else
                    {
                        discount = discountcurve.discount(EliteQuant.EQConverter.DateTimeToDate(paymentdate));
                    }
                    
                    pv = payment * discount;

                    // and return the matrix to vba
                    ret[i, 0] = (object)startdate;
                    ret[i, 1] = (object)enddate;
                    ret[i, 2] = (object)paymentdate;
                    ret[i, 3] = (s[3]=="") ? "":(object)resetdate;
                    ret[i, 4] = (object)(balance == 0 ? "" : (object)balance);
                    ret[i, 5] = (object)(rate == 0 ? "" : (object)rate);
                    ret[i, 6] = (object)(spread == 0 ? "" : (object)spread);
                    ret[i, 7] = (object)(payment == 0 ? "" : (object)payment);
                    ret[i, 8] = (object)(discount == 0 ? "" : (object)discount);
                    ret[i, 9] = (object)(pv == 0 ? "" : (object)pv);
                }
                for (int i = 0; i < inst.secondLegInfo().Count; i++)
                {
                    s = inst.secondLegInfo()[i].Split(',');
                    startdate = EliteQuant.EQConverter.DateToDateTime(new Date(Convert.ToInt32(s[0])));
                    enddate = EliteQuant.EQConverter.DateToDateTime(new Date(Convert.ToInt32(s[1])));
                    paymentdate = EliteQuant.EQConverter.DateToDateTime(new Date(Convert.ToInt32(s[2])));
                    resetdate = (s[3] == "") ? DateTime.MinValue : EliteQuant.EQConverter.DateToDateTime(new Date(Convert.ToInt32(s[3])));
                    balance = Convert.ToDouble(s[4]);
                    rate = Convert.ToDouble(s[5]);
                    spread = Convert.ToDouble(s[6]);
                    payment = Convert.ToDouble(s[7]);

                    // today's cashflow is not included
                    if (EliteQuant.EQConverter.DateTimeToDate(paymentdate).serialNumber() <= asofdate.serialNumber())
                    {
                        discount = 0.0;
                    }
                    else
                    {
                        discount = discountcurve.discount(EliteQuant.EQConverter.DateTimeToDate(paymentdate));
                    }
                    
                    pv = payment * discount;

                    // and return the matrix to vba
                    ret[i, 10] = (object)startdate;
                    ret[i, 11] = (object)enddate;
                    ret[i, 12] = (object)paymentdate;
                    ret[i, 13] = (s[3] == "") ? "" : (object)resetdate;
                    ret[i, 14] = (object)(balance == 0 ? "" : (object)balance);
                    ret[i, 15] = (object)(rate == 0 ? "" : (object)rate);
                    ret[i, 16] = (object)(spread == 0 ? "" : (object)spread);
                    ret[i, 17] = (object)(payment == 0 ? "" : (object)payment);
                    ret[i, 18] = (object)(discount == 0 ? "" : (object)discount);
                    ret[i, 19] = (object)(pv == 0 ? "" : (object)pv);
                }

                return ret;
            }
            catch (Exception e)
            {
                ExcelUtil.logError(callerAddress, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), e.Message);
                return "#EQ_ERR!";
            }
        }

        [ExcelFunction(Description = "Save interest rate swap to Disk ", Category = "EliteQuantExcel - Instruments")]
        public static object eqInstSaveIRSwapToDisk(
            [ExcelArgument(Description = "id of IR Swap ")] string tradeid,
            [ExcelArgument(Description = "trigger ")]object trigger)
        {
            if (ExcelUtil.CallFromWizard())
                return "";

            string callerAddress = "";
            callerAddress = ExcelUtil.getActiveCellAddress();

            try
            {
                Xl.Range rng = ExcelUtil.getActiveCellRange();

                string genswaptplid_ = tradeid;
                if (!genswaptplid_.Contains('@'))
                {
                    genswaptplid_ = "SWP@" + genswaptplid_;
                }
                if (!genswaptplid_.Contains("_TPL"))
                {
                    genswaptplid_ = genswaptplid_ + "_TPL";
                }

                EliteQuant.Instruments.InterestRateGenericSwap genswaptpl = OHRepository.Instance.getObject<EliteQuant.Instruments.InterestRateGenericSwap>(genswaptplid_);

                string path = EliteQuant.ConfigManager.Instance.IRRootDir + @"Trades\";

                EliteQuant.Instruments.InterestRateGenericSwap.Serialize(genswaptpl,
                    path + tradeid + ".xml");
                return tradeid;
            }
            catch (Exception e)
            {
                ExcelUtil.logError(callerAddress, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), e.Message);
                return "#EQ_ERR!";
            }
        }

        // load return load info, do not save to object repository. Let GenSWapTmpl to deal with repository
        [ExcelFunction(Description = "Load interest rate swap from Disk ", Category = "EliteQuantExcel - Instruments")]
        public static object eqInstLoadIRSwapFromDisk(
            [ExcelArgument(Description = "id of IR Swap ")] string tradeid, // should not include @ and _TPL
            [ExcelArgument(Description = "trigger ")]object trigger)
        {
            if (ExcelUtil.CallFromWizard())
                return "";

            string callerAddress = "";
            callerAddress = ExcelUtil.getActiveCellAddress();

            try
            {
                Xl.Range rng = ExcelUtil.getActiveCellRange();
                string genswaptplid_ = tradeid;
                if (!genswaptplid_.Contains('@'))
                {
                    genswaptplid_ = "SWP@" + genswaptplid_;
                }
                if (!genswaptplid_.Contains("_TPL"))
                {
                    genswaptplid_ = genswaptplid_ + "_TPL";
                }

                string path = EliteQuant.ConfigManager.Instance.IRRootDir + @"Trades\";

                EliteQuant.Instruments.InterestRateGenericSwap genswaptpl =
                    EliteQuant.Instruments.InterestRateGenericSwap.Deserialize(path + tradeid + ".xml");

                // preserve same order as excel
                List<object> ret = new List<object>();
                ret.Add(genswaptpl.ContractId);
                ret.Add(genswaptpl.Entity);
                ret.Add(genswaptpl.EntityID);
                ret.Add(genswaptpl.Counterparty);
                ret.Add(genswaptpl.CounterpartyID);
                ret.Add(genswaptpl.SwapType);
                ret.Add(genswaptpl.FixingDays.ToString());
                ret.Add(string.IsNullOrEmpty(genswaptpl.TradeDate) ? "" : (object)EliteQuant.EQConverter.StringToDateTime(genswaptpl.TradeDate));
                ret.Add(string.IsNullOrEmpty(genswaptpl.SettlementDate) ? "" : (object)EliteQuant.EQConverter.StringToDateTime(genswaptpl.SettlementDate));
                ret.Add(string.IsNullOrEmpty(genswaptpl.MaturityDate) ? "" : (object)EliteQuant.EQConverter.StringToDateTime(genswaptpl.MaturityDate));
                ret.Add(genswaptpl.Tenor);
                if ((genswaptpl.FirstLegNotionals.Count > 1) || (genswaptpl.SecondLegNotionals.Count > 1))
                {
                    ret.Add("TRUE");
                }
                else
                {
                    ret.Add(genswaptpl.IsScheduleGiven.ToString());    // always false for now
                }
                ret.Add(genswaptpl.FirstLegIndex);
                ret.Add(genswaptpl.FirstLegFrequency);
                ret.Add(genswaptpl.FirstLegConvention);
                ret.Add(genswaptpl.FirstLegCalendar);
                ret.Add(genswaptpl.FirstLegDayCounter);
                ret.Add(genswaptpl.FirstLegDateGenerationRule);
                ret.Add(genswaptpl.FirstLegEOM.ToString());
                ret.Add(genswaptpl.FirstLegSpread.ToString());
                if (genswaptpl.FirstLegNotionals.Count > 1)
                {
                    ret.Add("Collection...");
                }
                else
                {
                    ret.Add(genswaptpl.FirstLegNotionals[0].ToString());
                }   
                ret.Add("");
                ret.Add(genswaptpl.SecondLegIndex);
                ret.Add(genswaptpl.SecondLegFrequency);
                ret.Add(genswaptpl.SecondLegConvention);
                ret.Add(genswaptpl.SecondLegCalendar);
                ret.Add(genswaptpl.SecondLegDayCounter);
                ret.Add(genswaptpl.SecondLegDateGenerationRule);
                ret.Add(genswaptpl.SecondLegEOM.ToString());
                ret.Add(genswaptpl.SecondLegSpread.ToString());
                if (genswaptpl.SecondLegNotionals.Count > 1)
                {
                    ret.Add("Collection...");
                }
                else
                {
                    ret.Add(genswaptpl.SecondLegNotionals[0].ToString());
                }  
                ret.Add("");

                return ret.ToArray();
            }
            catch (Exception e)
            {
                ExcelUtil.logError(callerAddress, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), e.Message);
                return "#EQ_ERR!";
            }
        }
    }
}
