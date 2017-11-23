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
    public class Commodity
    {
        [ExcelFunction(Description = "Commodity Future", Category = "EliteQuantExcel - Instruments")]
        public static object eqInstCommodityFuture(
            [ExcelArgument(Description = "id of instrument ")] string ObjectId,
            [ExcelArgument(Description = "name of instrument ")] string name,       // given by user, could be the same as objectid
            [ExcelArgument(Description = "buy/sell (1/-1) ")] int buysell,
            [ExcelArgument(Description = "trade price ")] double tradePrice,
            [ExcelArgument(Description = "trade quantity ")] double quantity,
            [ExcelArgument(Description = "start date ")] DateTime startdate,
            [ExcelArgument(Description = "end date ")] DateTime enddate,
            [ExcelArgument(Description = "id of commodity index ")] string indexid,
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
                CommodityIndexExt idx = OHRepository.Instance.getObject<CommodityIndexExt>(indexid);
                YieldTermStructureHandle discountcurve = OHRepository.Instance.getObject<YieldTermStructureHandle>(discountId);
                Date refDate = discountcurve.referenceDate();

                EliteQuant.Date sd = EliteQuant.EQConverter.ConvertObject<EliteQuant.Date>(startdate);
                EliteQuant.Date ed = EliteQuant.EQConverter.ConvertObject<EliteQuant.Date>(enddate);

                PricingPeriodExt pp = new PricingPeriodExt(sd, ed, sd, quantity);         // pay at start date
                EnergyFutureExt ef = new EnergyFutureExt(buysell, pp, tradePrice, idx, name, discountcurve);

                // Store the futures and return its id
                string id = "Fut@" + ObjectId;
                OHRepository.Instance.storeObject(id, ef, callerAddress);
                id += "#" + (String)DateTime.Now.ToString(@"HH:mm:ss");
                return id;
            }
            catch (Exception e)
            {
                ExcelUtil.logError(callerAddress, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), e.Message);
                return "#EQ_ERR!";
            }
        }

        [ExcelFunction(Description = "Commodity vanilla swap", Category = "EliteQuantExcel - Instruments")]
        public static object eqInstCommodityVanillaSwap(
            [ExcelArgument(Description = "id of instrument ")] string ObjectId,
            [ExcelArgument(Description = "name of instrument ")] string name,
            [ExcelArgument(Description = "payer/receiver (1/0) ")] bool payer,
            [ExcelArgument(Description = "trade price ")] double fixedPrice,
            [ExcelArgument(Description = "trade quantity ")] double[] quantities,
            [ExcelArgument(Description = "start date ")] object[] startdates,
            [ExcelArgument(Description = "end date ")] object[] enddates,
            [ExcelArgument(Description = "id of commodity index ")] string indexid,
            [ExcelArgument(Description = "id of discount curve ")] string discountId,
            [ExcelArgument(Description = "trigger ")]object trigger)
        {
            if (ExcelUtil.CallFromWizard())
                return "";

            string callerAddress = "";
            callerAddress = ExcelUtil.getActiveCellAddress();

            try
            {
                //bool ispayer = string.Compare(payer.ToUpper(), "PAYER") == 0 ? true : false;
                bool ispayer = payer;
                if (startdates.Length != enddates.Length)
                {
                    return "size mismatch";
                }

                Xl.Range rng = ExcelUtil.getActiveCellRange();
                CommodityIndexExt idx = OHRepository.Instance.getObject<CommodityIndexExt>(indexid);
                YieldTermStructureHandle discountcurve = OHRepository.Instance.getObject<YieldTermStructureHandle>(discountId);
                Date refDate = discountcurve.referenceDate();

                PricingPeriodExts pps = new PricingPeriodExts(startdates.Length);

                for (int i = 0; i < startdates.Length; i++)
                {
                    if (ExcelUtil.isNull(startdates[i]))
                        continue;

                    //EliteQuant.Date sd = Conversion.ConvertObject<EliteQuant.Date>((DateTime)startdates[i], "date");
                    //EliteQuant.Date ed = Conversion.ConvertObject<EliteQuant.Date>((DateTime)enddates[i], "date");
                    Date sd = new Date(Convert.ToInt32(startdates[i]));
                    Date ed = new Date(Convert.ToInt32(enddates[i]));

                    PricingPeriodExt pp = new PricingPeriodExt(sd, ed, sd, quantities[i]);
                    pps.Add(pp);
                }

                EnergyVanillaSwapExt evs = new EnergyVanillaSwapExt(ispayer, fixedPrice, idx, pps, name, discountcurve, discountcurve,
                    Frequency.Monthly, new NullCalendar());

                // Store the futures and return its id
                string id = "Swp@" + ObjectId;
                OHRepository.Instance.storeObject(id, evs, callerAddress);
                id += "#" + (String)DateTime.Now.ToString(@"HH:mm:ss");
                return id;
            }
            catch (Exception e)
            {
                ExcelUtil.logError(callerAddress, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), e.Message);
                return "#EQ_ERR!";
            }
        }

        [ExcelFunction(Description = "Commodity basis swap", Category = "EliteQuantExcel - Instruments")]
        public static object eqInstCommodityBasisSwap(
            [ExcelArgument(Description = "id of instrument ")] string ObjectId,
            [ExcelArgument(Description = "name of instrument ")] string name,
            [ExcelArgument(Description = "trade quantity ")] double[] quantities,
            [ExcelArgument(Description = "start date ")] object[] startdates,
            [ExcelArgument(Description = "end date ")] object[] enddates,
            [ExcelArgument(Description = "payCoeff ")] double[] payCoeff,
            [ExcelArgument(Description = "recCoeff ")] double[] recCoeff,
            [ExcelArgument(Description = "paySpread ")] double[] paySprd,
            [ExcelArgument(Description = "recSpread ")] double[] recSprd,
            [ExcelArgument(Description = "id of pay leg index ")] string payeridxid,
            [ExcelArgument(Description = "id of rec leg index ")] string recidxid,
            [ExcelArgument(Description = "id of pay leg discount curve ")] string paydiscountId,
            [ExcelArgument(Description = "id of rec leg discount curve ")] string recdiscountId,
            [ExcelArgument(Description = "trigger ")]object trigger)
        {
            if (ExcelUtil.CallFromWizard())
                return "";

            string callerAddress = "";
            callerAddress = ExcelUtil.getActiveCellAddress();

            try
            {
                if (startdates.Length != enddates.Length)
                {
                    return "size mismatch";
                }

                Xl.Range rng = ExcelUtil.getActiveCellRange();
                CommodityIndexExt payeridx = OHRepository.Instance.getObject<CommodityIndexExt>(payeridxid);
                CommodityIndexExt recidx = OHRepository.Instance.getObject<CommodityIndexExt>(recidxid);
                YieldTermStructureHandle paydiscountcurve = OHRepository.Instance.getObject<YieldTermStructureHandle>(paydiscountId);
                YieldTermStructureHandle recdiscountcurve = OHRepository.Instance.getObject<YieldTermStructureHandle>(recdiscountId);
                Date refDate = paydiscountcurve.referenceDate();

                PricingPeriodExts pps = new PricingPeriodExts(startdates.Length);

                for (int i = 0; i < startdates.Length; i++)
                {
                    //EliteQuant.Date sd = Conversion.ConvertObject<EliteQuant.Date>((DateTime)startdates[i], "date");
                    //EliteQuant.Date ed = Conversion.ConvertObject<EliteQuant.Date>((DateTime)enddates[i], "date");
                    Date sd = new Date(Convert.ToInt32(startdates[i]));
                    Date ed = new Date(Convert.ToInt32(enddates[i]));

                    PricingPeriodExt pp = new PricingPeriodExt(sd, ed, sd, quantities[i], payCoeff[i], recCoeff[i], paySprd[i], recSprd[i]);
                    pps.Add(pp);
                }

                EnergyBasisSwapExt ebs = new EnergyBasisSwapExt(payeridx, recidx, pps, name, paydiscountcurve, recdiscountcurve, Frequency.Monthly, new NullCalendar());

                // Store the futures and return its id
                string id = "Swp@" + ObjectId;
                OHRepository.Instance.storeObject(id, ebs, callerAddress);
                id += "#" + (String)DateTime.Now.ToString(@"HH:mm:ss");
                return id;
            }
            catch (Exception e)
            {
                ExcelUtil.logError(callerAddress, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), e.Message);
                return "#EQ_ERR!";
            }
        }

        [ExcelFunction(Description = "Retrieve instrument NPV and others", Category = "EliteQuantExcel - Instruments")]
        public static object eqInstGetInstrumentNPV(
            [ExcelArgument(Description = "id of instrument ")] string ObjectId,
            [ExcelArgument(Description = "Greek type ")]string gtype,
            [ExcelArgument(Description = "is genericswap ")]bool isgenswap,
            [ExcelArgument(Description = "trigger ")]object trigger)
        {
            if (ExcelUtil.CallFromWizard())
                return "";

            string callerAddress = "";
            callerAddress = ExcelUtil.getActiveCellAddress();

            try
            {
                Xl.Range rng = ExcelUtil.getActiveCellRange();

                bool isgenswap_ = false;
                if (ExcelUtil.isNull(isgenswap))
                    isgenswap_ = false;
                else
                    isgenswap_ = (bool)isgenswap;

                Instrument inst = null;
                if (isgenswap_)
                {
                    EliteQuant.Instruments.InterestRateGenericSwap genswap = OHRepository.Instance.getObject<EliteQuant.Instruments.InterestRateGenericSwap>(ObjectId);
                    inst = genswap.eqswap_;
                }
                else
                {
                    inst = OHRepository.Instance.getObject<Instrument>(ObjectId);
                }
                
                
                double ret = inst.NPV();

                Type type = inst.GetType();
                if (type == typeof(VanillaSwap))
                {
                    if (gtype.ToUpper() == "FAIRRATE")
                    {
                        ret = (inst as VanillaSwap).fairRate();
                    }
                }
                else if (type == typeof(OvernightIndexedSwap))
                {
                    if (gtype.ToUpper() == "FAIRRATE")
                    {
                        ret = (inst as OvernightIndexedSwap).fairRate();
                    }
                }
                else if (type == typeof(EliteQuant.Instruments.InterestRateGenericSwap))
                {
                    if (gtype.ToUpper() == "FAIRRATE")
                    {
                        ret = (inst as GenericSwap).fairRate();
                    }
                    else if (gtype.ToUpper() == "FAIRSPREAD")
                    {
                        ret = (inst as GenericSwap).fairSpread();
                    }
                    else if (gtype.ToUpper() == "FIRSTLEGBPS")
                    {
                        ret = (inst as GenericSwap).firstLegBPS();
                    }
                    else if (gtype.ToUpper() == "SECONDLEGBPS")
                    {
                        ret = (inst as GenericSwap).secondLegBPS();
                    }
                }

                return ret;
            }
            catch (Exception e)
            {
                ExcelUtil.logError(callerAddress, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), e.Message);
                return "#EQ_ERR!";
            }
        }


        [ExcelFunction(Description = "Display energy swap structure", Category = "EliteQuantExcel - Instruments")]
        public static object eqInstDisplayEnergySwap(
            [ExcelArgument(Description = "id of energy ")] string ObjectId,
            [ExcelArgument(Description = "trigger ")]object trigger)
        {
            if (ExcelUtil.CallFromWizard())
                return "";

            string callerAddress = "";
            callerAddress = ExcelUtil.getActiveCellAddress();

            try
            {
                Xl.Range rng = ExcelUtil.getActiveCellRange();

                EnergyCommodityExt inst = OHRepository.Instance.getObject<EnergyCommodityExt>(ObjectId);
                if (inst.GetType() == typeof(EnergyFutureExt))
                {
                    PricingPeriodExt pp = (inst as EnergyFutureExt).pricingPeriod();

                    object[,] ret = new object[2, 10];
                    ret[0, 0] = "PayDate"; ret[0, 1] = "Quantity"; ret[0, 2] = "PayCoeff"; ret[0, 3] = "RecCoeff";
                    ret[0, 4] = "PaySprd"; ret[0, 5] = "RecSprd";
                    ret[0, 6] = "PayDelta"; ret[0, 7] = "DiscountedPayDelta";
                    ret[0, 8] = "RecDelta"; ret[0, 9] = "DiscountedRecDelta";

                    ret[1, 0] = pp.paymentDate().serialNumber();
                    ret[1, 1] = pp.quantity();
                    ret[1, 2] = pp.payCoeff();
                    ret[1, 3] = pp.recCoeff();
                    ret[1, 4] = pp.paySpread();
                    ret[1, 5] = pp.recSpread();
                    ret[1, 6] = pp.getuPayDelta();
                    ret[1, 7] = pp.getdPayDelta();
                    ret[1, 8] = pp.getuRecDelta();
                    ret[1, 9] = pp.getdRecDelta();

                    return ret;
                }
                else
                {
                    double npv = inst.NPV();
                    PricingPeriodExts pps = (inst as EnergySwapExt).pricingPeriods();

                    object[,] ret = new object[pps.Count + 1, 10];

                    ret[0, 0] = "PayDate"; ret[0, 1] = "Quantity"; ret[0, 2] = "PayCoeff"; ret[0, 3] = "RecCoeff";
                    ret[0, 4] = "PaySprd"; ret[0, 5] = "RecSprd";
                    ret[0, 6] = "PayDelta"; ret[0, 7] = "DiscountedPayDelta";
                    ret[0, 8] = "RecDelta"; ret[0, 9] = "DiscountedRecDelta";
                    for (int i = 0; i < pps.Count; i++)
                    {
                        ret[i + 1, 0] = pps[i].paymentDate().serialNumber();
                        ret[i + 1, 1] = pps[i].quantity();
                        ret[i + 1, 2] = pps[i].payCoeff();
                        ret[i + 1, 3] = pps[i].recCoeff();
                        ret[i + 1, 4] = pps[i].paySpread();
                        ret[i + 1, 5] = pps[i].recSpread();
                        ret[i + 1, 6] = pps[i].getuPayDelta();
                        ret[i + 1, 7] = pps[i].getdPayDelta();
                        ret[i + 1, 8] = pps[i].getuRecDelta();
                        ret[i + 1, 9] = pps[i].getdRecDelta();
                    }

                    return ret;
                }         
            }
            catch (Exception e)
            {
                ExcelUtil.logError(callerAddress, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), e.Message);
                return "#EQ_ERR!";
            }
        }
    }
}
