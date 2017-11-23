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
    public class OptionPricer
    {
        #region European Option
        [ExcelFunction(Description = "Vanilla Option ", Category = "EliteQuantExcel - Instruments")]
        public static string eqInstVanillaOption(
            [ExcelArgument(Description = "id of option to be constructed ")] string ObjectId,
            [ExcelArgument(Description = "Option type E(uropean), A(merican), B(ermudan) ")]string exercisetype,
            [ExcelArgument(Description = "CALL or PUT ")]string optype,
            [ExcelArgument(Description = "Strike price ")]double strikeprice,
            [ExcelArgument(Description = "Expiry Dates (E = 1, A = 2, B = many) ")]object[] dates)
        {
            if (ExcelUtil.CallFromWizard())
                return "";

            string callerAddress = "";            
            callerAddress = ExcelUtil.getActiveCellAddress();

            try
            {
                Exercise exercise = null;
                if (exercisetype.ToUpper() == "E")
                {
                    Date maturitydate = EliteQuant.EQConverter.DateTimeToDate(DateTime.FromOADate((double)dates[0]));     // assume first date
                    exercise = new EuropeanExercise(maturitydate);
                }
                else if (exercisetype.ToUpper() == "A")
                {
                    Date earliestdate  = EliteQuant.EQConverter.DateTimeToDate(DateTime.FromOADate((double)dates[0]));     // assume first date
                    Date lastdate  = EliteQuant.EQConverter.DateTimeToDate(DateTime.FromOADate((double)dates[1]));     // assume last date
                    exercise = new AmericanExercise(earliestdate, lastdate);
                }
                else if (exercisetype.ToUpper() == "B")
                {
                    DateVector dv = new DateVector();
                    foreach (var dt in dates)
                    {
                        Date dte = EliteQuant.EQConverter.DateTimeToDate(DateTime.FromOADate((double)dt));
                        dv.Add(dte);
                    }
                    exercise = new BermudanExercise(dv);
                }
                else
                    throw new Exception("Unknow exercise type ");

                Option.Type optiontype;
                if (optype.ToUpper() == "CALL")
                    optiontype = Option.Type.Call;
                else if (optype.ToUpper() == "PUT")
                    optiontype = Option.Type.Put;
                else
                    throw new Exception("Unknow option type");

                PlainVanillaPayoff payoff = new PlainVanillaPayoff(optiontype, strikeprice);

                VanillaOption europeanOption = new VanillaOption(payoff, exercise);

                // Store the option and return its id
                string id = "OPTION@" + ObjectId;
                OHRepository.Instance.storeObject(id, europeanOption, callerAddress);
                id += "#" + (String)DateTime.Now.ToString(@"HH:mm:ss");
                return id;
            }
            catch (Exception e)
            {
                ExcelUtil.logError(callerAddress, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), e.Message);
                return "#EQ_ERR!";
            }
        }

        /*
        [ExcelFunction(Description = "Black Scholes  ", Category = "EliteQuantExcel - Instruments")]
        public static string eqPricerAnalyticEuropeanEngine(
            [ExcelArgument(Description = "id of option option pricer ")] string ObjectId)
        {
            if (SystemUtil.CallFromWizard())
                return "";

            string callerAddress = "";            
            callerAddress = SystemUtil.getActiveCellAddress();

            try
            {
                if (date == DateTime.MinValue)
                    throw new Exception("Date must not be empty. ");

                if (SystemUtil.isNull(daycounter))
                    daycounter = "ACTUAL365";
                if (SystemUtil.isNull(calendar))
                    calendar = "NYC";

                Option.Type optiontype;
                if (callput.ToUpper() == "CALL")
                    optiontype = Option.Type.Call;
                else if (callput.ToUpper() == "PUT")
                    optiontype = Option.Type.Put;
                else
                    throw new Exception("Unknow option type");


                BlackScholesMertonProcess stochasticProcess = new BlackScholesMertonProcess(underlyingQuoteH,
                    flatDividendTSH, flatRateTSH, flatVolTSH);

                PricingEngine engine = new AnalyticEuropeanEngine(stochasticProcess);

                // Store the pricer and return its id
                string id = "PRICER@" + ObjectId;
                OHRepository.Instance.storeObject(id, europeanOption, callerAddress);
                id += "#" + (String)DateTime.Now.ToString(@"HH:mm:ss");
                return id;
            }
            catch (Exception e)
            {
                SystemUtil.logError(callerAddress, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), e.Message);
                return "#EQ_ERR!";
            }
        }*/


        [ExcelFunction(Description = "European Option with Black Scholes Pricer", Category = "EliteQuantExcel - Instruments")]
        public static string eqInstEuropeanOptionBlackScholes(
            [ExcelArgument(Description = "id of option to be constructed ")] string ObjectId,
            [ExcelArgument(Description = "Option type ")]string optype,
            [ExcelArgument(Description = "Spot price ")]double underlyingprice,
            [ExcelArgument(Description = "Strike price ")]double stirkeprice,
            [ExcelArgument(Description = "Expiry Date ")]DateTime date,
            [ExcelArgument(Description = "Risk free rate ")]double riskfreerate,
            [ExcelArgument(Description = "dividend/convenience rate ")]double dividendrate,
            [ExcelArgument(Description = "Black-Scholes Vol ")]double volatility,
            [ExcelArgument(Description = "DayCounter ")]string daycounter,
            [ExcelArgument(Description = "Calendar ")]string calendar,
            [ExcelArgument(Description = "trigger ")]object trigger)
        {
            if (ExcelUtil.CallFromWizard())
                return "";

            string callerAddress = "";            
            callerAddress = ExcelUtil.getActiveCellAddress();

            try
            {
                if (date == DateTime.MinValue)
                    throw new Exception("Date must not be empty. ");

                if (ExcelUtil.isNull(daycounter))
                    daycounter = "ACTUAL365";
                if (ExcelUtil.isNull(calendar))
                    calendar = "NYC";

                Option.Type optiontype;
                if (optype.ToUpper() == "CALL")
                    optiontype = Option.Type.Call;
                else if (optype.ToUpper() == "PUT")
                    optiontype = Option.Type.Put;
                else
                    throw new Exception("Unknow option type");

                EliteQuant.Calendar cal = EliteQuant.EQConverter.ConvertObject<EliteQuant.Calendar>(calendar);
                EliteQuant.DayCounter dc = EliteQuant.EQConverter.ConvertObject<EliteQuant.DayCounter>(daycounter);
                EliteQuant.Date maturitydate = EliteQuant.EQConverter.ConvertObject<EliteQuant.Date>(date);

                EliteQuant.Date today = EliteQuant.Settings.instance().getEvaluationDate();
                EliteQuant.Date settlementdate = today;           // T+2
                if (maturitydate.serialNumber() <= today.serialNumber())
                    throw new Exception("Option already expired.");

                EuropeanExercise europeanExercise = new EuropeanExercise(maturitydate);

                QuoteHandle underlyingQuoteH = new QuoteHandle(new EliteQuant.SimpleQuote(underlyingprice));

                YieldTermStructureHandle flatRateTSH = new YieldTermStructureHandle(
                    new FlatForward(settlementdate, riskfreerate, dc));
                YieldTermStructureHandle flatDividendTSH = new YieldTermStructureHandle(
                    new FlatForward(settlementdate, dividendrate, dc));

                BlackVolTermStructureHandle flatVolTSH = new BlackVolTermStructureHandle(
                    new BlackConstantVol(settlementdate, cal, volatility, dc));

                BlackScholesMertonProcess stochasticProcess = new BlackScholesMertonProcess(underlyingQuoteH,
                    flatDividendTSH, flatRateTSH, flatVolTSH);

                PlainVanillaPayoff payoff = new PlainVanillaPayoff(optiontype, stirkeprice);

                VanillaOption europeanOption = new VanillaOption(payoff, europeanExercise);

                europeanOption.setPricingEngine(new AnalyticEuropeanEngine(stochasticProcess));

                // Store the option and return its id
                string id = "OPTION@" + ObjectId;
                OHRepository.Instance.storeObject(id, europeanOption, callerAddress);
                id += "#" + (String)DateTime.Now.ToString(@"HH:mm:ss");
                return id;
            }
            catch (Exception e)
            {
                ExcelUtil.logError(callerAddress, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), e.Message);
                return "#EQ_ERR!";
            }
        }
        #endregion

        #region American Option
        [ExcelFunction(Description = "American Option with BAW approximate Pricer", Category = "EliteQuantExcel - Instruments")]
        public static string eqInstAmericanOptionBaroneAdesiWhaley(
            [ExcelArgument(Description = "id of option to be constructed ")] string ObjectId,
            [ExcelArgument(Description = "Option type ")]string optype,
            [ExcelArgument(Description = "Spot price ")]double underlyingprice,
            [ExcelArgument(Description = "Strike price ")]double stirkeprice,
            [ExcelArgument(Description = "Expiry Date ")]DateTime date,
            [ExcelArgument(Description = "Risk free rate ")]double riskfreerate,
            [ExcelArgument(Description = "dividend/convenience rate ")]double dividendrate,
            [ExcelArgument(Description = "Black-Scholes Vol ")]double volatility,
            [ExcelArgument(Description = "DayCounter ")]string daycounter,
            [ExcelArgument(Description = "Calendar ")]string calendar,
            [ExcelArgument(Description = "trigger ")]object trigger)
        {
            if (ExcelUtil.CallFromWizard())
                return "";

            string callerAddress = "";
            callerAddress = ExcelUtil.getActiveCellAddress();

            try
            {
                if (date == DateTime.MinValue)
                    throw new Exception("Date must not be empty. ");

                if (ExcelUtil.isNull(daycounter))
                    daycounter = "ACTUAL365";
                if (ExcelUtil.isNull(calendar))
                    calendar = "NYC";

                Option.Type optiontype;
                if (optype.ToUpper() == "CALL")
                    optiontype = Option.Type.Call;
                else if (optype.ToUpper() == "PUT")
                    optiontype = Option.Type.Put;
                else
                    throw new Exception("Unknow option type");

                EliteQuant.Calendar cal = EliteQuant.EQConverter.ConvertObject<EliteQuant.Calendar>(calendar);
                EliteQuant.DayCounter dc = EliteQuant.EQConverter.ConvertObject<EliteQuant.DayCounter>(daycounter);
                EliteQuant.Date maturitydate = EliteQuant.EQConverter.ConvertObject<EliteQuant.Date>(date);

                EliteQuant.Date today = EliteQuant.Settings.instance().getEvaluationDate();
                EliteQuant.Date settlementdate = today;           // T+2
                if (maturitydate.serialNumber() <= today.serialNumber())
                    throw new Exception("Option already expired.");

                AmericanExercise americanExercise = new AmericanExercise(today, maturitydate);

                QuoteHandle underlyingQuoteH = new QuoteHandle(new EliteQuant.SimpleQuote(underlyingprice));

                YieldTermStructureHandle flatRateTSH = new YieldTermStructureHandle(
                    new FlatForward(settlementdate, riskfreerate, dc));
                YieldTermStructureHandle flatDividendTSH = new YieldTermStructureHandle(
                    new FlatForward(settlementdate, dividendrate, dc));

                BlackVolTermStructureHandle flatVolTSH = new BlackVolTermStructureHandle(
                    new BlackConstantVol(settlementdate, cal, volatility, dc));

                BlackScholesMertonProcess stochasticProcess = new BlackScholesMertonProcess(underlyingQuoteH,
                    flatDividendTSH, flatRateTSH, flatVolTSH);

                PlainVanillaPayoff payoff = new PlainVanillaPayoff(optiontype, stirkeprice);

                VanillaOption europeanOption = new VanillaOption(payoff, americanExercise);

                europeanOption.setPricingEngine(new BaroneAdesiWhaleyEngine(stochasticProcess));

                // Store the option and return its id
                string id = "OPTION@" + ObjectId;
                OHRepository.Instance.storeObject(id, europeanOption, callerAddress);
                id += "#" + (String)DateTime.Now.ToString(@"HH:mm:ss");
                return id;
            }
            catch (Exception e)
            {
                ExcelUtil.logError(callerAddress, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), e.Message);
                return "#EQ_ERR!";
            }
        }
        #endregion

        #region Bermudan Option
        #endregion

        #region Asian Option
        #endregion

        #region Basket/Spread Option
        [ExcelFunction(Description = "European Spread Option with Kirk Pricer", Category = "EliteQuantExcel - Instruments")]
        public static string eqInstSpreadOptionKirk(
            [ExcelArgument(Description = "id of option to be constructed ")] string ObjectId,
            [ExcelArgument(Description = "Option type (Call/Put) ")]string optype,
            [ExcelArgument(Description = "Spot price leg 1")]double spot1,
            [ExcelArgument(Description = "Spot price leg 2")]double spot2,
            [ExcelArgument(Description = "Strike price ")]double stirkeprice,
            [ExcelArgument(Description = "Expiry Date ")]DateTime exdate,
            [ExcelArgument(Description = "Risk free rate ")]double riskfreerate,
            [ExcelArgument(Description = "Black-Scholes Vol for leg 1 ")]double vol1,
            [ExcelArgument(Description = "Black-Scholes Vol for leg 2 ")]double vol2,
            [ExcelArgument(Description = "correlation between leg 1 and leg 2 ")]double corr,
            [ExcelArgument(Description = "DayCounter ")]string daycounter,
            [ExcelArgument(Description = "Calendar ")]string calendar,
            [ExcelArgument(Description = "trigger ")]object trigger)
        {
            if (ExcelUtil.CallFromWizard())
                return "";

            string callerAddress = "";
            callerAddress = ExcelUtil.getActiveCellAddress();

            try
            {
                if (exdate == DateTime.MinValue)
                    throw new Exception("Date must not be empty. ");

                if (ExcelUtil.isNull(daycounter))
                    daycounter = "ACTUAL365";
                if (ExcelUtil.isNull(calendar))
                    calendar = "NYC";

                Option.Type optiontype;
                if (optype.ToUpper() == "CALL")
                    optiontype = Option.Type.Call;
                else if (optype.ToUpper() == "PUT")
                    optiontype = Option.Type.Put;
                else
                    throw new Exception("Unknow option type");

                EliteQuant.Calendar cal = EliteQuant.EQConverter.ConvertObject<EliteQuant.Calendar>(calendar);
                EliteQuant.DayCounter dc = EliteQuant.EQConverter.ConvertObject<EliteQuant.DayCounter>(daycounter);
                EliteQuant.Date maturitydate = EliteQuant.EQConverter.ConvertObject<EliteQuant.Date>(exdate);

                EliteQuant.Date today = EliteQuant.Settings.instance().getEvaluationDate();
                EliteQuant.Date settlementdate = today;           // T+2
                if (maturitydate.serialNumber() <= today.serialNumber())
                    throw new Exception("Option already expired.");


                YieldTermStructureHandle rTSH = new YieldTermStructureHandle(
                    new FlatForward(settlementdate, riskfreerate, dc));
                BlackVolTermStructureHandle flatVolTSH1 = new BlackVolTermStructureHandle(
                    new BlackConstantVol(settlementdate, cal, vol1, dc));
                BlackVolTermStructureHandle flatVolTSH2 = new BlackVolTermStructureHandle(
                    new BlackConstantVol(settlementdate, cal, vol2, dc));

                Quote qh1 = new SimpleQuote(spot1);
                Quote qh2 = new SimpleQuote(spot2);
                QuoteHandle s1 = new QuoteHandle(qh1);
                QuoteHandle s2 = new QuoteHandle(qh2);

                BlackProcess p1 = new BlackProcess(s1, rTSH, flatVolTSH1);
                BlackProcess p2 = new BlackProcess(s2, rTSH, flatVolTSH2);

                Payoff payoff1 = new PlainVanillaPayoff(optiontype, stirkeprice);
                Payoff payoff2 = new SpreadBasketPayoff(payoff1);

                Exercise exercise = new EuropeanExercise(maturitydate);

                PricingEngine engine = new KirkEngine(p1, p2, corr);
                
                BasketOption bo = new BasketOption(payoff2, exercise);

                bo.setPricingEngine(engine);

                // Store the option and return its id
                string id = "OPTION@" + ObjectId;
                OHRepository.Instance.storeObject(id, bo, callerAddress);
                id += "#" + (String)DateTime.Now.ToString(@"HH:mm:ss");
                return id;
            }
            catch (Exception e)
            {
                ExcelUtil.logError(callerAddress, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), e.Message);
                return "#EQ_ERR!";
            }
        }

        [ExcelFunction(Description = "European Spread Option with Kirk Pricer", Category = "EliteQuantExcel - Instruments")]
        public static string eqInstSpreadOptionMonteCarlo(
            [ExcelArgument(Description = "id of option to be constructed ")] string ObjectId,
            [ExcelArgument(Description = "Option type (Call/Put) ")]string optype,
            [ExcelArgument(Description = "Spot price leg 1")]double spot1,
            [ExcelArgument(Description = "Spot price leg 2")]double spot2,
            [ExcelArgument(Description = "Strike price ")]double stirkeprice,
            [ExcelArgument(Description = "Expiry Date ")]DateTime exdate,
            [ExcelArgument(Description = "Risk free rate ")]double riskfreerate,
            [ExcelArgument(Description = "Black-Scholes Vol for leg 1 ")]double vol1,
            [ExcelArgument(Description = "Black-Scholes Vol for leg 2 ")]double vol2,
            [ExcelArgument(Description = "correlation between leg 1 and leg 2 ")]double corr,
            [ExcelArgument(Description = "DayCounter ")]string daycounter,
            [ExcelArgument(Description = "Calendar ")]string calendar,
            [ExcelArgument(Description = "Pseudorandom (pr) or lowdiscrepancy (ld) ")]string traits,
            [ExcelArgument(Description = "trigger ")]object trigger)
        {
            if (ExcelUtil.CallFromWizard())
                return "";

            string callerAddress = "";
            callerAddress = ExcelUtil.getActiveCellAddress();

            try
            {
                if (exdate == DateTime.MinValue)
                    throw new Exception("Date must not be empty. ");

                if (ExcelUtil.isNull(daycounter))
                    daycounter = "ACTUAL365";
                if (ExcelUtil.isNull(calendar))
                    calendar = "NYC";
                if (ExcelUtil.isNull(traits))
                    traits = "pr";

                Option.Type optiontype;
                if (optype.ToUpper() == "CALL")
                    optiontype = Option.Type.Call;
                else if (optype.ToUpper() == "PUT")
                    optiontype = Option.Type.Put;
                else
                    throw new Exception("Unknow option type");

                EliteQuant.Calendar cal = EliteQuant.EQConverter.ConvertObject<EliteQuant.Calendar>(calendar);
                EliteQuant.DayCounter dc = EliteQuant.EQConverter.ConvertObject<EliteQuant.DayCounter>(daycounter);
                EliteQuant.Date maturitydate = EliteQuant.EQConverter.ConvertObject<EliteQuant.Date>(exdate);

                EliteQuant.Date today = EliteQuant.Settings.instance().getEvaluationDate();
                EliteQuant.Date settlementdate = today;           // T+2
                if (maturitydate.serialNumber() <= today.serialNumber())
                    throw new Exception("Option already expired.");


                YieldTermStructureHandle rTSH = new YieldTermStructureHandle(
                    new FlatForward(settlementdate, riskfreerate, dc));
                BlackVolTermStructureHandle flatVolTSH1 = new BlackVolTermStructureHandle(
                    new BlackConstantVol(settlementdate, cal, vol1, dc));
                BlackVolTermStructureHandle flatVolTSH2 = new BlackVolTermStructureHandle(
                    new BlackConstantVol(settlementdate, cal, vol2, dc));

                Quote qh1 = new SimpleQuote(spot1);
                Quote qh2 = new SimpleQuote(spot2);
                QuoteHandle s1 = new QuoteHandle(qh1);
                QuoteHandle s2 = new QuoteHandle(qh2);

                BlackProcess p1 = new BlackProcess(s1, rTSH, flatVolTSH1);
                BlackProcess p2 = new BlackProcess(s2, rTSH, flatVolTSH2);

                StochasticProcessVector spv = new StochasticProcessVector(2);
                spv.Add(p1);
                spv.Add(p2);

                Matrix corrmtrx = new Matrix(2,2);
                corrmtrx.set(0, 0, 1.0); corrmtrx.set(1, 1, 1.0);
                corrmtrx.set(0, 1, corr); corrmtrx.set(1, 0, corr); 
                StochasticProcessArray spa = new StochasticProcessArray(spv, corrmtrx);

                PricingEngine engine = new MCEuropeanBasketEngine(spa, traits, 100, 1, false, true, 5000, 1e-6);

                Payoff payoff1 = new PlainVanillaPayoff(optiontype, stirkeprice);
                Payoff payoff2 = new SpreadBasketPayoff(payoff1);

                Exercise exercise = new EuropeanExercise(maturitydate);

                BasketOption bo = new BasketOption(payoff2, exercise);

                bo.setPricingEngine(engine);

                // Store the option and return its id
                string id = "OPTION@" + ObjectId;
                OHRepository.Instance.storeObject(id, bo, callerAddress);
                id += "#" + (String)DateTime.Now.ToString(@"HH:mm:ss");
                return id;
            }
            catch (Exception e)
            {
                ExcelUtil.logError(callerAddress, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), e.Message);
                return "#EQ_ERR!";
            }
        }
        #endregion

        [ExcelFunction(Description = "European Option with Black Scholes Pricer", Category = "EliteQuantExcel - Instruments")]
        public static object eqInstGetOptionGreeks(
            [ExcelArgument(Description = "id of option ")] string ObjectId,
            [ExcelArgument(Description = "Greek type ")]string gtype,
            [ExcelArgument(Description = "Option type (VANILLA or MULTIASSET)")] string otype,
            [ExcelArgument(Description = "trigger ")]object trigger)
        {
            if (ExcelUtil.CallFromWizard())
                return "";

            string callerAddress = "";            
            callerAddress = ExcelUtil.getActiveCellAddress();

            try
            {
                Xl.Range rng = ExcelUtil.getActiveCellRange();

                if (ExcelUtil.isNull(gtype))
                    gtype = "NPV";
                if (ExcelUtil.isNull(otype))
                    otype = "VANILLA";

                if (otype == "VANILLA")
                {
                    VanillaOption option = OHRepository.Instance.getObject<VanillaOption>(ObjectId);
                    switch (gtype.ToUpper())
                    {
                        case "NPV":
                            return option.NPV();
                        case "DELTA":
                            return option.delta();
                        case "GAMMA":
                            return option.gamma();
                        case "VEGA":
                            return option.vega();
                        case "THETA":
                            return option.theta();
                        case "RHO":
                            return option.rho();
                        default:
                            return 0;
                    }
                }
                else if (otype == "MULTIASSET")
                {
                    BasketOption option = OHRepository.Instance.getObject<BasketOption>(ObjectId);
                    switch (gtype.ToUpper())
                    {
                        case "NPV":
                            return option.NPV();
                        case "DELTA":
                            return option.delta();
                        case "GAMMA":
                            return option.gamma();
                        case "VEGA":
                            return option.vega();
                        case "THETA":
                            return option.theta();
                        case "RHO":
                            return option.rho();
                        default:
                            return 0;
                    }
                }
                else
                {
                    return "Unknown option type";
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
