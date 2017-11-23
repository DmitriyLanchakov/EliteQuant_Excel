using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;
using System.Threading.Tasks;

using EliteQuant;

namespace EliteQuant
{
    // one factor simulation e.g. short rate
    // It serves as MonteCarloModel + MCSimulation in pricing
    public class RiskHullWhiteSimulationEngine
    {
        /// <summary>
        /// Pass in market data on asOfDate
        /// </summary>
        /// <param name="asOfDate"></param>
        /// <param name="yieldcurve"></param>
        /// <param name="a"></param>
        /// <param name="sigma"></param>
        public RiskHullWhiteSimulationEngine(DateTime asOfDate, YieldTermStructureHandle discountcurve, double a, double sigma, InstrumentVector instruments,
            IborIndex usd3mIndex, RelinkableYieldTermStructureHandle usdForwardingTSHandle)
        {
            asOfDate_ = calendar_.adjust(new Date((int)asOfDate.ToOADate()), BusinessDayConvention.ModifiedFollowing);
            terminationDate_ = calendar_.advance(asOfDate_, 1, TimeUnit.Years);        // 15 years simulation length

            a_ = a; sigma_ = sigma;
            discountCurve_ = discountcurve;
            usdIndex_ = usd3mIndex;
            usdForwardingTSHandle_ = usdForwardingTSHandle;
            process_ = new HullWhiteProcess(discountCurve_, a, sigma);

            Settings.instance().includeTodaysCashFlows(false);
            Settings.instance().includeReferenceDateEvents(false);
            Settings.instance().setEvaluationDate(asOfDate_);
            settingsToday_ = Settings.instance().getEvaluationDate();

            // setup grids
            schedule_ = new Schedule(asOfDate_, terminationDate_, new Period(Frequency.Monthly), calendar_, BusinessDayConvention.ModifiedFollowing,
                BusinessDayConvention.ModifiedFollowing, DateGeneration.Rule.Forward, endOfMonth_);
            DoubleVector simtimes = new DoubleVector();
            for (uint i = 0; i < schedule_.size(); i++)
            {
                simtimes.Add(daycounter_.yearFraction(asOfDate_, schedule_.date(i)));
            }
            timeGrid_ = new TimeGrid(simtimes);

            int seed = 1234;
            uniformRng_ = new UniformRandomGenerator(seed);
            // timegrid from 0 to 30Y, rsg dinension is size -1, excluding time 0;
            uniformRsg_ = new UniformRandomSequenceGenerator(timeGrid_.size() - 1, uniformRng_);
            gaussianRsg_ = new GaussianRandomSequenceGenerator(uniformRsg_);
            pathGenerator_ = new GaussianPathGenerator(process_, timeGrid_, gaussianRsg_, false);

            portfolio_ = instruments;
        }

        /// <summary>
        /// # of paths
        /// </summary>
        private int sampleSize_ = 1000;
        public int SampleSize
        {
            get { return sampleSize_; }
            set { sampleSize_ = value; }
        }
        private bool isAntithetic_ = false;
        public bool IsAntithetic
        {
            get {return isAntithetic_;}
            set {isAntithetic_ = value;}
        }

        private UniformRandomGenerator uniformRng_;
        private UniformRandomSequenceGenerator uniformRsg_;
        private GaussianRandomSequenceGenerator gaussianRsg_;
        private GaussianPathGenerator pathGenerator_;

        private Date asOfDate_;
        private Date settingsToday_;
        private Date terminationDate_;
        private double a_, sigma_;
        private YieldTermStructureHandle discountCurve_;
        private IborIndex usdIndex_;               // to change the forwarding index
        private RelinkableYieldTermStructureHandle usdForwardingTSHandle_;        // forwarding index
        private HullWhiteProcess process_;
        private HullWhite model_;

        private TimeGrid timeGrid_;
        bool endOfMonth_ = true;
        
        private InstrumentVector portfolio_;
        private Dictionary<int, List<List<double>>> pfeMatrix_;                     // instrument --> samplesize * steps, last one is portfolio

        private Calendar calendar_ = new UnitedStates(UnitedStates.Market.NYSE);        // NYSE for now
        private DayCounter daycounter_ = new ActualActual(ActualActual.Convention.Actual365);   // ActualActual for now
        private Schedule schedule_;

        public void calculate()
        {
            model_ = new HullWhite(discountCurve_, a_, sigma_);
            // initialize data structure
            // this includes time 0
            pfeMatrix_ = new Dictionary<int,List<List<double>>>(portfolio_.Count+1);        // last one is portfolio
            int nSteps = (int)timeGrid_.size();             // with mandatory steps, this may not equal to simsteps+1

            double npvPortfolio = 0;
            for (int v = 0; v < portfolio_.Count; v++)          // for each asset
            {
                DiscountingSwapEngine pricingEngine_ = new DiscountingSwapEngine(discountCurve_);
                portfolio_[v].setPricingEngine(pricingEngine_);
                double npv = portfolio_[v].NPV();
                npvPortfolio += npv;

                List<List<double>> oneDeal = new List<List<double>>(sampleSize_);
                for (int j = 0; j < sampleSize_; j++)
                {
                    List<double> oneDealOnePath = new List<double>(nSteps);
                    oneDealOnePath.Add(npv);
                    oneDeal.Add(oneDealOnePath);
                }
                pfeMatrix_.Add(v, oneDeal);
            }
            // portfolio
            List<List<double>> onePort = new List<List<double>>(sampleSize_);
            for (int j = 0; j < sampleSize_; j++)
            {
                List<double> onePortOnePath = new List<double>(nSteps);
                onePortOnePath.Add(npvPortfolio);
                onePort.Add(onePortOnePath);
            }
            pfeMatrix_.Add(portfolio_.Count, onePort);
            // add fixing
            usdIndex_.addFixing(asOfDate_, discountCurve_.zeroRate(0.25, Compounding.Simple).rate());

            // For each sample path
            for (int i = 0; i < sampleSize_; i++)
            {
                SamplePath path = pathGenerator_.next();
                
                int steps = (int)path.value().timeGrid().size();
                Debug.WriteLine("Path time grid size is: " + steps +
                    " from time " + path.value().time(0) +
                    " to time " + path.value().time((uint)steps-1));

                // for each simulation step
                for (uint j = 1; j < nSteps; j++ )
                {
                    Date d0 = schedule_.date(j);        // new start date
                    Date dT = calendar_.advance(d0, 30, TimeUnit.Years);
                    double r0 = path.value().value(j);      // new short rate


                    Schedule sch = new Schedule(d0, dT, new Period(Frequency.Monthly), calendar_,
                        BusinessDayConvention.ModifiedFollowing, BusinessDayConvention.ModifiedFollowing, DateGeneration.Rule.Forward, true);
                    
                    DateVector x = new DateVector();        // time
                    DoubleVector y = new DoubleVector();    // discount
                    for (uint ts = 0; ts < sch.size(); ts++)
                    {
                        x.Add(sch.date(ts));
                        y.Add(model_.discountBond(daycounter_.yearFraction(d0,d0),
                            daycounter_.yearFraction(d0, sch.date(ts)), r0));
                    }
                    
                    // 1. construct discount yield curve
                    YieldTermStructure newyieldcurve = new DiscountCurve(x, y, daycounter_);        // LogLinear by default
                    YieldTermStructureHandle newyieldcurvehandle = new YieldTermStructureHandle(newyieldcurve);
                    //forwardingTS_.linkTo(newyieldcurve);
                    DiscountingSwapEngine pricingEngine_ = new DiscountingSwapEngine(newyieldcurvehandle);
                    // 2. reprice each deal
                    Settings.instance().setEvaluationDate(d0);
                    double discount = discountCurve_.discount(d0);           // discount. Used for CVA/DVA

                    double sum = 0;
                    for (int v = 0; v < portfolio_.Count; v++)
                    {
                        portfolio_[v].setPricingEngine(pricingEngine_);
                        // add to path matrix
                        pfeMatrix_[v][i].Add(portfolio_[v].NPV());
                        sum += portfolio_[v].NPV();
                    }
                    pfeMatrix_[portfolio_.Count][i].Add(sum);

                    // clear fixings
                    usdIndex_.addFixing(asOfDate_, newyieldcurve.zeroRate(0.25, Compounding.Simple).rate());
                }   // for each simulation step j
            }   // for each sample path i

            // restore settings
            Settings.instance().setEvaluationDate(settingsToday_);
        } // function end

        public Dictionary<Date, double> getPFECurve(double percentile)
        {
            Dictionary<Date, double> ret = new Dictionary<Date, double>((int)schedule_.size());
            Statistics statistics_ = new Statistics();

            for (int i = 0; i < schedule_.size(); i++)
            {
                statistics_.reset();
                for (int j = 0; j < sampleSize_; j++ )
                {
                    statistics_.add(pfeMatrix_[portfolio_.Count][j][i]);
                }
                
                ret.Add(schedule_.date((uint)i), statistics_.percentile(percentile));
            }
            
            return ret;
        }
    }
}
