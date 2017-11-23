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
    public class Rates
    {
        #region Interest Rate
        [ExcelFunction(Description = "Interest Rate compounding factor", Category = "EliteQuantExcel - Rates")]
        public static object eqRatesCompoundFactor(
            [ExcelArgument(Description = "interest rate ")] double r,
            [ExcelArgument(Description = "time in years ")] double t,
            [ExcelArgument(Description = "DayCounter (e.g. Actual365) ")] string dc,
            [ExcelArgument(Description = "Compounding (e.g. Continuous) ")] string comp,
            [ExcelArgument(Description = "Frequency (e.g. Annual) ")] string freq,
            [ExcelArgument(Description = "trigger ")]object trigger)
        {
            if (ExcelUtil.CallFromWizard())
                return "";

            string callerAddress = ExcelUtil.getActiveCellAddress();

            try
            {
                DayCounter daycounter = EliteQuant.EQConverter.ConvertObject<DayCounter>(dc);
                Compounding compounding = EliteQuant.EQConverter.ConvertObject<Compounding>(comp);
                Frequency frequency = EliteQuant.EQConverter.ConvertObject<Frequency>(freq);

                InterestRate ir = new InterestRate(r, daycounter, compounding, frequency);

                return ir.compoundFactor(t);
            }
            catch (Exception e)
            {
                ExcelUtil.logError(callerAddress, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), e.Message);
                return "#EQ_ERR!";
            }
        }

        [ExcelFunction(Description = "Interest Rate discount factor", Category = "EliteQuantExcel - Rates")]
        public static object eqRatesDiscountFactor(
            [ExcelArgument(Description = "interest rate ")] double r,
            [ExcelArgument(Description = "time in years ")] double t,
            [ExcelArgument(Description = "DayCounter (default Actual365) ")] string dc,
            [ExcelArgument(Description = "Compounding (default Continuous) ")] string comp,
            [ExcelArgument(Description = "Frequency (default Actual365) ")] string freq,
            [ExcelArgument(Description = "trigger ")]object trigger)
        {
            if (ExcelUtil.CallFromWizard())
                return "";

            string callerAddress = ExcelUtil.getActiveCellAddress();

            try
            {
                DayCounter daycounter = EliteQuant.EQConverter.ConvertObject<DayCounter>(dc);
                Compounding compounding = EliteQuant.EQConverter.ConvertObject<Compounding>(comp);
                Frequency frequency = EliteQuant.EQConverter.ConvertObject<Frequency>(freq);

                InterestRate ir = new InterestRate(r, daycounter, compounding, frequency);

                return ir.discountFactor(t);
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
