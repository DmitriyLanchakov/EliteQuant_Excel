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
    public class eqTime
    {
        [ExcelFunction(Description = "today's date (non volatile function)", Category = "EliteQuantExcel - Time")]
        public static Object eqTimeToday(bool withTime)
        {
            if (withTime)
                return DateTime.Now;
            else
                return DateTime.Today;
        }

        //[ExcelFunction(IsVolatile=true)]
        [ExcelFunction(Description = "set the evaluation date of the whole spreadsheet", Category = "EliteQuantExcel - Time")]
        public static object eqTimeSetEvaluationDate(
            [ExcelArgument(Description = "Evaluation Date ")]DateTime date)
        {
            if (ExcelUtil.CallFromWizard())
                return false;

            string callerAddress = "";
            callerAddress = ExcelUtil.getActiveCellAddress();

            EliteQuant.Date todaysDate; 
            try
            {
                if (date == DateTime.MinValue)
                    todaysDate = EliteQuant.EQConverter.ConvertObject<EliteQuant.Date>(DateTime.Today);
                else
                    todaysDate = EliteQuant.EQConverter.ConvertObject<EliteQuant.Date>(date);

                EliteQuant.Settings.instance().setEvaluationDate(todaysDate);

                return EliteQuant.EQConverter.ConvertObject<DateTime>(todaysDate);
            }
            catch (Exception e)
            {
                ExcelUtil.logError(callerAddress, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), e.Message);
                return false;
            }
        }

        [ExcelFunction(Description = "Get the evaluation date of the whole spreadsheet", Category = "EliteQuantExcel - Time")]
        public static object eqTimeGetEvaluationDate()
        {
            if (ExcelUtil.CallFromWizard())
                return false;

            string callerAddress = "";
            callerAddress = ExcelUtil.getActiveCellAddress();

            try
            {
                Date dt = EliteQuant.Settings.instance().getEvaluationDate();
                return EliteQuant.EQConverter.ConvertObject<DateTime>(dt);
            }
            catch (Exception e)
            {
                ExcelUtil.logError(callerAddress, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), e.Message);
                return e.Message;
            }
        }

        [ExcelFunction(Description = "calculate year fraction between two dates", Category = "EliteQuantExcel - Time")]
        public static object eqTimeYearFraction(
            [ExcelArgument(Description = "Start Date ")]DateTime date1,
            [ExcelArgument(Description = "End Date ")]DateTime date2,
            [ExcelArgument(Description = "Day Counter (default ActualActual) ")]string dc = "ActualActual")
        {
            if (ExcelUtil.CallFromWizard())
                return "";

            string callerAddress = "";
            callerAddress = ExcelUtil.getActiveCellAddress();
            OHRepository.Instance.removeErrorMessage(callerAddress);

            try
            {
                if ((date1 == DateTime.MinValue) || (date2 == DateTime.MinValue))
                    throw new Exception("Date must not be empty. ");

                Date start = EliteQuant.EQConverter.ConvertObject<Date>(date1);
                Date end = EliteQuant.EQConverter.ConvertObject<Date>(date2);
                DayCounter daycounter = EliteQuant.EQConverter.ConvertObject<DayCounter>(dc);
                return daycounter.yearFraction(start, end);
            }
            catch (Exception e)
            {
                ExcelUtil.logError(callerAddress, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), e.Message);
                return "";
            }
        }

        [ExcelFunction(Description = "business days between two dates (doesn't include these two days)", Category = "EliteQuantExcel - Time")]
        public static object eqTimeBusinessDaysBetween(
            [ExcelArgument(Description = "Start Date ")]DateTime date1,
            [ExcelArgument(Description = "End Date ")]DateTime date2,
            [ExcelArgument(Description = "Calendar (default NYC) ")]string calendar)
        {
            if (ExcelUtil.CallFromWizard())
                return "";

            string callerAddress = "";
            callerAddress = ExcelUtil.getActiveCellAddress();
            OHRepository.Instance.removeErrorMessage(callerAddress);

            try
            {
                if ((date1 == DateTime.MinValue) || (date2 == DateTime.MinValue))
                    throw new Exception("Date must not be empty. ");
                Date start = EliteQuant.EQConverter.ConvertObject<Date>(date1);
                Date end = EliteQuant.EQConverter.ConvertObject<Date>(date2);

                if (string.IsNullOrEmpty(calendar)) calendar = "NYC";
                EliteQuant.Calendar can = EliteQuant.EQConverter.ConvertObject<EliteQuant.Calendar>(calendar);
                
                return can.businessDaysBetween(start, end, false, false);
            }
            catch (Exception e)
            {
                ExcelUtil.logError(callerAddress, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), e.Message);
                return "";
            }
        }

        [ExcelFunction(Description = "check if the given date is a business day", Category = "EliteQuantExcel - Time")]
        public static bool eqTimeIsBusinessDay(
            [ExcelArgument(Description = "Date ")]DateTime date,
            [ExcelArgument(Description = "Calendar (default NYC) ")]string calendar)
        {
            if (ExcelUtil.CallFromWizard())
                return false;

            string callerAddress = "";
            try
            {
                callerAddress = ExcelUtil.getActiveCellAddress();

                OHRepository.Instance.removeErrorMessage(callerAddress);
            }
            catch (Exception)
            {
            }
            try
            {
                if (date == DateTime.MinValue)
                    throw new Exception("Date must not be empty. ");
                EliteQuant.Date d = EliteQuant.EQConverter.ConvertObject<EliteQuant.Date>(date);

                if (string.IsNullOrEmpty(calendar)) calendar = "NYC";
                EliteQuant.Calendar can = EliteQuant.EQConverter.ConvertObject<EliteQuant.Calendar>(calendar);
                
                return can.isBusinessDay(d);
            }
            catch (Exception e)
            {
                ExcelUtil.logError(callerAddress, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), e.Message);
                return false;
            }
        }

        [ExcelFunction(Description = "adjust a date to business day", Category = "EliteQuantExcel - Time")]
        public static object eqTimeAdjustDate(
            [ExcelArgument(Description = "Date ")]DateTime date,
            [ExcelArgument(Description = "Calendar (default NYC) ")]string calendar,
            [ExcelArgument(Description = "BusinessDayConvention (default ModifiedFollowing) ")]string bdc)
        {
            if (ExcelUtil.CallFromWizard())
                return "";

            string callerAddress = "";
            callerAddress = ExcelUtil.getActiveCellAddress();
            OHRepository.Instance.removeErrorMessage(callerAddress);

            try
            {
                if (date == DateTime.MinValue)
                    throw new Exception("Date must not be empty. ");
                EliteQuant.Date d = EliteQuant.EQConverter.ConvertObject<EliteQuant.Date>(date);

                if (string.IsNullOrEmpty(calendar)) calendar = "NYC";
                EliteQuant.Calendar can = EliteQuant.EQConverter.ConvertObject<EliteQuant.Calendar>(calendar);

                if (string.IsNullOrEmpty(bdc)) bdc = "MF";                
                BusinessDayConvention bdc2 = EliteQuant.EQConverter.ConvertObject<BusinessDayConvention>(bdc);

                Date newday = can.adjust(d, bdc2);
                return newday.serialNumber();
            }
            catch(TypeInitializationException e)
            {
                ExcelUtil.logError(callerAddress, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), e.Message);
                object[,] ret = new object[5,1];
                ret[0,1] = e.ToString();
                ret[1,1] = e.Message;
                ret[2,1] = e.StackTrace;
                ret[3,3] = e.Source;
                ret[4,1] = e.InnerException.Message;
                return ret;
            }
            catch (Exception e)
            {
                ExcelUtil.logError(callerAddress, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), e.Message);
                return e.Message;
            }
        }

        [ExcelFunction(Description = "advance forward a date acording to tenor given", Category = "EliteQuantExcel - Time")]
        public static object eqTimeAdvanceDate(
            [ExcelArgument(Description = "Date ")]DateTime date,
            [ExcelArgument(Description = "Calendar (default NYC) ")]string calendar,
            [ExcelArgument(Description = "Tenor (e.g. '3D' or '2Y') ")]string tenor,
            [ExcelArgument(Description = "BusinessDayConvention (default ModifiedFollowing) ")]string bdc,
            [ExcelArgument(Description = "is endofmonth ")]bool eom)
        {
            if (ExcelUtil.CallFromWizard())
                return "";

            string callerAddress = "";
            callerAddress = ExcelUtil.getActiveCellAddress();
            OHRepository.Instance.removeErrorMessage(callerAddress);

            try
            {
                if (date == DateTime.MinValue)
                    throw new Exception("Date must not be empty. ");
                EliteQuant.Date d = EliteQuant.EQConverter.ConvertObject<EliteQuant.Date>(date);

                if (string.IsNullOrEmpty(calendar)) calendar = "NYC";
                EliteQuant.Calendar can = EliteQuant.EQConverter.ConvertObject<EliteQuant.Calendar>(calendar);

                if (string.IsNullOrEmpty(tenor))
                    tenor = "1D";
                EliteQuant.Period period = EliteQuant.EQConverter.ConvertObject<EliteQuant.Period>(tenor);

                if (string.IsNullOrEmpty(bdc)) bdc = "MF";
                BusinessDayConvention bdc2 = EliteQuant.EQConverter.ConvertObject<BusinessDayConvention>(bdc);

                Date newday = can.advance(d, period, bdc2, eom);
                return newday.serialNumber();
            }
            catch (Exception e)
            {
                ExcelUtil.logError(callerAddress, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), e.Message);
                return e.Message;
            }
        }

        [ExcelFunction(Description = "return next IMM date ", Category = "EliteQuantExcel - Time")]
        public static object eqTimeNextIMMDate(
            [ExcelArgument(Description = "Date ")]DateTime date,
            [ExcelArgument(Description = "is main cycle ? ")]bool maincycle)
        {
            if (ExcelUtil.CallFromWizard())
                return "";

            string callerAddress = "";
            callerAddress = ExcelUtil.getActiveCellAddress();
            OHRepository.Instance.removeErrorMessage(callerAddress);

            try
            {
                if (date == DateTime.MinValue)
                    throw new Exception("Date must not be empty. ");
                EliteQuant.Date d = EliteQuant.EQConverter.ConvertObject<EliteQuant.Date>(date);

                EliteQuant.Date immdate = EliteQuant.IMM.nextDate(d, maincycle);

                return immdate.serialNumber();
            }
            catch (Exception e)
            {
                ExcelUtil.logError(callerAddress, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), e.Message);
                return e.Message;
            }
        }

        [ExcelFunction(Description = "time schedule betwwen two days according to tenor", Category = "EliteQuantExcel - Time")]
        public static object eqTimeSchedule(
            [ExcelArgument(Description = "Start Date ")]DateTime date1,
            [ExcelArgument(Description = "End Date ")]DateTime date2,
            [ExcelArgument(Description = "Tenor (e.g. '3D' or '2Y') ")]string tenor,
            [ExcelArgument(Description = "Calendar (default NYC) ")]string calendar,
            [ExcelArgument(Description = "BusinessDayConvention (default ModifiedFollowing) ")]string bdc,
            [ExcelArgument(Description = "DateGenerationRule (default Backward) ")]string rule,
            [ExcelArgument(Description = "is endofmonth ")]bool eom)
        {
            if (ExcelUtil.CallFromWizard())
                return "";

            string callerAddress = "";
            callerAddress = ExcelUtil.getActiveCellAddress();
            OHRepository.Instance.removeErrorMessage(callerAddress);

            try
            {
                if ((date1 == DateTime.MinValue) || (date2 == DateTime.MinValue))
                    throw new Exception("Date must not be empty. ");
                Date start = EliteQuant.EQConverter.ConvertObject<Date>(date1);
                Date end = EliteQuant.EQConverter.ConvertObject<Date>(date2);

                EliteQuant.Period period = EliteQuant.EQConverter.ConvertObject<EliteQuant.Period>(tenor);

                if (string.IsNullOrEmpty(calendar)) calendar = "NYC";
                EliteQuant.Calendar can = EliteQuant.EQConverter.ConvertObject<EliteQuant.Calendar>(calendar);

                if (string.IsNullOrEmpty(bdc)) bdc = "MF";
                BusinessDayConvention bdc2 = EliteQuant.EQConverter.ConvertObject<BusinessDayConvention>(bdc);

                if (string.IsNullOrEmpty(rule)) rule = "BACKWARD";
                DateGeneration.Rule rule2 = EliteQuant.EQConverter.ConvertObject<DateGeneration.Rule>(rule);

                Schedule sch = new Schedule(start, end, period, can, bdc2, bdc2, rule2, eom);

                object[,] ret = new object[sch.size(), 1];
                for (uint i = 0; i < sch.size(); i++)
                {
                    ret[i, 0] = sch.date(i).serialNumber();
                }

                return ret;
            }
            catch (Exception e)
            {
                ExcelUtil.logError(callerAddress, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), e.Message);
                return "";
            }
        }

        [ExcelFunction(Description = "convert from futures symbols to date ", Category = "EliteQuantExcel - Time")]
        public static object eqTimeDateFromFuturesSymbol(
            [ExcelArgument(Description = "futures symbol (e.g. F16 ")]string fsymbol)
        {
            if (ExcelUtil.CallFromWizard())
                return "";

            string callerAddress = "";
            callerAddress = ExcelUtil.getActiveCellAddress();
            OHRepository.Instance.removeErrorMessage(callerAddress);

            try
            {
                DateTime ret = EliteQuant.Utils.DateFromFuturesSymbol(fsymbol);
                return ret;
            }
            catch (Exception e)
            {
                ExcelUtil.logError(callerAddress, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), e.Message);
                return "";
            }
        }
    }
}
