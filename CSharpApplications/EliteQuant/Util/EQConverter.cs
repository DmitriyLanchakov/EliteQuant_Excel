using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using System.Text;

namespace EliteQuant
{
    /// <summary>
    /// Convert string to EliteQuant class
    /// </summary>
    public class EQConverter
    {
        #region EQ Class <--> C# class converter
        public static T ConvertObject<T>(Object obj)
        {
            return ConvertObject<T>(obj, null);
        }

        public static T ConvertObject<T>(Object obj, Object defaultValue)
        {
            try
            {
                // pass through
                if (obj.GetType() == typeof(T))
                {
                    return (T)(Object)obj;
                }
                if (typeof(T) == typeof(DateTime))          // To DateTime
                {
                    if (obj is EliteQuant.Date)
                        return (T)(Object)DateTime.FromOADate(((EliteQuant.Date)obj).serialNumber());
                    else if (obj is String)
                        return (T)(Object)DateTime.ParseExact((String)obj, "M/d/yyyy", null);
                    else if (obj is double)
                        return (T)(Object)DateTime.FromOADate((double)obj);
                    else if (obj is object[,])
                        return (T)(Object)EliteQuant.EQConverter.ConvertObject2Array(obj, DateTime.MinValue)[0];
                    else
                        throw new ArgumentOutOfRangeException("Unknown type to convert to System.DateTime");
                }
                else if (typeof(T) == typeof(EliteQuant.Date))       // to QuanLib.Date
                {
                    if (obj is DateTime)
                    {
                        EliteQuant.Date d = new EliteQuant.Date(Convert.ToInt32(((DateTime)obj).ToOADate()));
                        return (T)(Object)d;
                    }
                    else if (obj is double)
                    {
                        EliteQuant.Date d = new EliteQuant.Date(Convert.ToInt32((double)obj));
                        return (T)(Object)d;
                    }
                    else
                        throw new ArgumentOutOfRangeException("Unknown type to convert to EliteQuant.Date");

                }
                else if (typeof(T) == typeof(EliteQuant.Calendar))        // to EliteQuant.Calendar
                {
                    if (obj is String)
                    {
                        string c = (string)obj;
                        return (T)(Object)(EliteQuant.EQConverter.ConvertStringToCalendar(c));
                    }
                    else
                        throw new ArgumentOutOfRangeException("Unknown type to convert to Calendar");
                }
                else if (typeof(T) == typeof(EliteQuant.DayCounter))              // to DayCounter
                {
                    if (obj is string)
                    {
                        string c = (string)obj;
                        return (T)(object)(EliteQuant.EQConverter.ConvertStringToDayCounter(c));
                    }
                    else
                        throw new ArgumentOutOfRangeException("Unknown type to convert to DayCounter");
                }
                else if (typeof(T) == typeof(EliteQuant.BusinessDayConvention))              // to bdc
                {
                    if (obj is string)
                    {
                        string c = (string)obj;
                        return (T)(object)(EliteQuant.EQConverter.ConvertStringToBDC(c));
                    }
                    else
                        throw new ArgumentOutOfRangeException("Unknown type to convert to BusinessDayConvention");
                }
                else if (typeof(T) == typeof(EliteQuant.DateGeneration.Rule))              // to rule
                {
                    if (obj is string)
                    {
                        string c = (string)obj;
                        return (T)(object)(EliteQuant.EQConverter.ConvertStringToDGRule(c));
                    }
                    else
                        throw new ArgumentOutOfRangeException("Unknown type to convert to DateGenerationRule");
                }
                else if (typeof(T) == typeof(EliteQuant.Frequency))              // to rule
                {
                    if (obj is string)
                    {
                        string c = (string)obj;
                        return (T)(object)(EliteQuant.EQConverter.ConvertStringToFrequency(c));
                    }
                    else
                        throw new ArgumentOutOfRangeException("Unknown type to convert to Frequency");
                }
                else if (typeof(T) == typeof(EliteQuant.Compounding))              // to rule
                {
                    if (obj is string)
                    {
                        string c = (string)obj;
                        return (T)(object)(EliteQuant.EQConverter.ConvertStringToCompounding(c));
                    }
                    else
                        throw new ArgumentOutOfRangeException("Unknown type to convert to Compounding");
                }
                else if (typeof(T) == typeof(EliteQuant.Period))
                {
                    if (obj is String)      // String (2D, 3M, 10Y)
                    {
                        return (T)(Object)(EliteQuant.EQConverter.ConvertStringToPeriod((string)obj));
                    }
                    else if (obj is EliteQuant.Frequency)             // Monthly, Annual
                    {
                        string[] freqNames = Enum.GetNames(typeof(EliteQuant.Frequency));
                        StringCollection sc = new StringCollection();
                        sc.AddRange(freqNames);
                        if (sc.Contains(obj as string))
                        {
                            EliteQuant.Frequency frq = (EliteQuant.Frequency)Enum.Parse(typeof(EliteQuant.Frequency), obj as string);
                            return (T)(Object)new EliteQuant.Period(frq);
                        }
                        else
                            return (T)(Object)new EliteQuant.Period((EliteQuant.Frequency)obj);
                    }
                    else if (obj is int)
                        return (T)(Object)new EliteQuant.Period((int)obj, EliteQuant.TimeUnit.Years);
                    else
                        throw new ArgumentOutOfRangeException("Unknown type to convert to EliteQuant.Period");
                }
                else if (typeof(T) == typeof(DateTime[]))
                {
                    if (obj is object[,] || obj is DateTime || obj is double)
                        return (T)(Object)EQConverter.ConvertObject2Array(obj, DateTime.MinValue);
                    else
                        throw new ArgumentOutOfRangeException("Unknown type to convert to DateTime[]");
                }
                else if (typeof(T) == typeof(EliteQuant.Period[]))
                {
                    if (obj is object[,] || obj is String)
                        return (T)(Object)EQConverter.ConvertObject2Array(obj, new EliteQuant.Period());
                    else
                        throw new ArgumentOutOfRangeException("Unknown type to convert to EliteQuant.Period[]");
                }
                else if (typeof(T) == typeof(double))
                {
                    if (obj is String)
                    {
                        double amount = 0.0;
                        double.TryParse((String)obj, out amount);
                        return (T)(Object)amount;
                    }
                }
                else if (typeof(T) == typeof(double[]))
                {
                    if (obj is object[,] || obj is double)
                    {
                        double def = (defaultValue is double) ? (double)defaultValue : 0.0;
                        return (T)(Object)EQConverter.ConvertObject2Array<double>(obj, def);
                    }
                }
                else if (typeof(T).IsEnum)
                {
                    if (obj is String)
                    {
                        if (EQEnum.EnumDictionary.ContainsKey(obj as string))
                            obj = EQEnum.EnumDictionary[obj as string];
                        return (T)Enum.Parse(typeof(T), (string)obj, true);
                    }
                }
                else if (typeof(T) == typeof(String[]))
                {
                    if (obj is object[,] || obj is string)
                    {
                        StringCollection tmp = EQConverter.ConvertObject2StringCollection(obj);
                        String[] ret = new String[tmp.Count];
                        tmp.CopyTo(ret, 0);
                        return (T)(Object)ret;
                    }
                }
                else if (typeof(T) == typeof(StringCollection))
                {
                    if (obj is object[,] || obj is string || obj is double)
                        return (T)(Object)EQConverter.ConvertObject2StringCollection(obj);
                }
                else if (typeof(T) == typeof(bool[]))
                {
                    if (obj is object[,] || obj is bool)
                    {
                        bool def = (defaultValue is bool) ? (bool)defaultValue : false;
                        return (T)(Object)EQConverter.ConvertObject2Array<bool>(obj, def);
                    }
                }
                else if (typeof(T) == typeof(int[]))
                {
                    if (obj is object[,] || obj is double)
                    {
                        double def = (defaultValue is int) ? Convert.ToDouble(defaultValue) : 0.0;
                        double[] tmp = EQConverter.ConvertObject2Array<double>(obj, def);
                        int[] ret = new int[tmp.Length];
                        for (int i = 0; i < ret.Length; i++) ret[i] = (int)tmp[i];
                        return (T)(Object)ret;
                    }
                }
            }
            catch (System.Exception e)
            {
                if (defaultValue != null)
                    return DefaultValue<T>(defaultValue);
                else
                    throw new Exception(" EQConverter failed: " + e.Message);
            }

            return DefaultValue<T>(defaultValue);
        }

        public static T DefaultValue<T>(Object defaultValue)
        {
            bool isNull = (defaultValue == null);

            if (!isNull && typeof(T) == defaultValue.GetType())
            {
                return (T)defaultValue;
            }
            else
            {
                if (typeof(T) == typeof(EliteQuant.Period))
                {
                    if (isNull) return (T)(Object)(new EliteQuant.Period(0, EliteQuant.TimeUnit.Days));
                    return (T)(Object)defaultValue;
                }
                else if (typeof(T) == typeof(DateTime))
                {
                    if (isNull) return (T)(Object)(DateTime.MinValue);
                    return (T)(Object)defaultValue;
                }
                else if (typeof(T) == typeof(double[]))
                {
                    if (isNull) return (T)(Object)(new double[0] { });
                    return (T)(Object)(new double[1] { (double)defaultValue });
                }
                else if (typeof(T) == typeof(int[]))
                {
                    if (isNull) return (T)(Object)(new int[0] { });
                    return (T)(Object)(new int[1] { (int)defaultValue });
                }
                else if (typeof(T) == typeof(DateTime[]))
                {
                    if (isNull) return (T)(Object)(new DateTime[0] { });
                    return (T)(Object)(new DateTime[1] { (DateTime)defaultValue });
                }
                else if (typeof(T) == typeof(EliteQuant.Period[]))
                {
                    if (isNull) return (T)(Object)(new double[0] { });
                    return (T)(Object)(new EliteQuant.Period[1] { (EliteQuant.Period)defaultValue });
                }
                else if (typeof(T) == typeof(StringCollection))
                {
                    if (isNull) return (T)(Object)(new StringCollection());
                    StringCollection ret = new StringCollection();
                    ret.Add((String)defaultValue);
                    return (T)(Object)ret;
                }
            }

            throw new Exception("unexpected default value. ");
        }
        #endregion

        #region object covnerter
        public static StringCollection ConvertObject2StringCollection(Object obj0)
        {
            StringCollection result = new StringCollection();
            if (obj0 is string)
            {
                result.Add((string)obj0);
                return result;
            }
            else if (obj0 is double)
            {
                result.Add(System.Convert.ToString(System.Convert.ToInt32((double)obj0)));
                return result;
            }
            Object[,] obj = (Object[,])obj0;
            int rows = obj.GetLength(0);
            int cols = obj.GetLength(1);
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    if (obj[i, j] is string)
                    {
                        result.Add((String)obj[i, j]);
                    }
                    else if (obj[i, j] is double)
                    {
                        result.Add(System.Convert.ToString(System.Convert.ToInt32((double)obj[i, j])));
                    }
                }
            }

            return result;
        }

        public static T[] ConvertObject2Array<T>(object arg, T def_value) where T : new()
        {
            try
            {
                if (arg is int || arg is double || arg is T)
                {
                    T[] ret = new T[1];
                    ret[0] = (T)arg;
                    return ret;
                }
                else if (arg is object[,])
                {
                    int row = ((object[,])arg).GetLength(0);
                    int col = ((object[,])arg).GetLength(1);

                    T[] ret = new T[row * col];
                    for (int i = 0; i < row; i++)
                    {
                        for (int j = 0; j < col; j++)
                        {
                            ret[i * col + j] = EQConverter.ConvertObject<T>(((object[,])arg)[i, j], def_value);
                        }
                    }
                    return ret;
                }
            }
            catch (Exception e)
            {
                throw e;
            }

            T[] retArr = new T[1];
            retArr[0] = def_value;
            return retArr;
        }
        #endregion

        #region Time
        static public EliteQuant.Date DateTimeToDate(DateTime dt)
        {
            int serial = (int)dt.ToOADate();
            return new EliteQuant.Date(serial);
        }

        static public DateTime DateToDateTime(EliteQuant.Date dt)
        {
            int serial = dt.serialNumber();
            return DateTime.FromOADate((double)serial);
        }

        static public string DateTimeToString(DateTime dt)
        {
            return dt.ToString("yyyyMMdd");
        }

        static public DateTime StringToDateTime(string s)
        {
            return DateTime.ParseExact(s, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
        }
        #endregion

        #region string to class converter
        public static Calendar ConvertStringToCalendar(string c)
        {
            switch (c.ToUpper())
            {
                case "NYC":
                    return (new EliteQuant.UnitedStates(EliteQuant.UnitedStates.Market.Settlement));
                case "LON":
                    return (new EliteQuant.UnitedKingdom(EliteQuant.UnitedKingdom.Market.Exchange));
                case "NYC|LON":
                case "LON|NYC":
                    return (new EliteQuant.JointCalendar(
                        new EliteQuant.UnitedStates(EliteQuant.UnitedStates.Market.Settlement),
                        new EliteQuant.UnitedKingdom(EliteQuant.UnitedKingdom.Market.Exchange),
                        EliteQuant.JointCalendarRule.JoinHolidays));
                default:
                    throw new ArgumentOutOfRangeException("Unknown Financial Center.");
            }
        }

        public static DayCounter ConvertStringToDayCounter(string c)
        {
            switch (c.ToUpper())
            {
                case "ACTUAL360":
                    return (new EliteQuant.Actual360());
                case "ACTUAL365":
                    return (new EliteQuant.Actual365Fixed());
                case "ACTUALACTUAL":
                    return (new EliteQuant.ActualActual());
                case "THIRTY360":
                    return (new Thirty360(Thirty360.Convention.USA));
                default:
                    throw new ArgumentOutOfRangeException("Unknow day counter.");
            }
        }

        public static BusinessDayConvention ConvertStringToBDC(string c)
        {
            switch (c.ToUpper())
            {
                case "F":
                case "FOLLOWING":
                    return (EliteQuant.BusinessDayConvention.Following);
                case "MF":
                case "MODIFIEDFOLLOWING":
                    return (EliteQuant.BusinessDayConvention.ModifiedFollowing);
                case "P":
                case "PRECEDING":
                    return (EliteQuant.BusinessDayConvention.Preceding);
                case "MP":
                case "MODIFIEDPRECEDING":
                    return (EliteQuant.BusinessDayConvention.ModifiedPreceding);
                case "NONE":
                    return (EliteQuant.BusinessDayConvention.Unadjusted);
                default:
                    throw new ArgumentOutOfRangeException("unknow business day convention.");
            }
        }

        public static DateGeneration.Rule ConvertStringToDGRule(string c)
        {
            switch (c.ToUpper())
            {
                case "BACKWARD":
                    return (EliteQuant.DateGeneration.Rule.Backward);
                case "FORWARD":
                    return (EliteQuant.DateGeneration.Rule.Forward);
                case "ZERO":
                    return (EliteQuant.DateGeneration.Rule.Zero);
                case "THIRDWEDNESDAY":
                    return (EliteQuant.DateGeneration.Rule.ThirdWednesday);
                case "TWENTIETH":
                    return (EliteQuant.DateGeneration.Rule.Twentieth);
                case "TWENTIETHIMM":
                    return (EliteQuant.DateGeneration.Rule.TwentiethIMM);
                case "CDS":
                    return (EliteQuant.DateGeneration.Rule.CDS);
                default:
                    throw new ArgumentOutOfRangeException("unknow date generation rule.");
            }
        }

        public static Frequency ConvertStringToFrequency(string c)
        {
            switch (c.ToUpper())
            {
                case "ANNUAL":
                    return Frequency.Annual;
                case "SEMIANNUAL":
                    return Frequency.Semiannual;
                case "QUARTERLY":
                    return Frequency.Quarterly;
                case "MONTHLY":
                    return Frequency.Monthly;
                case "WEEKLY":
                    return Frequency.Weekly;
                case "DAILY":
                    return Frequency.Daily;
                default:
                    throw new ArgumentOutOfRangeException("unknow frequency.");
            }
        }

        public static Compounding ConvertStringToCompounding(string c)
        {
            switch (c.ToUpper())
            {
                case "SIMPLE":
                    return Compounding.Simple;
                case "COMPOUNDED":
                    return Compounding.Compounded;
                case "CONTINUOUS":
                    return Compounding.Continuous;
                case "SIMPLETHENCOMPOUNDED":
                    return Compounding.SimpleThenCompounded;
                default:
                    throw new ArgumentOutOfRangeException("unknow frequency.");
            }
        }

        public static Period ConvertStringToPeriod(string obj)
        {
            return new EliteQuant.Period(obj);
        }
        #endregion
    }  // eod of class Converter
}
