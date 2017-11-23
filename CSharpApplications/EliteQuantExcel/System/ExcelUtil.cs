using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using System.Text;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Diagnostics;
using Xl = Microsoft.Office.Interop.Excel;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using ExcelDna.Logging;


namespace EliteQuantExcel
{
    sealed public class ExcelUtil
    {
        //////////////////////////////////////////////////////////////////////////////////////////////
        ///////////////////////////                Converter           ///////////////////////////////
        //////////////////////////////////////////////////////////////////////////////////////////////
        public static T ConvertObject<T>(Object obj)
        {
            return ConvertObject<T>(obj, null);
        }

        public static T ConvertObject<T>(Object obj, Object defaultValue)
        {
            try
            {
                if (ExcelUtil.isNull(obj))
                {
                    if (defaultValue != null)
                        return (T)defaultValue;
                    else
                        throw new Exception(" Conversion failed. " + typeof(T).FullName);
                }

                return EliteQuant.EQConverter.ConvertObject<T>(obj, defaultValue);
                // return default(T);
            }
            catch (System.Exception e)
            {
                throw new Exception(" Conversion failed: " + typeof(T).FullName + ", " + e.Message);
            }
        }

        //////////////////////////////////////////////////////////////////////////////////////////////
        ///////////////////////////       Auxiliary functions          ///////////////////////////////
        //////////////////////////////////////////////////////////////////////////////////////////////
            #region ExcelDNA related functions
            // int default to 0, string default to "", DateTiem default to 1899/1/1, bool default to false
            // for exmaple, empty gives DateTime 12/30/1899, string "", object ExcelMissing
            // with value, DateTime gives mm/dd/yyyy, string give "4xxxx", object gives 4xxxx.x
        public static bool isNull(Object o)
        {
            if (o == null || o is ExcelMissing || o is ExcelEmpty || o is ExcelError)
                return true;
            if (o is string && (o as string == string.Empty || o as string == ""))
                return true;
            return false;
        }

        public static bool isError(Object o)
        {
            return (o is ExcelError || (o is String && ((String)o == "#EQ_ERR!" || (String)o == "#EQ_TRG!")));
        }

        public static Xl.Range getActiveCellRange()
        {
            Xl.Application xlApp = (Xl.Application)ExcelDna.Integration.ExcelDnaUtil.Application;
            Xl.Range range;
            try
            {
                range = (Xl.Range)xlApp.get_Caller(System.Type.Missing);
            }
            catch (Exception)
            {
                range = xlApp.ActiveCell;
            }
            return range;
        }

        public static ExcelReference getActiveReference()
        {
            ExcelReference reference = null;

            // first check if calling from Excel
            Xl.Application xlApp = (Xl.Application)ExcelDna.Integration.ExcelDnaUtil.Application;
            Xl.Range range;
            try
            {
                range = (Xl.Range)xlApp.get_Caller(System.Type.Missing);
            }
            catch (System.Exception)
            {
                return null;
            }
            try
            {
                reference = (ExcelReference)XlCall.Excel(XlCall.xlfCaller);
            }
            catch (System.Exception)
            {
                reference = (ExcelReference)XlCall.Excel(XlCall.xlfSelection);
            }
            return reference;
        }

        public static string getActiveCellAddressLocal(ExcelReference reference)
        {
            string cellReference = (string)XlCall.Excel(XlCall.xlfAddress, 1 + reference.RowFirst, 1 + reference.ColumnFirst);
            if (reference.RowFirst != reference.RowLast || reference.ColumnFirst != reference.ColumnLast)
            {
                string bottomRight = (string)XlCall.Excel(XlCall.xlfAddress, 1 + reference.RowLast, 1 + reference.ColumnLast);
                cellReference += ":" + bottomRight;
            }
            return cellReference;
        }

        public static string getActiveCellAddressLocal()
        {
            return getActiveCellAddressLocal(getActiveReference());
        }

        public static string getActiveCellAddress()
        {
            if (Ops.Operation.CallerAddressFlag == false)
                return "getActiveCellAddress";

            ExcelReference reference = getActiveReference();
            if (isNull(reference))
                return "getActiveCellAddress";

            string sheetName = (string)XlCall.Excel(XlCall.xlSheetNm, reference);
            return sheetName + "!" + getActiveCellAddressLocal(reference);
        }

        public static string getSelectCellAddress()
        {
            ExcelReference reference = (ExcelReference)XlCall.Excel(XlCall.xlfSelection);
            string sheetName = (string)XlCall.Excel(XlCall.xlSheetNm, reference);

            return sheetName + "!" + getActiveCellAddressLocal(reference);
        }

        public static string getAddress(Xl.Range range)
        {
            int c = range.Column;
            int r = range.Row;
            Microsoft.Office.Interop.Excel.Worksheet sheet = (Microsoft.Office.Interop.Excel.Worksheet)range.Parent;
            string s = sheet.Name;
            string ps = ((Microsoft.Office.Interop.Excel.Workbook)sheet.Parent).Name;

            return string.Format("[{0}]{1}${2}${3}", ps, s, ExcelColumnIndexToName(c), r);
        }

        public static string ExcelColumnIndexToName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = string.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }

        public static Xl.Range ConvertExcelReferenceToRange(ExcelReference excelRef)
        {
            Xl.Application app = (Xl.Application)ExcelDnaUtil.Application;
            string sheetFullName = (string)XlCall.Excel(XlCall.xlSheetNm, excelRef);   // something like [book1]sheet1, we need to remove [book1]
            string sheetName = sheetFullName.Substring(sheetFullName.IndexOf(']') + 1);
            Xl.Worksheet sheet = app.Worksheets[sheetName];
            string topAddress = ExcelColumnIndexToName(excelRef.ColumnFirst + 1) + (excelRef.RowFirst + 1);
            string bottomAddress = ExcelColumnIndexToName(excelRef.ColumnLast + 1) + (excelRef.RowLast + 1);
            return sheet.get_Range(topAddress, bottomAddress);
        }

        public static Type GetTypeForComObject(Object comObject)
        {
            IntPtr iunknown = System.Runtime.InteropServices.Marshal.GetIUnknownForObject(comObject);
            System.Reflection.Assembly assembly = System.Reflection.Assembly.GetAssembly(typeof(Xl.Range));

            Type[] excelTypes = assembly.GetTypes();

            foreach (Type currType in excelTypes)
            {
                Guid iid = currType.GUID;
                if (!currType.IsInterface || iid == Guid.Empty)
                    continue;

                IntPtr ipointer = IntPtr.Zero;
                System.Runtime.InteropServices.Marshal.QueryInterface(iunknown, ref iid, out ipointer);
                if (ipointer != IntPtr.Zero)
                    return currType;
            }

            return null;
        }

        public static bool TestExcelType(Object obj)
        {
            if (!(obj is object[,]))
            {
                if (obj is ExcelError || ExcelUtil.isEQError(obj))
                    return false;
            }
            else
            {
                foreach (Object o in (object[,])obj)
                {
                    if (o is ExcelError || ExcelUtil.isEQError(o))
                        return false;
                }
            }

            return true;
        }

        public static void logError(string callerAddress, string callerName, string Message)
        {
            OHRepository.Instance.storeErrorMsg(callerAddress, Message);
            LogDisplay.RecordLine("{0}\t{1}\t{2}", callerAddress, callerName, Message);
        }
        #endregion

        public static bool isEQError(Object obj)
        {
            try
            {
                string v = (string)obj;

                if (v == "#EQ_ERR!" || v == "#EQ_TRG!")
                    return true;
                else
                    return false;
            }
            catch (Exception)
            {
                return false;
            }
        }

        #region Call from Wizard
        public delegate bool WindowEnumDelegate(IntPtr hwnd, IntPtr lParam);

        // declare the API function the enumerate child windows
        [DllImport("user32.dll")]
        public static extern int EnumChildWindows(IntPtr hwnd, WindowEnumDelegate del, IntPtr lParma);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern int GetClassName(IntPtr hwnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern int GetWindowThreadProcessId(IntPtr hwnd, out ulong lpdwProcessId);

        public static bool WindowEnumProc(IntPtr hwnd, IntPtr lParam)
        {
            int nRet;
            StringBuilder ClassName = new StringBuilder(100);
            // Get the window class name
            nRet = GetClassName(hwnd, ClassName, ClassName.Capacity);
            if (nRet == 0)
                return true;

            if (string.Compare(ClassName.ToString(), 0, "bosa_sdm_xl", 0, 11, true, System.Globalization.CultureInfo.InvariantCulture) == 0)
            {
                GCHandle gch = GCHandle.FromIntPtr(lParam);
                List<IntPtr> list = gch.Target as List<IntPtr>;
                if (list == null)
                {
                    throw new InvalidCastException("GCHandle Target could not be cast as List<IntPtr>");
                }

                ulong lpdwProcessId;
                GetWindowThreadProcessId(hwnd, out lpdwProcessId);
                list.Add(new IntPtr((long)lpdwProcessId));
                return true;
            }

            return true;
        }

        public static bool CallFromWizard()
        {
            if (Ops.Operation.CallFromWizardFlag == false)
                return false;

            Process currentProcess = Process.GetCurrentProcess();
            IntPtr currPtr = new IntPtr(currentProcess.Id);

            List<IntPtr> result = new List<IntPtr>();
            GCHandle listHandle = GCHandle.Alloc(result);

            try
            {
                ExcelUtil.WindowEnumDelegate childProc = new ExcelUtil.WindowEnumDelegate(ExcelUtil.WindowEnumProc);
                ExcelUtil.EnumChildWindows(IntPtr.Zero, childProc, GCHandle.ToIntPtr(listHandle));

                foreach (IntPtr ptr in result)
                {
                    if (ptr == currPtr)
                        return true;
                }
            }
            finally
            {
                if (listHandle.IsAllocated)
                    listHandle.Free();
            }
            return false;
        }
        #endregion

        /// <summary>
        /// Dump data to excel
        /// </summary>
        /// <param name="selection"></param>
        /// <param name="result"></param>
        public static void DumpDatabaseResult(ExcelReference selection, object[,] result)
        {
            int resultRows = result.GetLength(0);
            int resultCols = result.GetLength(1);
            ExcelReference target = new ExcelReference(selection.RowLast + 1, selection.RowLast + resultRows,
                selection.ColumnFirst, selection.ColumnFirst + resultCols - 1, selection.SheetId);
            target.SetValue(result);
        }
    }

    //////////////////////////////////////////////////////////////////////////////////////////////
    ///////////////////////////    Supporting Classes             ///////////////////////////////
    //////////////////////////////////////////////////////////////////////////////////////////////
    #region Supporting classes
    class FunctionWizardThread
    {
        private Xl.Range range_ = null;
        private System.Threading.SynchronizationContext xlcontext_ = null;

        public FunctionWizardThread(Xl.Range range, System.Threading.SynchronizationContext xlcontext)
        {
            range_ = range;
            xlcontext_ = xlcontext;
        }

        public void functionargumentsPopup()
        {
            xlcontext_.Post(
                delegate(object state)
                {
                    range_.FunctionWizard();
                }, null);
        }
    }
    #endregion
}
