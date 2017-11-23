using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using System.Text;
using System.Xml;
using System.IO;
using System.Threading;
using System.Reflection;
using Xl = Microsoft.Office.Interop.Excel;
using System.Data;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using ExcelDna.Logging;
using System.Data.SqlClient;


namespace EliteQuantExcel
{
    [ComVisible(true)]
    //[ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class RibbonMenu : ExcelRibbon
    {
        public static SynchronizationContext syncContext_;
        #region Event Handler
        public static void Login_Click()
        {
            try
            {
                System.Windows.Forms.MessageBox.Show("You are now logged into EliteQuant library");
            }
            catch
            {
                System.Windows.Forms.MessageBox.Show("Unable to log into EliteQuant library");
            }
        }

        public void Error_Click(IRibbonControl control_)
        {
            LogDisplay.Show();
        }

        // [ExcelCommand(MenuName = "Range Tools", MenuText = "Square Selection")]
        public static void ReadData_Click()
        {
            object[,] result = null;
            // Get a reference to the current selection
            ExcelReference selection = (ExcelReference)XlCall.Excel(XlCall.xlfSelection);

            try
            {
            }
            catch
            {
                result = new object[,] { { "Unable to retrieve data." } };
            }

            // Now create the target reference that will refer to Sheet 2, getting a reference that contains the SheetId first
            // ExcelReference sheet2 = (ExcelReference)XlCall.Excel(XlCall.xlSheetId, "Sheet2"); // Throws exception if no Sheet2 exists
            // ... then creating the reference with the right size as new ExcelReference(RowFirst, RowLast, ColFirst, ColLast, SheetId)
            int resultRows = result.GetLength(0);
            int resultCols = result.GetLength(1);
            //ExcelReference target = new ExcelReference(selection.RowFirst, selection.RowFirst + resultRows - 1,
            ExcelReference target = new ExcelReference(0, 0 + resultRows - 1,           // start from top
                selection.ColumnLast + 1, selection.ColumnLast + resultCols, selection.SheetId);
            // Finally setting the result into the target range.
            target.SetValue(result);
        }

        public void Help_Click(IRibbonControl control_)
        {
            //System.Diagnostics.Process.Start(xllDir + @"documents\EliteQuantExcel.chm");
            System.Windows.Forms.MessageBox.Show("Please contact letian.zj for help.");
        }

        public void About_Click(IRibbonControl control_)
        {
            // About abt = new About();
            // abt.ShowDialog();
            System.Windows.Forms.MessageBox.Show("EliteQuant v1.0");
        }

        public void Function_Click(IRibbonControl control_)
        {
            Xl.Application xlApp = (Xl.Application)ExcelDna.Integration.ExcelDnaUtil.Application;
            String fname = control_.Id;
            fname = "=" + fname + "()";

            Xl.Range rg = xlApp.ActiveCell;

            String cellName = ExcelUtil.ExcelColumnIndexToName(rg.Column) + rg.Row;

            Xl._Worksheet sheet = (Xl.Worksheet)xlApp.ActiveSheet;
            Xl.Range range = sheet.get_Range(cellName, System.Type.Missing);
            string previousFormula = range.FormulaR1C1.ToString();
            range.Value2 = fname;
            range.Select();

            syncContext_ = SynchronizationContext.Current;
            if (syncContext_ == null)
            {
                syncContext_ = new System.Windows.Forms.WindowsFormsSynchronizationContext();
            }

            FunctionWizardThread othread = new FunctionWizardThread(range, syncContext_);
            Thread thread = new Thread(new ThreadStart(othread.functionargumentsPopup));
            thread.Start();
        }

        public void excelFile_Click(IRibbonControl control_)
        {
            Xl.Application xlApp = (Xl.Application)ExcelDna.Integration.ExcelDnaUtil.Application;
            String fname = control_.Id;

            string file = "";
            switch (fname)
            {
                case "StockTrading":
                    file = @".\Workbooks\StockTrading.xlsm";
                    break;
                case "HistoricalData":
                    file = @".\Workbooks\HistoricalData.xlsm";
                    break;
                case "VanillaOptionPricer":
                    file = @".\Workbooks\VanillaOptionPricer.xlsx";
                    break;
                case "RatesDemo":
                    file = @".\Workbooks\RatesDemo.xlsx";
                    break;
                /*
                case "SABRModel":
                    file = @".\Workbooks\SABRModel.xlsm";
                    break;         
                case "PRNG":
                    file = @".\Workbooks\PRNG.xlsx";
                    break;
                case "BookTrades":
                    file = @".\Workbooks\BookTrades.xlsm";
                    break;
                case "LoadHistCurve":
                    file = @".\Workbooks\HistUSDCurves.xlsm";
                    break;
                case "PublishCurve":
                    file = @".\Workbooks\USDIRCurve.xlsm";
                    break;*/
                default:
                    break;
            }

            //file = @"C:\Workspace\Output\Debug\OptionPricer.xlsx";
            // rootPath = System.IO.Path.Combine(rootPath, @"..\..\EliteQuant_Excel\");
            // rootPath = System.IO.Path.GetFullPath(rootPath); 
            string filepath = System.IO.Path.Combine(EliteQuant.ConfigManager.Instance.RootDir, file);
            filepath = System.IO.Path.GetFullPath(filepath);
            xlApp.Workbooks.Open(filepath);
        }
        #endregion
    }
}
