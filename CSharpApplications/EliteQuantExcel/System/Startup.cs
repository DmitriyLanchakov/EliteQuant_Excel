using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using System.Diagnostics;
using Xl = Microsoft.Office.Interop.Excel;
using ExcelDna.Integration;
using ExcelDna.ComInterop;
using ExcelDna.Integration.CustomUI;
using ExcelDna.Logging;


namespace EliteQuantExcel
{
    [ComVisible(true)]
    public class Startup : IExcelAddIn
    {
        public void AutoOpen()
        {
            string xllName = (string)XlCall.Excel(XlCall.xlGetName);
            string rootPath = System.IO.Path.GetDirectoryName(xllName);
            rootPath = System.IO.Path.Combine(rootPath, @"..\..\");
            rootPath = System.IO.Path.GetFullPath(rootPath);
            EliteQuant.ConfigManager.Instance.RootDir = rootPath;

            System.Windows.Forms.MessageBox.Show("Welcome to EliteQuantExcel.");
            // System.Windows.Forms.MessageBox.Show("EliteQuantExcel Loaded from " + xllName);
            // System.Windows.Forms.MessageBox.Show("Root Path is " + rootPath);

            ComServer.DllRegisterServer();
        }

        public void AutoClose()
        {
            //EliteQuant.ConfigManager.Instance.REngine.Dispose();
            ComServer.DllUnregisterServer();
            System.Windows.Forms.MessageBox.Show("Thanks for using EliteQuantExcel");
        }
    }
}
