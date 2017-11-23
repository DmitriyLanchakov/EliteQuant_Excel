using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xl = Microsoft.Office.Interop.Excel;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using ExcelDna.Logging;

namespace EliteQuantExcel.Ops
{
    public class WindowsOp
    {
        [ExcelFunction(Description = "List Path sub folders", Category = "EliteQuantExcel - Operation")]
        public static object eqOpListFolders(
            [ExcelArgument(Description = "path ")] string path)
        {
            if (ExcelUtil.CallFromWizard())
                return "";

            string callerAddress = ExcelUtil.getActiveCellAddress();

            try
            {
                //path = @"S:\All\Risk\Simulation\Pds\" + asofdate.ToString("yyyyMMdd");
                string[] dirs = System.IO.Directory.GetDirectories(path).OrderByDescending(x => x).ToArray();
                object[,] ret = new object[dirs.Count(), 1];

                int i = 0;
                foreach (var s in dirs)
                {
                    ret[i++, 0] = (new System.IO.DirectoryInfo(s)).Name;
                }

                return ret;
            }
            catch (Exception e)
            {
                ExcelUtil.logError(callerAddress, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), e.Message);
                return "#EQ_ERR!";
            }
        }

        [ExcelFunction(Description = "List files in Path", Category = "EliteQuantExcel - Operation")]
        public static object eqOpListFiles(
            [ExcelArgument(Description = "path ")] string path,
            [ExcelArgument(Description = "with file extensions ")] bool withext)
        {
            if (ExcelUtil.CallFromWizard())
                return "";

            string callerAddress = ExcelUtil.getActiveCellAddress();

            try
            {
                string[] files = System.IO.Directory.GetFiles(path);
                object[,] ret = new object[files.Count(), 1];

                int i = 0;
                foreach (var f in files)
                {
                    if (withext)
                    {
                        ret[i++, 0] = System.IO.Path.GetFileName(f);
                    }
                    else
                    {
                        ret[i++, 0] = System.IO.Path.GetFileNameWithoutExtension(f);
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

        [ExcelFunction(Description = "List an environment variable", Category = "EliteQuantExcel - Operation")]
        public static object eqOpListEnvironmentVariable(
            [ExcelArgument(Description = "environment variable (e.g. Path or PythonPath)")] string evar)
        {
            if (ExcelUtil.CallFromWizard())
                return "";

            string callerAddress = ExcelUtil.getActiveCellAddress();

            try
            {
                if (string.IsNullOrEmpty(evar))
                {
                    int env_count = Environment.GetEnvironmentVariables().Count;
                    object[,] ret = new object[env_count, 3];
                    int i = 0;
                    foreach (System.Collections.DictionaryEntry env in Environment.GetEnvironmentVariables())
                    {
                        ret[i, 0] = (i + 1).ToString() + "/" + env_count.ToString();
                        ret[i, 1] = (string)env.Key;
                        ret[i, 2] = (string)env.Value;
                        i++;
                    }
                    return ret;
                }
                else
                {
                    var path = System.Environment.GetEnvironmentVariable(evar);
                    var path_array = path.Split(';');
                    object[,] ret = new object[path_array.Count(), 2];
                    for (int i = 0; i < path_array.Count(); i++)
                    {
                        ret[i, 0] = (i + 1).ToString() + "/" + path_array.Count();
                        ret[i, 1] = path_array[i];
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

        [ExcelFunction(Description = "Retrieve R Path", Category = "EliteQuantExcel - Operation")]
        public static object eqOpGetRPath()
        {
            if (ExcelUtil.CallFromWizard())
                return "";

            string callerAddress = ExcelUtil.getActiveCellAddress();

            try
            {
                Microsoft.Win32.RegistryKey registryKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(@"SOFTWARE\R-core\R");
                string rPath = (string)registryKey.GetValue("InstallPath");
                string rVersion = (string)registryKey.GetValue("Current Version");
                registryKey.Dispose();

                object[,] ret = new object[2, 2];
                ret[0, 0] = "R Version"; ret[0, 1] = rVersion;
                ret[1, 0] = "R Path"; ret[1, 1] = rPath;

                return ret;
            }
            catch (Exception e)
            {
                ExcelUtil.logError(callerAddress, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), e.Message);
                return "#EQ_ERR!";
            }
        }
    }
}
