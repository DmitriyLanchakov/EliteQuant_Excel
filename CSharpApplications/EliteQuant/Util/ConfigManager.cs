using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Configuration;
using System.IO;


namespace EliteQuant
{
    public sealed class ConfigManager
    {
        string _rootdir = @"c:\";

        private static ConfigManager _instance = null;
        private static readonly object padlock = new object();

        private ConfigManager()
        {
        }

        public static ConfigManager Instance
        {
            get
            {
                lock (padlock)
                {
                    if (_instance == null)
                    {
                        _instance = new ConfigManager();
                    }
                    return _instance;
                }
            }
        }


        #region global path
        public string RootDir
        {
            get { return _rootdir; }
            set { _rootdir = value; }
        }

        public string IRRootDir
        {
            get { return _rootdir + @"InterestRates\"; }
        }

        public string SimRootDir
        {
            get { return _rootdir + @"Simulation\"; }
        }
        #endregion

        public void Log(string msg)
        {
        }

        public delegate void VoidStringDelegate(string msg);
        public VoidStringDelegate debughandler_;

        public void Debug(string msg)
        {
            if (debughandler_ != null)
            {
                debughandler_(msg);
            }
        }
    }
}
