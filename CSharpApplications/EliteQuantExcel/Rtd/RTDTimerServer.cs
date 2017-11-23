using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using ExcelDna.Integration.Rtd;
using System.Threading.Tasks;
using System.Threading;

/// http://putridparrot.com/blog/a-real-time-data-rtd-server-using-excel-dna-in-excel/
/// https://code.google.com/p/finansu/
namespace EliteQuantExcel
{
    [ComVisible(true)]
    public class RTDTimerServer : IRtdServer
    {
        public class RtdTask
        {
            public bool IsUpdated = false;
            public double interval;
            public string functionName;
            public System.Timers.Timer timer;
            List<string> parameters;        // the first one is function name and the rest is its parameters

            public object result;

            public event EventHandler OnComplete;

            public RtdTask(double interval_, string function_, List<string> parameters_)
            {
                interval = interval_;
                functionName = function_;
                parameters = parameters_;
                timer = new System.Timers.Timer();
                timer.Interval = (int)interval * 1000;
                timer.Elapsed += new System.Timers.ElapsedEventHandler(timer_Elapsed);
                timer.Start();
            }

            void timer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
            {
                if (functionName == "eqRefreshVolParams")
                {
                    //eqEquityVolatility.eqEquityParamsFromDB();
                    //object resultInArray = eqEquityVolatility.eqDisplayEquityVolParams();
                    //result = ConvertToExcelFormat((Array)resultInArray);
                    result = "";
                    this.IsUpdated = true;
                    OnComplete(this, null);
                }
                else if (functionName == "GOOG")
                {
                    var query = @"https://finance.google.com/finance/info?q=" + parameters[0];
                    var quote = EliteQuant.Broker.GetRealGOOGTimeData(query);
                    result = quote[0].l;
                    this.IsUpdated = true;
                    OnComplete(this, null);
                }
                else if (functionName == "NOW")
                {
                    result = DateTime.Now;
                    this.IsUpdated = true;
                    OnComplete(this, null);
                }
            }

            private string ConvertToExcelFormat(Array array)
            {
                string result = string.Empty;
                for (int j = 0; j < array.GetLength(1); j++)
                {
                    for (int i = 0; i < array.GetLength(0); i++)
                    {
                        if (i == 0)
                            result = result + array.GetValue(i, j);
                        else
                            result = result + "," + array.GetValue(i, j);
                    }
                    if (j < array.GetLength(1) - 1)
                        result = result + ";";
                }
                result = "{ " + result + " }";
                return result;
            }

            public void stop()
            {
                timer.Stop();
            }
        } // end of class RtdTask

        private IRTDUpdateEvent _callback;
        private Dictionary<int, RtdTask> taskMap;   // we use forms.Timer as it calls back on native window's thread

        #region IRtdServer Members
        /* we assume that the input strings are: repeat interval in seconds, function name, and input parameters */
        public object ConnectData(int topicId, ref Array Strings, ref bool GetNewValues)
        {
            if (taskMap.ContainsKey(topicId))
            {
                // return "rtd task already set";
                return DateTime.Now;
            }

            if (Strings.Length < 2)
            {
                throw new Exception("input parameters should at least contains interval and function names");
            }

            double interval = Double.Parse((string)Strings.GetValue(0));
            string functionName = (string)Strings.GetValue(1);

            List<string> parameters = new List<string>();
            if (Strings.Length > 2)
            {
                for (int i = 2; i < Strings.Length; ++i)
                {
                    parameters.Add((string)Strings.GetValue(i));
                }
            }

            RtdTask task = new RtdTask(interval, functionName, parameters);
            task.OnComplete += new EventHandler(task_OnComplete);
            taskMap.Add(topicId, task);
            return DateTime.Now;
        }

        void task_OnComplete(object sender, EventArgs e)
        {
            _callback.UpdateNotify();
        }

        public void DisconnectData(int topicId)
        {

        }

        public int Heartbeat()
        {
            return 1;
        }

        public Array RefreshData(ref int TopicCount)
        {
            object[,] results = new object[2, taskMap.Keys.Count];
            TopicCount = 0;

            foreach(int topicID in taskMap.Keys)
            {
                if (taskMap[topicID].IsUpdated == true)
                {
                    results[0, TopicCount] = topicID;
                    results[1, TopicCount] = taskMap[topicID].result;
                    taskMap[topicID].IsUpdated = false;     // mark it back to false
                    TopicCount++;
                }
            }

            object[,] temp = new object[2, TopicCount];
            for (int i = 0; i < TopicCount; i++)
            {
                temp[0, i] = results[0, i];
                temp[1, i] = results[1, i];
            }

            return temp;
        }

        public int ServerStart(IRTDUpdateEvent CallbackObject)
        {
            _callback = CallbackObject;
            taskMap = new Dictionary<int, RtdTask>();
            return 1;
        }

        public void ServerTerminate()
        {
            foreach(RtdTask task in taskMap.Values)
            {
                task.stop();
            }
        }
        #endregion

        private string GetTime()
        {
            return DateTime.Now.ToString("HH:mm:ss.fff");
        }
    } // end of class RTDTimerServer

    [ComVisible(true)]
    public class RTDSimpleTimerServer : IRtdServer
    {
        private IRTDUpdateEvent _callback;
        private System.Windows.Forms.Timer _timer;
        private Dictionary<int, RoutineTask> _tasks = new Dictionary<int, RoutineTask>();

        // 1. Start the server
        public int ServerStart(IRTDUpdateEvent CallbackObject)
        {
            // Called when the first RTD topic is requested
            // Sets server to poll for updates every second
            _callback = CallbackObject;
            _timer = new System.Windows.Forms.Timer();
            _timer.Tick += Callback;
            _timer.Interval = 1000;
            return 1;
        }

        public void ServerTerminate()
        {
            // Kills timer and queries so they don't run after workbook closes
            _timer.Dispose();
            _tasks = null;
        }

        // 2. connect to the data
        public object ConnectData(int topicId, ref Array Strings, ref bool GetNewValues)
        {
            if (_tasks.ContainsKey(topicId))
            {
                // return "rtd task already set";
                return DateTime.Now;
            }

            if (Strings.Length < 2)
            {
                throw new Exception("input parameters should at least contains interval and function names");
            }

            double interval = Double.Parse((string)Strings.GetValue(0));
            string functionName = (string)Strings.GetValue(1);

            List<string> parameters = new List<string>();
            if (Strings.Length > 2)
            {
                for (int i = 2; i < Strings.Length; ++i)
                {
                    parameters.Add((string)Strings.GetValue(i));
                }
            }

            RoutineTask task = new RoutineTask(interval, functionName, parameters);
            _tasks.Add(topicId, task);

            //_timer.Interval = (int)Strings.GetValue(0);
            _timer.Start();
            return _tasks[topicId].ToString();
        }

        public void DisconnectData(int topicId)
        {
            // Removes query on disconnect
            _tasks.Remove(topicId);
        }

        public int Heartbeat()
        {
            // Called by Excel if a given interval has elapsed (returns true)
            return 1;
        }

        public Array RefreshData(ref int topicCount)
        {
            // Called when Excel requests refresh
            // Returns topic count and an array of IDs and values
            object[,] results = new object[2, _tasks.Count];
            topicCount = 0;

            // Prevent overwriting by any Query delegates
            lock (_tasks)
            {
                foreach (int topicID in _tasks.Keys)
                {
                    // Only return results if they've been updated
                    if (_tasks[topicID].Updated)
                    {
                        _tasks[topicID].Updated = false;
                        results[0, topicCount] = topicID;
                        results[1, topicCount] = _tasks[topicID].ToString();
                        topicCount++;
                    }
                }
            }

            _timer.Start();
            return results;
        }

        private void Callback(object sender, EventArgs e)
        {
            // Stops timer and tells all queries to run their async delegates
            _timer.Stop();
            lock (_tasks)
            {
                foreach (KeyValuePair<int, RoutineTask> t in _tasks)
                {
                    t.Value.AsyncImport();
                }
            }
            _callback.UpdateNotify();
        }

        public class RoutineTask
        {
            private const string WAIT_TEXT = "#WAIT";
            private delegate void AsyncImporter();
            private AsyncImporter _importer;
            private DateTime _nextUpdate = DateTime.Now.AddDays(-1);
            private bool _updated;
            private double _interval;
            private string _function;
            private List<string> _parameters;
            private object _result;

            public RoutineTask(double interval, string function, List<string> parameters)
            {
                _interval = interval;
                _function = function;
                _parameters = parameters;
                AsyncImport();
            }

            public bool Updated
            {
                get { return _updated; }
                set { _updated = value; }
            }
            public DateTime NextUpdate
            {
                get { return _nextUpdate; }
                set { _nextUpdate = value; }
            }

            public object Results
            {
                get { return _result; }
                set { _result = value; }
            }

            public void AsyncImport()
            {
                // Only runs async based on query frequency
                if (DateTime.Compare(DateTime.Now, _nextUpdate) > 0)
                {
                    _nextUpdate = DateTime.Now.AddSeconds(_interval);
                    _importer = new AsyncImporter(RunImportDelegate);
                    _importer.BeginInvoke(null, null);
                }
            }

            // real update
            private void RunImportDelegate()
            {
                if (_function == "eqRefreshVolParams")
                {
                    _result = "";
                }
                else if (_function == "NOW")
                {
                    _result = DateTime.Now;
                    
                }
                else if (_function == "RealTimeQuote")
                {
                    _result = EliteQuant.Broker.GetRealTimeQuote(_parameters[0], _parameters[1]);
                }
                _updated = true;
            }

            public override string ToString()
            {
                if (_result == null)
                {
                    return WAIT_TEXT;
                }
                else
                {
                    if (_function == "NOW")
                    {
                        //return ((DateTime)_result).ToString("HH:mm:ss.fff");
                        return ((DateTime)_result).ToString("HH:mm:ss");
                    }
                    else if (_function == "RealTimeQuote")
                    {
                        object[] temp = (object[])_result;
                        List<string> strtmp = new List<string>();
                        for (int i = 0; i < temp.Length;i++)
                        {
                            strtmp.Add((string)temp[i]);
                        }

                        string retstr = string.Join(",", strtmp);
                        return retstr;
                    }
                    else
                    {
                        return WAIT_TEXT;
                    }
                }
            }
        }
    } // End class RTDSimpleTimerServer
}
