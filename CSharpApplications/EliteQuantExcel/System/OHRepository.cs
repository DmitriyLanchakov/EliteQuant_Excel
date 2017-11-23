using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Text.RegularExpressions;
using System.Linq;
using System.Text;

namespace EliteQuantExcel
{
    public sealed class OHRepository
    {
        public static OHRepository Instance
        {
            get { return _instance; }
        }

        private static readonly OHRepository _instance = new OHRepository();
        private OHRepository() 
        {
            objectInfo_ = new Hashtable();
            errorMap_ = new Dictionary<string, string>();
            lockobj_ = new object();
        }

        private Hashtable objectInfo_;
        private Dictionary<string, string> errorMap_;
        private object lockobj_;

        public void storeObject(string objName, object obj, string callerAddress)
        {
            string objID = stripObjID(objName);
            DateTime now = DateTime.Now;
            DateTime creationTime;

            System.Threading.Monitor.Enter(lockobj_);
            try
            {
                if (objectInfo_.ContainsKey(objID))
                {
                    creationTime = (DateTime)((List<object>) (objectInfo_[objID]))[2];
                    objectInfo_.Remove(objID);
                }
                else
                {
                    creationTime = now;
                }
                List<object> newObj = new List<object>(4);
                newObj.Add(obj);
                newObj.Add(callerAddress);
                newObj.Add(creationTime);
                newObj.Add(now);

                objectInfo_.Add(objID, newObj);
                if(errorMap_.ContainsKey(callerAddress))
                    errorMap_.Remove(callerAddress);
            }
            finally
            {
                System.Threading.Monitor.Exit(lockobj_);
            }
        }

        public void removeObject(string objID)
        {
            System.Threading.Monitor.Enter(lockobj_);
            try
            {
                if (objectInfo_.ContainsKey(objID))
                    objectInfo_.Remove(objID);
            }
            finally
            {
                System.Threading.Monitor.Exit(lockobj_);
            }
        }

        public bool containsObject(string objID)
        {
            if (objID == "" || objID == "#VALUE!" || objID == "#NA!" || objID == "#EQ_ERR!")
                return false;

            string realID = stripObjID(objID);
            if (objectInfo_.ContainsKey(realID))
            {
                return true;
            }
            return false;
        }

        public T getObject<T>(string objID)
        {
            if (objID == "")
                throw new System.Exception("Empty ID");
            if (objID == "#VALUE!" || objID == "#NA!" || objID == "#EQ_ERR!")
                throw new System.Exception("Cannot identify object ID " + objID + " and get " + typeof(T).FullName);

            string realID = stripObjID(objID);
            System.Threading.Monitor.Enter(lockobj_);
            try
            {
                if(objectInfo_.ContainsKey(realID))
                {
                    return (T)((List<object>) (objectInfo_[realID]))[0];
                    // List<object> tmp = (List<object>)objectInfo_[realID];
                    // T ob = (T)tmp[0];
                    // return ob;
                }
            }
            catch (Exception e)
            {
                ExcelUtil.logError(objID, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), e.Message);
            }
            finally
            {
                System.Threading.Monitor.Exit(lockobj_);
            }

            throw new System.Exception("Cannot find object with ID: " + objID + " and get " + typeof(T).FullName);
        }

        public List<T> getObject<T>(List<string> objIDs)
        {
            // remove empty entries
            int ii = objIDs.IndexOf("");
            while (ii > -1)
            {
                objIDs.RemoveAt(ii);
                ii = objIDs.IndexOf("");
            }

            if (objIDs.Count == 0)
                throw new Exception("No object ID exists to be retrieved ");

            List<T> list = new List<T>(objIDs.Count);
            for (int i = 0; i < objIDs.Count; i++)
            {
                list[i] = getObject<T>(objIDs[i]);
            }
            return list;
        }

        public List<T> getObject<T>(StringCollection objIDs)
        {
            var list = objIDs.Cast<string>().ToList();

            return getObject<T>(list);
        }

        public void storeErrorMsg(string id, string errMsg)
        {
            errorMap_.Add(id, errMsg);
        }

        public string retrieveError(string id)
        {
            if (errorMap_.ContainsKey(id))
                return errorMap_[id];
            else
                return "";
        }

        public void removeErrorMessage(string callerAddress)
        {
            if (callerAddress == "") return;

            System.Threading.Monitor.Enter(lockobj_);
            try
            {
                if (errorMap_.ContainsKey(callerAddress))
                    errorMap_.Remove(callerAddress);
            }
            finally
            {
                System.Threading.Monitor.Exit(lockobj_);
            }
        }

        public List<string> listObjects(string pattern)
        {
            ArrayList matched = new ArrayList();

            Regex regex = new Regex(pattern);

            foreach (DictionaryEntry entry in objectInfo_)
            {
                string className = ((List<object>)entry.Value)[0].GetType().FullName;
                if (pattern.Length == 0 || regex.Match(className).Success)
                {
                    matched.Add(entry.Key);
                }
            }

            List<string> ret = new List<string>(matched.Count);
            for (int i = 0; i < matched.Count; i++)
            {
                ret.Add((string)matched[i]);
            }

            return ret;
        }

        public Type getObjectType(string objID)
        {
            string realID = stripObjID(objID);
            if (objectInfo_.ContainsKey(realID))
                return ((List<object>)objectInfo_[realID])[0].GetType();

            throw new System.Exception("Cannot find object id " + objID);
        }

        public object getObject(string objID)
        {
            System.Threading.Monitor.Enter(lockobj_);
            try
            {
                if (objectInfo_.ContainsKey(objID))
                {
                    return ((List<object>)objectInfo_[objID])[0];
                }
            }
            finally
            {
                System.Threading.Monitor.Exit(lockobj_);
            }

            throw new System.Exception("Cannot find object with ID: " + objID);
        }

        public string getCallerAddress(string objID)
        {
            string realID = stripObjID(objID);
            if (objectInfo_.ContainsKey(realID))
                return (string)((List<object>)objectInfo_[realID])[1];

            throw new System.Exception("Cannot find object id " + objID);
        }

        public DateTime getObjectCreationTime(string objID)
        {
            string realID = stripObjID(objID);
            if (objectInfo_.ContainsKey(realID))
                return (DateTime)((List<object>)objectInfo_[realID])[2];

            throw new System.Exception("Cannot find object id " + objID);
        }

        public DateTime getObjectUpdateTime(string objID)
        {
            string realID = stripObjID(objID);
            if (objectInfo_.ContainsKey(realID))
                return (DateTime)((List<object>)objectInfo_[realID])[3];

            throw new System.Exception("Cannot find object id " + objID);
        }

        /// <summary>
        /// retrive objectID before #
        /// That is Type@Name
        /// </summary>
        private string stripObjID(string objID)
        {
            string realID;
            if (objID.IndexOf('#') != -1)
            {
                realID = objID.Substring(0, objID.IndexOf('#'));
            }
            else
            {
                realID = objID;
            }

            return realID;
        }

        public void clear()
        {
            objectInfo_.Clear();
            errorMap_.Clear();
        }
    }
}
