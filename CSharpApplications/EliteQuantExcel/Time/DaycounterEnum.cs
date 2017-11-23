using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel;
using System.Threading.Tasks;

namespace EliteQuantExcel
{
    public enum DaycounterEnum
    {
        [Description("Actual360")]
        ACTUAL360,
        [Description("Actual365Fixed")]
        ACTUAL365,
        [Description("ActualActual")]
        ACTUALACTUAL
    }
}
