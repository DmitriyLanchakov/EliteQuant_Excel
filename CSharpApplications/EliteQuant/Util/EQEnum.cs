using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using System.Text;

using EliteQuant;

namespace EliteQuant
{
    /// <summary>
    /// Enum class
    /// </summary>
    public class EQEnum
    {
        public static Dictionary<string, string> EnumDictionary = new Dictionary<string, string>()
        {
            {"MF", "ModifiedFollowing" },
            {"Modified Following", "ModifiedFollowing" },
            {"F", "Following" },
            {"P", "Preceding" },
            {"MP", "ModifiedPreceding" },
            {"Modified Preceding", "ModifiedPreceding" }
        };
    } // end of class EQEnum
}
