using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System;

// General Information about an assembly is controlled through the following 
// set of attributes. Change these attribute values to modify the information
// associated with an assembly.
[assembly: AssemblyTitle("EliteQuantExcel")]
[assembly: AssemblyDescription("")]
[assembly: AssemblyConfiguration("")]
[assembly: AssemblyCompany("EliteQuant")]
[assembly: AssemblyProduct("EliteQuantExcel")]
[assembly: AssemblyCopyright("Copyright © EliteQuant 2017")]
[assembly: AssemblyTrademark("")]
[assembly: AssemblyCulture("")]

// Setting ComVisible to false makes the types in this assembly not visible 
// to COM components.  If you need to access a type in this assembly from 
// COM, set the ComVisible attribute to true on that type.
[assembly: ComVisible(false)]

// The following GUID is for the ID of the typelib if this project is exposed to COM
[assembly: Guid("eed9cb40-67dc-407e-8f22-99de5025c066")]

// Version information for an assembly consists of the following four values:
//
//      Major Version
//      Minor Version 
//      Build Number
//      Revision
//
// You can specify all the values or you can default the Build and Revision Numbers 
// by using the '*' as shown below:
// [assembly: AssemblyVersion("1.0.*")]
[assembly: AssemblyVersion("1.0.0.0")]
[assembly: AssemblyFileVersion("1.0.0.0")]


// The attribute class below is used in xll functions to display build date and user
// The attribute initialization is auto-generated and saved in AssemblyBuildInfo.cs
[AttributeUsage(AttributeTargets.Assembly, AllowMultiple = false)]
internal sealed class BuildInfoAttribute : Attribute
{
    public string UserName
    {
        get;
        private set;
    }

    public string BuildDate
    {
        get;
        private set;
    }

    public BuildInfoAttribute(string username_, string buildDate_)
    {
        UserName = username_;
        BuildDate = buildDate_;
    }
}