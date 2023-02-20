using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

[assembly: ComVisible(true)]
[assembly: Guid(COMCalculatorNet6Version.COM.AssemblyInfo.LibraryGuid)]

namespace COMCalculatorNet6Version.COM
{
    internal class AssemblyInfo
    {
        internal const string LibraryGuid = "B6225162-C6EB-4FDB-B932-35F62A4B6F7B";
        internal const string CalculatorInterfaceGuid = "954275BE-B766-4AFD-A8FD-18BFC0D1D9B5";
        internal const string CalculatorEventsGuid = "6C4DEE22-3A22-4C9E-BF7F-720E49196C3E";
        internal const string CalculatorClassGuid = "32792C02-A345-40BA-BF30-50CABD61155D";
        internal const string OnAdditionDoneDelegateGuid = "C8307475-1DCD-488F-B396-DFD0B9C3ED09";

        internal static T Attribute<T>()
            where T : Attribute
        {
            return typeof(AssemblyInfo).Assembly.GetCustomAttribute<T>();
        }
    }
}
