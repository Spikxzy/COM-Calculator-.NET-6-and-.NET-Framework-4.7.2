using COMCalculatorNet6Version.COM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace COMCalculatorNet6Version.Interfaces
{

    [ComVisible(true)]
    [Guid(AssemblyInfo.CalculatorEventsGuid)]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface CalculatorEvents
    {
        [DispId(1)]
        void OnAdditionDone();
    }
}
