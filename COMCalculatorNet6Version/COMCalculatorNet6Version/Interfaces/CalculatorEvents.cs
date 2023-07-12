namespace COMCalculatorNet6Version.Interfaces
{
    using System;
    using System.Runtime.InteropServices;
    using COMCalculatorNet6Version.Classes;
    using COMCalculatorNet6Version.COM;

    [ComVisible(true)]
    [Guid(AssemblyInfo.CalculatorEventsGuid)]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface CalculatorEvents
    {
        [DispId(1)]
        void OnAdditionDone();
    }
}
