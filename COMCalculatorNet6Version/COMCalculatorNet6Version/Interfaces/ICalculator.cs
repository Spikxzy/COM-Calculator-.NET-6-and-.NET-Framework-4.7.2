using COMCalculatorNet6Version.COM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace COMCalculatorNet6Version.Interfaces
{
    [ComImport]
    [Guid(AssemblyInfo.CalculatorInterfaceGuid)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICalculator
    {
        [DispId(1)]
        double Addition(double firstValue, double secondValue);

        [DispId(2)]
        void TriggerAdditionEvent();
    }
}
