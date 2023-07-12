namespace COMCalculatorNet6Version.Interfaces
{
    using System;
    using System.Runtime.InteropServices;
    using COMCalculatorNet6Version.COM;

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