using COMCalculatorFrameworkVersion.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace COMCalculatorFrameworkVersion.Classes
{
    [ComVisible(true)]
    [Guid("322C9DE3-89BD-4221-9809-4E31BF65EC99")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(CalculatorEvents))]
    [ProgId("COM Calculator .NET Framework version")]
    public class Calculator : ICalculator
    {
        [Guid("D2490451-1874-484F-849B-6F267682BA0E")]
        public delegate void OnAdditionDoneDelegate();
        
        private event OnAdditionDoneDelegate OnAdditionDone;
        
        public double Addition(double val1, double val2)
        {
            TriggerAdditionEvent();
            return val1 + val2;
        }
        
        public void TriggerAdditionEvent() => Task.Run(() => OnAdditionDone?.Invoke());
    }
}
