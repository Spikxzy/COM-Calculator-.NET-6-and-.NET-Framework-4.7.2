using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace COMCalculatorFrameworkVersion.Interfaces
{
    [ComVisible(true)]
    [Guid("180BA672-E0AD-4CEC-BD47-2E759FF9FE33")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICalculator
    {
        [DispId(1)]
        double Addition(double val1, double val2);

        [DispId(2)]
        void TriggerAdditionEvent();
    }
}
