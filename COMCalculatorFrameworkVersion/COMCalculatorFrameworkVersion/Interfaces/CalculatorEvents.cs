using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace COMCalculatorFrameworkVersion.Interfaces
{
    [ComVisible(true)]
    [Guid("D3959B7D-20FF-4579-B57B-1BC285738B9A")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface CalculatorEvents
    {
        [DispId(1)]
        void OnAdditionDone();
    }
}
