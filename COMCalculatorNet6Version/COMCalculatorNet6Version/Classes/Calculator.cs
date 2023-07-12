// <copyright file="Calculator.cs" company="PlaceholderCompany">
// Copyright (c) PlaceholderCompany. All rights reserved.
// </copyright>

namespace COMCalculatorNet6Version.Classes
{
    using System;
    using System.Reflection;
    using System.Runtime.InteropServices;
    using System.Runtime.InteropServices.ComTypes;
    using System.Threading.Tasks;
    using COMCalculatorNet6Version.COM;
    using COMCalculatorNet6Version.Interfaces;
    using Microsoft.Win32;

    [ComVisible(true)]
    [Guid(AssemblyInfo.CalculatorClassGuid)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(CalculatorEvents))]
    [ProgId("ComCalculatorTest")]
    public class Calculator : ICalculator
    {
        [ComRegisterFunction]
        public static void DllRegisterServer(Type t)
        {
            // Additional CLSID entries
            using (RegistryKey key = Registry.ClassesRoot.CreateSubKey(@"CLSID\{" + AssemblyInfo.CalculatorClassGuid + @"}"))
            {
                using (RegistryKey typeLib = key.CreateSubKey(@"TypeLib"))
                {
                    typeLib.SetValue(string.Empty, "{" + AssemblyInfo.LibraryGuid + "}", RegistryValueKind.String);
                }
            }

            // Interface entries
            using (RegistryKey key = Registry.ClassesRoot.CreateSubKey(@"Interface\{" + AssemblyInfo.CalculatorInterfaceGuid + @"}"))
            {
                using (RegistryKey typeLib = key.CreateSubKey(@"ProxyStubClsid32"))
                {
                    typeLib.SetValue(string.Empty, "{00020424-0000-0000-C000-000000000046}", RegistryValueKind.String);
                }

                using (RegistryKey typeLib = key.CreateSubKey(@"TypeLib"))
                {
                    typeLib.SetValue(string.Empty, "{" + AssemblyInfo.LibraryGuid + "}", RegistryValueKind.String);
                    Version version = typeof(AssemblyInfo).Assembly.GetName().Version;
                    typeLib.SetValue("Version", string.Format("{0}.{1}", version.Major, version.Minor), RegistryValueKind.String);
                }
            }

            // TLB entries
            using (RegistryKey key = Registry.ClassesRoot.CreateSubKey(@"TypeLib\{" + AssemblyInfo.LibraryGuid + @"}"))
            {
                Version version = typeof(AssemblyInfo).Assembly.GetName().Version;
                using (RegistryKey keyVersion = key.CreateSubKey(string.Format("{0}.{1}", version.Major, version.Minor)))
                {
                    // typelib key for 32 bit
                    keyVersion.SetValue(string.Empty, AssemblyInfo.Attribute<AssemblyDescriptionAttribute>().Description, RegistryValueKind.String);
                    using (RegistryKey keyWin32 = keyVersion.CreateSubKey(@"0\win32"))
                    {
                        keyWin32.SetValue(string.Empty, Path.ChangeExtension(Assembly.GetExecutingAssembly().Location, ".comhost.tlb"), RegistryValueKind.String);
                    }

                    // typelib key for 64 bit
                    keyVersion.SetValue(string.Empty, AssemblyInfo.Attribute<AssemblyDescriptionAttribute>().Description, RegistryValueKind.String);
                    using (RegistryKey keyWin64 = keyVersion.CreateSubKey(@"0\win64"))
                    {
                        keyWin64.SetValue(string.Empty, Path.ChangeExtension(Assembly.GetExecutingAssembly().Location, ".comhost.tlb"), RegistryValueKind.String);
                    }

                    using (RegistryKey keyFlags = keyVersion.CreateSubKey(@"FLAGS"))
                    {
                        keyFlags.SetValue(string.Empty, "0", RegistryValueKind.String);
                    }
                }
            }
        }

        [ComUnregisterFunction]
        public static void DllUnregisterServer(Type t)
        {
            Registry.ClassesRoot.DeleteSubKeyTree(@"CLSID\{" + AssemblyInfo.CalculatorClassGuid + @"}", false);
            Registry.ClassesRoot.DeleteSubKeyTree(@"Interface\{" + AssemblyInfo.CalculatorInterfaceGuid + @"}", false);
            Registry.ClassesRoot.DeleteSubKeyTree(@"TypeLib\{" + AssemblyInfo.LibraryGuid + @"}", false);
        }

        [Guid(AssemblyInfo.OnAdditionDoneDelegateGuid)]
        public delegate void OnAdditionDoneDelegate();

        public event OnAdditionDoneDelegate OnAdditionDone;

        public double Addition(double firstValue, double secondValue)
        {
            TriggerAdditionEvent();
            return firstValue + secondValue;
        }

        public void TriggerAdditionEvent() => Task.Run(() =>
        {
            if (this.OnAdditionDone != null && Marshal.IsComObject(this.OnAdditionDone))
            {
                var dispatchedInstance = Marshal.GetIDispatchForObject(this.OnAdditionDone);
                Guid emptyGuid = Guid.Empty;
                DISPPARAMS args = new();

                unsafe
                {
                    var dispatchInvoke = (delegate* unmanaged[Stdcall]<IntPtr, int, Guid*, short, short, void*, void*, void*, void*, int>)(*(*(void***)dispatchedInstance + 6));
                    int hr = dispatchInvoke(dispatchedInstance, 1, &emptyGuid, 0, 1, &args, null, null, null);
                }

                Marshal.Release(dispatchedInstance);
            }
            else if (this.OnAdditionDone != null)
            {
                this.OnAdditionDone.Invoke();
            }

            /*foreach (Delegate d in this.OnAdditionDone.GetInvocationList())
            {
                Console.WriteLine(d.ToString());
                if (d.Target != null && Marshal.IsComObject(d.Target))
                {
                    Console.WriteLine(d.ToString());
                    var inst = Marshal.GetIDispatchForObject(d.Target!);
                    Guid empty = Guid.Empty;
                    DISPPARAMS args = new();

                    unsafe
                    {
                        var dispatchInvoke = (delegate* unmanaged[Stdcall]<IntPtr, int, Guid*, short, short, void*, void*, void*, void*, int>)(*(*(void***)inst + 6));
                        int hr = dispatchInvoke(inst, 1, &empty, 0, 1, &args, null, null, null);
                    }

                    Marshal.Release(inst);
                }
                else
                {
                    d.DynamicInvoke();
                }
            }
            */
        });
    }
}
