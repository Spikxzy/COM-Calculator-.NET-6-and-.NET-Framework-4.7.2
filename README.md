# COM Calculator .NET 6 and .NET Framework 4.7.2
 
This repository contains two solutions and an Excel worksheet with macros enabled. It's goal is to show how COM components are created in .NET Framework 4.7.2 and .NET 6. It also provides a COM client that can be used to access the COM components in VBA. This is done by providing a simple calculator that can add two double values, as well as trigger a COM event.

## .NET 6 Solution

Just building this solution will create and register the COM component. By using the MIDL compiler with post build events a TLB is created with the usage of an IDL file. It is then also automatically registerd inside the Windows registry. However, for automatic registration to work Visual Studio has to be started as administrator.

## .NET Framework Solution

Building this solution will result in a .dll file. This file can afterwards be used to generate a TLB and register the COM component. RegAsm can be used to generate a TLB and the register the COM component. Run the following command 'regasm "\<Path-To-Dll\>" /tlb: "\<TLB-Name\>" /codebase' inside a command line tool which is running as administrator. Afterwards the COM component will be registered.

## Calculator Client
 
The Excel workbook contains a COM client written in VBA. This client has buttons which execute the basic functionality that the calculator supports ('Addition', 'TriggerAdditionEvent').
