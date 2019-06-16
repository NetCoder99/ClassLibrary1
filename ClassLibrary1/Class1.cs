using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;

// %SystemRoot%\Microsoft.NET\Framework\v4.0.30319\regasm.exe ClassLibrary1.dll
// %SystemRoot%\Microsoft.NET\Framework64\v4.0.30319\regasm.exe ClassLibrary1.dll

namespace ClassLibrary1
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [Guid("D24536F5-ECD9-482B-8C57-C9EC2195546D")]
    public class Class1 : IClass1
    {
        public string Hello()
        {
            return "Hello world, how are you doing?";
        }
        public ArrayList HelloList()
        {
            ArrayList rtn_list = new ArrayList
            {
                "String1",
                "String2",
                "String3",
                "String4",
                "String5",
                "String6"
            };
            return rtn_list;
        }
    }
}
