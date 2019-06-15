using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace ClassLibrary1
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [Guid("D24536F5-ECD9-482B-8C57-C9EC2195546D")]
    public class Class1 : IClass1
    {
        public string Hello()
        {
            return "Hello world, how are you?";
        }
        public List<string> HelloList()
        {
            List<string> rtn_list = new List<string>();
            rtn_list.Add("String1");
            rtn_list.Add("String2");
            rtn_list.Add("String3");
            rtn_list.Add("String4");
            return rtn_list;
        }
    }
}
