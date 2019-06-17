using System;
using System.Collections;
using Newtonsoft.Json;
using System.Runtime.InteropServices;

// %SystemRoot%\Microsoft.NET\Framework\v4.0.30319\regasm.exe ClassLibrary1.dll
// %SystemRoot%\Microsoft.NET\Framework64\v4.0.30319\regasm.exe ClassLibrary1.dll

namespace ClassLibrary1
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [Guid("D24536F5-ECD9-482B-8C57-C9EC2195546D")]
    public class Class1 // : IClass1
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
            rtn_list.Add("String99");
            return rtn_list;
        }

        public Person GetPerson()
        {
            Person iPerson = new Person() { Id=1, FName = "John", LName = "Smith" };
            return iPerson;
        }

        public ArrayList GetPersonList()
        {
            ArrayList rtn_list = new ArrayList();
            rtn_list.Add((Person)new Person() { Id = 1, FName = "John", LName = "Smith" });
            rtn_list.Add((Person)new Person() { Id = 2, FName = "Joan", LName = "Smith" });
            rtn_list.Add((Person)new Person() { Id = 3, FName = "Jack", LName = "Smith" });
            rtn_list.Add((Person)new Person() { Id = 4, FName = "Juan", LName = "Smith" });
            return rtn_list;
        }

        public String GetPersonListJSON()
        {
            ArrayList rtn_list = new ArrayList();
            rtn_list.Add((Person)new Person() { Id = 1, FName = "John", LName = "Smith" });
            rtn_list.Add((Person)new Person() { Id = 2, FName = "Joan", LName = "Smith" });
            rtn_list.Add((Person)new Person() { Id = 3, FName = "Jack", LName = "Smith" });
            rtn_list.Add((Person)new Person() { Id = 4, FName = "Juan", LName = "Smith" });
            string j1 = JsonConvert.SerializeObject(rtn_list);
            return JsonConvert.SerializeObject(rtn_list);
        }

    }
}
