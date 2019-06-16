using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace ClassLibrary1
{
    [ComVisible(true)]
    [Guid("67F6AA4C-A9A5-4682-98F9-15BDF2246A74")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IClass1
    {
        string Hello();
        ArrayList HelloList();
    }

}