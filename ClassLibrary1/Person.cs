using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace ClassLibrary1
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [Guid("85D4A03F-7B37-44B6-8EBF-24C689AE4A49")]
    public class Person //: IPerson
    {
        public int Id { get; set; }
        public string FName { get; set; }
        public string LName { get; set; }
    }
}
