using Bytescout.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp4
{
    internal interface IReadExel
    {
        public Worksheet sheet { get; set; }
        public Spreadsheet doc { get; set; }
        public void Close();

    }
}
