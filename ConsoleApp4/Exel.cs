using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp4
{
    internal class Exel: ReadExel
    {
        private int num = 0;
        private string user;
        public Exel(string user,string path, int numberSheet = 0):base(path, numberSheet)
        {
            this.user = user;  
            string textnum;
            if (sheet.Cell(sheet.NotEmptyRowMax, 0).Value == null)
                textnum = "0";
            else
                textnum = sheet.Cell(sheet.NotEmptyRowMax, 0).Value.ToString()??"0";
            num = int.Parse(textnum);
            num++;

        }

        public void Add(string product, string price)
        {
            sheet.Cell(sheet.NotEmptyRowMax + 1, 0).Value = num;
            num++;
            sheet.Cell(sheet.NotEmptyRowMax, 1).Value = user;
            sheet.Cell(sheet.NotEmptyRowMax, 2).Value = product;
            sheet.Cell(sheet.NotEmptyRowMax, 3).Value = price;
            sheet.Cell(sheet.NotEmptyRowMax, 4).Value = DateTime.Now.ToString();
        }
    }
}
