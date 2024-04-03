using Bytescout.Spreadsheet;

namespace ConsoleApp4
{
    internal class ReadExel : IReadExel
    {
        public Worksheet sheet { get; set; }
        public Spreadsheet doc { get; set; }
        private string path;

        public ReadExel(string path, int numberSheet = 0)
        {
            doc = new Spreadsheet();
            doc.LoadFromFile(path);
            this.path = path;
            sheet = doc.Workbook.Worksheets[numberSheet];
        }

        public void Close()
        {
            doc.SaveAs(path);
            doc.Close();
        }
    }
}
//sheet.Cell(0, 0).Value = "jjj";