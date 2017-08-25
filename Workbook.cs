using System.Collections.Generic;

namespace ExcelImport
{

    public class Workbook
    {        
        public List<Worksheet> Worksheets { get; set; }

        public Workbook()
        {
            Worksheets = new List<Worksheet>();
        }
    }
}
