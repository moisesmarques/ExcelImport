using System.Collections.Generic;
using System.Xml.Serialization;

namespace ExcelImport
{
    public class Row
    {
        [XmlElement("c")]
        public List<Cell> Cells;      

    }
}
