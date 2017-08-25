using System;
using System.Collections.Generic;
using System.Xml.Serialization;

namespace ExcelImport
{

    [Serializable()]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot("worksheet", Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class Worksheet
    {
        [XmlArray("sheetData")]
        [XmlArrayItem("row")]
        public List<Row> Rows;
    }
}
