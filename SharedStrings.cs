using System;
using System.Xml.Serialization;

namespace ExcelImport
{

    [Serializable()]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot("sst", Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class SharedStrings
    {
        [XmlAttribute]
        public string uniqueCount;
        [XmlAttribute]
        public string count;
        [XmlElement("si")]
        public si[] si;

        public SharedStrings()
        {
        }
    }
    public class si 
    {
        public string t;
        [XmlElement("r")]
        public r[] r;
    }
    public class r
    {
        public string t;

    }
    
}
