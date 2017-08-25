using System;
using System.Text.RegularExpressions;
using System.Xml.Serialization;

namespace ExcelImport
{

    public enum CellType
    {
        Number,
        String,
        Null

    }
    public class Cell
    {

        [XmlIgnore]
        private string _value = string.Empty;

        [XmlIgnore]
        private string _type = string.Empty;

        [XmlAttribute("t")]
        public string TypeElement
        {
            get { return _type; }
            set { _type = value; }
        }

        /// <summary>
        /// Used for converting from Excel column/row to column index starting at 0
        /// </summary>
        [XmlAttribute("r")]
        public string CellReference
        {
            get
            {
                return ColumnIndex.ToString();
            }
            set
            {
                ColumnIndex = GetColumnIndex(value);
            }
        }

        public CellType Type
        {
            get
            {
                if (string.IsNullOrEmpty(TypeElement))
                    return CellType.Number;
                else if (TypeElement.Equals("s") || Type.Equals("str"))
                    return CellType.String;
                else
                    return CellType.Null;

            }
        }
        /// <summary>
        /// Original value of the Excel cell
        /// </summary>
        [XmlElement("v")]
        public string Value
        {
            get
            {
                return _value;
            }
            set
            {
                if (TypeElement.Equals("s"))
                {
                    var si = WorkbookFactory.SharedStrings.si;
                    var index = Convert.ToInt32(value);

                    _value = string.Empty;

                    if (si[index].t != null)
                    {
                        _value = si[index].t;
                    }
                    else if (si[index].r != null)
                    {
                        foreach (var fragmentString in si[index].r)
                            _value += fragmentString.t;
                    }
                }
                else
                {
                    _value = value;
                }

            }
        }
        /// <summary>
        /// Index of the orignal Excel cell column starting at 0
        /// </summary>
        [XmlIgnore]
        public int ColumnIndex;

        private int GetColumnIndex(string CellReference)
        {
            string colLetter = new Regex("[A-Za-z]+").Match(CellReference).Value.ToUpper();
            int colIndex = 0;

            for (int i = 0; i < colLetter.Length; i++)
            {
                colIndex *= 26;
                colIndex += (colLetter[i] - 'A' + 1);
            }
            return colIndex - 1;
        }
    }
}
