using Microsoft.VisualBasic.FileIO;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace esdatconverter
{ 
    public class ElementInfoModel
    {
        public string ChemCode { get; set; }
        public string MethodType { get; set; }
        public string MethodName { get; set; }
    }

    public class ElementInformation
    {
        private Dictionary<string, ElementInfoModel> elementInformation;
        public ElementInformation(string filename)
        {
            using (TextFieldParser parser = new TextFieldParser(filename))
            {
                parser.Delimiters = new string[] { "," };
                parser.HasFieldsEnclosedInQuotes = true;
                elementInformation = new Dictionary<string, ElementInfoModel>();
                while (true)
                {
                    string[] values = parser.ReadFields();
                    if (values == null)
                    {
                        break;
                    }

                    var elementInfo = new ElementInfoModel()
                    {
                        ChemCode = values[2],
                        MethodType = values[3],
                        MethodName = values[4]
                    };

                    elementInformation.Add(values[0].Trim() + values[1].Trim(), elementInfo);
                }
            }
        }

        public ElementInfoModel this[string key]
        {
            get
            {
                if (!elementInformation.ContainsKey(key))
                    throw new KeyNotFoundException(key+" does not have any element information.");

                return elementInformation[key];
            }
        }
    }
}
