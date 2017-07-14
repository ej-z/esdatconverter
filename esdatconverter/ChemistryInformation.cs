using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace esdatconverter
{
    public class ChemistryInfoModel
    {
        public int Order { get; set; }
        public string Element { get; set; }
        public string ChemCode { get; set; }
        public string MethodType { get; set; }
        public string MethodName { get; set; }
        public string SampleCode { get; set; }
        public string Prefix { get; set; }
        public double Result { get; set; }
        public string Result_Unit { get; set; }
        public string Total_or_Filtered { get; set; }
        public string Result_Type { get; set; }
        public DateTime? Extraction_Date { get; set; }
        public DateTime? Analysed_Date { get; set; }
        public double? EQL { get; set; }
        public string EQL_Units { get; set; }
        public string Comments { get; set; }
        public string Lab_Qualifier { get; set; }
        public double? UCL { get; set; }
        public double? LCL { get; set; }

        public ChemistryInfoModel()
        {
            Total_or_Filtered = "T";
        }

    }
    class ChemistryInformation : IEnumerable<ChemistryInfoModel>
    {
        private List<ChemistryInfoModel> chemistryInformation;
        public ChemistryInformation(string filename, ElementInformation elementInformation)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;
            string element;
            int rCnt;
            int cCnt;
            int rw = 0;
            int cl = 0;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(filename, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets["Report Face"];

            range = xlWorkSheet.UsedRange;
            rw = 98;
            cl = 14;
            chemistryInformation = new List<ChemistryInfoModel>();
            string family = string.Empty;
            rCnt = 3;
            int order = 0;
            while(rCnt<rw)
            {
                element = GetValue(range.Cells[rCnt, 1] as Excel.Range).Trim();

                if (string.IsNullOrEmpty(element))
                {
                    while (GetValue(range.Cells[rCnt, 1] as Excel.Range) == string.Empty)
                        rCnt++;
                    family = GetValue(range.Cells[rCnt, 1] as Excel.Range).Trim();
                    order++;
                    rCnt = rCnt + 1;
                    element = GetValue(range.Cells[rCnt, 1] as Excel.Range).Trim();
                    if(string.IsNullOrEmpty(element))
                    {
                        element = family;
                        rCnt = rCnt - 1;
                    }
                }

                string Unit = string.Empty;
                string Type = string.Empty;
                double eql = double.NaN;
                

                var str = GetValue(range.Cells[rCnt, 2] as Excel.Range);
                if (str.Equals("surr.", StringComparison.InvariantCultureIgnoreCase))
                {
                    Unit = "%";
                    Type = "SUR";
                }
                else if (str.Equals("%", StringComparison.InvariantCultureIgnoreCase))
                {
                    Unit = "%";
                    Type = "REG";
                }
                else
                {
                    double.TryParse(str, out eql);
                    Unit = "mg/kg";
                    Type = "REG";
                }                 

                for (cCnt = 3; cCnt <= cl; cCnt++)
                {                    
                    var chemistryInfo = new ChemistryInfoModel();
                    chemistryInfo.Order = order;
                    chemistryInfo.Element = element;                  
                    chemistryInfo.Result_Unit = chemistryInfo.EQL_Units = Unit;
                    chemistryInfo.Result_Type = Type;

                    if (eql != double.NaN)
                        chemistryInfo.EQL = eql;
                    else
                        chemistryInfo.EQL = null;

                    chemistryInfo.SampleCode = GetValue(range.Cells[1, cCnt] as Excel.Range);

                    var result = GetValue(range.Cells[rCnt, cCnt] as Excel.Range);
                    ParseResult(result, chemistryInfo);

                    var key = family + element;
                    chemistryInfo.ChemCode = elementInformation[key].ChemCode;
                    chemistryInfo.MethodName = elementInformation[key].MethodName;
                    chemistryInfo.MethodType = elementInformation[key].MethodType;

                    chemistryInformation.Add(chemistryInfo);
                }
                rCnt++;
            }

            xlWorkBook.Close(true, null, null);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);
        }

        public int Length { get { return chemistryInformation.Count; } }

        public IEnumerator<ChemistryInfoModel> GetEnumerator()
        {
            return chemistryInformation.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        private string GetValue(Excel.Range value)
        {
            if (value == null || value.Value2 == null)
                return string.Empty;

            return value.Value2.ToString();
        }

        private void ParseResult(string result, ChemistryInfoModel chemistryInfo)
        {
            if (string.IsNullOrEmpty(result))
                return;

            int c;
            var prefix = result[0].ToString();
            var suffix = result[result.Length - 1].ToString();

            if(int.TryParse(prefix,out c) && int.TryParse(suffix, out c))
            {
                if(chemistryInfo.EQL_Units == "%")
                    chemistryInfo.Result = Math.Round(double.Parse(result)*100);
                else
                    chemistryInfo.Result = Math.Round(double.Parse(result),1);
            }
            else if(int.TryParse(prefix, out c))
            {
                chemistryInfo.Result = Math.Round(double.Parse(result.Remove(result.Length - 1))*100);
            }
            else
            {
                chemistryInfo.Result = Math.Round(double.Parse(result.Remove(0,1)),1);
                chemistryInfo.Prefix = prefix;
            }
        }
    }
}
