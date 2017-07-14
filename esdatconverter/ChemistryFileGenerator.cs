using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace esdatconverter
{
    public class ChemistryFileGenerator
    {
        private static string[] headers = { "SampleCode", "ChemCode", "OriginalChemName", "Prefix", "Result", "Result_Unit", "Total_or_Filtered", "Result_Type", "Method_Type", "Method_Name", "Extraction_Date", "Analysed_Date", "EQL", "EQL_Units", "Comments", "Lab_Qualifier", "UCL", "LCL" };

        public static void Generate(string LabInfoFileName, string LabDataFileName, string GeneratedFileName)
        {
            var elementInformation = new ElementInformation(LabInfoFileName);
            var chemistryInformation = new ChemistryInformation(LabDataFileName, elementInformation);
            Generate(GeneratedFileName, chemistryInformation);
        }
        private static void Generate(string filename, ChemistryInformation chemistryInformation)
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine(string.Join(",", headers));
            var cnt = headers.Length;

            var j = 2;
            foreach (var chemInfo in chemistryInformation.OrderBy(x=> x.Order))
            {
                var format = "\"{0}\",\"{1}\",\"{2}\",\"{3}\",{4},\"{5}\",\"{6}\",\"{7}\",\"{8}\",\"{9}\",\"{10}\",\"{11}\",{12},\"{13}\",\"{14}\",\"{15}\",{16},{17}";

                sb.AppendLine(string.Format(format, chemInfo.SampleCode, chemInfo.ChemCode, chemInfo.Element, chemInfo.Prefix, chemInfo.Result, chemInfo.Result_Unit, chemInfo.Total_or_Filtered, chemInfo.Result_Type, chemInfo.MethodType, chemInfo.MethodName, chemInfo.Extraction_Date, chemInfo.Analysed_Date, chemInfo.EQL, chemInfo.EQL_Units, chemInfo.Comments, chemInfo.Lab_Qualifier, chemInfo.UCL, chemInfo.LCL));
            }

            File.AppendAllText(filename, sb.ToString());
        }
    }
}
