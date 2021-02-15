using ClosedXML.Excel;
using StructGenerator.Help;
using System;
using System.Linq;
using System.Text;

namespace StructGenerator.GeneratorFactory
{
    public class StructFormat : ExcelFile, IFile
    {
 
        private AppSetting _appSetting;

        private string _decodeStr;

        public StructFormat(AppSetting appSetting)
        {
            _appSetting = appSetting;

            StarSheet = _appSetting.StarSheet;
            StarReadRaw = _appSetting.StarReadRow;
            StarReadCloumn = _appSetting.StarReadCloumn;
            ExcelPath = _appSetting.ExcelPath;
            OutFileName = _appSetting.OutFileName;
            ExportPath = _appSetting.ExportPath;
        }

        public void DecodeExcel()
        {
            var workbook = new XLWorkbook(ExcelPath);

            var decodeStr = new StringBuilder();

            decodeStr.Append("using System;\r\n");
            decodeStr.Append("using System.Runtime.InteropServices;\r\n");
            decodeStr.Append("namespace MsgStruct\r\n{\r\n");
            decodeStr.Append($"\tpublic class Msg_{OutFileName}\t\n" + "{");

            for (int i = StarSheet; i <= workbook.Worksheets.Count; i++)
            {
                var rows = workbook.Worksheet(i).RangeUsed().RowsUsed().Skip(StarReadRaw);
                decodeStr.Append("[Serializable]\r\n[StructLayout(LayoutKind.Sequential, Pack = 1)]\r\n");
                decodeStr.Append("public class " + "Msg_" + workbook.Worksheet(i).Name + " { \r\n");


                foreach (var row in rows)
                {
                    var rowNumber = row.RowNumber();
                    
                    var item = row.Cell(StarReadCloumn).GetString();
                    item = ClearInvaildSymbol(item);

                    var dataType = row.Cell(StarReadCloumn + 1).GetString();
                    var length = row.Cell(StarReadCloumn + 2).GetString();
                    var des = row.Cell(StarReadCloumn + 4).GetString();

                    string type = GetTypeCode(dataType);
                    {
                        decodeStr.Append(GetAttr(type, Convert.ToInt32(length)) + "\r\n");
                        decodeStr.Append($"    public {type}  {item}");
                        decodeStr.Append(";\r\n");
                    }
                }
                decodeStr.Append("\r\n}\r\n");
            }
            decodeStr.Append("\r\n}\r\n}");
            _decodeStr =  decodeStr.ToString();
        }

        public void GenFile()
        {
            var genOK = _decodeStr.WriteToFile(ExportPath + "DemoMsg.cs");

            if (!genOK)
                Console.WriteLine("Gen CS File Fail");
            else
                Console.WriteLine("Gen CS File Sucess");

        }

        private string ClearInvaildSymbol(string txt)
        {
            txt = txt.Replace(" ", "");
            txt = txt.Replace("-", "");
            txt = txt.Replace("(", "");
            txt = txt.Replace(")", "");
            txt = txt.Replace("-", "");
            txt = txt.Replace("~", "");
            txt = txt.Replace(".", "");

            return txt;
        }

        private string GetTypeCode(string type)
        {
            string rtn = "";
            type = type.ToLower();
            if (type.StartsWith("a")) type = "c";
            switch (type)
            {
                case "il":
                case "dw":
                    rtn = "int";
                    break;
                case "r":
                case "f":
                    rtn = "float";
                    break;
                case "i":
                case "w":
                    rtn = "short";
                    break;
                case "c":
                    rtn = "byte[]";
                    break;
                case "n":
                    rtn = "byte[]";
                    break;
            }
            return rtn;
        }

        public string GetAttr(string type, int strLen = -1)
        {
            string rtn = "";
            switch (type)
            {
                case "short":
                    rtn = $"    [MarshalAs(UnmanagedType.I2)]";
                    break;
                case "int":
                    rtn = $"    [MarshalAs(UnmanagedType.I4)]";
                    break;
                case "float":
                    rtn = $"    [MarshalAs(UnmanagedType.R4)]";
                    break;            
                case "byte[]":
                    rtn = $"    [MarshalAs(UnmanagedType.ByValArray, SizeConst = {strLen})]";
                    break;
            }
            return rtn;
        }
    }
}
