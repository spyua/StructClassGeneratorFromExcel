using System.Configuration;

namespace StructGenerator
{
    public class AppSetting
    {
        public static class SingletonHolder
        {
            static SingletonHolder() { }
            internal static readonly AppSetting INSTANCE = new AppSetting();
        }
        public static AppSetting Instance { get { return SingletonHolder.INSTANCE; } }

        public int StarSheet { get; set; }
        public int StarReadRow { get; set; }
        public int StarReadCloumn { get; set; }

        public string ExcelPath { get; set; }
        public string OutFileName { get; set; }

        public string ExportPath { get; set; }

        public AppSetting()
        {
            StarSheet = GetAppConfigIntVaule("StarSheet");
            StarReadRow = GetAppConfigIntVaule("StarReadRow");
            StarReadCloumn = GetAppConfigIntVaule("StarReadCloumn");

            ExcelPath = GetAppConfigValue("ExcelPath");
            ExportPath = GetAppConfigValue("ExportPath");
            OutFileName = GetAppConfigValue("OutFileName");
        }

        private string GetAppConfigValue(string value)
        {
            return ConfigurationManager.AppSettings[value];
        }
        private int GetAppConfigIntVaule(string value)
        {
            return int.Parse(ConfigurationManager.AppSettings[value]);
        }


    }
}
