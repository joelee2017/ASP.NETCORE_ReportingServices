using AspNetCoreSSRS.Conversion;
using ssrstest001;
using System;
using System.IO;
using System.Linq;

namespace AspNetCoreSSRS
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("ReportManager Go!!!");

            //讀取參數 json 文件檔
            string dir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            var data = JsonRead.FromFile<ReportsModel>(Path.Combine(dir, "Configuration\\appsettings.json"));

            //使用json 給參數，也能直接連db         
            foreach (var item in data.Reports)
            {
                ReportManager reportManager = new ReportManager(item.ReportServerWsdlUrl);

                var parameters = item.Parameters.ToDictionary(str => str.Name, str => str.Value);
                var result = reportManager.RenderReport(item.Rerpot_Path, parameters);

                FileStream stream = File.Create("D:\\report_aspnetcore.pdf", result.Result.Length);
                stream.Write(result.Result, 0, result.Result.Length);
                stream.Close();
            }
        }

        //暫不使用Configuration注入
        //public static IConfigurationRoot GetConfiguation()
        //{
        //    var builder = new ConfigurationBuilder()
        //        .SetBasePath(Directory.GetCurrentDirectory())
        //        .AddJsonFile("Configuration\\appsettings.json", false)
        //        .Build();

        //    return builder;
        //}
    }
}
