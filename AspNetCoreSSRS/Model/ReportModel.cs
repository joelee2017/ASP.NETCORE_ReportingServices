using System.Collections.Generic;

namespace ssrstest001
{
    public class ReportsModel
    {
        public List<ReportModel> Reports { get; set; }
    }

    public class ReportModel
    {
        /// <summary>
        /// 報表服務器路徑
        /// </summary>
        public string ReportServerWsdlUrl { get; set; }

        /// <summary>
        /// 報表路徑及名稱
        /// </summary>
        public string Rerpot_Path { get; set; }

        /// <summary>
        /// 報表回傳格式
        /// </summary>
        public string Render_Format { get; set; }

        /// <summary>
        /// 參數
        /// </summary>
        public List<ParameterModel> Parameters { get; set; }

        /// <summary>
        /// 指定下載路徑
        /// </summary>
        public string Destination_Path { get; set; }


        /// <summary>
        /// 檔名及副檔名
        /// </summary>
        public string File_Name { get; set; }

    }


   
}
