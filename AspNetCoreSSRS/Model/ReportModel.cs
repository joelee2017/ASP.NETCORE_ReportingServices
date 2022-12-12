using System.Collections.Generic;

namespace AspNetCoreSSRS
{
    public class ReportsModel
    {
        public List<ReportModel> Reports { get; set; }
    }

    public class ReportModel
    {
       /// <summary>
        /// 水流編號
        /// </summary>
        public int Id { get; set; }

        /// <summary>
        /// 任務名稱
        /// </summary>
        public string MissionName { get; set; }

        /// <summary>
        /// 排程任務規則
        /// </summary>
        public string Crontab { get; set; }
        /// <summary>
        /// 是否啟用該設定
        /// </summary>
        public bool Enabled { get; set; }

        /// <summary>
        /// 執行方式：立即執行 0 ，排程執行 1
        /// </summary>
        public int ExecuteType { get; set; }
        /// <summary>
        /// 報表服務器路徑
        /// </summary>
        public string ReportServerWsdlUrl { get; set; }
        /// <summary>
        /// 0 郵寄， 1 指定路徑、2 郵寄加指定路徑、3 服務重啟，
        /// 新增 docx、xlsx
        /// 4 指定路、5 指定路徑、6 郵寄加指定路徑
        /// </summary>
        public ActionType Action { get; set; }
        /// <summary>
        /// 收件人
        /// </summary>
        public string Mail_To { get; set; }
        /// <summary>
        /// 副本
        /// </summary>
        public string CC { get; set; }
        /// <summary>
        /// 密件副本
        /// </summary>
        public string Bcc { get; set; }
        /// <summary>
        /// 主旨
        /// </summary>
        public string Subject { get; set; }
        /// <summary>
        /// 信件內文
        /// </summary>
        public string Mail_Content { get; set; }
        /// <summary>
        /// 是否含附件
        /// </summary>
        public bool Attached { get; set; }
        /// <summary>
        /// 是否加密
        /// </summary>
        public bool IsEncrypt { get; set; }
        /// <summary>
        /// 加密密碼
        /// </summary>
        public string Password { get; set; }
        /// <summary>
        /// 是否為html
        /// </summary>
        public bool IsHtml { get; set; }
        /// <summary>
        /// 報表路徑及名稱
        /// </summary>
        public string Rerpot_Path { get; set; }
        /// <summary>
        /// 報表回傳格式
        /// </summary>
        public string Render_Format { get; set; }
        /// <summary>
        /// 檔名及副檔名
        /// </summary>
        public string File_Name { get; set; }
        /// <summary>
        /// 指定下載路徑
        /// </summary>
        public string File_Path { get; set; }
        /// <summary>
        /// 參數，物件陣列
        /// </summary>
        public List<ParameterModel> Parameters { get; set; }
        /// <summary>
        /// 備忘錄
        /// </summary>
        public string Memo { get; set; }

        /// <summary>
        /// 建立日期
        /// </summary>
        public DateTime CreateDate { get; set; }

        /// <summary>
        /// 修改日期
        /// </summary>
        public DateTime? ModifiedDate { get; set; }

    }


   
}
