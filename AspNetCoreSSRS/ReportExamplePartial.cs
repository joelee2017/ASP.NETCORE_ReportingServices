using KGI.MailComponent.Model;
using KGI.MicrosoftTeamsComponent.Model;
using KGI.ReportComponent.Helper;
using KGI.ReportComponent.Model;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace KGI.ReportComponent
{
    // ReportService 輔助工具
    public partial class ReportExample
    {
        /// <summary>
        /// 報表至指定位置
        /// </summary>
        /// <param name="model">報表參數</param>
        private async Task SendToPath(ReportModel model, byte[] report)
        {
            if (model.IsEncrypt)
            {
                var zip = FileToZip(model.File_Name, model.Password, report);
                _credentials.WriteFile(model.File_Path + model.File_Name.Split(".")[0] + ".zip", zip);
            }
            else
            {
                _credentials.WriteFile(model.File_Path + model.File_Name, report);
            }

            var filepath = model.File_Path + model.File_Name;
            await Task.CompletedTask;
            bool isSuccess = File.Exists(filepath); //檔案是否存在
        }

        /// <summary>
        /// 報表寄送郵件
        /// </summary>
        /// <param name="model">報表參數</param>
        private async Task SendToMail(ReportModel model, byte[] report)
        {
            MailMessageModel mailMessage = new MailMessageModel();
            mailMessage.MailSubject = model.Subject;
            mailMessage.MailIsHtml = model.IsHtml;
            mailMessage.MailBody = model.Mail_Content;
            mailMessage.MailRecipientAddress = model.Mail_To;
            mailMessage.MailSenderAddress = _smtpApi.SenderAddress;
            mailMessage.MailSenderDisplayName = _smtpApi.SenderDisplayName;

            if (!string.IsNullOrWhiteSpace(model.CC))
                mailMessage.MailCC = model.CC.Replace(",", ";").Split(';').ToList();

            if (!string.IsNullOrWhiteSpace(model.Bcc))
                mailMessage.MailBCC = model.Bcc.Replace(",", ";").Split(';').ToList();

            // 是否加附件
            if (model.Attached)
            {
                var fils = new Dictionary<string, byte[]>();

                if (model.IsEncrypt)
                {
                    var zip = FileToZip(model.File_Name, model.Password, report);
                    fils.Add(model.File_Name.Split(".")[0] + ".zip", zip);
                }
                else
                {
                    fils.Add(model.File_Name, report);
                }

                mailMessage.MailAttachmentByte = fils;
            }

            await _mailService.Send(mailMessage); //寄送是否成功
        }

        /// <summary>
        /// 報表寄送密碼郵件
        /// </summary>
        /// <param name="model">報表參數</param>
        private async Task SendPasswordMail(ReportModel model)
        {
            MailMessageModel mailMessage = new MailMessageModel();
            mailMessage.MailSubject = "[密碼通知]" + model.Subject;
            mailMessage.MailIsHtml = true;
            mailMessage.MailBody = "<h4 style='color:red;'> 密碼：" + model.Password + "</h4>";
            mailMessage.MailRecipientAddress = model.Mail_To;
            mailMessage.MailSenderAddress = _smtpApi.SenderAddress;
            mailMessage.MailSenderDisplayName = _smtpApi.SenderDisplayName;

            if (!string.IsNullOrWhiteSpace(model.CC))
                mailMessage.MailCC = model.CC.Replace(",", ";").Split(';').ToList();

            if (!string.IsNullOrWhiteSpace(model.Bcc))
                mailMessage.MailBCC = model.Bcc.Replace(",", ";").Split(';').ToList();

            await _mailService.Send(mailMessage); //寄送是否成功
        }

        /// <summary>
        /// 確認報表伺服器 - 是否為啟動
        /// </summary>
        /// <param name="model">報表參數</param>
        /// <returns>bool</returns>
        private async Task<bool> CallReportService(ReportModel model, byte[] report)
        {
            string reportUrl = model.ReportServerWsdlUrl ?? _configuration.GetSection("ReportServerWsdlUrl").Get<string>();

            _reportManager = new ReportManager(reportUrl, _userName, _passWrod, _domain);

            bool isSuccess = await _reportManager.CallReport(model.Rerpot_Path);

            await _microsoftTeamsService.HttpClientTeamsMessage(new TeamsModel()
            {
                Title = $"SSRS-WebApi 呼叫服務確認",
                Text = $"<strong>服務確認：</strong> " + (isSuccess ? "呼叫成功" : "呼叫失敗"),
                Url = _teamsBot
            });


            return isSuccess;
        }

        /// <summary>
        /// 檔案壓縮
        /// </summary>
        /// <param name="fileName">檔名</param>
        /// <param name="res">資料</param>
        /// <returns></returns>
        private byte[] FileToZip(string fileName, string passowrd, byte[] res)
        {
            byte[] zip;
            zip = DotNetZip.ByteArrayToZip(res, passowrd, fileName);
            return zip;
        }

        /// <summary>
        /// 產生密碼
        /// </summary>
        private string GetEncryptPassword()
        {
            string str = string.Empty;
            str = Enumerable.Range(0, 9)
               .OrderBy(i => Guid.NewGuid())
               .Aggregate(string.Empty, (current, next) => current + "" + next).Substring(0, 4);

            return str;
        }
    }
}
