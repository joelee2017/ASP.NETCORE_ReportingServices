using KGI.MailComponent.Interface;
using KGI.MailComponent.Model;
using KGI.MicrosoftTeamsComponent.Interface;
using KGI.MicrosoftTeamsComponent.Model;
using KGI.ReportComponent.Helper;
using KGI.ReportComponent.Interface;
using KGI.ReportComponent.Model;
using Microsoft.Extensions.Configuration;
using ReportExecution2005;
using SimpleImpersonation;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Threading.Tasks;
using Utility.Helper;

namespace KGI.ReportComponent
{
    public class ReportService : IReportService
    {
        IConfiguration _configuration;
        IMailService _mailService;
        SmtpApiModel _smtpApi;
        NasModel _nas;
        string _userName;
        string _passWrod;
        string _domain;
        UserCredentials _credentials;
        IMicrosoftTeamsService _microsoftTeamsService;
        string _teamsBot;

        ReportManager _reportManager;

        public ReportService(IConfiguration configuration, IMailService mailService, IMicrosoftTeamsService microsoftTeamsService)
        {
            _configuration = configuration;
            _mailService = mailService;

            _smtpApi = _configuration.GetSection("SmtpApi").Get<SmtpApiModel>();
            _mailService.SmtpApiUri = _smtpApi.Uri;

            _nas = _configuration.GetSection("NAS").Get<NasModel>();
            _userName = _nas.Username.DecryptionAES();
            _passWrod = _nas.Password.DecryptionAES();
            _domain = _nas.Domain.DecryptionAES();
            _credentials = new UserCredentials(_domain, _userName, _passWrod);
            _microsoftTeamsService = microsoftTeamsService;
            _teamsBot = _configuration.GetSection("SystemMessage").Get<string>();
        }

        public async Task ReportProcess(ReportModel model)
        {
            if (model.IsEncrypt)
            {
                model.Password = GetEncryptPassword();
            }

            byte[] report = default;
            if ((int)model.Action != 3)
            {
                switch ((int)model.Action)
                {
                    case 1:
                    case 2:
                        report = await this.GetReport(model);
                        break;
                    case 4:
                    case 5:
                    case 6:
                        report = await this.Report2010Process(model);
                        break;
                    default:
                        break;
                }
            }

            Dictionary<int, Func<ReportModel, byte[], Task>> act = DictionaryFuncMethod();

            await act[(int)model.Action].Invoke(model, report);

            if (model.IsEncrypt)
            {
                await SendPasswordMail(model);
            }
        }

        /// <summary>
        /// 報表流程種類方法
        /// 0 發送至信箱
        /// 1 發送至路徑資料夾
        /// 2 發送至信箱及路徑資料夾
        /// 3 呼叫 ssrs 確認是否有反應
        /// </summary>        
        private Dictionary<int, Func<ReportModel, byte[], Task>> DictionaryFuncMethod()
        {
            Dictionary<int, Func<ReportModel, byte[], Task>> act = new Dictionary<int, Func<ReportModel, byte[], Task>>();

            act.Add(0, async (rep, report) => { await SendToMail(rep, report); });
            act.Add(1, async (rep, report) => { await SendToPath(rep, report); });
            act.Add(2, async (rep, report) => { await SendToMail(rep, report); await SendToPath(rep, report); });
            act.Add(3, async (rep, report) => { await CallReportService(rep, report); });

            return act;
        }

        /// <summary>
        /// 取得報表資料
        /// </summary>
        /// <param name="model">報表參數</param>
        /// <returns>byte[]</returns>
        private async Task<byte[]> GetReport(ReportModel model)
        {
            string reportUrl = model.ReportServerWsdlUrl ?? _configuration.GetSection("ReportServerWsdlUrl").Get<string>();

            _reportManager = new ReportManager(reportUrl, _userName, _passWrod, _domain);

            var parameters = _reportManager.ConvertParameterValues(model.Parameters).Result;

            byte[] res = await GetReportByteArray(parameters, model.Rerpot_Path, model.Render_Format);

            return res;
        }

        // 呼叫報表
        private async Task<byte[]> GetReportByteArray(ParameterValue[] parameters, string path, string renderFormat, int count = 0)
        {
            byte[] res = default;
            try
            {
                res = await _reportManager.RenderReport(path, renderFormat, parameters);
            }
            catch (TimeoutException) // SSRS 呼叫有可能會失敗，失敗後在呼叫最多三次。
            {
                //  距離上次呼叫間隔 1 分鐘
                await Task.Delay(60000);

                if (count < 3)
                {
                    count++;
                    if (res == null || res.Length == 0)
                    {
                        res = await this.GetReportByteArray(parameters, path, renderFormat, count);
                    }

                    if (res != null)
                    {
                        return res;
                    }
                }

                throw;
            }

            return res;
        }

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

        public Task<byte[]> Report2010Process(ReportModel model)
        {
            string reportUrl = model.ReportServerWsdlUrl ?? _configuration.GetSection("ReportServerWsdlUrl").Get<string>();

            Dictionary<string, ReportFormats> format = new Dictionary<string, ReportFormats>();
            format.Add("PDF", ReportFormats.Pdf);
            format.Add("WORD", ReportFormats.Docx);
            format.Add("EXCEL", ReportFormats.Xlsx);

            using (ReportManager2010 report = new ReportManager2010
            {
                ReportServerPath = reportUrl,
                Format = format.TryGetValue(model.Render_Format, out ReportFormats reportFormat) == true ? reportFormat : ReportFormats.Pdf,
                ReportPath = model.Rerpot_Path
            })
            {

                if (model.Parameters.Any())
                {
                    List<ParameterModel> newparameter = new List<ParameterModel>();
                    foreach (var par in model.Parameters)
                    {
                        // {"Name":"WFUND_ID","Value":"L010,K002,H003"} 單一key，帶多參時使用
                        //var aa = par.Value.ToString().Split(",");
                        //if (par.Value.ToString().Split(",").Length > 0)
                        //{
                        //    foreach (var par2 in par.Value.ToString().Split(","))
                        //    {
                        //        report.Params.Add(par.Name, par2);
                        //    }
                        //}
                        //else
                        //{
                        //    report.Params.Add(par.Name, par.Value.ToString());
                        //}
                        report.Params.Add(par.Name, par.Value.ToString());
                    }
                }

                report.Credentials = new NetworkCredential(_userName, _passWrod, _domain);
                MemoryStream ms = new MemoryStream();
                report.Render().CopyTo(ms);
                byte[] bytes = ms.ToArray();
                ms.Close();
                return Task.FromResult(bytes);
            }
        }
    }
}
