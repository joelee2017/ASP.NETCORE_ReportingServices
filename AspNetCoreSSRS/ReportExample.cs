using KGI.MailComponent.Interface;
using KGI.MicrosoftTeamsComponent.Interface;
using KGI.ReportComponent.Helper;
using KGI.ReportComponent.Interface;
using KGI.ReportComponent.Model;
using Microsoft.Extensions.Configuration;
using ReportExecution2005;
using SimpleImpersonation;
using System;
using System.Collections.Generic;
using System.Diagnostics.Metrics;
using System.IO;
using System.Linq;
using System.Net;
using System.Threading.Tasks;
using Utility.Helper;

namespace KGI.ReportComponent
{
    public partial class ReportServiceExample : IReportService
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
                    case 0:
                    case 1:
                    case 2:
                        report = await this.GetReport2005(model);
                        break;
                    case 4:
                    case 5:
                    case 6:
                        report = await this.GetReport2010(model);
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
            act.Add(4, async (rep, report) => { await SendToMail(rep, report); });
            act.Add(5, async (rep, report) => { await SendToPath(rep, report); });
            act.Add(6, async (rep, report) => { await SendToMail(rep, report); await SendToPath(rep, report); });

            return act;
        }

        /// <summary>
        /// 取得報表資料
        /// </summary>
        /// <param name="model">報表參數</param>
        /// <returns>byte[]</returns>
        private async Task<byte[]> GetReport2005(ReportModel model)
        {
            string reportUrl = model.ReportServerWsdlUrl ?? _configuration.GetSection("ReportServerWsdlUrl").Get<string>();

            _reportManager = new ReportManager(reportUrl, _userName, _passWrod, _domain);

            var parameters = _reportManager.ConvertParameterValues(model.Parameters).Result;

            byte[] res = await GetReportByteArray2005(parameters, model.Rerpot_Path, model.Render_Format);

            return res;
        }

        /// <summary>
        /// 取得報表資料
        /// </summary>
        /// <param name="model">報表參數</param>
        /// <returns>byte[]</returns>
        private async Task<byte[]> GetReport2010(ReportModel model) 
            => await TestCallReportLoop((model) 
                => GetReportByteArray2010(model), model, 0);

        // 透過委派方式整合呼叫失敗時遞迴重新呼叫行為
        private async Task<byte[]> TestCallReportLoop(Func<ReportModel, Task<byte[]>> func, ReportModel model, int count = 0)
        {
            Task<byte[]> res = default;
            try
            {
                res = func(model);
            }
            catch (TimeoutException) // SSRS 呼叫有可能會失敗，失敗後在呼叫最多三次。
            {
                //  距離上次呼叫間隔 1 分鐘
                await Task.Delay(60000);

                if (count < 3)
                {
                    count++;

                    // Timeout 後確認是否有取得資料
                    byte[] r =  await res;
                    if (r == null || r.Length == 0)
                    {
                        res = this.TestCallReportLoop(func,model,count);
                    }

                    if (r != null)
                    {
                        return await res;
                    }
                }

                throw;
            }

            return await res;
        }

        // 呼叫報表 2005
        private async Task<byte[]> GetReportByteArray2005(ParameterValue[] parameters, string path, string renderFormat, int count = 0)
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
                        res = await this.GetReportByteArray2005(parameters, path, renderFormat, count);
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

        // 呼叫報表 2010
        private Task<byte[]> GetReportByteArray2010(ReportModel model, int count = 0)
        {
            string reportUrl = model.ReportServerWsdlUrl ?? _configuration.GetSection("ReportServerWsdlUrl").Get<string>();

            Dictionary<string, ReportFormats> format = new Dictionary<string, ReportFormats>();
            format.Add("PDF", ReportFormats.Pdf);
            format.Add("WORD", ReportFormats.Docx);
            format.Add("EXCEL", ReportFormats.Xlsx);

            byte[] bytes = default;
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
                bytes = ms.ToArray();
                ms.Close();
            }

            return Task.FromResult(bytes);
        }
    }
}
