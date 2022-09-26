using KGI.ReportComponent.Model;
using ReportExecution2005;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.ServiceModel;
using System.Threading.Tasks;

namespace KGI.ReportComponent.Helper
{
    public class ReportManager
    {
        private readonly ReportExecutionServiceSoapClient _reportServerExecutionService;

        /// <summary>
        /// 報表管理
        /// </summary>
        /// <param name="reportServerWsdlUrl">報表url</param>
        /// <param name="username">帳號</param>
        /// <param name="_password">密碼</param>
        /// <param name="domain">domain</param>
        public ReportManager(string reportServerWsdlUrl, string username = null, string password = null, string domain = null)
        {
            var binding = new BasicHttpBinding(BasicHttpSecurityMode.TransportCredentialOnly);
            binding.Security.Transport.ClientCredentialType = HttpClientCredentialType.Ntlm;

            binding.MaxReceivedMessageSize = 10485760; //I wanted a 10MB size limit on response to allow for larger PDFs

            // 設定服務超時
            binding.OpenTimeout = new System.TimeSpan(0, 05, 0);
            binding.CloseTimeout = new System.TimeSpan(0, 05, 0);
            binding.SendTimeout = new System.TimeSpan(0, 05, 0);
            binding.ReceiveTimeout = new System.TimeSpan(0, 05, 0);

            _reportServerExecutionService = new ReportExecutionServiceSoapClient(binding, new EndpointAddress(reportServerWsdlUrl));
            _reportServerExecutionService.ClientCredentials.Windows.AllowedImpersonationLevel = System.Security.Principal.TokenImpersonationLevel.Impersonation;
            _reportServerExecutionService.ClientCredentials.Windows.ClientCredential = new NetworkCredential(username, password, domain);
        }

        /// <summary>
        /// 轉換參數
        /// </summary>
        /// <param name="parameterModels"></param>
        /// <returns></returns>
        public async Task<ParameterValue[]> ConvertParameterValues(List<ParameterModel> parameterModels)
        {
            ParameterValue[] reportParam = default;
            List<ParameterModel> newparameter = new List<ParameterModel>();
            if (parameterModels == null)
            {
                return await Task.FromResult(reportParam);
            }

            List<ParameterModel> parAry = parameterModels.Where(s => s.Value?.ToString() is string && s.Value?.ToString().Contains(',') == true).ToList();

            foreach (var item in parAry)
            {
                newparameter = item.Value.ToString().Split(',').Select(ar => new ParameterModel { Name = item.Name, Value = ar }).ToList();
            }

            List<ParameterModel> other = parameterModels.Where(x => x.Value?.ToString().Contains(',') != true).ToList();

            newparameter = newparameter.Concat(other).ToList();

            reportParam = newparameter.Any() ? ToParameterValueArray(newparameter) : ToParameterValueArray(parameterModels);

            return await Task.FromResult(reportParam);
        }

        /// <summary>
        /// 轉換參數 ParameterValue[]
        /// </summary>
        /// <param name="parameterModels">model</param>

        private static ParameterValue[] ToParameterValueArray(List<ParameterModel> parameterModels)
        {
            return parameterModels.Select(s => new ParameterValue { Name = s.Name, Value = s.Value.ToString(), Label = s.Value.ToString() }).ToArray();
        }

        /// <summary>
        /// 取得報表
        /// </summary>
        /// <param name="report">報表路徑及名稱</param>
        /// <param name="exportFormat">輸出格式</param>
        /// <param name="parameters">報表參數</param>
        /// <returns></returns>
        public async Task<byte[]> RenderReport(string report, string exportFormat, ParameterValue[] parameters = null)
        {
            RenderResponse response;
            _reportServerExecutionService.Endpoint.EndpointBehaviors.Add(new ReportingServiceEndPointBehavior());

            //Load the report
            TrustedUserHeader trusted = new TrustedUserHeader();
            var taskLoadReport = await _reportServerExecutionService.LoadReportAsync(trusted, report, null);

            await SetParameters(parameters, trusted, taskLoadReport);

            //run the report
            const string deviceInfo = @"<DeviceInfo><Toolbar>False</Toolbar></DeviceInfo>";
            response = await _reportServerExecutionService.RenderAsync(new RenderRequest(taskLoadReport.ExecutionHeader, trusted, exportFormat, deviceInfo));

            //spit out the result
            return response.Result;
        }

        public async Task<bool> CallReport(string report)
        {
            bool result = default;
            _reportServerExecutionService.Endpoint.EndpointBehaviors.Add(new ReportingServiceEndPointBehavior());

            //Load the report
            TrustedUserHeader trusted = new TrustedUserHeader();
            var taskLoadReport = await _reportServerExecutionService.LoadReportAsync(trusted, report, null);

            //var taskLoadReport = await _reportServerExecutionService.ListSecureMethodsAsync(trusted);

            //spit out the result
            result = taskLoadReport != null ? true : false;
            return result;
        }

        /// <summary>
        /// 參數設定
        /// </summary>
        /// <param name="reportParameters"></param>
        /// <param name="trusted"></param>
        /// <param name="taskLoadReport"></param>
        /// <returns></returns>
        private async Task SetParameters(ParameterValue[] reportParameters, TrustedUserHeader trusted, LoadReportResponse taskLoadReport)
        {
            if (reportParameters.Any() && taskLoadReport.executionInfo.Parameters.Any())
            {
                var reportP = reportParameters.Select(s => s.Name);
                var taskR = taskLoadReport.executionInfo.Parameters.Select(t => t.Name);
                var reportIntersect = reportP.Intersect(taskR);

                // 確認參數是否有對上
                if (reportIntersect.Any())
                {
                    //Set the parameteres asked for by the report            
                    await _reportServerExecutionService.SetExecutionParametersAsync(taskLoadReport.ExecutionHeader, trusted, reportParameters, "en-us");
                }
            }
        }
    }



}
