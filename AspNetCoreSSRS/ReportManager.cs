using ReportExecution2005;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Threading.Tasks;

namespace AspNetCoreSSRS
{
    public class ReportManager
    {
        private readonly ReportExecutionServiceSoapClient _reportServerExecutionService;

        /// <summary>
        /// 報表管理
        /// </summary>
        /// <param name="reportServerWsdlUrl">報表url</param>
        /// <param name="username">帳號</param>
        /// <param name="password">密碼</param>
        /// <param name="domain">domain</param>
        public ReportManager(string reportServerWsdlUrl, string username = null, string password = null)
        {
            //My binding setup, since ASP.NET Core apps don't use a web.config file
            var binding = new BasicHttpBinding(BasicHttpSecurityMode.TransportCredentialOnly);
            binding.Security.Transport.ClientCredentialType = HttpClientCredentialType.Ntlm;

            binding.MaxReceivedMessageSize = 10485760; //I wanted a 10MB size limit on response to allow for larger PDFs

            _reportServerExecutionService = new ReportExecutionServiceSoapClient(binding, new EndpointAddress(reportServerWsdlUrl));
            _reportServerExecutionService.ClientCredentials.Windows.AllowedImpersonationLevel = System.Security.Principal.TokenImpersonationLevel.Impersonation;

            _reportServerExecutionService.ClientCredentials.ServiceCertificate.DefaultCertificate = new System.Security.Cryptography.X509Certificates.X509Certificate2();

            _reportServerExecutionService.ClientCredentials.UserName.UserName = username;
            _reportServerExecutionService.ClientCredentials.UserName.Password = password;

        }

        public async Task<byte[]> RenderReport(string report, IDictionary<string, object> parameters, string exportFormat = null)
        {
            _reportServerExecutionService.Endpoint.EndpointBehaviors.Add(new ReportingServiceEndPointBehavior());

            //Load the report
            TrustedUserHeader trusted = new TrustedUserHeader();
            //ExecutionHeader execution = new ExecutionHeader();
            var taskLoadReport = await _reportServerExecutionService.LoadReportAsync(trusted, report, null);

            await SetParameters(parameters, trusted, taskLoadReport);

            //run the report
            const string deviceInfo = @"<DeviceInfo><Toolbar>False</Toolbar></DeviceInfo>";
            var response = await _reportServerExecutionService.RenderAsync(new RenderRequest(taskLoadReport.ExecutionHeader, trusted, exportFormat ?? "PDF", deviceInfo));

            //spit out the result
            return response.Result;
        }

        /// <summary>
        /// 參數設定
        /// </summary>
        /// <param name="parameters"></param>
        /// <param name="trusted"></param>
        /// <param name="taskLoadReport"></param>
        /// <returns></returns>
        private async Task SetParameters(IDictionary<string, object> parameters, TrustedUserHeader trusted, LoadReportResponse taskLoadReport)
        {
            if(parameters.Any() && taskLoadReport.executionInfo.Parameters.Any())
            {
                //Set the parameteres asked for by the report
                var reportParameters = taskLoadReport.executionInfo.Parameters.Where(x => parameters.ContainsKey(x.Name)).Select(x => new ParameterValue() { Name = x.Name, Value = parameters[x.Name].ToString() }).ToArray();
                await _reportServerExecutionService.SetExecutionParametersAsync(taskLoadReport.ExecutionHeader, trusted, reportParameters, "en-us");
            }          
        }
    }
}
