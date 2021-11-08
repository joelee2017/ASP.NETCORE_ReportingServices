//using ReportExecution2005;
//using System.Collections.Generic;
//using System.Linq;
//using System.Net;
//using System.ServiceModel;
//using System.Threading.Tasks;

//namespace AspNetCoreSSRS
//{
//    public class ReportMangerExample
//    {
//        //public async Task<IActionResult> GetPDFReport()
//        //{
//        //    string reportName = "YourReport";
//        //    IDictionary<string, object> parameters = new Dictionary<string, object>();
//        //    parameters.Add("companyId", "2");
//        //    parameters.Add("customerId", "123");
//        //    string languageCode = "en-us";

//        //    byte[] reportContent = await this.RenderReport(reportName, parameters, languageCode, "PDF");

//        //    Stream stream = new MemoryStream(reportContent);

//        //    return new FileStreamResult(stream, "application/pdf");

//        //}

//        /// <summary>
//        /// </summary>
//        /// <param name="reportName">
//        ///  report name.
//        /// </param>
//        /// <param name="parameters">report's required parameters</param>
//        /// <param name="exportFormat">value = "PDF" or "EXCEL". By default it is pdf.</param>
//        /// <param name="languageCode">
//        ///   value = 'en-us', 'fr-ca', 'es-us', 'zh-chs'. 
//        /// </param>
//        /// <returns></returns>
//        private async Task<byte[]> RenderReport(string reportName, IDictionary<string, object> parameters, string languageCode, string exportFormat)
//        {
//            //
//            // SSRS report path. Note: Need to include parent folder directory and report name.
//            // Such as value = "/[report folder]/[report name]".
//            //
//            string reportPath = string.Format("{0}{1}", ConfigSettings.ReportingServiceReportFolder, reportName);

//            //
//            // Binding setup, since ASP.NET Core apps don't use a web.config file
//            //
//            var binding = new BasicHttpBinding(BasicHttpSecurityMode.TransportCredentialOnly);
//            binding.Security.Transport.ClientCredentialType = HttpClientCredentialType.Ntlm;
//            binding.MaxReceivedMessageSize = this.ConfigSettings.ReportingServiceReportMaxSize; //It is 10MB size limit on response to allow for larger PDFs

//            //Create the execution service SOAP Client
//            ReportExecutionServiceSoapClient reportClient = new ReportExecutionServiceSoapClient(binding, new EndpointAddress(this.ConfigSettings.ReportingServiceUrl));

//            //Setup access credentials. Here use windows credentials.
//            var clientCredentials = new NetworkCredential(this.ConfigSettings.ReportingServiceUserAccount, this.ConfigSettings.ReportingServiceUserAccountPassword, this.ConfigSettings.ReportingServiceUserAccountDomain);
//            reportClient.ClientCredentials.Windows.AllowedImpersonationLevel = System.Security.Principal.TokenImpersonationLevel.Impersonation;
//            reportClient.ClientCredentials.Windows.ClientCredential = clientCredentials;

//            //This handles the problem of "Missing session identifier"
//            reportClient.Endpoint.EndpointBehaviors.Add(new ReportingServiceEndPointBehavior());

//            string historyID = null;
//            TrustedUserHeader trustedUserHeader = new TrustedUserHeader();
//            ExecutionHeader execHeader = new ExecutionHeader();

//            trustedUserHeader.UserName = clientCredentials.UserName;

//            //
//            // Load the report
//            //
//            var taskLoadReport = await reportClient.LoadReportAsync(trustedUserHeader, reportPath, historyID);
//            // Fixed the exception of "session identifier is missing".
//            execHeader.ExecutionID = taskLoadReport.executionInfo.ExecutionID;

//            //
//            //Set the parameteres asked for by the report
//            //
//            ParameterValue[] reportParameters = null;
//            if (parameters != null && parameters.Count > 0)
//            {
//                reportParameters = taskLoadReport.executionInfo.Parameters.Where(x => parameters.ContainsKey(x.Name)).Select(x => new ParameterValue() { Name = x.Name, Value = parameters[x.Name].ToString() }).ToArray();
//            }

//            await reportClient.SetExecutionParametersAsync(execHeader, trustedUserHeader, reportParameters, languageCode);
//            // run the report
//            const string deviceInfo = @"<DeviceInfo><Toolbar>False</Toolbar></DeviceInfo>";

//            var response = await reportClient.RenderAsync(new RenderRequest(execHeader, trustedUserHeader, exportFormat ?? "PDF", deviceInfo));

//            //spit out the result
//            return response.Result;
//        }
//    }
//}
