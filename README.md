# ASP.NETCORE_ReportingServices
Asp.Net Core_SQL Server Reporting Services (SSRS) 實作練習

文章參考：https://stackoverflow.com/questions/44036903/using-reporting-services-ssrs-as-a-reference-in-an-asp-net-core-site

文章參考：https://www.codeproject.com/Articles/1110411/How-to-Write-a-Csharp-Wrapper-Class-for-the-SSRS-R

文章參考：https://csharp.hotexamples.com/examples/-/ReportExecutionService/-/php-reportexecutionservice-class-examples.html

文章參考：https://csharp.hotexamples.com/examples/-/ReportExecution2005/-/php-reportexecution2005-class-examples.html

使用 asp.net core 建立wfc服務呼叫 ssrs 服務，可以依json檔中設定參數。

一、使用ssrs服務連結建立wcf 服務

二、建立 ReportExecutionServiceSoapClient 實體，使用時需要修改底層 GetEndpointAddress  路徑，或從新建立wcf服務

三、參考文章後將 ReportExecutionServiceSoapClient  包裝成 ReportManager，提供兩種使用辦法，回傳byte[] 及檔案下載。