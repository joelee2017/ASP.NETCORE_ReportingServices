SSRS - 筆記

[TOC]



------



##### ODBC使用

ReportingServices - 需使用 32/64位 driver 建議兩種都裝

MariaDB 就須必使用 mariadb-connector-c-3.2.4-win32

------



##### Visual Studio 2019 - 直接使用商業智慧方案

Visual Studio 2019 - 直接使用商業智慧方案( SSIS or SSRS 的專案)

參考文章： https://dotblogs.com.tw/jamesfu/2019/04/15/visualstudio_2019_bi

使用「延伸模組」→「管理擴充功能」來進行安裝，輸入 Integration 當成關鍵字，就可以很容易地找到 「SQL Server Integration Service Projects[預覽]」的擴充功能

<img src="https://dotblogsfile.blob.core.windows.net/user/jamesfu/ade1dc4d-cd61-4c32-acb2-4a84973996ca/1555320225_58907.png" alt="img" style="zoom: 50%;" />

<img src="https://dotblogsfile.blob.core.windows.net/user/jamesfu/ade1dc4d-cd61-4c32-acb2-4a84973996ca/1555320438_61856.png" alt="img" style="zoom:50%;" />

<img src="https://dotblogsfile.blob.core.windows.net/user/jamesfu/ade1dc4d-cd61-4c32-acb2-4a84973996ca/1555320467_7779.png" alt="img" style="zoom:50%;" />

而安裝完 SSIS 的支援，接著我們再重新進入 Visual Studio，依然再去使用「延伸模組」→「管理擴充功能」來進行安裝，此時我們用 Report 當關鍵字去搜尋，就可以很容易的找到「Microsoft Reporting Services Projects」的擴充功能

<img src="https://dotblogsfile.blob.core.windows.net/user/jamesfu/ade1dc4d-cd61-4c32-acb2-4a84973996ca/1555320637_43683.png" alt="img" style="zoom:50%;" />

<img src="https://dotblogsfile.blob.core.windows.net/user/jamesfu/ade1dc4d-cd61-4c32-acb2-4a84973996ca/1555320754_13045.png" alt="img" style="zoom:50%;" />

當順利安裝完畢之後，此時我們再重新進入 Visual Studio 2019，目前這兩個擴充功能，似乎沒有像其他專案有設定好一些分類標籤，因此只能再新增專案的時候，將選項拉到最後面，此時我們就可以看到 Visual Studio 可以支援 SSIS 和 SSRS 了

<img src="https://dotblogsfile.blob.core.windows.net/user/jamesfu/ade1dc4d-cd61-4c32-acb2-4a84973996ca/1555321030_88204.png" alt="img" style="zoom:50%;" />

接下來就可以順利地使用 Visual Studio 來開發商業智慧方案了。

<img src="https://dotblogsfile.blob.core.windows.net/user/jamesfu/ade1dc4d-cd61-4c32-acb2-4a84973996ca/1555321117_99903.png" alt="img" style="zoom:50%;" />

<img src="https://dotblogsfile.blob.core.windows.net/user/jamesfu/ade1dc4d-cd61-4c32-acb2-4a84973996ca/1555321210_28899.png" alt="img" style="zoom:50%;" />

------

##### 使用C#呼叫SSRS服務

asp.net core 不支援 ssrs 叫用，需用舊版來加入服務 > 加入web 參考加入服務網址

![image-20211105090320579](C:\Users\0011185\AppData\Roaming\Typora\typora-user-images\image-20211105090320579.png)

於url中輸入：http://{youservername}/ReportServer/ReportService2010.asmx?wsdl

輸入完成後點選移至，即可找到服務，給予web參考名稱後加入參考

<img src="C:\Users\0011185\AppData\Roaming\Typora\typora-user-images\image-20211105091402568.png" alt="image-20211105091402568" style="zoom:80%;" />

完成後即會在專案中生成 Web References 資料夾，於資料下就會有剛剛建立的服務

實作呼叫端參考文件：https://www.codeproject.com/Articles/1110411/How-to-Write-a-Csharp-Wrapper-Class-for-the-SSRS-R



服務url 配置兩種

ReportService2012 http://{youservername}/ReportServer/ReportService2010.asmx?wsdl
ReportExecutionService2005：http://{youservername}/ReportServer/ReportExecution2005.asmx?wsdl