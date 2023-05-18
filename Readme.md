<!-- default badges list -->
![](https://img.shields.io/endpoint?url=https://codecentral.devexpress.com/api/v1/VersionRange/128613865/23.1.2%2B)
[![](https://img.shields.io/badge/Open_in_DevExpress_Support_Center-FF7200?style=flat-square&logo=DevExpress&logoColor=white)](https://supportcenter.devexpress.com/ticket/details/T409684)
[![](https://img.shields.io/badge/📖_How_to_use_DevExpress_Examples-e9f6fc?style=flat-square)](https://docs.devexpress.com/GeneralInformation/403183)
<!-- default badges end -->
<!-- default file list -->
*Files to look at*:

* [Form1.cs](./CS/MailMergeEFData/Form1.cs) (VB: [Form1.vb](./VB/MailMergeEFData/Form1.vb))
<!-- default file list end -->
# How to prevent loading an untrusted custom assembly in Entity Framework data model


This example addresses securirty considerations that are specific to mail merge using data from Entity Framework data models contained in a compiled assembly.<br>Before loading the data assembly you should have an option to perform a path check to ensure that an assembly is obtained from a trusted source and that the path length is within valid limits.<br><br>When this example runs, it asks the user for permission to load the assembly. The first request is initiated by the <a href="http://help.devexpress.com/#CoreLibraries/DevExpressSpreadsheetIWorkbook_GenerateMailMergeDocumentstopic">GenerateMailMergeDocuments</a> method, the second request occurs when the Spreadsheet Control attempts to populate its Field List. Each request fires the <a href="http://help.devexpress.com/#WindowsForms/DevExpressXtraSpreadsheetSpreadsheetControl_CustomAssemblyLoadingtopic">CustomAssemblyLoading</a> event. The event handler asks the end-user for the permission and prompts whether the decision is final. If the decision is not final, the <a href="http://help.devexpress.com/#CoreLibraries/clsDevExpressXtraSpreadsheetServicesICustomAssemblyLoadingNotificationServicetopic">ICustomAssemblyLoadingNotificationService</a> is used to grant or deny the permission to load the assembly.

<br/>


