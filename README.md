# Project
To Start - Clone Project locally. This project has been verified to work in Visual Studio 2019. Once project is open, restore NuGet packages and you should be able to build the solution. Common is the main project to build, but all the distributing DLLs are part of the Root solution (Root\bin\Release). If you need to distribute, you will need the following DLLs from Root\bin\Release:
* Discovery.dll
* Microsoft.Services.WorkfIowAssessment.Common.dll
* Microsoft.Services.WorkfIowAssessment.Root.dll
* Microsoft.SharePoint.CIient.dll
* Microsoft.SharePoint.CIient.Runtime.dll
* Microsoft. SharePoint.CIient.WorkfIowServices.dll
* OfficeDevPnP.Core.dll

![image](https://user-images.githubusercontent.com/63272213/136854578-da4def7f-e22a-4541-ae74-f1d3dc328494.png) 

To use code import module 
`import-module .\Microsoft.Services.WorkflowAssessment.Root.dll -verbose`

Run to get workflows using the sites.csv using CSOM code. Permissions needed is at least site collection admin for each of the site in the sites.csv, and can be done from any PC that has connection to SharePoint web server  
`Get-WorkflowAssociationsForOnprem -SiteCollectionURLFilePath .\sites.csv  -DomainName contoso -AssessmentOutputFolder .\Output`

Output
Folders Review
This utility produces 4 folders, but only Logs and Summary have contents from this utility.

* Logs – contains the process logs for the utility and any errors.
* Summary – contains the WorkflowDiscovery.csv that has workflow information for all the site collections that were supplied in the sites.csv file.

![Folder_Structure](https://user-images.githubusercontent.com/63272213/137014648-a9ce8eb4-6e00-4bdd-aa39-2dde31a412a0.png)

## Contributing

This project welcomes contributions and suggestions.  Most contributions require you to agree to a
Contributor License Agreement (CLA) declaring that you have the right to, and actually do, grant us
the rights to use your contribution. For details, visit https://cla.opensource.microsoft.com.

When you submit a pull request, a CLA bot will automatically determine whether you need to provide
a CLA and decorate the PR appropriately (e.g., status check, comment). Simply follow the instructions
provided by the bot. You will only need to do this once across all repos using our CLA.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/).
For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or
contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

## Trademarks

This project may contain trademarks or logos for projects, products, or services. Authorized use of Microsoft 
trademarks or logos is subject to and must follow 
[Microsoft's Trademark & Brand Guidelines](https://www.microsoft.com/en-us/legal/intellectualproperty/trademarks/usage/general).
Use of Microsoft trademarks or logos in modified versions of this project must not cause confusion or imply Microsoft sponsorship.
Any use of third-party trademarks or logos are subject to those third-party's policies.

## Origins

This base code was developed by the Modern Work Team from the Industry Solutions Deliver group. https://www.microsoft.com/en-us/msservices/modern-work
