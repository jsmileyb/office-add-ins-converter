
# **Deploy and publish Office Add-ins**

08/15/2025

You can use one of several methods to deploy your Office Add-in for testing or distribution to users. The deployment method can also affect which platforms surface your add-in.

#### 7 **Note**

For information about how end users acquire, insert, and run add-ins, see **[Start using your](https://support.microsoft.com/office/82e665c4-6700-4b56-a3f3-ef5441996862) [Office Add-in](https://support.microsoft.com/office/82e665c4-6700-4b56-a3f3-ef5441996862)** .

## **Primary publication methods**

The following table summarizes the primary publication methods that can be used regardless of which type of manifest the add-in uses. If the add-in uses the add-in only manifest, see also Additional publication methods for the add-in only manifest.

ノ **Expand table**

| Method                                                      | Use                                                                                                                                       |
|-------------------------------------------------------------|-------------------------------------------------------------------------------------------------------------------------------------------|
| Sideloading                                                 | As part of your development process, to test your add-in running on<br>Windows, iPad, Mac, or in a browser. (Not for production add-ins.) |
| AppSource                                                   | To distribute your add-in publicly to users.                                                                                              |
| Integrated apps portal in the<br>Microsoft 365 admin center | To distribute your add-in to users in your organization.                                                                                  |

## **Production deployment methods**

The following sections provide additional information about the deployment methods that are most commonly used to distribute production Office Add-ins to users.

### **AppSource**

You can make your add-in available through [AppSource](https://appsource.microsoft.com/marketplace/apps?product=office) , Microsoft's online app store which is accessible through a browser and through the UI of Office applications. Distribution through AppSource gives you the option of including installation of your add-in with the installation of your Windows app or a COM or VSTO add-in. For more information, see Publish to your Office Add-in to AppSource.


#### 7 **Note**

If you plan to **publish** your add-in to AppSource and make it available within the Office experience, make sure that you conform to the **Commercial marketplace certification policies**. For example, to pass validation, your add-in must work across all platforms that support the methods that you define (for more information, see **section 1120.3** and the **[Office Add-in application and availability page](https://learn.microsoft.com/en-us/javascript/api/requirement-sets)**).

#### **Integrated apps portal in the Microsoft 365 admin center**

The Microsoft 365 admin center makes it easy for an administrator to deploy Office Add-ins to users and groups in their organization. Add-ins deployed via the admin center are available to users in their Office applications right away, with no client configuration required. You can use integrated apps portal to deploy internal add-ins as well as add-ins provided by independent software vendors (ISVs). The integrated apps portal also shows admins add-ins and other apps bundled together by same ISV, giving them exposure to the entire experience across the Microsoft 365 platform.

When you link your Office Add-ins, Teams apps, SharePoint Framework (SPFx) apps, and [other](https://learn.microsoft.com/en-us/microsoft-365/admin/manage/test-and-deploy-microsoft-365-apps#what-apps-can-i-deploy-from-integrated-apps) [apps](https://learn.microsoft.com/en-us/microsoft-365/admin/manage/test-and-deploy-microsoft-365-apps#what-apps-can-i-deploy-from-integrated-apps) together, you create a single software as a service (SaaS) offering for your customers. For general information about this process, see [How to plan a SaaS offer for the commercial](https://learn.microsoft.com/en-us/azure/marketplace/plan-saas-offer) [marketplace](https://learn.microsoft.com/en-us/azure/marketplace/plan-saas-offer). For specifics on how to create the offer, see [Create the offer.](https://learn.microsoft.com/en-us/azure/marketplace/create-new-saas-offer)

For more information on the deployment process, see [Get started with the integrated apps](https://learn.microsoft.com/en-us/microsoft-365/admin/manage/test-and-deploy-microsoft-365-apps) [portal.](https://learn.microsoft.com/en-us/microsoft-365/admin/manage/test-and-deploy-microsoft-365-apps)

#### 7 **Note**

If your add-in uses the **unified manifest for Microsoft 365** and is distributed as an internal add-in in the integrated apps portal (instead of being acquired by the administrator from AppSource), it won't be installable by users with certain versions of Office. For more information, see **Office Add-ins with the unified app manifest for Microsoft 365 - Client and platform support**.

#### ) **Important**

Customers in sovereign or government clouds don't have access to the integrated apps portal. They use Centralized Deployment instead. (See **Additional publication methods for the add-in only manifest** later in this article.) Centralized Deployment is a similar deployment method, but doesn't expose connected add-ins and apps to the admin. For


more information, see **[Determine if Centralized Deployment of add-ins works for your](https://learn.microsoft.com/en-us/microsoft-365/admin/manage/centralized-deployment-of-add-ins) [organization](https://learn.microsoft.com/en-us/microsoft-365/admin/manage/centralized-deployment-of-add-ins)**.

### **Deploy updates**

When you add features or fix bugs in your add-in, you'll need to deploy the updates. If your add-in is deployed by one or more admins to their organizations, some manifest changes will require the admin to consent to the updates. Users remain on the existing version of the addin until the admin consents to the updates. The following manifest changes will require the admin to consent again.

- Changes to requested permissions. See Requesting permissions for API use in add-ins and Understanding Outlook add-in permissions.
- Additional or changed [Scopes.](https://learn.microsoft.com/en-us/javascript/api/manifest/scopes) (Not applicable if the add-in uses the unified manifest for Microsoft 365.)
- Additional or changed Events.

7 **Note**

Whenever you make a change to the manifest, you must raise the version number of the manifest.

- If the add-in uses the add-in only manifest, see **[Version element](https://learn.microsoft.com/en-us/javascript/api/manifest/version)**.
- If the add-in uses the unified manifest, see **[version property](https://learn.microsoft.com/en-us/microsoft-365/extensibility/schema/root#version)**.

## **Additional publication methods for the add-in only manifest**

The following table summarizes publication methods that are available *only when the add-in uses the add-in only manifest*.

- ノ **Expand table**

| Method           | Use                                                                                                                                                                                                          | Support limitations                                                                                                                      |
|------------------|--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|------------------------------------------------------------------------------------------------------------------------------------------|
| Network<br>share | As part of your development process, to<br>test your add-in running on Windows<br>computers other than your development<br>computer after you have published the add<br>in to a server other than localhost. | Not supported for production add<br>ins.<br>Not supported for Outlook add-ins.<br>Not supported for testing on iPad,<br>Mac, or the web. |


| Method                    | Use                                                                                           | Support limitations                                                                                                                                                                                                       |
|---------------------------|-----------------------------------------------------------------------------------------------|---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| SharePoint<br>catalog     | In an on-premises environment, to<br>distribute your add-in to users in your<br>organization. | Not supported for Outlook add-ins.<br>Not supported for Office on Mac.<br>Not supported for add-ins with any<br>feature that requires a<br><versionoverrides> element in the<br/>add-in only manifest.</versionoverrides> |
| Exchange<br>server        | In an on-premises or online environment, to<br>distribute Outlook add-ins to users.           | Only supported for Outlook add-ins.                                                                                                                                                                                       |
| Centralized<br>Deployment | To distribute your add-in to users in your<br>organization.                                   |                                                                                                                                                                                                                           |

### **SharePoint app catalog deployment**

A SharePoint app catalog is a special SharePoint site collection that you can create to host the manifests (add-in only manifest type) of a Word, Excel, or PowerPoint add-in. If you're deploying add-ins in an on-premises environment and none of the add-in users use a Mac, consider using a SharePoint catalog. For details, see Publish task pane and content add-ins to a SharePoint catalog.

Because SharePoint catalogs don't support new add-in features implemented in the VersionOverrides node of the manifest, including add-in commands, for these add-ins, we recommend that you use Centralized Deployment via the admin center if possible.

### **Outlook add-in Exchange server deployment**

For on-premises and online environments that don't use the [Microsoft Entra](https://learn.microsoft.com/en-us/entra/fundamentals/what-is-entra) identity service, you can deploy Outlook add-ins via the Exchange server.

Outlook add-in deployment requires:

- Microsoft 365, Exchange Online, or Exchange Server 2016 or later
- Outlook 2016 or later

To assign and manage add-ins for your tenants and users, use [Exchange PowerShell](https://learn.microsoft.com/en-us/powershell/module/exchangepowershell). For more information, see [Add-ins for Outlook in Exchange Server](https://learn.microsoft.com/en-us/exchange/add-ins-for-outlook-2013-help) and [Add-ins for Outlook in Exchange](https://learn.microsoft.com/en-us/exchange/clients-and-mobile-in-exchange-online/add-ins-for-outlook/add-ins-for-outlook) [Online](https://learn.microsoft.com/en-us/exchange/clients-and-mobile-in-exchange-online/add-ins-for-outlook/add-ins-for-outlook).

It's important to note that some versions of Outlook clients and Exchange servers may only support certain Mailbox requirement sets. For details about supported requirement sets, see [Requirement sets supported by Exchange servers and Outlook clients](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients).


## **GoDaddy Microsoft 365 SKUs**

[Microsoft 365 subscriptions provided by GoDaddy](https://www.godaddy.com/business/office-365) have limited support for add-ins. The following options are **not** supported.

- Deployment through the Microsoft 365 admin center.
- Deployment through Exchange servers.
- Acquiring add-ins from AppSource.

## **See also**

- Sideload Outlook add-ins for testing
- Publish to your Office Add-in to AppSource
- [AppSource](https://appsource.microsoft.com/marketplace/apps?product=office)
- Design guidelines for Office Add-ins
- [Create effective AppSource listings](https://learn.microsoft.com/en-us/partner-center/marketplace-offers/create-effective-office-store-listings)
- Troubleshoot user errors with Office Add-ins
- [What is the Microsoft commercial marketplace?](https://learn.microsoft.com/en-us/azure/marketplace/overview)
- [Microsoft Dev Center app publishing page](https://developer.microsoft.com/microsoft-teams/app-publishing)

# **Publish an add-in developed with Visual Studio Code**

Article • 05/19/2025

This article describes how to publish an Office Add-in that you created using the Yeoman generator and developed with [Visual Studio Code (VS Code)](https://code.visualstudio.com/) or any other editor.

#### 7 **Note**

- For information about publishing an Office Add-in that you created using Visual Studio, see **Publish your add-in using Visual Studio**.
- The process described in this article doesn't apply to add-ins that use the **unified manifest for Microsoft 365**. Add-ins created using Microsoft 365 Agents Toolkit use the unified manifest. For information about publishing an add-in that you created using Agents Toolkit, see **[Deploy Teams app to the cloud](https://learn.microsoft.com/en-us/microsoftteams/platform/toolkit/deploy?pivots=visual-studio-code)** and **[Deploy your first](https://learn.microsoft.com/en-us/microsoftteams/platform/sbs-gs-javascript?tabs=vscode%2Cvsc%2Cviscode) [Teams app](https://learn.microsoft.com/en-us/microsoftteams/platform/sbs-gs-javascript?tabs=vscode%2Cvsc%2Cviscode)**. The latter article is about Teams tab apps, but it is applicable to Office Add-ins created with Agents Toolkit.

### **Publishing an add-in for other users to access**

The simplest Office Add-in is made up of a manifest file and an HTML page. The manifest file describes the add-in's characteristics, such as its name, what Office applications it can run in, and the URL for the add-in's HTML page. The HTML page is contained in a web app that users interact with when they install and run your add-in within an Office application. You can host the web app of an Office Add-in on any web hosting platform, including Azure.

While you're developing, you can run the add-in on your local web server ( localhost ). When you're ready to publish it for other users to access, you'll need to deploy the web application and update the manifest to specify the URL of the deployed application.

When your add-in is working as desired, you can publish it directly through Visual Studio Code using the Azure Storage extension.

### **Using Visual Studio Code to publish**


These steps only work for projects created with the Yeoman generator, and that use the add-in only manifest. They don't apply if you created the add-in using Agents Toolkit or created it with the Yeoman generator and it uses the unified manifest for Microsoft 365.

- 1. Open your project from its root folder in Visual Studio Code (VS Code).
- 2. Select **View** > **Extensions** ( Ctrl + Shift + X ) to open the Extensions view.
- 3. Search for the **Azure Storage** extension and install it.
- 4. Once installed, an Azure icon is added to the **Activity Bar**. Select it to access the extension. If the **Activity Bar** is hidden, open it by selecting **View** > **Appearance** > **Activity Bar**.
- 5. Select **Sign in to Azure** to sign in to your Azure account. If you don't already have an Azure account, create one by selecting **Create an Azure Account**. Follow the provided steps to set up your account.

- 6. Once you're signed in, you'll see your Azure storage accounts appear in the extension. If you don't already have a storage account, create one using the **Create Storage Account** option in the command palette. Name your storage account a globally unique name, using only 'a-z' and '0-9'. Note that by default, this creates a storage account and a resource group with the same name. It automatically puts the storage account in West US. This can be adjusted online through [your Azure account](https://portal.azure.com/) .


- 7. Right-click (or select and hold) your storage account and select **Configure Static Website**. You'll be asked to enter the index document name and the 404 document name. Change the index document name from the default index.html to **taskpane.html** . You may also change the 404 document name but aren't required to.
- 8. Right-click (or select and hold) your storage account again and this time select **Browse Static Website**. From the browser window that opens, copy the website URL.
- 9. Open your project's manifest file and change all references to your localhost URL (such as https://localhost:3000 ) to the URL you've copied. This endpoint is the static website URL for your newly created storage account. Save the changes to your manifest file.
- 10. Open a command line prompt or terminal window and go to the root directory of your add-in project. Run the following command to prepare all files for production deployment.

| command line  |  |  |  |
|---------------|--|--|--|
| npm run build |  |  |  |

When the build completes, the **dist** folder in the root directory of your add-in project will contain the files that you'll deploy in subsequent steps.

- 11. In VS Code, go to the Explorer and right-click (or select and hold) the **dist** folder, and select **Deploy to Static Website via Azure Storage**. When prompted, select the storage account you created previously.


|  | EXPLORER                 |         | Validate this folder                       |                     |  |
|--|--------------------------|---------|--------------------------------------------|---------------------|--|
|  | OPEN EDITORS             |         | New File                                   |                     |  |
|  | > TEST-DEPLOY            | 9 0 0 0 | New Folder                                 |                     |  |
|  | > .vscode                |         | Reveal in File Explorer                    | Shift +Alt + R      |  |
|  | assets<br>A              |         |                                            |                     |  |
|  | > dist                   |         | Open in Integrated Terminal                |                     |  |
|  | node_modules<br>>        |         | Cleanup files in folder                    |                     |  |
|  | V SIC                    |         | Collapse relative links in folder          |                     |  |
|  | commands<br>>            |         |                                            |                     |  |
|  | v taskpane               |         | Compress all images in folder              |                     |  |
|  | taskpane.css<br>#        |         | Learn: Create module in CURRENT folder     |                     |  |
|  | taskpane.html<br><>      |         | Learn: Create module in NEW folder         |                     |  |
|  | JS taskpane.js           |         | Learn: Update module folder name           |                     |  |
|  | @ .eslintrc.json         |         |                                            |                     |  |
|  | babel.config.json<br>B   |         | Find in Folder                             | Shift+Alt+F         |  |
|  | manifest.xml             |         |                                            |                     |  |
|  | package-lock.json<br>1 } |         | Cut                                        | Ctrl+X              |  |
|  | package.json<br>11       |         | Сору                                       | Ctrl+C              |  |
|  | tsconfig.json<br>TS      |         | Paste                                      | Ctrl+V              |  |
|  | @ webpack.config.js      |         |                                            |                     |  |
|  |                          |         | Copy Path                                  | Shift+Alt+C         |  |
|  |                          |         | Copy Relative Path                         | Ctrl+K Ctrl+Shift+C |  |
|  |                          |         | Rename                                     | F2                  |  |
|  |                          |         | Delete                                     | Delete              |  |
|  |                          |         | Deploy to Static Website via Azure Storage |                     |  |
|  |                          |         | Upload to Azure Storage                    |                     |  |

- 12. When deployment is complete, right-click (or select and hold) the storage account that you created previously and select **Browse Static Website**. This opens the static web site and displays the task pane.
- 13. Finally, sideload the manifest file and the add-in will load from the static web site you just deployed.

### **Deploy custom functions for Excel**

If your add-in has custom functions, there are a few more steps to enable them on the Azure Storage account. First, enable CORS so that Office can access the functions.json file.

- 1. Right-click (or select and hold) the Azure storage account and select **Open in Portal**.
- 2. In the Settings group, select **Resource sharing (CORS)**. You can also use the search box to find this.


- 3. Create a new CORS rule for the **Blob service** with the following settings.
#### ノ **Expand table**

| Property        | Value                       |  |
|-----------------|-----------------------------|--|
| Allowed origins | *                           |  |
| Allowed methods | GET                         |  |
| Allowed headers | *                           |  |
| Exposed headers | Access-Control-Allow-Origin |  |
| Max age         | 200                         |  |

#### 4. Select **Save**.

#### U **Caution**

This CORS configuration assumes all files on your server are publicly available to all domains.

Next, add a MIME type for JSON files.

- 1. Create a new file in the /src folder named **web.config**.
- 2. Insert the following XML and save the file.

```
XML
<?xml version="1.0"?>
<configuration>
 <system.webServer>
 <staticContent>
 <mimeMap fileExtension=".json" mimeType="application/json" />
 </staticContent>
 </system.webServer>
</configuration>
```
- 3. Open the **webpack.config.js** file.
- 4. Add the following code in the list of plugins to copy the web.config into the bundle when the build runs.

JavaScript


```
new CopyWebpackPlugin({
 patterns: [
 {
 from: "src/web.config",
 to: "src/web.config",
 },
],
}),
```
- 5. Open a command line prompt and go to the root directory of your add-in project. Then, run the following command to prepare all files for deployment.
command line npm run build

When the build completes, the **dist** folder in the root directory of your add-in project will contain the files that you'll deploy.

- 6. To deploy, in the VS Code **Explorer**, right-click (or select and hold) the **dist** folder and select **Deploy to Static Website via Azure Storage**. When prompted, select the storage account you created previously. If you already deployed the **dist** folder, you'll be prompted if you want to overwrite the files in the Azure storage with the latest changes.
### **Deploy updates**

You'll deploy updates to your web application in the same manner as described previously. Changes to the manifest require redistributing your manifest to users. The process to do so depends on your publishing method. For more information on updating your add-in, see Maintain your Office Add-in.

## **See also**

- Develop Office Add-ins with Visual Studio Code
- Deploy and publish your Office Add-in
- [Cross-Origin Resource Sharing (CORS) support for Azure Storage](https://learn.microsoft.com/en-us/rest/api/storageservices/cross-origin-resource-sharing--cors--support-for-the-azure-storage-services)

# **Publish your add-in using Visual Studio**

Article • 08/13/2024

Your Office Add-in package contains an XML manifest file that you'll use to publish the add-in. You'll have to publish the web application files of your project separately. This article describes how to deploy your web project and package your add-in by using Visual Studio 2019.

#### 7 **Note**

For information about publishing an Office Add-in that you created using the Yeoman generator and developed with Visual Studio Code or any other editor, see **Publish an add-in developed with Visual Studio Code**.

### **To deploy your web project using Visual Studio 2019**

Complete the following steps to deploy your web project using Visual Studio 2019.

- 1. From the **Build** tab, choose **Publish [Name of your add-in]**.
- 2. In the **Pick a publish target** window, choose one of the options to publish to your preferred target. Each publish target requires you to include more information to get started, such as an Azure Virtual Machine or folder location. Once you have specified a publish location and filled in all of the information required, select **Publish**

#### 7 **Note**

Picking a publish target specifies the server you are deploying to, the credentials needed to sign in to the server, the databases to deploy, and other deployment options.

- 3. For more information about deployment steps for each publish target option, see [First look at deployment in Visual Studio](https://learn.microsoft.com/en-us/visualstudio/deployment/deploying-applications-services-and-components?view=vs-2019&preserve-view=true).
### **To package and publish your add-in using IIS, FTP, or Web Deploy using Visual Studio 2019**


Complete the following steps to package your add-in using Visual Studio 2019.

- 1. From the **Build** tab, choose **Publish [Name of your add-in]**.
- 2. In the **Pick a publish target** window, choose **IIS, FTP, etc**, and select **Configure**. Next, select **Publish**.
- 3. A wizard appears that will help guide you through the process. Ensure the publish method is your preferred method, such as Web Deploy.
- 4. In the **Destination URL** box, enter the URL of the website that will host the content files of your add-in, and then select **Next**. If you plan to submit your add-in to AppSource, you can choose the **Validate Connection** button to identify any issues that will prevent your add-in from being accepted. You should address all issues before you submit your add-in to the store.
- 5. Confirm any settings desired including **File Publish Options** and select **Save**.

#### ) **Important**

While not strictly required in all add-in scenarios, using an HTTPS endpoint for your add-in is strongly recommended. Add-ins that are not SSL-secured (HTTPS) generate unsecure content warnings and errors during use. If you plan to run your add-in in Office on the web or publish your add-in to AppSource, it must be SSL-secured. If your add-in accesses external data and services, it should be SSL-secured to protect data in transit. Self-signed certificates can be used for development and testing, so long as the certificate is trusted on the local machine. Azure websites automatically provide an HTTPS endpoint.

You can now upload your manifest to the appropriate location to publish your add-in. You can find the manifest in OfficeAppManifests in the app.publish folder. For example:

```
%UserProfile%\Documents\Visual Studio
2019\Projects\MyApp\bin\Debug\app.publish\OfficeAppManifests
```
## **Deploy updates**

You'll deploy updates to your web application in the same manner as described previously. Changes to the manifest require redistributing your manifest to users. The process to do so depends on your publishing method. For more information on updating your add-in, see Maintain your Office Add-in.


### **See also**

- Publish your Office Add-in
- [Make your solutions available in AppSource and within Office](https://learn.microsoft.com/en-us/partner-center/marketplace-offers/submit-to-appsource-via-partner-center)

#### 6 **Collaborate with us on GitHub**

The source for this content can be found on GitHub, where you can also create and review issues and pull requests. For more information, see [our](https://github.com/OfficeDev/office-js-docs-pr/blob/main/Contributing.md) [contributor guide](https://github.com/OfficeDev/office-js-docs-pr/blob/main/Contributing.md).

#### **Office Add-ins feedback**

Office Add-ins is an open source project. Select a link to provide feedback:

- [Open a documentation issue](https://github.com/OfficeDev/office-js-docs-pr/issues/new?template=3-customer-feedback.yml&pageUrl=https%3A%2F%2Flearn.microsoft.com%2Fen-us%2Foffice%2Fdev%2Fadd-ins%2Fpublish%2Fpackage-your-add-in-using-visual-studio&pageQueryParams=&contentSourceUrl=https%3A%2F%2Fgithub.com%2FOfficeDev%2Foffice-js-docs-pr%2Fblob%2Fmain%2Fdocs%2Fpublish%2Fpackage-your-add-in-using-visual-studio.md&documentVersionIndependentId=1466e4a5-cdeb-7a2b-5b70-182efcd5bb80&feedback=%0A%0A%5BEnter+feedback+here%5D%0A&author=%40o365devx&metadata=*+ID%3A+a5695d91-ece0-a4a5-c713-bec3c4aa1364+%0A*+Service%3A+**microsoft-365**%0A*+Sub-service%3A+**add-ins**)
- [Provide product feedback](https://aka.ms/office-addins-dev-questions)