{0}------------------------------------------------

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

{1}------------------------------------------------

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

{2}------------------------------------------------

### **See also**

- Publish your Office Add-in
- [Make your solutions available in AppSource and within Office](https://learn.microsoft.com/en-us/partner-center/marketplace-offers/submit-to-appsource-via-partner-center)

#### 6 **Collaborate with us on GitHub**

The source for this content can be found on GitHub, where you can also create and review issues and pull requests. For more information, see [our](https://github.com/OfficeDev/office-js-docs-pr/blob/main/Contributing.md) [contributor guide](https://github.com/OfficeDev/office-js-docs-pr/blob/main/Contributing.md).

#### **Office Add-ins feedback**

Office Add-ins is an open source project. Select a link to provide feedback:

- [Open a documentation issue](https://github.com/OfficeDev/office-js-docs-pr/issues/new?template=3-customer-feedback.yml&pageUrl=https%3A%2F%2Flearn.microsoft.com%2Fen-us%2Foffice%2Fdev%2Fadd-ins%2Fpublish%2Fpackage-your-add-in-using-visual-studio&pageQueryParams=&contentSourceUrl=https%3A%2F%2Fgithub.com%2FOfficeDev%2Foffice-js-docs-pr%2Fblob%2Fmain%2Fdocs%2Fpublish%2Fpackage-your-add-in-using-visual-studio.md&documentVersionIndependentId=1466e4a5-cdeb-7a2b-5b70-182efcd5bb80&feedback=%0A%0A%5BEnter+feedback+here%5D%0A&author=%40o365devx&metadata=*+ID%3A+a5695d91-ece0-a4a5-c713-bec3c4aa1364+%0A*+Service%3A+**microsoft-365**%0A*+Sub-service%3A+**add-ins**)
- [Provide product feedback](https://aka.ms/office-addins-dev-questions)