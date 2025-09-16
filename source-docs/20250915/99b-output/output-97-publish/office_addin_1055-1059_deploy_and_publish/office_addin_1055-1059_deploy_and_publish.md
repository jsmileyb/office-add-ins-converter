{0}------------------------------------------------

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

{1}------------------------------------------------

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

{2}------------------------------------------------

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

{3}------------------------------------------------

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

{4}------------------------------------------------

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