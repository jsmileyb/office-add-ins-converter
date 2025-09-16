
# **Test Office Add-ins**

Article • 07/12/2024

This article contains guidance about testing, debugging, and troubleshooting issues with Office Add-ins.

## **Test cross-platform and for multiple versions of Office**

Office Add-ins run across major platforms, so you need to test an add-in in all the platforms where your users might be running Office. This usually includes Office on the web, Office on Windows (both perpetual and Microsoft 365 subscription), Office on Mac, Office on iOS, and (for Outlook add-ins) Office on Android. However, there may be some situations in which you can be sure that none of your users will be working on some platforms. For example, if you're making an add-in for a company that requires its users to work with Windows computers and subscription Office, then you don't need to test for Office on Mac or perpetual Office on Windows.

#### 7 **Note**

On Windows computers, the version of Windows and Office will determine which browser or webview control is used by add-ins. For more information, see **Browsers and webview controls used by Office Add-ins**. For brevity hereafter, this article uses "browser control" to mean "browser or webview control".

#### **Add-ins tested for Office on the web**

Add-ins are tested for Office on the web with all major modern browsers, including Microsoft Edge (Chromium-based WebView2), Chrome, and Safari. Accordingly, you should test on these platforms and browsers before you submit to [AppSource](https://learn.microsoft.com/en-us/partner-center/marketplace-offers/submit-to-appsource-via-partner-center). For more information about validation, see Commercial marketplace certification policies, especially section 1120.3, and the [Office Add-in application and availability page.](https://learn.microsoft.com/en-us/javascript/api/requirement-sets)

Office on the web no longer opens in Internet Explorer or Microsoft Edge Legacy (EdgeHTML). Consequently, AppSource doesn't test Office on the web on these browsers. Office still supports these browsers for add-in runtimes, so if you think you've encountered a bug in how add-ins run in them, please create an issue in the [office-js](https://github.com/OfficeDev/office-js/issues)


repository. For more information, see Support older Microsoft webviews and Office versions and Troubleshoot EdgeHTML and WebView2 (Microsoft Edge) issues.

#### **Add-ins tested for Office on Windows**

Some Office versions on Windows still use the webview controls that come with Internet Explorer and Microsoft Edge Legacy. AppSource tests whether your add-in supports these browser controls. If your add-in doesn't support these browser controls, AppSource only issues a warning and doesn't reject your add-in. In this instance, we recommend configuring a graceful failure message on your add-in for a smoother user experience. For further guidance, see Support older Microsoft webviews and Office versions.

## **Sideload an Office Add-in for testing**

You can use sideloading to install an Office Add-in for testing without having to first put it in an add-in catalog. The procedure for sideloading an add-in varies by platform, and in some cases, by product as well. The following articles each describe how to sideload Office Add-ins on a specific platform or within a specific product.

#### 7 **Note**

Office Add-ins that use the unified manifest for Microsoft 365 are *directly* supported in Office on the web, in **[new Outlook on Windows](https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627)** , and in Office on Windows connected to a Microsoft 365 subscription, Version 2304 (Build 16320.00000) or later. When the app package that contains the unified manifest is sideloaded to a platform that doesn't directly support that type of manifest then an add-in only manifest is generated from the unified manifest and this manifest is the one that's sideloaded.

- Sideload Office Add-ins in Office on the web
- Sideload Office Add-ins on Windows
- Sideload Office Add-ins on Mac
- Sideload Office Add-ins on iPad
- Sideload Outlook add-ins for testing

## **Unit testing**

For information about how to add unit tests to your add-in project, see Unit testing in Office Add-ins.


## **Debug an Office Add-in**

The procedure for debugging an Office Add-in varies based on your platform and environment. For more information, see Debug Office Add-ins.

### **Validate an Office Add-in manifest**

For information about how to validate the manifest file that describes your Office Add-in and troubleshoot issues with the manifest file, see Validate and troubleshoot issues with your manifest.

### **Troubleshoot user errors**

For information about how to resolve common issues that users may encounter with your Office Add-in, see Troubleshoot user errors with Office Add-ins.


# **Overview of debugging Office Add-ins**

Article • 04/02/2025

Debugging Office Add-ins is essentially the same as debugging any web application. However, a single set of tools won't work for all add-in developers. This is because addins can be developed on different operating systems and run cross-platform. This article helps you find the detailed debugging guidance for your development environment.

#### **Tip**

This article is concerned with debugging in the narrow sense of setting breakpoints and stepping through code. For guidance on testing and troubleshooting, start with **Test Office Add-ins** and **Troubleshoot development errors with Office Addins**.

#### 7 **Note**

Although you should *test* your add-in on all the platforms that you want to support, you'll only very rarely need to *debug* on an environment different from your development computer. For this reason, this article uses "your development computer" and "your development environment" to refer to the environment on which you're debugging. If a problem in the code occurs only on a platform other than the one on your development computer, and you need to set breakpoints or step through code to solve it, then the environment on which you're debugging isn't literally your development environment.

### **Server-side or client-side?**

Debugging the server-side code of an Office Add-in is the same as debugging the server-side of any web application. See the debugging instructions for your IDE or other tools. The following are examples for some of the most popular tools.

- [Debug ASP.NET or ASP.NET Core apps in Visual Studio](https://learn.microsoft.com/en-us/visualstudio/debugger/how-to-enable-debugging-for-aspnet-applications)
- [Debugging Express](https://expressjs.com/en/guide/debugging.html)
- [Node.js Debugging Guide](https://nodejs.org/en/learn/getting-started/debugging)
- [Node.js debugging in VS Code](https://code.visualstudio.com/docs/nodejs/nodejs-debugging)
- [Webpack Debugging](https://webpack.js.org/contribute/debugging/)


The rest of this article is concerned only with debugging client-side JavaScript (which may be transpiled from TypeScript).

## **Special cases**

There are some special cases in which the debugging process differs from normal for a given combination of platform, Office application, and development environment. If you're debugging any of these special cases, use the links in this section to find the proper guidance. Otherwise, continue to General guidance.

- **Debugging the Office.initialize or Office.onReady function**: Debug the initialize and onReady functions.
- **Debugging an Excel custom function in a** *non-shared* **runtime**: Custom functions debugging in a non-shared runtime.
- **Debugging a function command in a** *non-shared* **runtime**:
	- Outlook add-ins on a Windows development computer: Debug function commands in Outlook add-ins
	- Other Office application add-ins or Outlook on a Mac development computer: Debug a function command with a non-shared runtime.
- **Debugging an event-based or spam-reporting Outlook add-in**: [Debug event](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/debug-autolaunch)[based and spam-reporting add-ins.](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/debug-autolaunch)
- **Debugging an add-in in the new Outlook on Windows desktop client (preview)**: See the "Debug your add-in" section of Develop Outlook add-ins for the new Outlook on Windows.
- **Debugging a Blazor-based add-in**: Debug the add-in the same way you would debug a Blazor web application. See [Debug ASP.NET Core Blazor WebAssembly.](https://learn.microsoft.com/en-us/aspnet/core/blazor/debug/)

## **General guidance**

To find guidance for debugging client-side code, the first variable is the operating system of your development computer.

- Windows
- Mac
- Linux or other Unix variant

### **Debug on Windows**

The following provides general guidance to debugging on Windows. Debugging on Windows depends on your IDE.


- **Visual Studio**: Debug using the browser's F12 tools. See Debug Office Add-ins in Visual Studio.
- **Any other IDE** (or you don't want to debug inside your IDE): Use the developer tools that are associated with the webview control that add-ins use on your development computer. See one of the following:
	- For the Trident webview: Debug add-ins using developer tools for Internet Explorer
	- For the EdgeHTML webview: Debug add-ins using developer tools for Edge Legacy
	- For the WebView2 webview: Debug add-ins using developer tools in Microsoft Edge (Chromium-based)

For information about which runtime is being used, see Browsers and webview controls used by Office Add-ins and Runtimes in Office Add-ins.

#### **Tip**

In recent versions of Office, one way to identify the webview control that Office is using is through the **personality menu** on any add-in where it's available. (The personality menu isn't supported in Outlook.) Open the menu and select **Security Info**. In the **Security Info** dialog on Windows, the **Runtime** reports **Microsoft Edge**, **Microsoft Edge Legacy**, or **Internet Explorer**. The runtime isn't included on the dialog in older versions of Office.

#### **Debug on Mac**

Use the Safari Web Inspector. Instructions are in Debug Office Add-ins on a Mac.

#### **Debug on Linux**

There is no desktop version of Office for Linux, so you'll need to sideload the add-in to Office on the web to test and debug it. Debugging guidance is in Debug add-ins in Office on the web.

#### 7 **Note**

We don't recommend that you develop Office Add-ins on a Linux computer except in the unusual case where you can be sure that all the add-in's users will be accessing the add-in through Office on the web from a Linux computer.


## **Debug add-ins in staging or production**

To debug an add-in that is already in staging or production, attach a debugger from the UI of the add-in. For instructions, see Attach a debugger from the task pane.

### **Versions of office.js for debugging**

There are debug versions of the Office JavaScript libraries. These versions are more human readable and easier to step through with a debugger. Use them when the Office JavaScript APIs aren't working as expected. Avoid using them when you publish and deploy your add-in.

The debug versions are found at the following CDN locations.

- Office JavaScript API library: https://appsforoffice.microsoft.com/lib/1/hosted/office.debug.js
- Office JavaScript API (preview) library: https://appsforoffice.microsoft.com/lib/beta/hosted/office.debug.js

### **See also**

- Runtimes in Office Add-ins


# **Sideload Office Add-ins that use the unified manifest for Microsoft 365**

08/13/2025

The process of sideloading an add-in that uses the Unified manifest for Microsoft 365 varies depending on the tool you want to use and on how the add-in project was created.

#### 7 **Note**

An add-in that uses the unified manifest can be sideloaded on Office on Windows, Version 2304 (Build 16320.20000) or later. Sideloading on Windows has the effect of sideloading to Office on the web too. Currently, it can't be sideloaded on Mac or iPad.

## **Sideload add-ins created with the Yeoman generator for Office Add-ins (Yo Office)**

Use the process described in Sideload with a system prompt, bash shell, or terminal.

### **Sideload with Microsoft 365 Agents Toolkit**

- 1. First, *make sure Office desktop application that you want to sideload into is closed.*
- 2. In Visual Studio Code, open Agents Toolkit.
- 3. Required for Outlook only: in the **ACCOUNTS** section, verify that you're signed into Microsoft 365.
- 4. Select **View** | **Run** in Visual Studio Code. In the **RUN AND DEBUG** dropdown menu, select one of these options as appropriate for your add-in.
	- **Excel Desktop (Edge Chromium)**
	- **Outlook Desktop (Edge Chromium)**
	- **PowerPoint Desktop (Edge Chromium)**
	- **Word Desktop (Edge Chromium)**
- 5. Press F5 . The project builds and a Node dev-server window opens. This process may take a couple of minutes and then the desktop version of the Office application that you selected opens. You can now work with your add-in. For an Outlook add-in, be sure you're working in the **Inbox** of *your Microsoft 365 account identity*.


- 6. To stop debugging and uninstall the add-in, select **Run** | **Stop Debugging** in Visual Studio Code. Closing the server window doesn't reliably stop the server and closing the Office application doesn't reliably cause Office to unacquire the add-in.
#### 7 **Note**

If the preceding step seems to have no effect, uninstall the add-in by opening a **TERMINAL** in Visual Studio Code, and then complete the uninstall step — the *last* step — of the section **Sideload with a system prompt, bash shell, or terminal**.

### **Sideload with a system prompt, bash shell, or terminal**

- 1. First, *make sure the Office desktop application that you want to sideload into is closed.*
- 2. Open a system prompt, bash shell, or the Visual Studio Code **TERMINAL**, and navigate to the root of the project.
- 3. The command to sideload the add-in depends on when the project was created. If the "scripts" section of the project's package.json file has a "start:desktop" script, then run npm run start:desktop ; otherwise, run npm run start . The project builds and a Node dev-server window opens. This process may take a couple of minutes then the Office host application (Excel, Outlook, PowerPoint, or Word) desktop opens.
- 4. On some versions of Office, the add-in may not fully activate. For example, the add-in's buttons may not appear on the ribbon. If this happens, select the **Add-ins** button on the **Home** ribbon. On the flyout that opens, select the add-in. This completes the installation.
- 5. You can now work with your add-in.
- 6. When you're done working with your add-in, make sure to run the command npm run stop . Closing the server window doesn't reliably stop the server and closing the Office application doesn't reliably cause Office to unacquire the add-in.

### **Sideload other NodeJS and npm projects**

There are two tools you can use to sideload.

#### **Sideload with the Office-Addin-Debugging tool**

- 1. To sideload the add-in, run the following command. This command puts the unified manifest and the two icon image files that are referenced in the manifest's "icons" property into a zip file and sideloads it to the Office application. It also starts a server in a


separate NodeJS window to host the add-in files on localhost. For more details about this command, see [Office-Addin-Debugging](https://www.npmjs.com/package/office-addin-debugging) .

command line

npx office-addin-debugging start <relative-path-to-unified-manifest> desktop

- 2. When you use office-addin-debugging to start an add-in, *always stop the session with the following command*. Closing the server window doesn't reliably stop the server and closing the Office application doesn't reliably cause Office to unacquire the add-in.

```
command line
npx office-addin-debugging stop <relative-path-to-unified-manifest>
```
#### **Sideload with Microsoft 365 Agents Toolkit CLI (commandline interface)**

- 1. Create a zip package. See Manually create the add-in package file.
- 2. In the root of the project, open a command prompt or bash shell and run the following command to install the Agents Toolkit CLI.

command line

npm install -g @microsoft/m365agentstoolkit-cli

- 3. Run the following command to sideload the add-in.
command line

atk install --file-path <relative-path-to-zip-file>

#### ) **Important**

This command returns some information about the add-in including an autogenerated title ID as shown in the following example.


- 4. When you use the Agents Toolkit CLI to start an add-in, *always stop the session with the following command*. Closing the server window doesn't reliably stop the server and closing the Office application doesn't reliably cause Office to unacquire the add-in. Replace "{title ID}" with the title ID of the add-in including the "U_" prefix; for example, U_90d141c6-cf4f-40ee-b714-9df9ea593f39 .

```
command line
```
atk uninstall --mode title-id --title-id {title ID} --interactive false

#### ) **Important**

The **[documentation for the uninstall command](https://learn.microsoft.com/en-us/microsoftteams/platform/toolkit/teams-toolkit-cli?pivots=version-three#teamsapp-uninstall)** describes a way to use the add-in's manifest ID instead of the title ID. Due to a bug in an API that the CLI calls, this option doesn't currently work. You must use the uninstall command given above and you must include the --interactive false option.

## **Sideload through the Teams app store**


Add-ins that use the unified manifest can be manually sideloaded through the Teams app store, even if they have no Teams-related functionality. The steps are as follows.

- 1. Create an app package manually if it hasn't already been created by a tool. See Manually create the add-in package file.
- 2. Close all Office applications, and then clear the Office cache following the instructions at Manually clear the cache.
- 3. Open Teams and select **Apps** from the app bar, then select **Manage your apps** at the bottom of the **Apps** pane.
- 4. Select **Upload an app** in the **Apps** dialog, and then in the dialog that opens, select **Upload a custom app**.
- 5. In the **Open** dialog, navigate to, and select, the app package.
- 6. Select **Add** in the dialog that opens.
- 7. When you're prompted that the app was added, *don't* open it in Teams. Instead, close Teams.
- 8. The next task is to start a local web server that hosts your project's HTML and JavaScript files. How you do this depends on several factors including the folder structure of your project, the tools you use, such as a bundler, task manager, server application, and how you have configured those tools. The following instruction applies only to projects that meet the following conditions.
	- There's a **webpack.config.js** file in the root of the project that is similar to the ones in add-in projects that are created with the Yeoman Generator for Office Add-ins or Microsoft 365 Agent Toolkit.
	- There's a **package.json** file in the root of the project similar to the ones created by the same two tools and the file has a "scripts" section with the following script in it.

JSON "dev-server": "webpack serve --mode development"

- 9. In a command prompt or Visual Studio Code **TERMINAL** in the root of the project, run npm run dev-server to start the server on localhost.
- 10. Open the Office application that the add-in targets. Wait until the add-in has loaded. This may take as much as two minutes. Depending on your version of Office, ribbon buttons and other artifacts may appear automatically. In some versions, you need to manually


activate the add-in: Select the **Add-ins** button on the **Home** ribbon, and then in the flyout that opens, select your add-in. It will have the name specified in the ["name.short"](https://learn.microsoft.com/en-us/microsoft-365/extensibility/schema/root-name) property of the manifest.

#### ) **Important**

When you want to end a testing session and make changes to the add-in that you sideloaded through the Teams app store, be sure to remove the add-in completely with the following steps.

- 1. Close the Office application.
- 2. Shut down the server. See the documentation for your server application for how to do this. For the webpack dev-server application, shutting it down depends on whether the server is running in the same window in which you ran npm run devserver or a different window. If it's the same window, give the terminal focus and press Ctrl + C . Choose "Y" in response to the prompt to end the process. If it's in a different window, then in the window where you ran npm run dev-server , run npm run stop .
- 3. Clear the Office cache following the instructions at **Manually clear the cache**.
- 4. Open Teams and select **Apps** from the app bar, then select **Manage your apps** at the bottom of the **Apps** pane.
- 5. Find your add-in in the list of apps. It will have the name specified in the "name.short" property of the manifest.
- 6. Select the add-in from the list of apps to expand its row.
- 7. Select the trash can icon and then select **Remove** in the prompt.

Make your changes and then sideload the add-in again.

## **Manually create the add-in package file**

When the unified manifest is used, the unit of installation and sideloading is a zip-formatted package file. This file is usually created for you by the tools you use to create and test your add-in, but there are scenarios in which you create it manually. To do so, use any zip utility to create a zip file that contains the following files.

- The unified manifest, which goes in the root of the zip file.
- The two image files referenced in the "icons" property of the manifest.
- Any localization files that are referenced in the "localizationInfo" property of the manifest.


- Any declarative agent files that are referenced in the "copilotAgents" property.
- Any second-level supplementary files. For example, declarative agent configuration files sometimes reference second-level supplementary files, such as plugin configuration files. These should be included too.

#### ) **Important**

*All of these files must have the same relative path in the zip file as specified in the manifest.* For example, if the path of the two image files is **assets/icon-64.png** and **assets/icon-128.png**, then you must include an **assets** folder with the two files in the zip package. Second-level files, such as plugin configuration files for declarative agents, must have the same relative path in the zip file as they do in the first-level file that references them. For example, if the relative path of a declarative agent file specified in the manifest is **agents/myAgent.json**, then you must include an **agents** folder in the zip package and put the **myAgent.json** file in it. If the declarative agent file, in turn, gives the relative path of **plugins/myPlugin.json** for a plugin configuration file, then you must include a **plugins** subfolder under the **agents** folder and put the **myPlugin.json** file in it.

To maximize compatibility with Microsoft 365 development tools, we recommend that you keep the files that will be included in the package in a folder called **appPackage** in the root of your project, and that you put the package file in a subfolder named **build** in the **appPackage** folder.

The following are examples of the recommended structure. The structure inside the **\build\appPackage.zip** file must mirror the structure of the **appPackage** folder, except for the **build** folder itself.

| Console                                                                                         |  |
|-------------------------------------------------------------------------------------------------|--|
| \appPackage<br>\assets<br>color.png<br>outline.png<br>\build<br>appPackage.zip<br>manifest.json |  |
|                                                                                                 |  |
| Console                                                                                         |  |

\appPackage \agents myAgent.json \plugins myPlugin.json


```
 \assets
 color.png
 outline.png
 \build
 appPackage.zip
 \languages
 fr-FR.json
 es-MX.json
 manifest.json
```


# **Sideload Office Add-ins for testing from a network share**

Article • 05/21/2025

You can test an Office Add-in in an Office client that's on Windows by publishing the manifest to a network file share (instructions follow). This deployment option is intended to be used when you've completed development and testing on a localhost and want to test the add-in from a non-local server or cloud account.

#### ) **Important**

Deployment by network share isn't supported for production add-ins. This method has the following limitations.

- The add-in can only be installed on Windows computers.
- Add-ins that use the **unified manifest for Microsoft 365** aren't supported when published to a network share.
- If a new version of an add-in changes the ribbon, such as by adding a custom tab or custom button to it, each user will have to reinstall the add-in.

#### 7 **Note**

If your add-in project was created with a sufficiently recent version of the **Yeoman generator for Office Add-ins**, the add-in will automatically sideload in the Office desktop client when you run npm start .

This article applies only to testing Word, Excel, PowerPoint, and Project add-ins and only on Windows. If you want to test on another platform or want to test an Outlook add-in, see one of the following topics to sideload your add-in.

- Sideload Office Add-ins in Office on the web for testing
- Sideload Office Add-ins on Mac for testing
- Sideload Office Add-ins on iPad for testing
- Sideload Outlook add-ins for testing

The following video walks you through the process of sideloading your add-in in Office on the web or desktop using a shared folder catalog.

<https://www.youtube-nocookie.com/embed/XXsAw2UUiQo>


### **Share a folder**

- 1. In File Explorer on the Windows computer where you want to host your add-in, go to the parent folder, or drive letter, of the folder you want to use as your shared folder catalog.
- 2. Open the context menu for the folder you want to use as your shared folder catalog (for example, right-click the folder) and choose **Properties**.
- 3. Within the **Properties** dialog window, open the **Sharing** tab and then choose the **Share** button.

| my-folder Properties                                                                       |                                 |  |                   |           |  |  |  |
|--------------------------------------------------------------------------------------------|---------------------------------|--|-------------------|-----------|--|--|--|
| Sharing<br>General                                                                         | Security                        |  | Previous Versions | Customize |  |  |  |
|                                                                                            | Network File and Folder Sharing |  |                   |           |  |  |  |
| my-folder<br>Not Shared                                                                    |                                 |  |                   |           |  |  |  |
| Network Path:<br>Not Shared                                                                |                                 |  |                   |           |  |  |  |
| Share                                                                                      |                                 |  |                   |           |  |  |  |
| Advanced Sharing                                                                           |                                 |  |                   |           |  |  |  |
| Set custom permissions, create multiple shares, and set other<br>advanced sharing options. |                                 |  |                   |           |  |  |  |
| Advanced Sharing                                                                           |                                 |  |                   |           |  |  |  |
|                                                                                            |                                 |  |                   |           |  |  |  |
|                                                                                            |                                 |  |                   |           |  |  |  |
|                                                                                            |                                 |  |                   |           |  |  |  |
|                                                                                            |                                 |  |                   |           |  |  |  |
|                                                                                            |                                 |  |                   |           |  |  |  |

- 4. Within the **Network access** dialog window, add yourself and any other users and/or groups with whom you want to share your add-in. You'll need at least **Read/Write**


permission to the folder. After you've finished choosing people to share with, choose the **Share** button.

- 5. When you see the **Your folder is shared** confirmation, make note of the full network path that's displayed immediately following the folder name. (You'll need to enter this value as the **Catalog Url** when you specify the shared folder as a trusted catalog, as described in the next section of this article.) Choose the **Done** button to close the **Network access** dialog window.

|  | Network access                                                                                                   | × |
|--|------------------------------------------------------------------------------------------------------------------|---|
|  | Your folder is shared.                                                                                           |   |
|  | You can e-mail someone links to these shared items, or copy and paste the links into another<br>app.             |   |
|  | Individual Items<br>my-folder<br>\\KBRANDL-2017\my-folder                                                        | 1 |
|  | Shared items aren't accessible when your computer is asleep.<br>Show me all the network shares on this computer. |   |
|  | Done                                                                                                             |   |

- 6. Choose the **Close** button to close the **Properties** dialog window.
### **Specify the shared folder as a trusted catalog**

There are two options for how you specify this trust. Follow the instructions for the option that works better for your setup.

- Configure the trust manually.
- Configure the trust with a Registry script.

#### **Configure the trust manually**


- 1. Open a new document in Excel, Word, PowerPoint, or Project.
- 2. Choose the **File** tab, and then choose **Options**.
- 3. Choose **Trust Center**, and then choose the **Trust Center Settings** button.
- 4. Choose **Trusted Add-in Catalogs**.
- 5. In the **Catalog Url** box, enter the full network path to the folder that you shared previously. If you failed to note the folder's full network path when you shared the folder, you can get it from the folder's **Properties** dialog window, as shown in the following screenshot.

| my-folder Properties |                                                                                                     |          |  |                   | ×                                                             |  |
|----------------------|-----------------------------------------------------------------------------------------------------|----------|--|-------------------|---------------------------------------------------------------|--|
| General              | Sharing                                                                                             | Security |  | Previous Versions | Customize                                                     |  |
|                      | Network File and Folder Sharing<br>my-folder<br>Shared<br>Network Path:<br>\\KBRANDL-2017\my-folder |          |  |                   |                                                               |  |
|                      | Share<br>Advanced Sharing<br>advanced sharing options.<br>Advanced Sharing                          |          |  |                   | Set custom permissions, create multiple shares, and set other |  |
|                      |                                                                                                     |          |  |                   |                                                               |  |
|                      |                                                                                                     | OK       |  | Cancel            | Apply                                                         |  |


- 6. After you've entered the full network path of the folder into the **Catalog Url** box, choose the **Add catalog** button.
- 7. Select the **Show in Menu** check box for the newly-added item, and then choose the **OK** button to close the **Trust Center** dialog window.

| Trust Center            |                                                                                       | 2           | × |
|-------------------------|---------------------------------------------------------------------------------------|-------------|---|
| Trusted Publishers      | Trusted Web Add-in Catalogs                                                           |             |   |
| Trusted Locations       | Use these settings to manage your web add-in catalogs.                                |             |   |
| Trusted Documents       | Don't allow any web add-ins to start.                                                 |             |   |
| Trusted Add-in Catalogs | Don't allow web add-ins from the Qffice Store to start.                               |             |   |
| Add-ins                 | Trusted Catalogs Table                                                                |             |   |
| ActiveX Settings        | You should only add a catalog if you trust its owner. You may also select one of each |             |   |
| Macro Settings          | catalog type to show in the insert add-in menu. We will automatically start web       |             |   |
| Protected View          | add-ins from your insert add-in menu catalogs when opening documents.                 |             |   |
| Message Bar             | Catalog Url:                                                                          | Add catalog |   |
| File Block Settings     | Catalog Type<br>Show in Menu<br>Trusted Catalog Address<br>V                          |             |   |
| Privacy Options         | \\KBRANDL-2017\my-folder<br>Network share                                             |             |   |
|                         |                                                                                       |             |   |
|                         |                                                                                       |             |   |
|                         |                                                                                       |             |   |
|                         |                                                                                       | Remove      |   |
|                         |                                                                                       | Clear       |   |
|                         | OK                                                                                    | Cancel      |   |

- 8. Choose the **OK** button to close the **Options** dialog window.
- 9. Close and reopen the Office application so your changes will take effect.

#### **Configure the trust with a Registry script**

- 1. In a text editor, create a file named **TrustNetworkShareCatalog.reg**.
- 2. Add the following content to the file.

```
text
Windows Registry Editor Version 5.00
[HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{-
random-GUID-here-}]
"Id"="{-random-GUID-here-}"
```


- 3. Use one of the many online GUID generation tools, such as [GUID Generator](https://guidgenerator.com/) , to generate a random GUID, and within the TrustNetworkShareCatalog.reg file, replace the string "-random-GUID-here-" *in both places* with the GUID. (The enclosing {} symbols should remain.)
- 4. Replace the Url value with the full network path to the folder that you shared previously. (Note that any \ characters in the URL must be doubled.) If you failed to note the folder's full network path when you shared the folder, you can get it from the folder's **Properties** dialog window, as shown in the following screenshot.

| my-folder Properties                                                                                                                                                          |  |                  |  |                                                               |  |       |  |
|-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|--|------------------|--|---------------------------------------------------------------|--|-------|--|
| General                                                                                                                                                                       |  | Security         |  |                                                               |  |       |  |
| Sharing<br>Previous Versions<br>Customize<br>Network File and Folder Sharing<br>my-folder<br>Shared<br>Network Path:<br>\\KBRANDL-2017\my-folder<br>Share<br>Advanced Sharing |  |                  |  |                                                               |  |       |  |
| advanced sharing options.                                                                                                                                                     |  | Advanced Sharing |  | Set custom permissions, create multiple shares, and set other |  |       |  |
|                                                                                                                                                                               |  |                  |  |                                                               |  |       |  |
|                                                                                                                                                                               |  | OK               |  | Cancel                                                        |  | Apply |  |


- 5. The file should now look like the following. Save it.

```
text
Windows Registry Editor Version 5.00
[HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\
{01234567-89ab-cedf-0123-456789abcedf}]
"Id"="{01234567-89ab-cedf-0123-456789abcedf}"
"Url"="\\\\TestServer\\OfficeAddinManifests"
"Flags"=dword:00000001
```
- 6. Close *all* Office applications.
- 7. Run the TrustNetworkShareCatalog.reg just as you would any executable, such as doubleclicking it.

### **Sideload your add-in**

- 1. Put the manifest XML file of any add-in that you're testing into the shared folder catalog. Note that you deploy the web application itself to a web server. Be sure to specify the URL in the **<SourceLocation>** element of the manifest file.
#### ) **Important**

While not strictly required in all add-in scenarios, using an HTTPS endpoint for your add-in is strongly recommended. Add-ins that are not SSL-secured (HTTPS) generate unsecure content warnings and errors during use. If you plan to run your add-in in Office on the web or publish your add-in to AppSource, it must be SSL-secured. If your add-in accesses external data and services, it should be SSL-secured to protect data in transit. Self-signed certificates can be used for development and testing, so long as the certificate is trusted on the local machine.

#### 7 **Note**

For Visual Studio projects, use the manifest built by the project in the {projectfolder}\bin\Debug\OfficeAppManifests folder.

- 2. In Excel, Word, or PowerPoint, select **Home** > **Add-ins** from the ribbon, then select **Advanced**. In Project, select **My Add-ins** on the **Project** tab of the ribbon.
- 3. Choose **SHARED FOLDER** at the top of the **Office Add-ins** dialog box.


4. Select the name of the add-in and choose **Add** to insert the add-in.

### **Remove a sideloaded add-in**

You can remove a previously sideloaded add-in by clearing the Office cache on your computer. Details on how to clear the cache on Windows can be found in the article Clear the Office cache.

## **See also**

- Validate an Office Add-in's manifest
- Clear the Office cache
- Publish your Office Add-in


# **Attach a debugger from the task pane**

Article • 05/20/2023

In some environments, a debugger can be attached on an Office Add-in that is already running. This can be useful when you want to debug an add-in that is already in staging or production. If you are still developing and testing the add-in, see Overview of debugging Office Add-ins.

The technique described in this article can be used only when the following conditions are met.

- The add-in is running in Office on Windows.
- The computer is using a combination of Windows and Office versions that use the Edge (Chromium-based) webview control, WebView2. To determine which webview you're using, see Browsers and webview controls used by Office Add-ins.

#### **Tip**

In recent versions of Office, one way to identify the webview control that Office is using is through the **personality menu** on any add-in where it's available. (The personality menu isn't supported in Outlook.) Open the menu and select **Security Info**. In the **Security Info** dialog on Windows, the **Runtime** reports **Microsoft Edge**, **Microsoft Edge Legacy**, or **Internet Explorer**. The runtime isn't included on the dialog in older versions of Office.

To launch the debugger, choose the top right corner of the task pane to activate the **Personality** menu (as shown in the red circle in the following image).


Select **Attach Debugger**. This launches the Microsoft Edge (Chromium-based) developer tools. Use the tools as described in Debug add-ins using developer tools in Microsoft Edge (Chromium-based).

### **See also**

- Overview of debugging Office Add-ins

# **Debug add-ins using developer tools in Microsoft Edge (Chromium-based)**

Article • 07/14/2024

This article shows how to debug the client-side code (JavaScript or TypeScript) of your add-in when the following conditions are met.

- You can't, or don't wish to, debug using tools built into your IDE; or you are encountering a problem that only occurs when the add-in is run outside the IDE.
- Your computer is using a combination of Windows and Office versions that use the Edge (Chromium-based) webview control, WebView2.

### **Tip**

For information about debugging with Edge WebView2 (Chromium-based) inside Visual Studio Code, see **Debug add-ins on Windows using Visual Studio Code and Microsoft Edge WebView2 (Chromium-based)**.

To determine which webview you're using, see Browsers and webview controls used by Office Add-ins.

### **Tip**

In recent versions of Office, one way to identify the webview control that Office is using is through the **personality menu** on any add-in where it's available. (The personality menu isn't supported in Outlook.) Open the menu and select **Security Info**. In the **Security Info** dialog on Windows, the **Runtime** reports **Microsoft Edge**, **Microsoft Edge Legacy**, or **Internet Explorer**. The runtime isn't included on the dialog in older versions of Office.

# **Debug a task pane add-in using Microsoft Edge (Chromium-based) developer tools**

### 7 **Note**

If your add-in has an **add-in command** that executes a function, the function runs in a hidden browser runtime process that the Microsoft Edge (Chromium-based)


developer tools can't be launched from, so the technique described in this article can't be used to debug code in the function.

- 1. Sideload and run the add-in.
7 **Note**

To sideload an add-in in Outlook, see **Sideload Outlook add-ins for testing**.

- 2. Run the Microsoft Edge (Chromium-based) developer tools by one of these methods:
	- Be sure the add-in's task pane has focus and press Ctrl + Shift + I .
	- Right-click (or select and hold) the task pane to open the context menu and select **Inspect**, or open the personality menu and select **Attach Debugger**. (The personality menu isn't supported in Outlook.)

### 7 **Note**

The new Outlook on Window desktop client (preview) doesn't support the context menu or the keyboard shortcut to access the Microsoft Edge developer tools. Instead, you must run olk.exe --devtools from a command prompt. For more information, see the "Debug your add-in" section of **Develop Outlook add-ins for the new Outlook on Windows**.

- 3. Open the **Sources** tab.
- 4. Open the file that you want to debug with the following steps.
	- a. On the far right of the tool's top menu bar, select the **...** button and then select **Search**.
	- b. Enter a line of code from the file you want to debug in the search box. It should be something that's not likely to be in any other file.
	- c. Select the refresh button.
	- d. In the search results, select the line to open the code file in the pane above the search results.


- 5. To set a breakpoint, select the line number of the line in the code file. A red dot appears by the line in the code file. In the debugger window to the right, the breakpoint is registered in the **Breakpoints** drop down.
- 6. Execute functions in the add-in as needed to trigger the breakpoint.

### **Tip**

For more information about using the tools, see **[Microsoft Edge Developer Tools](https://learn.microsoft.com/en-us/microsoft-edge/devtools-guide-chromium/) [overview](https://learn.microsoft.com/en-us/microsoft-edge/devtools-guide-chromium/)**.

# **Debug a dialog in an add-in**

If your add-in uses the Office Dialog API, the dialog runs in a separate process from the task pane (if any) and the tool must be started from that separate process. Follow these steps.

- 1. Run the add-in.
- 2. Open the dialog and be sure it has focus.


- 3. Open the Microsoft Edge (Chromium-based) developer tools by one of these methods:
	- Press Ctrl + Shift + I or F12 .
	- Right-click (or select and hold) the dialog to open the context menu and select **Inspect**.
- 4. Use the tool the same as you would for code in a task pane. See Debug a task pane add-in using Microsoft Edge (Chromium-based) developer tools earlier in this article.


# **Debug Office Add-ins on Windows using Visual Studio Code and Microsoft Edge WebView2 (Chromium-based)**

Article • 04/15/2024

Office Add-ins running on Windows can debug against the Edge Chromium WebView2 runtime directly in Visual Studio Code.

#### ) **Important**

This article only applies when Office runs add-ins in the Microsoft Edge Chromium WebView2 runtime, as explained in **Browsers and webview controls used by Office Add-ins**. For instructions about debugging in Visual Studio Code against Microsoft Edge Legacy with the original WebView (EdgeHTML) runtime, see **Debug add-ins using developer tools in Microsoft Edge Legacy**.

### **Tip**

If you can't, or don't wish to, debug using tools built into Visual Studio Code; or you're encountering a problem that only occurs when the add-in is run outside Visual Studio Code, you can debug Edge Chromium WebView2 runtime by using the Edge (Chromium-based) developer tools as described in **Debug add-ins using developer tools for Microsoft Edge WebView2**.

This debugging mode is dynamic, allowing you to set breakpoints while code is running. See changes in your code immediately while the debugger is attached, all without losing your debugging session. Your code changes also persist, so you see the results of multiple changes to your code. The following image shows this extension in action.


# **Prerequisites**

- [Visual Studio Code](https://code.visualstudio.com/)
- [Node.js (version 10+)](https://nodejs.org/)
- Windows 10, 11
- A combination of platform and Office application that supports Microsoft Edge with WebView2 (Chromium-based) as explained in Browsers and webview controls used by Office Add-ins. If your version of Office from a Microsoft 365 subscription is earlier than Version 2101, you'll need to install WebView2. For instructions to install WebView2, see [Microsoft Edge WebView2 / Embed web content ... with](https://developer.microsoft.com/microsoft-edge/webview2/) [Microsoft Edge WebView2.](https://developer.microsoft.com/microsoft-edge/webview2/)

# **Debug a project created with Yo Office**

These instructions assume you have experience using the command line, understand basic JavaScript, and have created an Office Add-in project before using the Yeoman generator for Office Add-ins. If you haven't done this before, consider visiting one of our tutorials, such as the Excel Office Add-in tutorial.

- 1. The first step depends on the project and how it was created.
	- If you want to create a project to experiment with debugging in Visual Studio Code, use the Yeoman generator for Office Add-ins. Follow any of the Yo Office quick start guides, such as the Outlook add-in quick start.


- If you want to debug an existing project that was created with Yo Office, skip to the next step.
- 2. Open VS Code and open your project in it.
- 3. Choose **View** > **Run** or enter Ctrl + Shift + D to switch to debug view.
- 4. From the **RUN AND DEBUG** options, choose the Edge Chromium option for your host application, such as **Outlook Desktop (Edge Chromium)**. Select F5 or choose **Run** > **Start Debugging** from the menu to begin debugging. This action automatically launches a local server in a Node window to host your add-in and then automatically opens the host application, such as Excel or Word. This may take several seconds.

### **Tip**

If you aren't using a project created with Yo Office, you may be prompted to adjust a registry key. While in the root folder of your project, run the following in the command line.

command line

npx office-addin-debugging start <your manifest path>

#### ) **Important**

If your project was created with older versions of Yo Office, you may see the following error dialog box about 10 - 30 seconds after you start debugging (at which point you may have already gone on to another step in this procedure) and it may be hidden behind the dialog box described in the next step.

Complete the tasks in the **Appendix** and then restart this procedure.

- 5. In the host application, your add-in is now ready to use. Select **Show Taskpane** or run any other add-in command. A dialog box will appear with text similar to the


#### following:

WebView Stop On Load. To debug the webview, attach VS Code to the webview instance using the Microsoft Debugger for Edge extension, and click OK to continue. To prevent this dialog from appearing in the future, click Cancel.

Select **OK**.

### 7 **Note**

If you select **Cancel**, the dialog won't be shown again while this instance of the add-in is running. However, if you restart your add-in, you'll see the dialog again.

- 6. You're now able to set breakpoints in your project's code and debug. To set breakpoints in Visual Studio Code, hover next to a line of code and select the red circle that appears.

|    | 16 | export async function run() {                           |
|----|----|---------------------------------------------------------|
|    | 17 | try {                                                   |
|    | 18 | await Excel.run(async context => {{                     |
|    | 19 | **                                                      |
|    | 20 | * Insert your Excel code here                           |
|    | 21 | * /                                                     |
|    | 22 | const range = context.workbook.getSelectedRange();      |
|    | 23 |                                                         |
|    | 24 | / / Read the range address                              |
|    | 25 | range.load("address");                                  |
|    | 26 |                                                         |
|    | 27 | / / Update the fill color                               |
| 80 | 28 | range.format.fill.color = "yellow";                     |
|    | 29 |                                                         |
|    | 30 | await context.sync();                                   |
|    | 31 | console.log(`The range address was ${range.address}.`); |
|    | 32 | });                                                     |
|    | 33 | catch (error) {<br>8                                    |
|    | 34 | console.error(error);                                   |
|    | 35 | }                                                       |
|    | 36 |                                                         |
|    | 37 |                                                         |

- 7. Run functionality in your add-in that calls the lines with breakpoints. You'll see that breakpoints have been hit and you can inspect local variables.


#### 7 **Note**

Breakpoints in calls of Office.initialize or Office.onReady are ignored. For details about these functions, see **Initialize your Office Add-in**.

#### ) **Important**

The best way to stop a debugging session is to select Shift + F5 or choose **Run** > **Stop Debugging** from the menu. This action should close the Node server window and attempt to close the host application, but there'll be a prompt on the host application asking you whether to save the document or not. Make an appropriate choice and let the host application close. Avoid manually closing the Node window or host application. Doing so can cause bugs especially when you are stopping and starting debugging sessions repeatedly.

If debugging stops working---for example, if breakpoints are being ignored---stop debugging. Then, if necessary, close all host application windows and the Node window. Finally, close Visual Studio Code and reopen it.

## **Debug a project not created with Yo Office**

If your project wasn't created with Yo Office, you need to create a debug configuration for Visual Studio Code.

## **Configure package.json file**

JSON

- 1. Ensure you have a package.json file. If you don't already have a package.json file, run npm init in the root folder of your project and answer the prompts.
- 2. Run npm install office-addin-debugging . This package sideloads your add-in for debugging.
- 3. Open the package.json file. In the scripts section, add the following script.

```
"start:desktop": "office-addin-debugging start $MANIFEST_FILE$
desktop",
"dev-server": "$SERVER_START$"
```


- 4. Replace $MANIFEST_FILE$ with the correct file name and folder location of your manifest.
- 5. Replace $SERVER_START$ with the command to start your web server. Later in these steps, the office-addin-debugging package will specifically look for the dev-server script to launch your web server.
- 6. Save and close the package.json file.

## **Configure launch.json file**

- 1. Create a file named launch.json in the \.vscode folder of the project if there isn't one there already.
- 2. Copy the following JSON into the file.

```
JSON
{
 // Other properties may be here.
 "configurations": [
 {
 "name": "$HOST$ Desktop (Edge Chromium)",
 "type": "msedge",
 "request": "attach",
 "useWebView": true,
 "port": 9229,
 "timeout": 600000,
 "webRoot": "${workspaceRoot}",
 "preLaunchTask": "Debug: Excel Desktop"
 }
 ]
 // Other properties may be here.
}
```
#### 7 **Note**

If you already have a launch.json file, just add the single configuration to the configurations section.

- 3. Replace the placeholder $HOST$ with the name of the Office application that the add-in runs in. For example, Outlook or Word .
- 4. Save and close the file.


## **Configure tasks.json**

- 1. Create a file named tasks.json in the \.vscode folder of the project.
- 2. Copy the following JSON into the file. It contains a task that starts debugging for your add-in.

```
JSON
{
 "version": "2.0.0",
 "tasks": [
 {
 "label": "Debug: $HOST$ Desktop",
 "type": "shell",
 "command": "npm",
 "args": ["run", "start:desktop", "--", "--app", "$HOST$"],
 "presentation": {
 "clear": true,
 "panel": "dedicated"
 },
 "problemMatcher": []
 }
 ]
}
```
7 **Note**

If you already have a tasks.json file, just add the single task to the tasks section.

- 3. Replace both instances of the placeholder $HOST$ with the name of the Office application that the add-in runs in. For example, Outlook or Word .
You can now debug your project using the VS Code debugger (F5).

## **Appendix**

- 1. In the error dialog box, select the **Cancel** button.
- 2. If debugging doesn't stop automatically, select Shift + F5 or choose **Run** > **Stop Debugging** from the menu.
- 3. Close the Node window where the local server is running, if it doesn't close automatically.
- 4. Close the Office application if it doesn't close automatically.
- 5. Open the \.vscode\launch.json file in the project.


- 6. In the configurations array, there are several configuration objects. Find the one whose name has the pattern $HOST$ Desktop (Edge Chromium) , where $HOST$ is an Office application that your add-in runs in; for example, Outlook Desktop (Edge Chromium) or Word Desktop (Edge Chromium) .
- 7. Change the value of the "type" property from "edge" to "pwa-msedge" .
- 8. Change the value of the "useWebView" property from the string "advanced" to the boolean true (note there are no quotation marks around the true ).
- 9. Save the file.
- 10. Close VS Code.

# **See also**

- Test and debug Office Add-ins
- Debug add-ins using developer tools for Internet Explorer
- Debug add-ins using developer tools for Edge Legacy
- Debug add-ins using developer tools in Microsoft Edge (Chromium-based)
- Attach a debugger from the task pane
- Runtimes in Office Add-ins

# **Debug Office Add-ins in Visual Studio**

Article • 02/18/2025

This article describes how to debug client-side code in Office Add-ins that are created with one of the Office Add-in project templates in Visual Studio 2022. For information about debugging server-side code in Office Add-ins, see Overview of debugging Office Add-ins - Server-side or client-side?.

#### 7 **Note**

You can't use Visual Studio to debug add-ins in Office on Mac. For information about debugging on a Mac, see **Debug Office Add-ins on a Mac**.

# **Review the build and debug properties**

Before you start debugging, review the properties of each project in the solution to confirm that Visual Studio will open the desired Office application and that other build and debug properties are set appropriately.

# **Add-in project properties**

Open the **Properties** window for the add-in project to review project properties.

- 1. In **Solution Explorer**, choose the add-in project (*not* the web application project).
- 2. From the menu bar, choose **View** > **Properties Window**.

The following table describes the properties of the add-in project.

#### ノ **Expand table**

| Property                                                              | Description                                                                                                                                                                                                                                                                            |  |  |  |  |  |
|-----------------------------------------------------------------------|----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|--|--|--|--|--|
| Start Action                                                          | Specifies the debug mode for your add-in. This should be set to Office<br>Desktop Client when you want to debug in Microsoft 365 on Windows. To<br>debug in Microsoft 365 on the web, it should be set to Microsoft Edge.                                                              |  |  |  |  |  |
| Start Document<br>(Excel,<br>PowerPoint, and<br>Word add-ins<br>only) | Specifies what document to open when you start the project. In a new<br>project, this is set to [New Excel Workbook], [New Word Document], or<br>[New PowerPoint Presentation]. To specify a particular document, follow the<br>steps in Use an existing document to debug the add-in. |  |  |  |  |  |


| Property                                                                                                                                                                                                                                         | Description                                                                                                                                                       |  |  |  |  |
|--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|-------------------------------------------------------------------------------------------------------------------------------------------------------------------|--|--|--|--|
| Web Project                                                                                                                                                                                                                                      | Specifies the name of the web project associated with the add-in.                                                                                                 |  |  |  |  |
| Email Address                                                                                                                                                                                                                                    | Specifies the email address of the user account in Exchange Server or                                                                                             |  |  |  |  |
| (Outlook add-ins                                                                                                                                                                                                                                 | Exchange Online that you want to use to test your Outlook add-in. If left                                                                                         |  |  |  |  |
| only)                                                                                                                                                                                                                                            | blank, you'll be prompted for the email address when you start debugging.                                                                                         |  |  |  |  |
| EWS Url<br>(Outlook add-ins<br>only)                                                                                                                                                                                                             | Specifies the Exchange Web Services URL (For example:<br>https://www.contoso.com/ews/exchange.aspx ). This property can be left blank.                            |  |  |  |  |
| OWA Url<br>(Outlook add-ins<br>only)                                                                                                                                                                                                             | Specifies the Outlook on the web URL (For example:<br>https://www.contoso.com/owa ). This property can be left blank.                                             |  |  |  |  |
| Use multi-factor                                                                                                                                                                                                                                 | Specifies the boolean value that indicates whether multi-factor                                                                                                   |  |  |  |  |
| auth                                                                                                                                                                                                                                             | authentication should be used. The default is false, but the property has no                                                                                      |  |  |  |  |
| (Outlook add-ins                                                                                                                                                                                                                                 | practical effect. If you normally have to provide a second factor to login to                                                                                     |  |  |  |  |
| only)                                                                                                                                                                                                                                            | the email account, you'll be prompted to when you start debugging.                                                                                                |  |  |  |  |
| User Name                                                                                                                                                                                                                                        | Specifies the name of the user account in Exchange Server or Exchange                                                                                             |  |  |  |  |
| (Outlook add-ins                                                                                                                                                                                                                                 | Online that you want to use to test your Outlook add-in. This property can be                                                                                     |  |  |  |  |
| only)                                                                                                                                                                                                                                            | left blank.                                                                                                                                                       |  |  |  |  |
| Project File                                                                                                                                                                                                                                     | Specifies the name of the file containing build, configuration, and other<br>information about the project.                                                       |  |  |  |  |
| Project Folder                                                                                                                                                                                                                                   | Specifies the location of the project file.                                                                                                                       |  |  |  |  |
| Active<br>Deployment<br>Configuration<br>(present only<br>when debugging<br>Excel,<br>PowerPoint, or<br>Word on the web)                                                                                                                         | Specifies the deployment configuration. This should be set to Default.                                                                                            |  |  |  |  |
| Server<br>Specifies whether the project connects to the SharePoint service specified in<br>Connection<br>the Site URL property. This should be set to Online.<br>(present only<br>when debugging<br>Excel,<br>PowerPoint, or<br>Word on the web) |                                                                                                                                                                   |  |  |  |  |
| Site URL<br>(present only<br>when debugging<br>Excel,                                                                                                                                                                                            | Specifies the full, absolute URL of the SharePoint tenant that you want to<br>host the add-in pages when debugging. For example<br>https://mysite.sharepoint.com/ |  |  |  |  |


| Property |  |
|----------|--|
|----------|--|

#### **Property Description**

PowerPoint, or Word on the web)

#### 7 **Note**

For an Outlook add-in, you may choose to specify values for one or more of the *Outlook add-ins only* properties in the **Properties** window, but doing so isn't required.

# **Web application project properties**

Open the **Properties** window for the web application project to review project properties.

- 1. In **Solution Explorer**, choose the web application project.
- 2. From the menu bar, choose **View** > **Properties Window**.

The following table describes the properties of the web application project that are most relevant to Office Add-in projects.

|  | ノ | Expand table |  |
|--|---|--------------|--|

| Property          | Description                                                                                                                                                                                    |
|-------------------|------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| SSL               | Specifies whether SSL is enabled on the site. This property should be set to True for                                                                                                          |
| Enabled           | Office Add-in projects.                                                                                                                                                                        |
| SSL URL           | Specifies the secure HTTPS URL for the site. Read-only.                                                                                                                                        |
| URL               | Specifies the HTTP URL for the site. Read-only.                                                                                                                                                |
| Project           | Specifies the name of the file containing build, configuration, and other information                                                                                                          |
| File              | about the project.                                                                                                                                                                             |
| Project<br>Folder | Specifies the location of the project file. Read-only. The manifest file that Visual<br>Studio generates at runtime is written to the bin\Debug\OfficeAppManifests folder in<br>this location. |

# **Debug an add-in project on Windows desktop**


This section describes how to start and debug an add-in in desktop Office on Windows; that is, when the **Start Action** property of the add-in project is set to **Office Desktop Client**.

# **Start the add-in project**

Start the project by choosing **Debug** > **Start Debugging** from the menu bar or press the F5 button. Visual Studio automatically builds the solution and starts the Office host application.

When Visual Studio builds the project, it performs the following tasks:

- 1. Creates a copy of the add-in only manifest file and adds it to the _ProjectName_\bin\Debug\OfficeAppManifests directory. The Office application that hosts your add-in consumes this copy when you start Visual Studio and debug the add-in.
- 2. Creates a set of registry entries on your Windows computer that enables the addin to appear in the Office application.
- 3. Builds the web application project, and then deploys it to the local IIS web server ( https://localhost ).
- 4. If this is the first add-in project that you have deployed to the local IIS web server, you may be prompted to install a Self-Signed Certificate to the current user's Trusted Root Certificate store. This is required for IIS Express to display the content of your add-in correctly.

### 7 **Note**

If Office uses the Edge Legacy webview control (EdgeHTML) to run add-ins on your Windows computer, Visual Studio may prompt you to add a local network loopback exemption. This is required for the webview control to be able to access the website deployed to the local IIS web server. You can also change this setting anytime in Visual Studio under **Tools** > **Options** > **Office Tools (Web)** > **Web Add-In Debugging**. To find out what webview control is used on your Windows computer, see **Browsers and webview controls used by Office Add-ins**.

## **Tip**

In recent versions of Office, one way to identify the webview control that Office is using is through the **personality menu** on any add-in where it's available. (The


personality menu isn't supported in Outlook.) Open the menu and select **Security Info**. In the **Security Info** dialog on Windows, the **Runtime** reports **Microsoft Edge**, **Microsoft Edge Legacy**, or **Internet Explorer**. The runtime isn't included on the dialog in older versions of Office.

Next, Visual Studio does the following:

- 1. Modifies the [SourceLocation](https://learn.microsoft.com/en-us/javascript/api/manifest/sourcelocation) element of the add-in only manifest file (that was copied to the _ProjectName_\bin\Debug\OfficeAppManifests directory) by replacing the ~remoteAppUrl token with the fully qualified address of the start page (for example, https://localhost:44302/Home.html ).
- 2. Starts the web application project in IIS Express.
- 3. Validates the manifest.

#### ) **Important**

If you get validation errors for the manifest, it may be that Visual Studio's manifest schema files haven't been updated to support the latest features. Your first troubleshooting step should be to replace one or more of these files with the latest versions. For detailed instructions, see **Manifest schema validation errors in Visual Studio projects**.

- 4. Opens the Office application and sideloads your add-in.
# **Debug the add-in**

The best method for debugging an add-in in Visual Studio 2022 depends on whether the add-in is running in WebView2, which is the webview control that is associated with Microsoft Edge (Chromium), or an older webview control. If your computer is using WebView2, see Use the built-in Visual Studio debugger to debug on the desktop. For any other webview control, see Use the browser developer tools to debug on the desktop. To determine which webview control is being used, see Browsers and webview controls used by Office Add-ins.

## **Tip**

In recent versions of Office, one way to identify the webview control that Office is using is through the **personality menu** on any add-in where it's available. (The personality menu isn't supported in Outlook.) Open the menu and select **Security** 


**Info**. In the **Security Info** dialog on Windows, the **Runtime** reports **Microsoft Edge**, **Microsoft Edge Legacy**, or **Internet Explorer**. The runtime isn't included on the dialog in older versions of Office.

### **Use the built-in Visual Studio debugger to debug on the desktop**

- 1. Set breakpoints, as needed, in the source JavaScript or TypeScript files. You can do this either before or after you start the add-in as described in the earlier section Start the add-in project. If setting a breakpoint causes the Internet Information Services (IIS) server to shut down, restart debugging after you have set your breakpoints.
- 2. When the add-in is running, use the add-in's UI to run the code that contains your breakpoints.

#### ) **Important**

Breakpoints set in Office.initialize or Office.onReady aren't hit. To debug these methods, see **Debug the initialize and onReady functions**.

#### **Tip**

If you encounter any problems, there's more information at **[Debug a JavaScript or](https://learn.microsoft.com/en-us/visualstudio/javascript/debug-nodejs?view=vs-2022&preserve-view=true) [TypeScript app in Visual Studio](https://learn.microsoft.com/en-us/visualstudio/javascript/debug-nodejs?view=vs-2022&preserve-view=true)**.

#### **Use the browser developer tools to debug on the desktop**

- 1. Follow the steps in the earlier section Start the add-in project.
- 2. Launch the add-in in the Office application if it isn't already open. For example, if it's a task pane add-in, it'll have added a button (for example, a **Show Taskpane** button) to the **Home** ribbon or to a custom ribbon tab that's installed with the add-in. Select the button on the ribbon.
- 3. Open the personality menu and then choose **Attach a debugger**. This action opens the debugging tools for the webview control that Office is using to run add-ins on your Windows computer. You can set breakpoints and step through code as described in one of the following articles:
	- Debug add-ins using developer tools for Internet Explorer


- Debug add-ins using developer tools for Edge Legacy
- Debug add-ins using developer tools in Microsoft Edge (Chromium-based)
- 4. To make changes to your code, first stop the debugging session in Visual Studio and close the Office application. Make your changes, and start a new debugging session.

# **Debug an add-in project in Microsoft 365 on the web**

This section describes how to start and debug an add-in in desktop Office on the web; that is, when the **Start Action** property of the add-in project is set to **Microsoft Edge**.

# **Start the add-in project on the web**

Start the project by choosing **Debug** > **Start Debugging** from the menu bar or press the F5 button. Visual Studio automatically builds the solution and launches the Office application host page of your Microsoft 365 tenancy.

#### 7 **Note**

When you're debugging an add-in on the web, you may get an AADSTS50011 error similar to the following:

"The redirect URI {Full absolute URL to add-in home page} specified in the request doesn't match the redirect URIs configured for the application ... "

This occurs because new web applications that are deployed to SharePoint may take up to 24 hours to be available. To make your add-in debuggable immediately, take the following steps:

- 1. Stop debugging in Visual Studio.
- 2. **[Create a PowerShell script](https://learn.microsoft.com/en-us/powershell/scripting/windows-powershell/ise/how-to-write-and-run-scripts-in-the-windows-powershell-ise)** with the following lines. Replace the placeholder {Full absolute URL to add-in home page} with the redirect URL in the error message; for example, "https://contoso-

79d42f062409ae.sharepoint.com/_forms/default.aspx".

PowerShell


```
'00000003-0000-0ff1-ce00-000000000000'"
$sharepointPrincipal | fl
$replyUrls = $sharepointPrincipal.ReplyUrls
$replyUrls += "{Full absolute URL to add-in home page}"
Update-MgServiceprincipal -ServicePrincipalId
$sharepointPrincipal.Id -ReplyUrls $replyUrls
```
- 3. Run the script in PowerShell.
- 4. Restart the project by choosing **Debug** > **Start Debugging** from the menu bar or press the F5 button.

When Visual Studio builds the project it performs the following tasks.

- 1. Prompts you for login credentials. If you're asked to sign in repeatedly or if you receive an error that you're unauthorized, then Basic Auth may be disabled for accounts on your Microsoft 365 tenant. In this case, try using a Microsoft account instead. You can also try setting the property **Use multi-factor auth** to **True** in the add-in project properties pane. See Add-in project properties.
- 2. Creates a copy of the add-in only manifest file and adds it to the _ProjectName_\bin\Debug\OfficeAppManifests directory. Microsoft 365 consumes this copy when you start Visual Studio and debug the add-in.
- 3. Builds the web application project, and then deploys it to the Microsoft 365 tenancy.

Next, Visual Studio does the following:

- 1. Modifies the [SourceLocation](https://learn.microsoft.com/en-us/javascript/api/manifest/sourcelocation) element of the add-in only manifest file (that was copied to the _ProjectName_\bin\Debug\OfficeAppManifests directory) by replacing the ~remoteAppUrl token with the fully qualified address of the start page (for example, https://contoso-79d42f062409ae.sharepoint.com/_forms/default.aspx ).
- 2. Starts the web application project.
- 3. Validates the manifest.

### ) **Important**

If you get validation errors for the manifest, it may be that Visual Studio's manifest schema files haven't been updated to support the latest features. Your first troubleshooting step should be to replace one or more of these files 


with the latest versions. For detailed instructions, see **Manifest schema validation errors in Visual Studio projects**.

- 4. Opens the Office application host page of your Microsoft 365 tenancy in Microsoft Edge.
#### **Tip**

If for any reason, Visual Studio doesn't fully sideload the add-in and none of the fixes earlier works, you can manually sideload it. Follow the steps in **Sideload an add-in to Microsoft 365**. When you're instructed to browse to the manifest, navigate to the XML file in the folder _ProjectName_\bin\Debug\OfficeAppManifests directory.

## **Debug the add-in on the web**

The best method for debugging an add-in in Visual Studio 2022 depends on whether the add-in is running in WebView2, which is the webview control that is associated with Microsoft Edge (Chromium), or an older webview control. If your computer is using WebView2, see Use the built-in Visual Studio debugger to debug on the web. For any other webview control, see Use the browser developer tools to debug on the web. To determine which webview control is being used, see Browsers and webview controls used by Office Add-ins.

## **Tip**

In recent versions of Office, one way to identify the webview control that Office is using is through the **personality menu** on any add-in where it's available. (The personality menu isn't supported in Outlook.) Open the menu and select **Security Info**. In the **Security Info** dialog on Windows, the **Runtime** reports **Microsoft Edge**, **Microsoft Edge Legacy**, or **Internet Explorer**. The runtime isn't included on the dialog in older versions of Office.

## **Use the built-in Visual Studio debugger to debug on the web**

- 1. Set breakpoints, as needed, in the source JavaScript or TypeScript files. You can do this either before or after you start the add-in as described in the earlier section Start the add-in project on the web.


- 2. When the add-in is running, use the add-in's UI to run the code that contains your breakpoints.
## **Tip**

- Sometimes in Outlook on the web, the Visual Studio debugger doesn't attach. If you get errors by the breakpoints that indicate they won't be hit, use the browser developer tools to attach to the Visual Studio debugger: After you have pressed F5 to start debugging and Outlook on the web has opened, follow the first four steps in the **Use the browser developer tools to debug on the web**. (Use the instructions for Microsoft Edge (Chromium-based).) After you set a breakpoint in the browser tools and it's hit, execution pauses on the breakpoint in *both* the browser tools *and* in Visual Studio. This indicates that the Visual Studio debugger is attached. At this point, you can close the browser tools and add breakpoints in Visual Studio as you normally would.
- If you encounter any problems, there's more information at **[Debug a](https://learn.microsoft.com/en-us/visualstudio/javascript/debug-nodejs?view=vs-2022&preserve-view=true) [JavaScript or TypeScript app in Visual Studio](https://learn.microsoft.com/en-us/visualstudio/javascript/debug-nodejs?view=vs-2022&preserve-view=true)**.

## **Use the browser developer tools to debug on the web**

- 1. For an add-in in any host except Outlook, in the Office host application page, press F12 to open the debugging tool.
- 2. For an Outlook add-in, if the add-in's manifest is configured for a read surface, select an email message or appointment item to open it in its own window. If the add-in is configured for only a compose surface, open a new message, reply to message, or new appointment window. Ensure that the appropriate window has focus and press F12 to pen the debugging tool.
- 3. After the tool is open, launch the add-in. The exact steps vary depending on the design of your add-in. Typically, you press a button to open a task pane. In Outlook, in the toolbar at the top of the window, select the **More apps** button, and then select your add-in from the callout that opens.


| list. | OutlookWeb | N<br>Send to<br>OneNote | Share to<br>Teams | SPFx template | w blocked content |     |  |  |  |  |
|-------|------------|-------------------------|-------------------|---------------|-------------------|-----|--|--|--|--|
|       |            |                         |                   |               | 出                 | ം പ |  |  |  |  |
|       |            |                         |                   | Get add-ins - |                   |     |  |  |  |  |

- 4. Use the instructions in one of the following articles to set breakpoints and step through code. They each have a link to more detailed guidance.
	- Debug add-ins using developer tools for Edge Legacy
	- Debug add-ins using developer tools in Microsoft Edge (Chromium-based)

### **Tip**

To debug code that runs in the Office.initialize function or an Office.onReady function that runs when the add-in opens, set your breakpoints, and then close and reopen the add-in. For more information about these functions, see **Initialize your Office Add-in**.

- 5. To make changes to your code, first stop the debugging session in Visual Studio and close the Office on the web page. Make your changes, and start a new debugging session.
# **Use an existing document to debug the add-in**

If you have a document that contains test data you want to use while debugging your Excel, PowerPoint, or Word add-in, Visual Studio can be configured to open that document when you start the project. To specify an existing document to use while debugging the add-in, complete the following steps.

- 1. In **Solution Explorer**, choose the add-in project (*not* the web application project).
- 2. From the menu bar, choose **Project** > **Add Existing Item**.


- 3. In the **Add Existing Item** dialog box, locate and select the document that you want to add.
- 4. Choose the **Add** button to add the document to your project.
- 5. In **Solution Explorer**, choose the add-in project (*not* the web application project).
- 6. From the menu bar, choose **View** > **Properties Window**.
- 7. In the **Properties** window, choose the **Start Document** list, and then select the document that you added to the project. The project is now configured to start the add-in in that document.

# **Next steps**

After your add-in is working as desired, see Deploy and publish your Office Add-in to learn about the ways you can distribute the add-in to users.

# **Sideload Office Add-ins to Office on the web**

07/29/2025

When you sideload an add-in, you're able to install the add-in without first putting it in an addin catalog. This is useful when testing and developing your add-in because you can see how your add-in will appear and function.

### 7 **Note**

- This article applies to **Excel**, **OneNote**, **PowerPoint**, and **Word** add-ins. For information on sideloading **Outlook** add-ins, see the article **Sideload Outlook addins for testing**.
- This article applies to add-ins that use the add-in only manifest. For information about sideloading add-ins that use the **unified manifest for Microsoft 365**, see **Sideload Office Add-ins that use the unified manifest for Microsoft 365**.

When you sideload an add-in on the web, the add-in's manifest is stored in the browser's local storage, so if you clear the browser's cache, or switch to a different browser, you have to sideload the add-in again.

The steps to sideload an add-in on the web vary based on the following factors.

- The host application (for example, Excel, Word, Outlook)
- What tool created the add-in project (for example, Visual Studio, Yeoman generator for Office Add-ins, or neither)
- Whether you are sideloading to Office on the web with a Microsoft account or with an account in a Microsoft 365 tenant

In the following list, go to the section or article that matches your scenario. Note the first scenario in the list applies to Outlook add-ins. The remaining scenarios apply to non-Outlook add-ins.

- If you're sideloading an Outlook add-in, see the article Sideload Outlook add-ins for testing.
- If you created the add-in using the Yeoman generator for Office Add-ins, see Sideload a Yeoman-created add-in to Office on the web.


- If you created the add-in using Visual Studio, see Sideload an add-in on the web when using Visual Studio.
- For all other cases, see one of the following sections.
	- If you're sideloading to Office on the web with a Microsoft account, see Manually sideload an add-in to Office on the web.
	- If you're sideloading to Office on the web with an account in a Microsoft 365 tenant, see Sideload an add-in to Microsoft 365.

# **Sideload a Yeoman-created add-in to Office on the web**

This process is supported for **Excel**, **OneNote**, **PowerPoint**, and **Word** only. This example project assumes you're using a project created with the Yeoman generator for Office Add-ins.

- 1. Open [Office on the web](https://office.live.com/) or OneDrive. Using the **Create** option, make a document in **Excel**, **OneNote**, **PowerPoint**, or **Word**. In this new document, select **Share**, select **Copy Link**, and copy the URL.
- 2. Open a Command Prompt as an administrator. In the command line starting at the root directory of your project, run the following command. Replace "{url}" with the URL that you copied.

```
7 Note
```
If you are developing on a Mac, enclose the {url} in single quotation marks. Do *not* do this on Windows.

command line

npm run start -- web --document {url}

The following are examples.

- npm run start -- web --document https://contoso.sharepoint.com/:t:/g/EZGxP7ksiE5DuxvY638G798BpuhwluxCMfF1WZQj3V YhYQ?e=F4QM1R
- npm run start -- web --document https://1drv.ms/x/s!jkcH7spkM4EGgcZUgqthk4IK3NOypVw?e=Z6G1qp


- npm run start -- web --document https://contoso-my.sharepointdf.com/:t:/p/user/EQda453DNTpFnl1bFPhOVR0BwlrzetbXvnaRYii2lDr_oQ?e=RSccmNP
If your add-in doesn't sideload in the document, manually sideload it by following the instructions in Manually sideload add-ins to Office on the web.

- 3. The first time you use this method to sideload an add-in on the web, you'll see a dialog asking you to enable developer mode. Select the checkbox for **Enable Developer Mode now** and select **OK**.
- 4. You'll see a second dialog box, asking if you wish to register an Office Add-in manifest from your computer. Select **Yes**.
- 5. Your add-in is installed. If it has an add-in command, it should appear on either the ribbon or the context menu. If it's a task pane add-in without any add-in commands, the task pane should appear.

# **Sideload an add-in on the web when using Visual Studio**

If you're using Visual Studio to develop your add-in, press F5 to open an Office document in *desktop* Office, create a blank document, and sideload the add-in. When you want to sideload to *Office on the web*, the process to sideload is similar to manual sideloading to the web. The only difference is that you must update the value of the **SourceURL** element, and possibly other elements, in your manifest to include the full URL where the add-in is deployed.

- 1. In Visual Studio, choose **View** > **Properties Window**.
- 2. In the **Solution Explorer**, select the web project. This displays properties for the project in the **Properties** window.
- 3. In the Properties window, copy the **SSL URL**.
- 4. In the add-in project, open the manifest XML file. Be sure you're editing the source XML. For some project types, Visual Studio will open a visual view of the XML which won't work for the next step.
- 5. Search and replace all instances of **~remoteAppUrl/** with the SSL URL you just copied. You'll see several replacements depending on the project type, and the new URLs will appear similar to https://localhost:44300/Home.html .
- 6. **Save** the XML file.


- 7. In the **Solution Explorer**, open the context menu of the web project (for example, by right clicking on it) then choose **Debug** > **Start new instance**. This runs the web project without launching Office.
- 8. From Office on the web, sideload the add-in using steps described in Manually sideload an add-in to Office on the web.

# **Manually sideload an add-in to Office on the web**

This method doesn't use the command line and can be accomplished using commands only within the host application (such as Excel).

- 1. Open [Office on the web](https://office.com/) . Open a document in **Excel**, **OneNote**, **PowerPoint**, or **Word**.
- 2. Select **Home** > **Add-ins**, then select **More Settings**.
- 3. On the **Office Add-ins** dialog, select **Upload My Add-in**.
- 4. **Browse** to the add-in manifest file, and then select **Upload**.

| Upload Add-in                                         |        |        |
|-------------------------------------------------------|--------|--------|
| This feature is for developers to test their add-ins. |        |        |
| Choose your add-in manifest                           |        |        |
|                                                       |        | Browse |
|                                                       |        |        |
|                                                       | Upload | Cancel |

- 5. Verify that your add-in is installed. For example, if it has an add-in command, it should appear on either the ribbon or the context menu. If it's a task pane add-in that has no add-in commands, the task pane should appear.
#### 7 **Note**

To test your Office Add-in with EdgeHTML (Microsoft Edge Legacy), an additional configuration step is required. In a Windows Command Prompt, run the following line: npx office-addin-dev-settings appcontainer EdgeWebView --loopback --yes . This isn't required when Office is using the Chromium-based Edge WebView2. For more information, see **Browsers and webview controls used by Office Add-ins**.


#### ) **Important**

The office-addin-dev-settings tool is not supported on Mac.

### **Sideload an add-in to Microsoft 365**

- 1. Sign in to your Microsoft 365 account.
- 2. Open the App Launcher on the left end of the toolbar and select **Excel**, **OneNote**, **PowerPoint**, or **Word**, and then create a new document.
- 3. Follow steps 2 5 of the section Manually sideload an add-in to Office on the web.

### **Remove a sideloaded add-in**

If you ran the npm start command and your add-in was automatically sideloaded, then run npm stop when you're ready to stop the dev server and uninstall your add-in.

Otherwise, to remove an add-in sideloaded to Office on the web, simply clear your browser's cache. If you make changes to your add-in's manifest (for example, update file names of icons or text of add-in commands), you may need to clear your browser's cache and then re-sideload the add-in using the updated manifest. Doing so allows Office on the web to render the add-in as it's described by the updated manifest.

# **See also**

- Sideload Office Add-ins on Mac
- Sideload Office Add-ins on iPad
- Sideload Outlook add-ins for testing
- Clear the Office cache


# **Debug add-ins in Office on the web**

Article • 12/20/2023

This article describes how to use Office on the web to debug your add-ins. Use this technique:

- To debug add-ins on a computer that isn't running Windows or the Office desktop client—for example, if you're developing on a Mac or Linux.
- As an alternative debugging process if you can't, or don't wish to, debug in an IDE, such as Visual Studio or Visual Studio Code.

This article assumes that you have an add-in project that needs to be debugged. If you just want to practice debugging on the web, create a new project using one of the quick starts for specific Office applications, such as this quick start for Word.

# **Debug your add-in**

To debug your add-in by using Office on the web:

- 1. Run the project on localhost and sideload it to a document in Office on the web. For detailed sideloading instructions, see Manually sideload Office Add-ins on the web.
- 2. Open the browser's developer tools. This is usually done by pressing F12 . Open the debugger tool and use it to set breakpoints and watch variables. For detailed help in using your browser's tool, see one of the following:
	- [Firefox](https://firefox-source-docs.mozilla.org/devtools-user/index.html)
	- [Safari](https://support.apple.com/guide/safari/use-the-developer-tools-in-the-develop-menu-sfri20948/mac)
	- Debug add-ins using developer tools in Microsoft Edge (Chromium-based)
	- Debug add-ins using developer tools for Edge Legacy

### 7 **Note**

- Office on the web won't open in Internet Explorer.
- The new Outlook on Window desktop client (preview) doesn't support the context menu or the keyboard shortcut to access the Microsoft Edge developer tools. Instead, you must run olk.exe --devtools from a command prompt. For more information, see the "Debug your add-in" section of **Develop Outlook add-ins for the new Outlook on Windows**.


# **Potential issues**

The following are some issues that you might encounter as you debug.

- Some JavaScript errors that you see might originate from Office on the web.
- The browser might show an invalid certificate error that you'll need to bypass. The process for doing this varies with the browser and the various browsers' UIs for doing this change periodically. You should search the browser's help or search online for instructions. (For example, search for "Microsoft Edge invalid certificate warning".) Most browsers will have a link on the warning page that enables you to click through to the add-in page. For example, Microsoft Edge has a "Go on to the webpage (Not recommended)" link. But you'll usually have to go through this link every time the add-in reloads. For a longer lasting bypass, see the help as suggested.
- If you set breakpoints in your code, Office on the web might throw an error indicating that it's unable to save.

# **See also**

- Best practices for developing Office Add-ins
- Troubleshoot user errors with Office Add-ins

# **Test and debug Office Add-ins on a nonlocal server**

Article • 05/19/2025

When you've completed development and testing on a localhost and want to stage and test the add-in from a non-local server or cloud account, you can use the tool [office-addin](https://github.com/OfficeDev/Office-Addin-Scripts/tree/master/packages/office-addin-debugging)[debugging](https://github.com/OfficeDev/Office-Addin-Scripts/tree/master/packages/office-addin-debugging) for any Node.js-based add-in project. (The tool isn't supported in projects created with Visual Studio.)

#### 7 **Note**

If you're working on a Windows computer, you may have another option for non-local testing. See **Sideload Office Add-ins for testing from a network share**.

# **Projects created with Microsoft 365 Agents Toolkit or the Office Yeoman Generator (Yo Office)**

If your project was created with Agents Toolkit or Office Yeoman Generator (Yo Office), then the office-addin-debugging tool is already installed and your package.json file has start and stop scripts that invoke the tool. To use it for non-local testing, update the domain part of the URLs in your manifest to point to your staging server (or CDN as needed). Then run npm run start at the command line (or Visual Studio Code TERMINAL) to sideload the add-in for testing and debugging.

#### ) **Important**

The office-addin-debugging tool registers the add-in in the Windows registry or a special folder on a Mac. For an Outlook add-in, it also registers the add-in in Exchange. To avoid subtle bugs when developing, always end a testing session by running npm run stop to ensure that these registrations are removed and that the server process is fully stopped. *Manually closing the server, the command line window (or TERMINAL), Visual Studio Code, or the Office application doesn't remove these registrations.*

# **Other projects**

If your project wasn't created with Agents Toolkit or Yo Office, run the tool with npx in the root of the project. Invoke it with its start command followed by the relative path to the manifest.


The following is an example.

command line

npx office-addin-debugging start manifest.json

This command sideloads the add-in for testing and debugging. The tool also works with an add-in only manifest.

There are many options for the start command. For details, see the README for the tool at [office-addin-debugging](https://github.com/OfficeDev/Office-Addin-Scripts/tree/master/packages/office-addin-debugging) .

### ) **Important**

The office-addin-debugging tool registers the add-in in the Windows registry or a special folder on a Mac. For an Outlook add-in, it also registers the add-in in Exchange. To avoid subtle bugs when developing, always end a testing session by running npx office-addindebugging stop to ensure that these registrations are removed and that the server process is fully stopped. *Manually closing the server, the command line window (or TERMINAL), Visual Studio Code, or the Office application doesn't remove these registrations.* If you used the --prod option with the start command, use the same option with the stop command.


# **Clear the Office cache**

#### 06/25/2025

The Office cache stores resources and data used by Office Add-ins. By accessing stored resources, an add-in's performance is improved as it avoids redownloading these resources when needed.

You should clear the Office cache in the following scenarios.

- When you want to remove an add-in that you've previously sideloaded on Windows, Mac, or iOS.
- When you've updated the manifest (for example, to update the file names of icons or text of add-in commands). This ensures that you're using the latest version of the add-in.

### **Tip**

For add-ins that implement a task pane, if you only want the sideloaded add-in to reflect recent changes to its HTML or JavaScript source files, you shouldn't need to clear the cache. Instead, put focus in the add-in's task pane (by selecting anywhere within the task pane). Then, select Ctrl + F5 to reload the add-in.

- When you want to resolve issues or errors when running the add-in.
#### 7 **Note**

To remove a sideloaded add-in from Excel, OneNote, PowerPoint, or Word on the web, see **Sideload Office Add-ins in Office on the web for testing: Remove a sideloaded add-in**.

To remove a sideloaded add-in from Outlook on the web, see **Sideload Outlook add-ins for testing**.

### U **Caution**

When you clear the Office cache, clear it completely. Don't delete individual manifest files. This can cause all add-ins to stop loading.

# **Types of caches**

The Office cache can refer to either the web cache or the Wef cache.


- The **web cache** temporarily stores web-based resources and data used by an individual Office Add-in.
- The **Wef cache** locally stores resources and data for all installed Office Add-ins.

The following table outlines which Office cache types can be cleared on different platforms. It also provides links to instructions on how to clear a specific cache.

#### ノ **Expand table**

| Platform | Types of caches to clear                                                                                   | Options to clear the cache                                                                                           |
|----------|------------------------------------------------------------------------------------------------------------|----------------------------------------------------------------------------------------------------------------------|
| Windows  | Both the web and Wef caches. There's currently no<br>option to clear one cache without clearing the other. | Automatically clear the cache<br>Manually clear the cache<br>Use the Microsoft Edge<br>developer tools on Windows 10 |
| Mac      | Web<br>Both web and Wef caches                                                                             | Web: Use the personality menu<br>to clear the web cache<br>Web and Wef: Clear the web<br>and Wef caches on Mac       |
| iOS      | Web                                                                                                        | Use JavaScript to clear the<br>cache on iOS                                                                          |

### **Clear the Office cache on Windows**

Depending on your Office host and operating system, you can automatically or manually clear both the web and Wef caches on a Windows computer.

### ) **Important**

On Windows, the automatic and manual options clear both the web and Wef caches. There's currently no option to clear one cache without clearing the other.

### **Automatically clear the cache**

#### 7 **Note**

The automatic option is only supported for Excel, PowerPoint, and Word. Outlook only supports the **manual option**.


This method is recommended for add-in development computers. If your Office on Windows version is 2108 or later, the following steps configure the Office cache to be cleared the next time Office is reopened.

- 1. From the ribbon of Excel, PowerPoint, or Word, navigate to **File** > **Options** > **Trust Center** > **Trust Center Settings** > **Trusted Add-in Catalogs**.
- 2. Select the **Next time Office starts, clear all previously-started web add-ins cache** checkbox.
- 3. Select **OK**.
- 4. Restart Excel, PowerPoint, or Word.

### **Manually clear the cache**

### **Manually clear the cache in Excel, Word, and PowerPoint**

To remove all sideloaded add-ins from Excel, Word, and PowerPoint, delete the contents of the following folder.

%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\

If the following folder exists, delete its contents, too.

```
%userprofile%\AppData\Local\Packages\Microsoft.Win32WebViewHost_cw5n1h2txyewy\AC\#
!123\INetCache\
```
### **Manually clear the cache in Outlook**

Before attempting to clear the cache in Outlook, first try to remove the sideloaded add-in using the steps outlined in Sideload Outlook add-ins for testing.

If this add-in removal doesn't work, then delete the contents of the Wef folder as noted for Excel, Word, and PowerPoint in Manually clear the cache in Excel, Word, and PowerPoint.

If your Outlook add-in uses the Unified manifest for Microsoft 365, also delete the following folder.


To clear the cache in [new Outlook on Windows](https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627) , perform the following steps.

- 1. Close the Outlook client if it's open.
- 2. From a command line, run the following:

command line

olk.exe --devtools

This opens the new Outlook on Windows client and an instance of the Microsoft Edge DevTools.

- 3. In the Microsoft Edge DevTools window, select the **Network** tab.
- 4. Select and hold (or right-click) anywhere in the **Requests** table. Then, select **Clear browser cache**.

### **Use the Microsoft Edge developer tools on Windows 10**

To clear the Office cache on Windows 10 when the add-in is running in Microsoft Edge, use the Microsoft Edge DevTools.

#### 7 **Note**

To clear the Office cache using the following steps, your add-in must have a task pane. If your add-in is a UI-less add-in -- for example, one that uses the **on-send** feature -- you'll need to add a task pane to your add-in that uses the same domain for **[SourceLocation](https://learn.microsoft.com/en-us/javascript/api/manifest/sourcelocation)**, before you can use the following steps to clear the cache.

- 1. Install the [Microsoft Edge DevTools](https://apps.microsoft.com/detail/9mzbfrmz0mnj) .
- 2. Open your add-in in the Office client.
- 3. Run the Microsoft Edge DevTools.
- 4. In the Microsoft Edge DevTools, open the **Local** tab. Your add-in will be listed by its name.
- 5. Select the add-in name to attach the debugger to your add-in. A new Microsoft Edge DevTools window will open when the debugger attaches to your add-in.


- 6. On the **Network** tab of the new window, select **Clear cache**.

| ■<br>×<br>taskpane.html - Microsoft Edge DevTools Preview                     |                 |          |    |        |             |              |          |              |                 |           |                     |                                                           |
|-------------------------------------------------------------------------------|-----------------|----------|----|--------|-------------|--------------|----------|--------------|-----------------|-----------|---------------------|-----------------------------------------------------------|
| Elements                                                                      | Console × 1     | Debugger |    |        | Network (P) | Performance  | Memory   | Storage      | Service Workers | Emulation |                     | 2 8 ?                                                     |
|                                                                               | ్రామం<br>H<br>일 | प्र      | 20 | ুটি    | 세           | Content type |          |              |                 |           |                     | Find in files (Ctrl+F)                                    |
| Name                                                                          |                 |          |    |        | Protocol    | Method       | Result   | Content type | Received        | Time      | Initiator           | Bodv<br>Cookies<br>Headers<br>Parameters<br>Timings       |
|                                                                               |                 |          |    |        |             |              |          |              |                 |           |                     | Request URL: https://static2.sharepointonline.com/files/f |
| fabric.min.css<br>https://static2.sharepointonline.com/files/fabric/office-ui |                 |          |    | HTTP/2 | GET         | 200          | text/css | 24.73 KB     | 284.27 ms       |           | Request Method: GET |                                                           |
| fabric.min.css                                                                |                 |          |    |        | HTTP/2      | GET          | 200      | text/css     | 24.73 KB        | 271.33 ms |                     | Status Code: 200 /                                        |
| bakan llakukia ) ahanamainkantina anne thilan thalamalakhina iii              |                 |          |    |        |             |              |          |              |                 |           |                     | 1 Daminat Llandana                                        |

- 7. If completing these steps doesn't produce the desired result, try selecting **Always refresh from server**.

| taskpane.html - Microsoft Edge DevTools Preview     |                                                             |          |           |          |             |        |              |                 |           |                        | ■<br>×                                                                           |
|-----------------------------------------------------|-------------------------------------------------------------|----------|-----------|----------|-------------|--------|--------------|-----------------|-----------|------------------------|----------------------------------------------------------------------------------|
| Elements                                            | Console × 1                                                 | Debugger | Network ( |          | Performance | Memory | Storage      | Service Workers | Emulation |                        | 2 8                                                                              |
| प्र<br>发<br>简<br>U<br>10<br>는<br>Content type<br>ST |                                                             |          |           |          |             |        |              |                 |           | Find in files (Ctrl+F) |                                                                                  |
| Name                                                |                                                             |          |           | Protocol | Method      | Result | Content type | Received        | Time      | Initiator              | Headers<br>Bodv<br>Parameters<br>Cookies<br>Timinas                              |
| fabric.min.css                                      | https://static2.sharepointonline.com/files/fabric/office-ui |          |           | HTTP/2   | GET         | 200    | text/css     | 24.73 KB        | 284.27 ms |                        | Request URL: https://static2.sharepointonline.com/files/f<br>Request Method: GET |
| fabric.min.css                                      |                                                             |          |           | HTTP/2   | GET         | 200    | text/css     | 24.73 KB        | 271.33 ms |                        | Status Code: 200 /<br>1 Daminat Liandana                                         |

## **Clear the Office cache on Mac**

You can choose to clear the web or both the web and Wef caches on Mac.

### **Clear the web cache**

Normally, the web cache is cleared by reloading the add-in. If more than one add-in exists in the same document, the process of automatically clearing the cache on reload might not be reliable.

To clear the web cache in Excel, PowerPoint, and Word, use the personality menu of any task pane add-in.

7 **Note**

- The personality menu in task panes is only supported in Excel, PowerPoint, and Word. Because it isn't supported in Outlook, you must use the **option to clear both the web and Wef caches instead**.
- The personality menu is only shown in macOS Version 10.13.6 or later.

From the add-in's task pane, choose the personality menu. Then, choose **Clear Web Cache**.


### **Clear the web and Wef caches**

To clear both the web and Wef caches on Mac, delete the contents of the

~/Library/Containers/com.Microsoft.OsfWebHost/Data/ and

~/Library/Containers/com.microsoft.{host}/Data/Documents/wef folders. Replace {host} with the Office application, such as Excel .

### **Tip**

Use the terminal or Finder to search for the specified folders. To look for these folders via Finder, you must set Finder to show hidden files. Finder displays the folders inside the **Containers** directory by product name, such as **Microsoft Excel** instead of **com.microsoft.Excel**.

Deleting the contents of the ~/Library/Containers/com.microsoft.{host}/Data/Documents/wef folder removes all sideloaded add-ins from an application.

### 7 **Note**

If the ~/Library/Containers/com.Microsoft.OsfWebHost/Data/ folder doesn't exist, check for the following folders via terminal or Finder. If found, delete the contents of each folder.

- ~/Library/Containers/com.microsoft.{host}/Data/Library/Caches/ where {host} is the Office application (e.g., Excel )
- ~/Library/Containers/com.microsoft.{host}/Data/Library/Application Support/Microsoft/Office/16.0/Wef/ where {host} is the Office application (e.g., Excel )


- ~/Library/Containers/com.microsoft.Office365ServiceV2/Data/Caches/com.microsoft .Office365ServiceV2/
- ~/Library/Containers/com.microsoft.Office365ServiceV2/Data/Library/Caches/com.m icrosoft.Office365ServiceV2/

# **Clear the Office cache on iOS**

To clear the web cache on iOS, call window.location.reload(true) from JavaScript in the addin. This forces the add-in to reload. Alternatively, reinstall Office.

## **See also**

- Troubleshoot development errors with Office Add-ins
- Debug add-ins using developer tools for Internet Explorer
- Debug add-ins using developer tools for Edge Legacy
- Debug add-ins using developer tools in Microsoft Edge (Chromium-based)
- Debug your add-in with runtime logging
- Sideload Office Add-ins for testing
- Office Add-ins manifest
- Validate an Office Add-in's manifest
- Uninstall add-ins under development

# **Uninstall add-ins under development**

Article • 05/19/2025

Incompletely removed add-ins can leave artifacts on your computer, such as custom ribbon buttons or registry entries, during development. In this article, we call these "ghost add-ins".

Outlook add-ins also might add these artifacts to other computers when you sign into Outlook on them with the same ID as you used to develop the add-in.

#### ) **Important**

When you sign into Outlook, it downloads from Exchange, and sideloads, all the Outlook add-in manifests that are associated with your ID, *including add-ins that you are developing on a different computer using the same ID*. For example, any custom ribbon buttons defined in the manifest will appear for the add-in.

If the URLs in the manifest point to a non-localhost server and that server is running and accessible to the non-development computer, then Outlook caches the add-in's files in the local file system and the add-in usually runs normally on the computer. Otherwise, the add-in doesn't function, but visible parts of it, such as custom ribbon buttons appear. They have the labels defined in the manifest. The add-in's button icons also appear if they were ever cached locally on the non-development computer and the cache was never cleared. Icon files aren't stored with Exchange, so if they were never cached on the nondevelopment computer (or the cache has been cleared), then the buttons have default icons.

Until the add-in's registration is removed from Exchange, the add-in will continue to appear. See **Remove a ghost add-in** for information about removing the registration in Exchange.

This article provides some guidance to minimize the chance of these problems and to resolve them if they occur.

# **Prevent the problems**

When an add-in is sideloaded, several things happen:

- A web server, usually on localhost, is started to serve the add-in's files (such as the HTML, CSS, and JavaScript files).
- These same files are cached on your development computer.


- The add-in is registered with the development computer. The registration is done with Registry entries on a Windows computer or with certain files saved to the file system on a Mac.
- Most tools for sideloading add-ins automatically open the Office application that the add-in targets. The tools also populate the application with any custom ribbon buttons or context menu items that are defined in the add-in's manifest.
- For an Outlook add-in, the add-in's manifest is registered with the Exchange service.

### **Use your tool's uninstall facility**

To prevent ghost add-ins, end every testing, debugging, and sideloading session by using the uninstall (also called unacquire) option that is provided by the tool that you used to start the session. Doing this reverses the effects of sideloading, as stated earlier in this article.

The following list identifies, for each tool, how to uninstall but doesn't describe the procedures or syntax in detail. *Be sure to use the links to get complete instructions.*

#### 7 **Note**

Some of these tools don't close the Office application that opened automatically. In that case, close the application manually immediately after ending the session.

- **Yeoman generator for Office Add-ins (Yo Office)**: Use the npm stop script at the same command line where you started the session with npm start . For more information, see the various articles in the **Get started** and **Quick starts** sections and Remove a sideloaded add-in.
- **Microsoft 365 Agents Toolkit for Visual Studio Code**: Select **Run** | **Stop Debugging** in Visual Studio Code. For more information, see the last step of Create an Outlook Add-in project which also applies to non-Outlook add-ins.
- **Office Add-in Development Kit for Visual Studio Code**: With the Office Add-in Development Kit extension open, select **Stop Previewing Your Office Add-in**. For more information, see [Stop testing your add-in](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/development-kit-overview?tabs=vscode#stop-testing-your-office-add-in).
- **office-addin-debugging tool**: Use the office-addin-debugging stop command at the same command line where you started the session with office-addin-debugging start . For more information, see Sideload with the Office-Addin-Debugging tool.
- **Microsoft 365 Agents Toolkit CLI**: Use the atk uninstall command at the same command line where you started the session with atk install . For more information, see Sideload with Microsoft 365 Agents Toolkit CLI.
- **Visual Studio**: Select **DEBUG** | **Stop debugging** in the menu, or press Shift + F5 , or click the square red "stop" button on the debugging bar. Alternatively, closing the Office


application also stops the session and uninstalls the add-in. For more information, see [First look at the Visual Studio debugger.](https://learn.microsoft.com/en-us/visualstudio/debugger/debugger-feature-tour)

# **Remove a ghost add-in**

To remove a ghost add-in, you need to remove the artifacts that were created when it was last sideloaded, remove its local registration, and for Outlook add-ins remove its registration in Exchange.

### **Tip**

There's a fast way to remove a ghost add-in on Windows computers if the add-in was installed with the Teams Toolkit CLI. Try this first, and if it works, you can skip the remainder of this section.

- 1. Obtain the add-in's title ID from the Registry key
**HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\Outlook SideloadManifestPath\TitleId**. (The string "Outlook" is in the key name for historical reasons, but it applies to any add-in installed with the Agents Toolkit CLI.)

- 2. Run the following command in a command prompt, bash shell, or terminal. Replace " {title ID}" with the title ID of the add-in including the "U_" prefix; for example, U_90d141c6-cf4f-40ee-b714-9df9ea593f39 .
command line

atk uninstall --mode title-id --title-id {title ID} --interactive false

The process for removing the add-in varies depending on whether the add-in is for Outlook or some other Office application.

#### 7 **Note**

In the **unified manifest for Microsoft 365**, an add-in can be configured to support Outlook and one or more other Office applications; that is, there's more than one member of the **["extensions.requirements.scopes"](https://learn.microsoft.com/en-us/microsoft-365/extensibility/schema/requirements-extension-element#scopes)** array in the manifest and one of the members is "mail" (or the "extensions.requirements.scopes" property isn't present). Treat an add-in that is configured in this way as an Outlook add-in.


If the ghost add-in is not an Outlook add-in, skip to the section Remove the add-in artifacts.

### **Remove the Exchange registration of a ghost Outlook add-in**

- 1. Log into Outlook with the same ID you used when you sideloaded the add-in.
- 2. Open PowerShell as an Administrator.
- 3. Run the following commands. Answer "Yes" to all confirmation prompts.

```
PowerShell
Install-Module -Name ExchangeOnlineManagement -RequiredVersion 3.4.0
Set-ExecutionPolicy RemoteSigned
Connect-ExchangeOnline
```
7 **Note**

If the Connect-ExchangeOnline command returns the error "ActiveX control '8856f961-340a-11d0-a96b-00c04fd705a2' cannot be instantiated because the current thread is not in a single-threaded apartment", just run the command a second time. This is a well-known bug.

- 4. Run the following command. Answer "Yes" to all confirmation prompts.
A list of the add-ins installed on Outlook displays. These will include built-in Microsoft add-ins and add-ins you have installed. Any ghost Outlook add-ins will also be listed.

- 5. Find the ghost add-in in the list. If it was created with Yo Office or another Microsoft tool, it probably has the name "Contoso Task Pane Add-in".
- 6. Copy the App ID (a GUID) of the add-in. You need it for later steps.
- 7. Run the command Remove-App -Identity {{The GUID OF YOUR ADD-IN HERE}} (e.g., Remove-App -Identity 26ead0cb-10dd-4ba2-86c6-4db111876652 ). This command removes the addin from Exchange.

2 **Warning**


The removal of the registration needs to propagate to all Exchange servers. Wait at least three hours before continuing with the next step.

- 8. Continue with the section Remove the add-in artifacts.
### **Remove the add-in artifacts**

#### ) **Important**

Carry out this procedure on all devices on which you have had the add-in sideloaded.

- 1. Log out from all Office applications and then close them all, including Outlook.
- 2. Clear the Office cache. If the ghost add-in supports Outlook, use Clear the cache in Outlook manually.
- 3. Continue with the section Remove the local registration.

### **Remove the local registration**

### ) **Important**

Carry out this procedure on all computers on which you have had the add-in sideloaded.

- 1. Delete the local registration of the ghost add-in. The process varies depending on the operating system.
Windows

- a. Open the **Registry Editor**.
- b. Navigate to

**Computer\HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\WEF\Develop er**. This key lists the add-ins that are currently sideloaded, or were sideloaded in the past and weren't fully uninstalled. The **Data** value for each entry is the path to the add-in's manifest. The **Name** value varies depending on which version of which tool was used to create and sideload the add-in. If Visual Studio was used, the name is typically is also the path to the manifest. For other tools, the name is typically the add-in's ID. When an Office application launches, it reloads all add-ins listed in this key (that support the Office application). Reloading may have no practical or discernable effect if the add-in's artifacts have been deleted from the


cache, or the manifest no longer exists at the path, or the add-in's files aren't being served by a server.

Find the entry for the ghost add-in and delete it. If it is an Outlook add-in, then you have the ID from removing the Exchange registration. You can also use the path in the **Data** column to find the manifest to help identify the add-in the entry refers to and read the ID from the manifest. If any manifests listed in the **Data** column no longer exist at the specified path, then delete the entries for those manifests.

| Registry Editor<br>File<br>Edit<br>View<br>Favorites Help                                                                                              |                                                                                                                                                                                                                       |                                                                     |                                                                                                                                                                                         |
|--------------------------------------------------------------------------------------------------------------------------------------------------------|-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|---------------------------------------------------------------------|-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| Computer\HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\WEF\Developer                                                                                |                                                                                                                                                                                                                       |                                                                     |                                                                                                                                                                                         |
| I WEF<br>><br>AddinLifecycle<br>AllowedAppDomains<br>AutoInstallAddins<br>Cache<br>Developer<br>LastUpdate<br>Most Popular Addins<br>Dreinstalled Anns | Name<br>ab] (Default)<br>ab 1b8f5990-aaf7-43fc-88b1-f42cad5620da<br>ab 26ead0cb-10dd-4ba2-86c6-4db111876652<br>ap 688ee6d5-c218-42d2-bcb0-071eef71b0bf<br>ap c33782d0-b862-4ff5-b4f0-6259b4863a70<br>no RefreshAddins | Type<br>REG SZ<br>REG SZ<br>REG SZ<br>REG_SZ<br>REG_SZ<br>REG DWORD | Data<br>(value not set)<br>D:\gh\DevKitTest\manifest.xml<br>D:\gh\test4bug\manifest.xml<br>D:\gh\UllessCustomFunction\manifest.xml<br>D:\gh\Test4json3\manifest.json<br>0x000000000 (0) |

- c. Expand the **... Developer** node in the registry tree. Look for a subkey whose name is the same ghost add-in's ID. If it is there, delete it.
	-
- d. Navigate to **Computer\HKEY_USERS\**

**{SID}\Software\Microsoft\Office\16.0\WEF\Developer**, where **{SID}** is the [SID](https://learn.microsoft.com/en-us/windows-server/identity/ad-ds/manage/understand-security-identifiers) of the user you were signed in with when you sideloaded the add-in, and repeat the preceding two steps.

- e. Navigate to
**Computer\HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Common\Cus tomUIValidationCache**. In the **Name** column, find all the entries that begin with the add-in's ID (a GUID) and delete them. Then navigate to **Computer\HKEY_USERS\**

**{SID}\Software\Microsoft\Office\16.0\Common\CustomUIValidationCache**, where **{SID}** is the SID of the user you were signed in with when you sideloaded the add-in, and repeat the process.


- 2. If you are removing an Outlook add-in, continue with the section Test for removal of Outlook add-ins.
### **Test for removal of Outlook add-ins**

Open Outlook with the same identity you used when you created the add-in. If artifacts from the add-in (such as custom ribbon buttons) reappear after a few minutes or if event handlers from the add-in seem to be active, then the removal of the add-in's registration from Exchange hasn't propagated to all Exchange servers. Wait at least three hours and then repeat the procedures in the sections Remove the add-in artifacts and Remove the local registration on the computer where you observed the artifacts.

# **See also**

- Troubleshoot development errors with Office Add-ins
- Clear the Office cache
- The PowerShell reference for [Install-Module,](https://learn.microsoft.com/en-us/powershell/module/powershellget/install-module) [Set-ExecutionPolicy,](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.security/set-executionpolicy) [Connect-](https://learn.microsoft.com/en-us/powershell/exchange/connect-to-exchange-online-powershell)[ExchangeOnline](https://learn.microsoft.com/en-us/powershell/exchange/connect-to-exchange-online-powershell), and [Get-App](https://learn.microsoft.com/en-us/powershell/module/exchange/get-app).

# **Debug your add-in with runtime logging**

Article • 11/12/2024

You can use runtime logging to debug your add-in's manifest as well as several installation errors. This feature can help you identify and fix issues with your manifest that are not detected by XSD schema validation, such as a mismatch between resource IDs. Runtime logging is particularly useful for debugging add-ins that implement add-in commands and Excel custom functions.

#### 7 **Note**

The runtime logging feature is currently available for Office 2016 or later on desktop.

#### ) **Important**

Runtime Logging affects performance. Turn it on only when you need to debug issues with your add-in manifest.

## **Use runtime logging from the command line**

Enabling runtime logging from the command line is the fastest way to use this logging tool.

#### ) **Important**

The office-addin-dev-settings tool is not supported on Mac. See the section **Runtime logging on Mac** for Mac-specific instructions.

- To enable runtime logging:
command line

```
npx office-addin-dev-settings runtime-log --enable
```
- To enable runtime logging only for a specific file, use the same command with a filename:


```
command line
```
npx office-addin-dev-settings runtime-log --enable [filename.txt]

- To disable runtime logging:
command line

npx office-addin-dev-settings runtime-log --disable

- To display whether runtime logging is enabled:
command line

```
npx office-addin-dev-settings runtime-log
```
- To display help within the command line for runtime logging:
command line

```
npx office-addin-dev-settings runtime-log --help
```
## **Runtime logging on Mac**

- 1. Make sure that you are running Office 2016 desktop build **16.27.19071500** or later.
- 2. Open **Terminal** and set a runtime logging preference by using the defaults command:

```
command line
defaults write <bundle id> CEFRuntimeLoggingFile -string <file_name>
```
<bundle id> identifies which the host for which to enable runtime logging. <file_name> is the name of the text file to which the log will be written.

Set <bundle id> to one of the following values to enable runtime logging for the corresponding application.

- com.microsoft.Word
- com.microsoft.Excel
- com.microsoft.Powerpoint


- com.microsoft.Outlook
The following example enables runtime logging for Word and then opens the log file.

command line

```
defaults write com.microsoft.Word CEFRuntimeLoggingFile -string
"runtime_logs.txt"
open ~/library/Containers/com.microsoft.Word/Data/runtime_logs.txt
```
#### 7 **Note**

You'll need to restart Office after running the defaults command to enable runtime logging.

To turn off runtime logging, use the defaults delete command:

command line

defaults delete <bundle id> CEFRuntimeLoggingFile

The following example will turn off runtime logging for Word.

command line

defaults delete com.microsoft.Word CEFRuntimeLoggingFile

### **Use runtime logging to troubleshoot issues with your manifest**

To use runtime logging to troubleshoot issues loading an add-in:

- 1. Sideload your add-in for testing.
#### 7 **Note**

We recommend that you sideload only the add-in that you are testing to minimize the number of messages in the log file.

- 2. If nothing happens and you don't see your add-in (and it's not appearing in the add-ins dialog box), open the log file.


- 3. Search the log file for your add-in ID, which you define in your manifest. In the log file, this ID is labeled SolutionId .
### **Known issues with runtime logging**

You might see messages in the log file that are confusing or that are classified incorrectly. For example:

- The message Medium Current host not in add-in's host list followed by Unexpected Parsed manifest targeting different host is incorrectly classified as an error.
- If you see the message Unexpected Add-in is missing required manifest fields DisplayName and it doesn't contain a SolutionId, the error is most likely not related to the add-in you are debugging.
- Any Monitorable messages are expected errors from a system point of view. Sometimes they indicate an issue with your manifest, such as a misspelled element that was skipped but didn't cause the manifest to fail.

- Office Add-ins manifest
- Validate an Office Add-in's manifest
- Clear the Office cache
- Sideload Office Add-ins for testing
- Debug add-ins using developer tools for Internet Explorer
- Debug add-ins using developer tools for Edge Legacy
- Debug add-ins using developer tools in Microsoft Edge (Chromium-based)
- Runtimes in Office Add-ins


# **Debug a function command with a nonshared runtime**

06/17/2025

#### ) **Important**

If your add-in is **configured to use a shared runtime**, debug the code behind the function command just as you would the code behind a task pane. See **Debug Office Add-ins** and note that a function command in an add-in with a **shared runtime** is *not* a special case as described in that article.

#### 7 **Note**

This article assumes that you're familiar with **function commands**.

Function commands don't have a UI, so a debugger can't be attached to the process in which the function runs on desktop Office. (Outlook add-ins being developed on Windows are an exception to this. See Debug function commands in Outlook add-ins on Windows later in this article.) So function commands, in add-ins with a non-shared runtime, must be debugged in Office on the web where the function runs in the overall browser process. Use the following steps.

- 1. Sideload the add-in in Office on the web, and then select the button or menu item that runs the function command. This is necessary to load the code file for the function command.
- 2. Open the browser's developer tools. This is usually done by pressing F12 . The debugger in the tools attaches to the browser process.
- 3. Apply breakpoints to the code as needed for the function command.
- 4. Rerun the function command. The process stops on your breakpoints.

#### **Tip**

For more detailed information, see **Debug add-ins in Office on the web**.

### **Debug function commands in Outlook add-ins on Windows**


If your development computer is Windows, there's a way that you can debug a function command on Outlook desktop. See Debug function commands in Outlook add-ins.

- Runtimes in Office Add-ins


# **Debug the initialize and onReady functions**

06/17/2025

#### 7 **Note**

This article assumes that you're familiar with **Initialize your Office Add-in**.

The paradox of debugging the [Office.initialize](https://learn.microsoft.com/en-us/javascript/api/office#office-office-initialize-function(1)) and [Office.onReady](https://learn.microsoft.com/en-us/javascript/api/office#office-office-onready-function(1)) functions is that a debugger can only attach to a process that's running, but these functions run immediately as the add-in's runtime process starts up, before a debugger can attach. In most situations, restarting the addin after a debugger is attached doesn't help because restarting the add-in closes the original runtime process *and the attached debugger* and starts a new process that has no debugger attached.

Fortunately, there's an exception. You can debug these functions using Office on the web, with the following steps.

- 1. Sideload and run the add-in in Office on the web. This is usually done by opening an addin's task pane or running a function command. *The add-in runs in the overall browser process, not a separate process as it would in desktop Office.*
- 2. Open the browser's developer tools. This is usually done by pressing F12 . The debugger in the tools attaches to the browser process.
- 3. Apply breakpoints as needed to the code in the Office.initialize or Office.onReady function.
- 4. *Relaunch the add-in's task pane or the function command* just as you did in step 1. This action does *not* close the browser process or the debugger. The Office.initialize or Office.onReady function runs again and processing stops on your breakpoints.

#### **Tip**

For more detailed information, see **Debug add-ins in Office on the web**.

- Runtimes in Office Add-ins


# **Error handling with the application-specific JavaScript APIs**

06/23/2025

When you build an add-in using the application-specific Office JavaScript APIs, be sure to include error handling logic to account for runtime errors. Doing so is critical, due to the asynchronous nature of the APIs.

## **Best practices**

In our [code samples](https://github.com/OfficeDev/Office-Add-in-samples) and Script Lab snippets, you'll notice that every call to Excel.run , PowerPoint.run , or Word.run is accompanied by a catch statement to catch any errors. We recommend that you use the same pattern when you build an add-in using the applicationspecific APIs.

JavaScript

```
$("#run").on("click", () => tryCatch(run));
async function run() {
 await Excel.run(async (context) => {
 // Add your Excel JavaScript API calls here.
 // Await the completion of context.sync() before continuing.
 await context.sync();
 console.log("Finished!");
 });
}
/** Default helper for invoking an action and handling errors. */
async function tryCatch(callback) {
 try {
 await callback();
 } catch (error) {
 // Note: In a production add-in, you'd want to notify the user through your
add-in's UI.
 console.error(error);
 }
}
```
### **API errors**


When an Office JavaScript API request doesn't run successfully, the API returns an error object that contains the following properties.

- **code**: The code property of an error message contains a string that is part of OfficeExtension.ErrorCodes or {application}.ErrorCodes where *{application}* represents Excel, PowerPoint, or Word. For example, the error code "InvalidReference" indicates that the reference is not valid for the specified operation. Error codes are not localized.
- **message**: The message property of an error message contains a summary of the error in the localized string. The error message isn't intended for consumption by end users; you should use the error code and appropriate business logic to determine the error message that your add-in shows to end users.
- **debugInfo**: When present, the debugInfo property of the error message provides additional information that you can use to understand the root cause of the error.

#### 7 **Note**

If you use console.log() to print error messages to the console, those messages are only visible on the server. End users don't see those error messages in the add-in task pane or anywhere in the Office application. To report errors to the user, see **Error notifications**.

#### **Error codes and messages**

The following tables list the errors that application-specific APIs may return.

#### 7 **Note**

The following tables list error messages you may encounter while using the applicationspecific APIs. If you're working with the Common API, see **Office Common API error codes** to learn about relevant error messages.

| Error code   | Error message                                  | Notes                                                                                                                                                                                 |
|--------------|------------------------------------------------|---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| AccessDenied | You cannot perform the<br>requested operation. | This may be caused by a user's antivirus<br>software blocking parts of Office. See<br>the Common errors and<br>troubleshooting steps for "Error: Access<br>denied" for more guidance. |


| Error code            | Error message                                                                                                                                                                                                                                | Notes                                                                                                                                                                   |
|-----------------------|----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|-------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| ActivityLimitReached  | Activity limit has been<br>reached.                                                                                                                                                                                                          | None                                                                                                                                                                    |
| ApiNotAvailable       | The requested API is not<br>available.                                                                                                                                                                                                       | None                                                                                                                                                                    |
| ApiNotFound           | The API you are trying to<br>use could not be found. It<br>may be available in a<br>newer version of the<br>Office application. See<br>Office client application<br>and platform availability<br>for Office Add-ins for<br>more information. | None                                                                                                                                                                    |
| BadPassword           | The password you<br>supplied is incorrect.                                                                                                                                                                                                   | None                                                                                                                                                                    |
| Conflict              | Request could not be<br>processed because of a<br>conflict.                                                                                                                                                                                  | None                                                                                                                                                                    |
| ContentLengthRequired | A Content-length HTTP<br>header is missing.                                                                                                                                                                                                  | None                                                                                                                                                                    |
| GeneralException      | There was an internal<br>error while processing the<br>request.                                                                                                                                                                              | None                                                                                                                                                                    |
| HostRestartNeeded     | The Office application<br>needs to be restarted.                                                                                                                                                                                             | This error is thrown by the<br>Office.ribbon.requestUpdate() method<br>if the add-in that calls the method has<br>been updated since the Office<br>application started. |
| InsertDeleteConflict  | The insert or delete<br>operation attempted<br>resulted in a conflict.                                                                                                                                                                       | None                                                                                                                                                                    |
| InvalidArgument       | The argument is invalid or<br>missing or has an<br>incorrect format.                                                                                                                                                                         | None                                                                                                                                                                    |
| InvalidBinding        | This object binding is no<br>longer valid due to<br>previous updates.                                                                                                                                                                        | None                                                                                                                                                                    |


| Error code                       | Error message                                                                                                                                                            | Notes                                                                                                                                                                                                                |
|----------------------------------|--------------------------------------------------------------------------------------------------------------------------------------------------------------------------|----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| InvalidOperation                 | The operation attempted<br>is invalid on the object.                                                                                                                     | None                                                                                                                                                                                                                 |
| InvalidReference                 | This reference is not valid<br>for the current operation.                                                                                                                | None                                                                                                                                                                                                                 |
| InvalidRequest                   | Cannot process the<br>request.                                                                                                                                           | None                                                                                                                                                                                                                 |
| InvalidRibbonDefinition          | Office has been given an<br>invalid ribbon definition.                                                                                                                   | This error is thrown if an invalid<br>RibbonUpdateObject is passed to the<br>Office.ribbon.requestUpdate() method.                                                                                                   |
| InvalidSelection                 | The current selection is<br>invalid for this operation.                                                                                                                  | None                                                                                                                                                                                                                 |
| ItemAlreadyExists                | The resource being<br>created already exists.                                                                                                                            | None                                                                                                                                                                                                                 |
| ItemNotFound                     | The requested resource<br>doesn't exist.                                                                                                                                 | None                                                                                                                                                                                                                 |
| MemoryLimitReached               | The memory limit has<br>been reached. Your action<br>could not be completed.                                                                                             | None                                                                                                                                                                                                                 |
| NotImplemented                   | The requested feature<br>isn't implemented.                                                                                                                              | This could mean the API is in preview or<br>only supported on a particular platform<br>(such as online-only). See Office client<br>application and platform availability for<br>Office Add-ins for more information. |
| RequestAborted                   | The request was aborted<br>during run time.                                                                                                                              | None                                                                                                                                                                                                                 |
| RequestPayloadSizeLimitExceeded  | The request payload size<br>has exceeded the limit.<br>See the Resource limits<br>and performance<br>optimization for Office<br>Add-ins article for more<br>information. | This error only occurs in Office on the<br>web.                                                                                                                                                                      |
| ResponsePayloadSizeLimitExceeded | The response payload<br>size has exceeded the<br>limit. See the Resource<br>limits and performance<br>optimization for Office                                            | This error only occurs in Office on the<br>web.                                                                                                                                                                      |


| Error code           | Error message                                                                                            | Notes |
|----------------------|----------------------------------------------------------------------------------------------------------|-------|
|                      | Add-ins article for more<br>information.                                                                 |       |
| ServiceNotAvailable  | The service is unavailable.                                                                              | None  |
| Unauthenticated      | Required authentication<br>information is either<br>missing or invalid.                                  | None  |
| UnsupportedFeature   | The operation failed<br>because the source<br>worksheet contains one<br>or more unsupported<br>features. | None  |
| UnsupportedOperation | The operation being<br>attempted is not<br>supported.                                                    | None  |

#### **Excel-specific error codes and messages**

| Error code                | Error message                                                                                                                                                                                                                                        | Notes                                                         |
|---------------------------|------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|---------------------------------------------------------------|
| EmptyChartSeries          | The attempted<br>operation failed<br>because the chart<br>series is empty.                                                                                                                                                                           | None                                                          |
| FilteredRangeConflict     | The attempted<br>operation causes a<br>conflict with a filtered<br>range.                                                                                                                                                                            | None                                                          |
| FormulaLengthExceedsLimit | The bytecode of the<br>applied formula<br>exceeds the maximum<br>length limit. For Office<br>on 32-bit machines,<br>the bytecode length<br>limit is 16384<br>characters. On 64-bit<br>machines, the<br>bytecode length limit<br>is 32768 characters. | This error occurs in both Excel on the web<br>and on desktop. |


| Error code                     | Error message                                                                                                                                                                                                                                                                                                                                  | Notes                                                                                                                                                                                                                                                                      |
|--------------------------------|------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| GeneralException               | Various.                                                                                                                                                                                                                                                                                                                                       | The data types APIs return GeneralException<br>errors with dynamic error messages. These<br>messages reference the cell that is the<br>source of the error, and the problem that is<br>causing the error, such as: "Cell A1 is missing<br>the required property type<br>." |
| InactiveWorkbook               | The operation failed<br>because multiple<br>workbooks are open<br>and the workbook<br>being called by this API<br>has lost focus.                                                                                                                                                                                                              | None                                                                                                                                                                                                                                                                       |
| InvalidOperationInCellEditMode | The operation isn't<br>available while Excel is<br>in Edit cell mode. Exit<br>Edit mode by using the<br>Enter or Tab keys, or<br>by selecting another<br>cell, and then try again.                                                                                                                                                             | None                                                                                                                                                                                                                                                                       |
| MergedRangeConflict            | Cannot complete the<br>operation. A table can't<br>overlap with another<br>table, a PivotTable<br>report, query results,<br>merged cells, or an<br>XML Map.                                                                                                                                                                                    | None                                                                                                                                                                                                                                                                       |
| NonBlankCellOffSheet           | Microsoft Excel can't<br>insert new cells<br>because it would push<br>non-empty cells off the<br>end of the worksheet.<br>These non-empty cells<br>might appear empty<br>but have blank values,<br>some formatting, or a<br>formula. Delete<br>enough rows or<br>columns to make room<br>for what you want to<br>insert and then try<br>again. | None                                                                                                                                                                                                                                                                       |
| OperationCellsExceedLimit      | The attempted<br>operation affects more                                                                                                                                                                                                                                                                                                        | If the TableColumnCollection.add<br>API<br>triggers this error, confirm that there is no                                                                                                                                                                                   |


| Error code                  | Error message                                                                                                                                                                                        | Notes                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                              |
|-----------------------------|------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|                             | than the limit of<br>33554000 cells.                                                                                                                                                                 | unintentional data within the worksheet but<br>outside of the table. In particular, check for<br>data in the right-most columns of the<br>worksheet. Remove the unintended data to<br>resolve this error. One way to verify how<br>many cells that an operation processes is to<br>run the following calculation: (number of<br>table rows) x (16383 - (number of table<br>columns)) . The number 16383 is the<br>maximum number of columns that Excel<br>supports.<br>This error only occurs in Excel on the web. |
| PivotTableRangeConflict     | The attempted<br>operation causes a<br>conflict with a<br>PivotTable range.                                                                                                                          | None                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                               |
| RangeExceedsLimit           | The cell count in the<br>range has exceeded<br>the maximum<br>supported number.<br>See the Resource limits<br>and performance<br>optimization for Office<br>Add-ins article for<br>more information. | None                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                               |
| RefreshWorkbookLinksBlocked | The operation failed<br>because the user<br>hasn't granted<br>permission to refresh<br>external workbook<br>links.                                                                                   | None                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                               |
| UndoNotSupported            | The JavaScript API<br>request failed due to<br>lack of support for the<br>undo operation.                                                                                                            | None                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                               |
| UnsupportedSheet            | This sheet type does<br>not support this<br>operation, since it is a<br>Macro or Chart sheet.                                                                                                        | None                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                               |

#### **Word-specific error codes and messages**


| Error code                   | Error message                                | Notes                                           |
|------------------------------|----------------------------------------------|-------------------------------------------------|
| SearchDialogIsOpen           | The search dialog is open.                   | None                                            |
| SearchStringInvalidOrTooLong | The search string is invalid or<br>too long. | The search string maximum is 255<br>characters. |

### **Error notifications**

How you report errors to users depends on the UI system you're using.

- If you're using React as the UI system, use the [Fluent UI](https://react.fluentui.dev/) components and design elements. We recommend that error messages be conveyed with a [Dialog](https://react.fluentui.dev/?path=/docs/components-dialog--default) component. If the error is in the user's input, configure the [Input](https://react.fluentui.dev/?path=/docs/components-input--default) component to display the error as bold red text.
#### 7 **Note**

The **[Alert](https://react.fluentui.dev/?path=/docs/preview-components-alert--default)** component can also be used to report errors to users, but it's currently in preview and shouldn't be used in a production add-in. For information about its release status, see the **[Fluent UI React v9 Component Roadmap](https://github.com/microsoft/fluentui/wiki/Fluent-UI-React-v9-Component-Roadmap)** .

- If you're not using React for the UI, consider using the older [Fabric UI](https://developer.microsoft.com/fluentui#/get-started/web#fabric-core) components implemented directly in HTML and JavaScript. Some example templates are in the [Office-](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates)[Add-in-UX-Design-Patterns-Code](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates) repository. Take a look especially in the dialog and navigation subfolders. The sample [Excel-Add-in-SalesLeads](https://github.com/OfficeDev/Excel-Add-in-SalesLeads) uses a message banner.
- [OfficeExtension.Error object](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.error)
- Office Common API error codes


# **Office Common API error codes**

Article • 12/27/2022

This article documents the error messages you might encounter while using the Common API model. These error codes don't apply to application-specific APIs, such as the Excel JavaScript API or the Word JavaScript API.

See API models to learn more about the differences between the Common API and application-specific API models.

#### **Error codes**

The following table lists the error codes, names, and messages displayed, and the conditions they indicate.

| Error.code | Error.name                  | Error.message                                                                        | Condition                                                                                                                                       |
|------------|-----------------------------|--------------------------------------------------------------------------------------|-------------------------------------------------------------------------------------------------------------------------------------------------|
| 1000       | Invalid<br>Coercion<br>Type | The specified<br>coercion type is<br>not supported                                   | The coercion type is not supported in the<br>Office application. (For example, OOXML and<br>HTML coercion types are not supported in<br>Excel.) |
| 1001       | Data Read<br>Error          | The current<br>selection is not<br>supported.                                        | The user's current selection is not supported<br>(that is, it is something different than the<br>supported coercion types).                     |
| 1002       | Invalid<br>Coercion<br>Type | The specified<br>coercion type is<br>not compatible for<br>this binding type.        | The solution developer provided an<br>incompatible combination of coercion type<br>and binding type.                                            |
| 1003       | Data Read<br>Error          | The specified<br>rowCount or<br>columnCount<br>values are invalid.                   | The user supplies invalid column or row<br>counts.                                                                                              |
| 1004       | Data Read<br>Error          | The current<br>selection is not<br>compatible for the<br>specified coercion<br>type. | The current selection is not supported for the<br>specified coercion type by this application.                                                  |
| 1005       | Data Read                   | The specified                                                                        | The user supplies invalid startRow or startCol                                                                                                  |
|            | Error                       | startRow or                                                                          | values.                                                                                                                                         |


| Error.code | Error.name          | Error.message                                                                                                           | Condition                                                                                                                                                                                                   |                   |                                                 |                                                |                                      |
|------------|---------------------|-------------------------------------------------------------------------------------------------------------------------|-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|-------------------|-------------------------------------------------|------------------------------------------------|--------------------------------------|
|            |                     | startColumn values<br>are invalid.                                                                                      |                                                                                                                                                                                                             |                   |                                                 |                                                |                                      |
| 1006       | Data Read<br>Error  | Coordinate<br>parameters cannot<br>be used with<br>coercion type<br>"Table" when the<br>table contains<br>merged cells. | The user tries to get partial data from a non<br>uniform table (that is, a table that has merged<br>cells).                                                                                                 |                   |                                                 |                                                |                                      |
| 1007       | Data Read<br>Error  | The size of the<br>document is too<br>large.                                                                            | The user tries to get a document larger than<br>the size currently supported.                                                                                                                               |                   |                                                 |                                                |                                      |
| 1008       | Data Read           | The requested data                                                                                                      | The user requests to read data beyond the                                                                                                                                                                   |                   |                                                 |                                                |                                      |
|            | 1009                | Error                                                                                                                   | Data Read<br>Error                                                                                                                                                                                          | set is too large. | The specified file<br>type is not<br>supported. | data limits defined by the Office application. | The user sends an invalid file type. |
| 2000       | Data Write<br>Error | The supplied data<br>object type is not<br>supported.                                                                   | An unsupported data object is supplied.                                                                                                                                                                     |                   |                                                 |                                                |                                      |
| 2001       | Data Write<br>Error | Cannot write to the<br>current selection.                                                                               | The user's current selection is not supported<br>for a write operation. (For example, when the<br>user selects an image.)                                                                                   |                   |                                                 |                                                |                                      |
| 2002       | Data Write<br>Error | The supplied data<br>object is not<br>compatible with<br>the shape or<br>dimensions of the<br>current selection.        | Multiple cells are selected (and the selection<br>shape does not match the shape of the data).<br>Multiple cells are selected (and the selection<br>dimensions do not match the dimensions of<br>the data). |                   |                                                 |                                                |                                      |
| 2003       | Data Write<br>Error | The set operation<br>failed because the<br>supplied data<br>object will<br>overwrite data.                              | A single cell is selected and the supplied data<br>object overwrites data in the worksheet.                                                                                                                 |                   |                                                 |                                                |                                      |
| 2004       | Data Write<br>Error | The supplied data<br>object does not<br>match the size of<br>the current<br>selection.                                  | The user supplies an object larger than the<br>current selection size.                                                                                                                                      |                   |                                                 |                                                |                                      |


| Error.code | Error.name                   | Error.message                                                                                      | Condition                                                                                                                             |                      |                                                                                                                       |                                           |                                                                                                             |
|------------|------------------------------|----------------------------------------------------------------------------------------------------|---------------------------------------------------------------------------------------------------------------------------------------|----------------------|-----------------------------------------------------------------------------------------------------------------------|-------------------------------------------|-------------------------------------------------------------------------------------------------------------|
| 2005       | Data Write<br>Error          | The specified<br>startRow or<br>startColumn values<br>are invalid.                                 | The user supplies invalid startRow or startCol<br>values.                                                                             |                      |                                                                                                                       |                                           |                                                                                                             |
| 2006       | Invalid<br>Format Error      | The format of the<br>specified data<br>object is not valid.                                        | The solution developer supplies an invalid<br>HTML or OOXML string, a malformed HTML<br>string, or an invalid OOXML string.           |                      |                                                                                                                       |                                           |                                                                                                             |
| 2007       | Invalid Data<br>Object       | The type of the<br>specified data<br>object is not<br>compatible with<br>the current<br>selection. | The solution developer supplies a data object<br>not compatible with the specified coercion<br>type.                                  |                      |                                                                                                                       |                                           |                                                                                                             |
| 2008       | Data Write<br>Error          | TBD                                                                                                | TBD                                                                                                                                   |                      |                                                                                                                       |                                           |                                                                                                             |
| 2009       | Data Write                   | The specified data                                                                                 | The user tries to set data beyond the data                                                                                            |                      |                                                                                                                       |                                           |                                                                                                             |
|            | 2010                         | Error                                                                                              | Data Write<br>Error                                                                                                                   | object is too large. | Coordinate<br>parameters cannot<br>be used with<br>coercion type Table<br>when the table<br>contains merged<br>cells. | limits defined by the Office application. | The user tries to set partial data from a non<br>uniform table (that is, a table that has merged<br>cells). |
| 3000       | Binding<br>Creation<br>Error | Cannot bind to the<br>current selection.                                                           | The user's selection is not supported for<br>binding. (For example, the user is selecting an<br>image or other non-supported object.) |                      |                                                                                                                       |                                           |                                                                                                             |
| 3001       | Binding<br>Creation<br>Error | TBD                                                                                                | TBD                                                                                                                                   |                      |                                                                                                                       |                                           |                                                                                                             |
| 3002       | Invalid<br>Binding Error     | The specified<br>binding does not<br>exist.                                                        | The developer tries to bind to a non-existing<br>or removed binding.                                                                  |                      |                                                                                                                       |                                           |                                                                                                             |
| 3003       | Binding<br>Creation<br>Error | Noncontiguous<br>selections are not<br>supported.                                                  | The user is making multiple selections.                                                                                               |                      |                                                                                                                       |                                           |                                                                                                             |
| 3004       | Binding<br>Creation<br>Error | A binding cannot<br>be created with the<br>current selection                                       | There are several conditions under which this<br>might happen. Please see the "Binding                                                |                      |                                                                                                                       |                                           |                                                                                                             |


| Error.code | Error.name                          | Error.message                                                                                              | Condition                                                                                                                                              |                   |                                                       |                                            |                                                                                                                      |
|------------|-------------------------------------|------------------------------------------------------------------------------------------------------------|--------------------------------------------------------------------------------------------------------------------------------------------------------|-------------------|-------------------------------------------------------|--------------------------------------------|----------------------------------------------------------------------------------------------------------------------|
|            |                                     | and the specified<br>binding type.                                                                         | creation error conditions" section later in this<br>article.                                                                                           |                   |                                                       |                                            |                                                                                                                      |
| 3005       | Invalid                             | Operation is not                                                                                           | The developer sends an add row or add                                                                                                                  |                   |                                                       |                                            |                                                                                                                      |
|            | 3006                                | Binding                                                                                                    | Binding<br>Creation<br>Error                                                                                                                           | supported on this | The named item<br>does not exist.                     | column operation on a binding type that is | The named item cannot be found. No content<br>control or table with that name exists.                                |
|            | 3007                                | Operation                                                                                                  | Binding<br>Creation<br>Error                                                                                                                           | binding type.     | Multiple objects<br>with the same<br>name were found. | not of coercion type table .               | Collision error: more than one content control<br>with the same name exists, and fail on<br>collision is set to true |
| 3008       | Binding<br>Creation<br>Error        | The specified<br>binding type is not<br>compatible with<br>the supplied<br>named item.                     | Named item can't be bound to type. For<br>example, a content control contains text, but<br>the developer tried to bind by using coercion<br>type table |                   |                                                       |                                            |                                                                                                                      |
| 3009       | Invalid<br>Binding<br>Operation     | The binding type is<br>not supported.                                                                      | Used for backward compatibility.                                                                                                                       |                   |                                                       |                                            |                                                                                                                      |
| 3010       | Unsupported<br>Binding<br>Operation | The selected<br>content needs to<br>be in table format.<br>Format the data as<br>a table and try<br>again. | The developer is trying to use the<br>addRowsAsync or deleteAllDataValuesAsync<br>method of the TableBinding object on data<br>of coercion type matrix |                   |                                                       |                                            |                                                                                                                      |
| 4000       | Read Settings<br>Error              | The specified<br>setting name does<br>not exist.                                                           | A nonexistent setting name is supplied.                                                                                                                |                   |                                                       |                                            |                                                                                                                      |
| 4001       | Save Settings<br>Error              | The settings could<br>not be saved.                                                                        | Settings could not be saved.                                                                                                                           |                   |                                                       |                                            |                                                                                                                      |
| 4002       | Settings Stale<br>Error             | Settings could not<br>be saved because<br>they are stale.                                                  | Settings are stale and developer indicated not<br>to override settings.                                                                                |                   |                                                       |                                            |                                                                                                                      |
| 5000       | Settings Stale<br>Error             | The operation is<br>not supported.                                                                         | The operation is not supported in the current<br>Office application. For example,<br>document.getSelectionAsync is called from<br>Outlook.             |                   |                                                       |                                            |                                                                                                                      |


| Error.code | Error.name     | Error.message                      | Condition                                                                                                                                                    |
|------------|----------------|------------------------------------|--------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 5001       | Internal Error | An internal error<br>has occurred. | Refers to an internal error condition, which<br>can occur for any of the following reasons.                                                                  |
|            |                |                                    | Expand table<br>ノ                                                                                                                                            |
|            |                |                                    | An add-in being used by another user<br>sharing the workbook created a binding at<br>approximately the same time, and your<br>add-in needs to retry binding. |
|            |                |                                    | An unknown error occurred.                                                                                                                                   |
|            |                |                                    | The operation failed.                                                                                                                                        |
|            |                |                                    | Access was denied because the user is not<br>a member of an authorized role.                                                                                 |
|            |                |                                    | Access was denied because secure,<br>encrypted communication is required.                                                                                    |
|            |                |                                    | Data is stale and the user needs to confirm<br>enabling the queries to refresh it.                                                                           |
|            |                |                                    | The site collection CPU quota has been<br>exceeded.                                                                                                          |
|            |                |                                    | The site collection memory quota has been<br>exceeded.                                                                                                       |
|            |                |                                    | The session memory quota has been<br>exceeded.                                                                                                               |
|            |                |                                    | The workbook is in an invalid state and the<br>operation can't be performed.                                                                                 |
|            |                |                                    | The session has timed out due to inactivity<br>and the user needs to reload the<br>workbook.                                                                 |
|            |                |                                    | The maximum number of allowed sessions<br>per user has been exceeded.                                                                                        |
|            |                |                                    | The operation was canceled by the user.                                                                                                                      |
|            |                |                                    | The operation can't be completed because<br>it is taking too long.                                                                                           |
|            |                |                                    | The request can't be completed and needs<br>to be retried.                                                                                                   |


| Error.code | Error.name                     | Error.message                                                                    | Condition                                                                                                                                             |
|------------|--------------------------------|----------------------------------------------------------------------------------|-------------------------------------------------------------------------------------------------------------------------------------------------------|
|            |                                | The trial period of the product has expired.                                     |                                                                                                                                                       |
|            |                                |                                                                                  | The session has timed out due to inactivity.                                                                                                          |
|            |                                |                                                                                  | The user doesn't have permission to<br>perform the operation on the specified<br>range.                                                               |
|            |                                |                                                                                  | The user's regional settings don't match<br>the current collaboration session.                                                                        |
|            |                                |                                                                                  | The user is no longer connected and must<br>refresh or re-open the workbook.                                                                          |
|            |                                |                                                                                  | The requested range doesn't exist in the<br>sheet.                                                                                                    |
|            |                                |                                                                                  | The user doesn't have permission to edit<br>the workbook.                                                                                             |
|            |                                |                                                                                  | The workbook can't be edited because it is<br>locked.                                                                                                 |
|            |                                |                                                                                  | The session can't auto save the workbook.                                                                                                             |
|            |                                |                                                                                  | The session can't refresh its lock on the<br>workbook file.                                                                                           |
|            |                                |                                                                                  | The request can't be processed and needs<br>to be retried.                                                                                            |
|            |                                |                                                                                  | The user's sign-in information couldn't be<br>verified and needs to be re-entered.                                                                    |
|            |                                |                                                                                  | The user has been denied access.                                                                                                                      |
|            |                                |                                                                                  | The shared workbook needs to be<br>updated.                                                                                                           |
| 5002       | Permission<br>Denied           | The requested<br>operation is not<br>allowed on the<br>current document<br>mode. | The solution developer submits a set<br>operation, but the document is in a mode<br>that does not allow modifications, such as<br>'Restrict Editing'. |
| 5003       | Event<br>Registration<br>Error | The specified event<br>type is not<br>supported by the<br>current object.        | The solution developer tries to register or<br>unregister a handler to an event that does<br>not exist.                                               |


| Error.code | Error.name                         | Error.message                                                                                     | Condition                                                                                                                                     |
|------------|------------------------------------|---------------------------------------------------------------------------------------------------|-----------------------------------------------------------------------------------------------------------------------------------------------|
| 5004       | Invalid API<br>call                | Invalid API call in<br>the current context.                                                       | An invalid call is made for the context, for<br>example, trying to use a CustomXMLPart object<br>in Excel.                                    |
| 5005       | Data Stale                         | Operation failed<br>because the data is<br>stale on the server.                                   | The data on the server needs to be refreshed.                                                                                                 |
| 5006       | Session<br>Timeout                 | The document<br>session timed out.<br>Reload the<br>document.                                     | The session has timed out.                                                                                                                    |
| 5007       | Invalid API<br>call                | The enumeration is<br>not supported in<br>the current context.                                    | The enumeration is not supported in the<br>current context.                                                                                   |
| 5009       | Permission<br>Denied               | Access Denied                                                                                     | The add-in does not have permission to call<br>the specific API.                                                                              |
| 5012       | Invalid Or<br>Timed Out<br>Session | Your Office browser<br>session has expired<br>or is invalid. To<br>continue, refresh<br>the page. | The session between the Office client and<br>server has expired or the date, time, or time<br>zone is incorrect on your computer.             |
| 6000       | Invalid node                       | The specified node<br>was not found.                                                              | The CustomXmlPart node was not found.                                                                                                         |
| 6100       | Custom XML<br>error                | Custom XML error                                                                                  | Invalid API call.                                                                                                                             |
| 7000       | Invalid Id                         | The specified Id<br>does not exist.                                                               | Invalid ID.                                                                                                                                   |
| 7001       | Invalid<br>navigation              | The object is<br>located in a place<br>where navigation is<br>not supported.                      | The user can find the object, but cannot<br>navigate to it. (For example, in Word, the<br>binding is to the header, footer, or a<br>comment.) |
| 7002       | Invalid<br>navigation              | The object is<br>locked or<br>protected.                                                          | The user is trying to navigate to a locked or<br>protected range.                                                                             |
| 7004       | Invalid<br>navigation              | The operation<br>failed because the<br>Index is out of<br>range.                                  | The user is trying to navigate to an index that<br>is out of range.                                                                           |


| Error.code | Error.name           | Error.message                                                                                                                                | Condition                                                                                                                         |
|------------|----------------------|----------------------------------------------------------------------------------------------------------------------------------------------|-----------------------------------------------------------------------------------------------------------------------------------|
| 8000       | Missing<br>Parameter | We couldn't format<br>the table cell<br>because some<br>parameter values<br>are missing.<br>Double-check the<br>parameters and try<br>again. | The cellFormat method is missing some<br>parameters. For example, there are missing<br>cells, format, or tableOptions parameters. |
| 8010       | Invalid value        | One or more of the<br>cells parameters<br>have values that<br>aren't allowed.<br>Double-check the<br>values and try<br>again.                | The common cells reference enumeration is<br>not defined. For example, All, Data, Headers.                                        |
| 8011       | Invalid value        | One or more of the<br>tableOptions<br>parameters have<br>values that aren't<br>allowed. Double<br>check the values<br>and try again.         | One of the values in tableOptions is invalid.                                                                                     |
| 8012       | Invalid value        | One or more of the<br>format parameters<br>have values that<br>aren't allowed.<br>Double-check the<br>values and try<br>again.               | One of the values in the format is invalid.                                                                                       |
| 8020       | Out of range         | The row index<br>value is out of the<br>allowed range. Use<br>a positive value (0<br>or higher) that's<br>less than the<br>number of rows.   | The row index is more than the biggest row<br>index of the table or less than 0.                                                  |
| 8021       | Out of range         | The column index<br>value is out of the<br>allowed range. Use<br>a positive value (0<br>or higher) that's<br>less than the                   | The column index is more than the biggest<br>column index of the table or less than 0.                                            |


| Error.code | Error.name                             | Error.message                                                                                                         | Condition                                                                                                                                                                                                                                                                            |
|------------|----------------------------------------|-----------------------------------------------------------------------------------------------------------------------|--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|            |                                        | number of<br>columns.                                                                                                 |                                                                                                                                                                                                                                                                                      |
| 8022       | Out of range                           | The value is out of<br>the allowed range.                                                                             | Some of the values in the format are out of<br>the supported ranges.                                                                                                                                                                                                                 |
| 9016       | Permission<br>denied                   | Permission denied                                                                                                     | Access is denied.                                                                                                                                                                                                                                                                    |
| 9020       | Generic<br>Response<br>Error           | An internal error<br>has occurred.                                                                                    | Refers to an internal error condition, which<br>can occur for any number of reasons.                                                                                                                                                                                                 |
| 9021       | Save Error                             | Connection error<br>occurred while<br>trying to save the<br>item on the server.                                       | The item couldn't be saved. This could be due<br>to a server connection error if using Online<br>Mode in Outlook desktop, or due to an<br>attempt to re-save a draft item that was<br>deleted from the Exchange server.                                                              |
| 9022       | Message In<br>Different<br>Store Error | The EWS ID cannot<br>be retrieved<br>because the<br>message is saved in<br>another store.                             | The EWS ID for the current message couldn't<br>be retrieved as the message may have been<br>moved or the sending mailbox may have<br>changed.                                                                                                                                        |
| 9041       | Network<br>error                       | The user is no<br>longer connected<br>to the network.<br>Please check your<br>network<br>connection and try<br>again. | The user no longer has network or internet<br>access.                                                                                                                                                                                                                                |
| 9043       | Attachment<br>Type Not<br>Supported    | The attachment<br>type is not<br>supported.                                                                           | The API doesn't support the attachment type.<br>For example, item.getAttachmentContentAsync<br>throws this error if the attachment is an<br>embedded image in Rich Text Format, or if it's<br>an item type other than an email or calendar<br>item (such as a contact or task item). |
| 9057       | Size Limit<br>Exceeded                 | A maximum of<br>32KB is available<br>for the settings of<br>each add-in.                                              | When updating roaming settings via<br>Office.context.roamingSettings.set, the size<br>cannot exceed 32KB. See<br>Office.RoamingSettings interface.                                                                                                                                   |
| 12002      | Not<br>applicable                      | Not applicable                                                                                                        | One of the following:<br>- No page exists at the URL that was passed<br>to displayDialogAsync<br>- The page that was passed to                                                                                                                                                       |


| Error.code | Error.name        | Error.message  | Condition                                                                                                                                                                                                                                                                                   |
|------------|-------------------|----------------|---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|            |                   |                | displayDialogAsync loaded, but the dialog<br>box was directed to a page that it cannot find<br>or load, or it has been directed to a URL with<br>invalid syntax. Thrown within the dialog and<br>triggers a DialogEventReceived event in the<br>host page.                                  |
| 12003      | Not<br>applicable | Not applicable | The dialog box was directed to a URL with the<br>HTTP protocol. HTTPS is required. Thrown<br>within the dialog and triggers a<br>DialogEventReceived event in the host page.                                                                                                                |
| 12004      | Not<br>applicable | Not applicable | The domain of the URL passed to<br>displayDialogAsync is not trusted. The<br>domain must be the same domain as the host<br>page (including protocol and port number).<br>Thrown by call of displayDialogAsync                                                                               |
| 12005      | Not<br>applicable | Not applicable | The URL passed to displayDialogAsync uses<br>the HTTP protocol. HTTPS is required. Thrown<br>by call of displayDialogAsync . (In some<br>versions of Office, the error message returned<br>with 12005 is the same one returned for<br>12004.)                                               |
| 12006      | Not<br>applicable | Not applicable | The dialog box was closed, usually because<br>the user chooses the X button. Thrown within<br>the dialog and triggers a<br>DialogEventReceived event in the host page.                                                                                                                      |
| 12007      | Not<br>applicable | Not applicable | A dialog box is already opened from this host<br>window. A host window, such as a task pane,<br>can only have one dialog box open at a time.<br>Thrown by call of displayDialogAsync                                                                                                        |
| 12009      | Not<br>applicable | Not applicable | The user chose to ignore the dialog box. This<br>error can occur in online versions of Office,<br>where users may choose not to allow an add<br>in to present a dialog. Thrown by call of<br>displayDialogAsync .                                                                           |
| 12011      | Not<br>applicable | Not applicable | The user's browser is configured in a way that<br>blocks popups. This error can occur in Office<br>on the web if the browser is Safari and it's<br>configured to block popups or the browser is<br>Edge Legacy and the add-in domain is in a<br>different security zone from the domain the |


| Error.code | Error.name        | Error.message  | Condition                                                           |
|------------|-------------------|----------------|---------------------------------------------------------------------|
|            |                   |                | dialog is trying to open. Thrown by call of<br>displayDialogAsync . |
| 13nnn      | Not<br>applicable | Not applicable | See Causes and handling of errors from<br>getAccessToken.           |

### **Binding creation error conditions**

When a binding is created in the API, indicate the binding type that you want to use. The following tables lists the binding types and the resulting binding behaviors that are expected.

#### **Behavior in Excel**

The following table summarizes binding behavior in Excel.

| Specified<br>Binding<br>Type | Actual Selection                                                                                                               | Behavior                                                                                                         |
|------------------------------|--------------------------------------------------------------------------------------------------------------------------------|------------------------------------------------------------------------------------------------------------------|
| Matrix                       | Range of cells (including within a table,<br>and single cell)                                                                  | A binding of type matrix is created<br>on the selected cells. No<br>modification in the document is<br>expected. |
| Matrix                       | Text selected in the cell                                                                                                      | A binding of type matrix is created<br>on the whole cell. No modification in<br>the document is expected.        |
| Matrix                       | Multiple selection/invalid selection (For<br>example, user selects a picture, object, or<br>Word Art.)                         | The binding cannot be created.                                                                                   |
| Table                        | Range of cells (includes single cell)                                                                                          | The binding cannot be created.                                                                                   |
| Table                        | Range of cell within a table (includes<br>single cell within a table, or the whole<br>table, or text within a cell in a table) | A binding is created in the whole<br>table.                                                                      |
| Table                        | Half selection in a table and half selection<br>outside the table                                                              | The binding cannot be created.                                                                                   |
| Table                        | Text selected in the cell (not in the table.)                                                                                  | The binding cannot be created.                                                                                   |


| Specified<br>Binding<br>Type | Actual Selection                                                                                         | Behavior                                                |
|------------------------------|----------------------------------------------------------------------------------------------------------|---------------------------------------------------------|
| Table                        | Multiple selection/invalid selection (For<br>example, user selects a picture, object,<br>Word Art, etc.) | The binding cannot be created.                          |
| Text                         | Range of cells                                                                                           | The binding cannot be created.                          |
| Text                         | Range of cells within a table                                                                            | The binding cannot be created.                          |
| Text                         | Single cell                                                                                              | A binding of type text is created.                      |
| Text                         | Single cell within a table                                                                               | A binding of type text is created.                      |
| Text                         | Text selected in the cell                                                                                | A binding of type text in the whole<br>cell is created. |

#### **Behavior in Word**

The following table summarizes binding behavior in Word.

| Specified<br>Binding Type | Actual Selection                                                       | Behavior                                                                                              |
|---------------------------|------------------------------------------------------------------------|-------------------------------------------------------------------------------------------------------|
| Matrix                    | Text                                                                   | The binding cannot be created.                                                                        |
| Matrix                    | Whole table                                                            | A binding of type matrix is created.Document is<br>changed and a content control must wrap the table. |
| Matrix                    | Range within a table                                                   | The binding cannot be created.                                                                        |
| Matrix                    | Invalid selection (for<br>example, multiple, invalid<br>objects, etc.) | The binding cannot be created.                                                                        |
| Table                     | Text                                                                   | The binding cannot be created.                                                                        |
| Table                     | Whole table                                                            | A binding of type text is created.                                                                    |
| Table                     | Range within a table                                                   | The binding cannot be created.                                                                        |
| Table                     | Invalid selection (for<br>example, multiple, invalid<br>objects, etc.) | The binding cannot be created.                                                                        |


| Specified<br>Binding Type | Actual Selection                                                       | Behavior                                                                                                                                  |
|---------------------------|------------------------------------------------------------------------|-------------------------------------------------------------------------------------------------------------------------------------------|
| Text                      | Whole table                                                            | A binding of type text is created.                                                                                                        |
| Text                      | Range within a table                                                   | The binding cannot be created.                                                                                                            |
| Text                      | Multiple selection                                                     | The last selection will be wrapped with a content<br>control and a binding to that control. A content<br>control of type text is created. |
| Text                      | Invalid selection (for<br>example, multiple, invalid<br>objects, etc.) | The binding cannot be created.                                                                                                            |

#### **See also**

- Office Add-ins development lifecycle
- Understanding the Office JavaScript API
- Error handling with the application-specific JavaScript APIs
- Troubleshoot error messages for single sign-on (SSO)
- Troubleshoot development errors with Office Add-ins

#### 6 **Collaborate with us on GitHub**

The source for this content can be found on GitHub, where you can also create and review issues and pull requests. For more information, see [our](https://github.com/OfficeDev/office-js-docs-pr/blob/main/Contributing.md) [contributor guide](https://github.com/OfficeDev/office-js-docs-pr/blob/main/Contributing.md).

#### **Office Add-ins feedback**

Office Add-ins is an open source project. Select a link to provide feedback:

- [Open a documentation issue](https://github.com/OfficeDev/office-js-docs-pr/issues/new?template=3-customer-feedback.yml&pageUrl=https%3A%2F%2Flearn.microsoft.com%2Fen-us%2Foffice%2Fdev%2Fadd-ins%2Freference%2Fjavascript-api-for-office-error-codes&pageQueryParams=&contentSourceUrl=https%3A%2F%2Fgithub.com%2FOfficeDev%2Foffice-js-docs-pr%2Fblob%2Fmain%2Fdocs%2Freference%2Fjavascript-api-for-office-error-codes.md&documentVersionIndependentId=78e25f2a-c710-6c16-32d7-f16d28b08e31&feedback=%0A%0A%5BEnter+feedback+here%5D%0A&author=%40o365devx&metadata=*+ID%3A+ad14a1fc-e890-3292-0a53-8a755db35f6e+%0A*+Service%3A+**microsoft-365**%0A*+Sub-service%3A+**add-ins**)
- [Provide product feedback](https://aka.ms/office-addins-dev-questions)


# **Troubleshoot user errors with Office Addins**

06/25/2025

At times your users might encounter issues with Office Add-ins that you develop. For example, an add-in fails to load or is inaccessible. Use the information in this article to help resolve common issues that your users encounter with your Office Add-in.

You can also use tools to intercept HTTP messages to identify and debug issues with your addins. Popular choices include [Fiddler](https://www.telerik.com/fiddler) , [Charles](https://www.charlesproxy.com/) , and [Requestly](https://requestly.com/downloads) .

### **Common errors and troubleshooting steps**

The following table lists common error messages that users might encounter and steps that your users can take to resolve the errors.

| Error message                                                                                                                     | Resolution                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                          |
|-----------------------------------------------------------------------------------------------------------------------------------|-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| App error: Catalog                                                                                                                | Verify firewall settings."Catalog" refers to AppSource. This message indicates                                                                                                                                                                                                                                                                                                                                                                                                                                                                      |
| could not be reached                                                                                                              | that the user cannot access AppSource.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                              |
| APP ERROR: This app<br>could not be started.<br>Close this dialog to<br>ignore the problem or<br>click "Restart" to try<br>again. | Verify that the latest Office updates are installed, or update with the Windows<br>Installer.                                                                                                                                                                                                                                                                                                                                                                                                                                                       |
| Error: Access denied.<br>E_ACCESSDENIED<br>(0x80070005)                                                                           | The antivirus software installed on the machine might prevent the host app<br>from creating a WebView2 process. To resolve this issue, add an exemption or<br>exclusion to the antivirus for the .exe files in the Office root folder (<br>C:\Program<br>Files\Microsoft Office\root\Office16 ) or for the entire Office root folder. If<br>this does not fix the issue, add an exemption or exclusion for the WebView2<br>process (<br>C:\Program Files (x86)\Microsoft\EdgeWebView\Application[latest<br>installed version]\msedgewebview2.exe ). |
| Error: Object doesn't<br>support property or<br>method<br>'defineProperty'                                                        | Confirm that Internet Explorer is not running in Compatibility Mode. Go to<br>Tools > Compatibility View Settings.                                                                                                                                                                                                                                                                                                                                                                                                                                  |
| Sorry, we couldn't load                                                                                                           | Make sure that the browser supports HTML5 local storage, or reset your                                                                                                                                                                                                                                                                                                                                                                                                                                                                              |
| the app because your                                                                                                              | Internet Explorer settings. For information about supported browsers, see                                                                                                                                                                                                                                                                                                                                                                                                                                                                           |


| Error message                                                                                   | Resolution                               |
|-------------------------------------------------------------------------------------------------|------------------------------------------|
| browser version is not<br>supported. Click here<br>for a list of supported<br>browser versions. | Requirements for running Office Add-ins. |

### **When installing an add-in, you see "Error loading add-ins" in the status bar**

- 1. Close Office.
- 2. Check that the time and date are set correctly on your computer. An incorrect time and date can cause issues when verifying the add-in's manifest.
- 3. Verify that the manifest is valid. See Validate an Office Add-in's manifest.
- 4. Restart the add-in.
- 5. Install the add-in again.

If the add-in package was tampered with before installation, this error will occur. Download the add-in again and try to reinstall it. Alternatively, contact the publisher of the add-in for help.

You can also give us feedback: if using Office on Windows or Mac, you can report feedback to the Office extensibility team directly from Office. To do this, select **Help** > **Feedback** > **Report a problem**. Sending a report provides necessary information to understand the issue.

### **Outlook add-in doesn't work correctly**

If an Outlook add-in running on Windows and using Internet Explorer is not working correctly, try turning on script debugging in Internet Explorer.

- Go to **Tools** > **Internet Options** > **Advanced**.
- Under **Browsing**, uncheck **Disable script debugging (Internet Explorer)** and **Disable script debugging (Other)**.

We recommend that you uncheck these settings only to troubleshoot the issue. If you leave them unchecked, you will get prompts when you browse. After the issue is resolved, check **Disable script debugging (Internet Explorer)** and **Disable script debugging (Other)** again.

## **Add-in doesn't activate in Office**

If the add-in doesn't activate when the user performs the following steps.

- 1. Signs in with their Microsoft account in the Office application.


- 2. Enables two-step verification for their Microsoft account.
- 3. Verifies their identity when prompted when they try to insert an add-in.

Verify that the latest Office updates are installed, or update with the [Windows Installer.](https://learn.microsoft.com/en-us/officeupdates/office-updates-msi)

## **Add-in dialog box cannot be displayed**

When using an Office Add-in, the user is asked to allow a dialog box to be displayed. The user chooses **Allow**, and the following error message occurs.

"The security settings in your browser prevent us from creating a dialog box. Try a different browser, or configure your browser so that [URL] and the domain shown in your address bar are in the same security zone."

| Notification                                                                                                                                                                                                                                                             |
|--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| The security settings in your browser prevent us from<br>creating a dialog box. Try a different browser, or<br>configure your browser so that 'https://dialog-<br>test.azurewebsites.net:443' and the domain shown in<br>your address bar are in the same security zone. |
|                                                                                                                                                                                                                                                                          |

#### ノ **Expand table**

| Affected browsers | Affected platforms |
|-------------------|--------------------|
| Microsoft Edge    | Office on the web  |

To resolve the issue, end users or administrators can add the domain of the add-in to the list of trusted sites in the Microsoft Edge browser.

To add a URL to your list of trusted sites:

- 1. In **Control Panel**, go to **Internet options** > **Security**.
- 2. Select the **Trusted sites** zone, and choose **Sites**.


- 3. Enter the URL that appears in the error message, and choose **Add**.
- 4. Try to use the add-in again. If the problem persists, verify the settings for the other security zones and ensure that the add-in domain is in the same zone as the URL that is displayed in the address bar of the Office application.

This issue occurs when the Dialog API is used in pop-up mode. To prevent this issue from occurring, use the [displayInFrame](https://learn.microsoft.com/en-us/javascript/api/office/office.ui) flag. This requires that your page support display within an iframe. The following example shows how to use the flag.

```
JavaScript
Office.context.ui.displayDialogAsync(startAddress, {displayInIFrame:true},
callback);
```
## **Add-in won't upgrade**

You may see the following error when deploying an updated manifest for your add-in: ADD-IN WARNING: This add-in is currently upgrading. Please close the current message or appointment, and re-open in a few moments.

When you add features or fix bugs in your add-in, you'll need to deploy the updates. If your add-in is deployed by one or more admins to their organizations, some manifest changes will require the admin to consent to the updates. Users remain on the existing version of the addin until the admin consents to the updates. The following manifest changes will require the admin to consent again.

- Changes to requested permissions. See Requesting permissions for API use in add-ins and Understanding Outlook add-in permissions.
- Additional or changed [Scopes.](https://learn.microsoft.com/en-us/javascript/api/manifest/scopes) (Not applicable if the add-in uses the unified manifest for Microsoft 365.)
- Additional or changed [Outlook events.](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/autolaunch)

#### 7 **Note**

Whenever you make a change to the manifest, you must raise the version number of the manifest.

- If the add-in uses the add-in only manifest, see **[Version element](https://learn.microsoft.com/en-us/javascript/api/manifest/version)**.
- If the add-in uses the unified manifest, see **[version property](https://learn.microsoft.com/en-us/microsoft-365/extensibility/schema/root#version)**.


- Troubleshoot development errors with Office Add-ins


# **Troubleshoot development errors with Office Add-ins**

Article • 02/12/2025

Here's a list of common issues you may encounter while developing an Office Add-in.

#### **Tip**

Clearing the Office cache often fixes issues related to stale code. This guarantees the latest manifest is uploaded, using the current file names, menu text, and other command elements. To learn more, see **Clear the Office cache**.

### **Add-in doesn't load in task pane or other issues with the add-in manifest**

See Validate an Office Add-in's manifest and Debug your add-in with runtime logging to debug add-in manifest issues.

## **Ribbon customizations are not rendering as expected**

- With the add-in sideloaded and running, paste the URLs for the add-in's ribbon icons into a browser's navigation bar and see if the icon files open.
- By default, add-in errors connected to the Office UI are suppressed. You can turn on these error messages with the following steps.
	- 1. With the add-in removed, open the **File** tab of the Office application.
	- 2. Select **Options**.
	- 3. In the **Options** dialog, select **Advanced**.
	- 4. In the **General** section (the **Developers** section for Outlook), enable **Show add-in user interface errors**.

Sideload the add-in again and see if there are any errors.

## **Changes to add-in commands including ribbon buttons and menu items do not take effect**


Clearing the cache helps ensure the latest version of your add-in's manifest is being used. To clear the Office cache, follow the instructions in Clear the Office cache. If you're using Office on the web, clear your browser's cache through the browser's UI.

## **Add-in commands from old development addins stay on ribbon even after the cache is cleared**

Sometimes buttons or menus from an add-in that you were developing in the past appears on the ribbon when you run an Office application even after you have cleared the cache. Try these techniques:

- If you develop add-ins on more than one computer and your user settings are synchronized across the computers, try clearing the Office cache on all the computers. Shut down all Office applications on all the computers, and then clear the cache on all of them before you open any Office application on any of them.
- If you published the manifest of the old add-in to a network share, shut down all Office applications, clear the cache, and then *be sure that the manifest for the addin is removed from the shared folder*.

## **Changes to static files, such as JavaScript, HTML, and CSS do not take effect**

The browser may be caching these files. To prevent this, turn off client-side caching when developing. The details will depend on what kind of server you are using. In most cases, it involves adding certain headers to the HTTP Responses. We suggest the following set.

- Cache-Control: "private, no-cache, no-store"
- Pragma: "no-cache"
- Expires: "-1"

For an example of doing this in an Node.JS Express server, see [this app.js file](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-NodeJS-SSO/Complete/app.js) . For an example in an ASP.NET project, see [this cshtml file](https://github.com/OfficeDev/Office-Add-in-samples/blob/main/Samples/auth/Office-Add-in-ASPNET-SSO/Complete/Office-Add-in-ASPNETCore-WebAPI/Views/Shared/_Layout.cshtml) .

If your add-in is hosted in Internet Information Server (IIS), you could also add the following to the web.config.


```
<system.webServer>
 <staticContent>
 <clientCache cacheControlMode="UseMaxAge"
cacheControlMaxAge="0.00:00:00" cacheControlCustom="must-revalidate" />
 </staticContent>
```
If these steps don't seem to work at first, you may need to clear the browser's cache. Do this through the UI of the browser. Sometimes the Edge cache isn't successfully cleared when you try to clear it in the Edge UI. If that happens, run the following command in a Windows Command Prompt.

```
Bash
del /s /f /q
%LOCALAPPDATA%\Packages\Microsoft.Win32WebViewHost_cw5n1h2txyewy\AC\#!123\IN
etCache\
```
## **Changes made to property values don't happen and there is no error message**

Check the reference documentation for the property to see if it is read-only. Also, the TypeScript definitions for Office JS specify which object properties are read-only. If you attempt to set a read-only property, the write operation will fail silently, with no error thrown. The following example erroneously attempts to set the read-only property [Chart.id.](https://learn.microsoft.com/en-us/javascript/api/excel/excel.chart#excel-excel-chart-id-member) See also Some properties cannot be set directly.

```
JavaScript
// This will do nothing, since `id` is a read-only property.
myChart.id = "5";
```
## **Getting error: "This add-in is no longer available"**

The following are some of the causes of this error. If you discover additional causes, please tell us with the feedback tool at the bottom of the page.

- If you're using Visual Studio, there may be a problem with the sideloading. Close all instances of the Office host and Visual Studio. Restart Visual Studio and try pressing F5 again.


- The add-in's manifest has been removed from its deployment location, such as Centralized Deployment, a SharePoint catalog, or a network share.
- If the add-in only manifest is being used, one of the following may apply.
	- The value of the [ID](https://learn.microsoft.com/en-us/javascript/api/manifest/id) element in the manifest has been changed directly in the deployed copy. If for any reason, you want to change this ID, first remove the add-in from the Office host, then replace the original manifest with the changed manifest. You many need to clear the Office cache to remove all traces of the original. See the Clear the Office cache article for instructions on clearing the cache for your operating system.
	- The add-in's manifest has a resid that isn't defined anywhere in the [Resources](https://learn.microsoft.com/en-us/javascript/api/manifest/resources) section of the manifest, or there is a mismatch in the spelling of the resid between where it is used and where it is defined in the **<Resources>** section.
	- There is a resid attribute somewhere in the manifest with more than 32 characters. A resid attribute, and the id attribute of the corresponding resource in the **<Resources>** section, cannot be more than 32 characters.
- The add-in has a custom Add-in Command but you are trying to run it on a platform that doesn't support them. For more information, see [Add-in commands](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/common/add-in-commands-requirement-sets) [requirement sets.](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/common/add-in-commands-requirement-sets)

### **Add-in doesn't work on Edge but it works on other browsers**

See Troubleshoot EdgeHTML and WebView2 (Microsoft Edge) issues.

#### **Excel add-in throws errors, but not consistently**

See Troubleshoot Excel add-ins for possible causes.

### **Word add-in throws errors or displays broken behavior**

See Troubleshoot Word add-ins for possible causes.

## **Add-in only manifest schema validation errors in Visual Studio projects**


If you're using newer features that require changes to the add-in only manifest file, you may get validation errors in Visual Studio. For example, when adding the **<Runtimes>** element to implement the shared runtime, you may see the following validation error.

#### **The element 'Host' in namespace**

**'http://schemas.microsoft.com/office/taskpaneappversionoverrides' has invalid child element 'Runtimes' in namespace**

**'http://schemas.microsoft.com/office/taskpaneappversionoverrides'**

If this occurs, you can update the XSD files that Visual Studio uses to the latest versions. The latest schema versions are at [\[MS-OWEMXML\]: Appendix A: Full XML Schema.](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8)

#### **Locate the XSD files**

- 1. Open your project in Visual Studio.
- 2. In **Solution Explorer**, open the manifest.xml file. The manifest is typically in the first project under your solution.
- 3. Select **View** > **Properties Window** ( F4 ).
- 4. Set the cursor selection in the manifest.xml so that the **Properties** window shows the **XML Document** properties.
- 5. In the **Properties** window, select the **Schemas** property, then select the ellipsis (...) to open the **XML Schemas** editor. Here you can find the exact folder location of all schema files your project uses.

| Properties                             |                                      |
|----------------------------------------|--------------------------------------|
| XML Document                           |                                      |
|                                        |                                      |
| Misc<br>l                              |                                      |
| Encoding                               | Unicode (UTF-8)                      |
| Output                                 |                                      |
| Schemas                                | "C:\Program Files\Microsoft Visual S |
| Stylesheet                             |                                      |
|                                        |                                      |
|                                        |                                      |
|                                        |                                      |
|                                        |                                      |
|                                        |                                      |
|                                        |                                      |
| Schemas                                |                                      |
| Schemas used to validate the document. |                                      |

**Update the XSD files**


- 1. Open the XSD file you want to update in a text editor. The schema name from the validation error will correlate to the XSD file name. For example, open **TaskPaneAppVersionOverridesV1_0.xsd**.
- 2. Locate the updated schema at [\[MS-OWEMXML\]: Appendix A: Full XML Schema](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8). For example, TaskPaneAppVersionOverridesV1_0 is at [taskpaneappversionoverrides](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-owemxml/82e93ec5-de22-42a8-86e3-353c8336aa40) [Schema](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-owemxml/82e93ec5-de22-42a8-86e3-353c8336aa40).
- 3. Copy the text into your text editor.
- 4. Save the updated XSD file.
- 5. Restart Visual Studio to pick up the new XSD file changes.

You can repeat the previous process for any additional schemas that are out-of-date.

## **When working offline, no Office APIs work**

When you're loading the Office JavaScript Library from a local copy instead of from the CDN, the APIs may stop working if the library isn't up-to-date. If you have been away from a project for a while, reinstall the library to get the latest version. The process varies according to your IDE. Choose one of the following options based on your environment.

- **Visual Studio**: Follow these steps to update the NuGet package.
	- 1. Choose **Tools** > **NuGet Package Manager** > **Manage Nuget Packages for Solution**.
	- 2. Choose the **Updates** tab.
	- 3. Select "Microsoft.Office.js". Ensure the package source is from nuget.org.
	- 4. In the left pane, choose **Install** and complete the package update process.
- **Any other IDE**: Get the latest npm packages [@microsoft/office-js](https://www.npmjs.com/package/@microsoft/office-js) and [@types/office-js](https://www.npmjs.com/package/@types/office-js) .

- Debug add-ins in Office on the web
- Sideload an Office Add-in on Mac
- Sideload an Office Add-in on iPad
- Debug Office Add-ins on a Mac
- Validate an Office Add-in's manifest
- Debug your add-in with runtime logging
- Troubleshoot user errors with Office Add-ins
- Runtimes in Office Add-ins


- [Microsoft Q&A (Office Development)](https://aka.ms/office-addins-dev-questions)


# **Debug event-based or spam-reporting add-ins**

07/16/2025

This article discusses the key debugging stages to enable and set breakpoints in your code as you implement event-based activation or integrated spam reporting in your add-in. Before you proceed, we recommend reviewing the troubleshooting guide for additional steps on how to resolve development errors.

To begin debugging, select the tab for your applicable client.

Windows (classic)

If you used the Yeoman generator for Office Add-ins to create your add-in project (for example, by completing an event-based activation walkthrough), follow the **Created with Yeoman generator** option throughout this article. Otherwise, follow the **Other** steps.

## **Mark your add-in for debugging and set the debugger port**

1. Get your add-in's ID from the manifest.

- **Add-in only manifest**: Use the value of the **<Id>** element child of the root **<OfficeApp>** element.
- **Unified manifest for Microsoft 365**: Use the value of the "id" property of the root anonymous { ... } object.
- 2. In the registry, mark your add-in for debugging.
	- **Created with Yeoman generator**: In a command line window, navigate to the root of your add-in folder then run the following command.

command line

npm start

In addition to building the code and starting the local server, this command sets the data of the

HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\[Add-in


ID]\UseDirectDebugger registry DWORD value for this add-in to 1 . [Add-in ID] is your add-in's ID from the manifest.

- **Other**: In the
HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\[Add-in ID]\UseDirectDebugger registry DWORD value, where [Add-in ID] is your addin's ID from the manifest, set its data to 1 .

#### 7 **Note**

If the Developer key (folder) doesn't already exist under

HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\ , complete the following steps to create it.

- a. Right-click (or select and hold) the **WEF** key (folder) and select **New** > **Key**.
- b. Name the new key **Developer**.
- 3. In the registry key

HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\[Add-in ID] , where [Add-in ID] is your add-in's ID from the manifest, create a new DWORD value with the following configuration.

- **Value name**: DebuggerPort
- **Value data (hexadecimal)**: 00002407

This sets the debugger port to 9223 .

- 4. Start your Office application or restart it if it's already open.
- 5. Perform the action to initiate the event you're developing for, such as creating a new message to initiate the OnNewMessageCompose event or reporting spam messages. The **Debug Event-based handler** dialog should appear. Do *not* interact with the dialog yet.


### **Configure and attach the debugger**

You can debug your add-in using the Microsoft Edge Inspect tool or Visual Studio Code.

#### **Debug with Microsoft Edge**

- 1. Open Microsoft Edge and go to **edge://inspect/#devices**.
- 2. In the **Remote Target** section, look for your add-in using its ID from the manifest. Then, select **Inspect**.

The DevTools window appears.

#### 7 **Note**

It may take some time for your add-in to appear in the **Remote Target** section. You may need to refresh the page for the add-in to appear.

#### 3. In the **Sources** tab, go to **file://** >

**Users/[User]/AppData/Local/Microsoft/Office/16.0/Wef/{[Office profile GUID]}/[Office account encoding]/Javascript/[Add-in ID]_[Add-in Version]_[locale]** > **bundle.js**. For readability, this article refers to the file name as **bundle.js**, but exact name depends on the Office application.

- Excel: **bundle_excel.js**
- Outlook: **bundle.js**
- PowerPoint: **bundle_powerpoint.js**
- Word: **bundle_word.js**


There's no direct method to determine the Office profile GUID or mail account encoding used in the **bundle.js** file path. If you're debugging multiple add-ins simultaneously, the easiest way to access an add-in's **bundle.js** file from the DevTools window is to locate the add-in's ID in the file path.

- 4. In the **bundle.js** file, place breakpoints where you want the debugger to stop.
- 5. Run the debugger.

#### **Debug with Visual Studio Code**

To debug your add-in in Visual Studio Code, you must have at least version 1.56.1 installed.

#### **Configure the debugger**

Configure the debugger in Visual Studio Code. Follow the steps applicable to your add-in project.

#### **Created with Yeoman generator**

- 1. In the command line, run the following to open your add-in project in Visual Studio Code.

```
command line
code .
```
- 2. In Visual Studio Code, open the **./.vscode/launch.json** file and add the following excerpt to your list of configurations. Save your changes.

```
JSON
{
 "name": "Direct Debugging",
 "type": "node",
 "request": "attach",
 "port": 9223,
 "timeout": 600000,
 "trace": true
}
```


#### **Other**

- 1. Create a new folder called **Debugging** (perhaps in your **Desktop** folder).
- 2. Open Visual Studio Code.
- 3. Go to **File** > **Open Folder**, navigate to the folder you created, then choose **Select Folder**.
- 4. On the Activity Bar, select **Run and Debug** ( Ctrl + Shift + D ).

- 5. Select the **create a launch.json file** link.
- 6. In the **Select Environment** dropdown, select **Edge: Launch** to create a launch.json file.
- 7. Add the following excerpt to your list of configurations. Save your changes.

```
JSON
{
 "name": "Direct Debugging",
 "type": "node",
 "request": "attach",
 "port": 9223,
 "timeout": 600000,
```


#### **Attach the debugger**

}

The **bundle.js** file of an add-in contains the JavaScript code of your add-in. It's created when an Office on Windows application is opened. When Office starts, the **bundle.js** file of each installed add-in is cached in the **Wef** folder of your machine.

- 1. To find the add-in's **bundle.js** file, navigate to the following folder in File Explorer. The text enclosed in [] represents your applicable Office and add-in information.
text

%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\{[Office profile GUID]}\[Office account encoding]\Javascript\[Add-in ID]_[Add-in Version]_[locale]

#### **Tip**

- For readability, this article refers to the file name as **bundle.js**, but exact name depends on the Office application.
	- Excel: **bundle_excel.js**
	- Outlook: **bundle.js**.
	- PowerPoint: **bundle_powerpoint.js**
	- Word: **bundle_word.js**
- There's no direct method to determine the Office profile GUID and account encoding used in the **bundle.js** file path. The most effective approach to locate your add-in's **bundle.js** file is to manually inspect each folder until you locate the **Javascript** folder that contains your add-in's ID.
- The **bundle.js** file is downloaded to the local **Wef** folder when the add-in is first installed. It's refreshed every time the Office application starts or is restarted. If the **bundle.js** file doesn't appear in the **Wef** folder and your add-in is installed or sideloaded, restart Office. For Outlook, you may need to **remove your add-in**, then **sideload** it again.
- 2. Open **bundle.js** in Visual Studio Code.
- 3. Place breakpoints in **bundle.js** where you want the debugger to stop.


- 4. In the **DEBUG** dropdown, select **Direct Debugging**, then select the **Start Debugging** icon.
## **Run the debugger**

After confirming that the debugger is attached, return to the Office application. In the **Debug Event-based handler** dialog, select **OK**.

You can now reach your breakpoints to debug your event-based activation or spamreporting code.

#### ) **Important**

Starting in Version 2403 (Build 17425.20000), event-based and spam-reporting addins use the **[V8 JavaScript engine](https://v8.dev/)** to run JavaScript, regardless of whether debugging is turned on or off. In earlier versions, the Chakra JavaScript engine is used when debugging is off, but the V8 engine may be used when debugging is turned on.

## **Stop the debugger**

To stop debugging the rest of the current Office on Windows session, in the **Debug Eventbased handler** dialog, choose **Cancel**. To re-enable debugging, restart the Office application.

To prevent the **Debug Event-based handler** dialog from popping up and stop debugging for subsequent sessions, delete the associated registry key,

HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\[Add-in ID]\UseDirectDebugger , or set its value to 0 .


## **Stop the local server**

When you want to stop the local web server and uninstall the add-in, follow the applicable instructions:

- To stop the server, run the following command. If you used npm start , the following command should also uninstall the add-in.

| command line |  |
|--------------|--|
|--------------|--|

npm stop

- If you manually sideloaded the add-in, see Remove a sideloaded add-in.
- Activate add-ins with events
- Implement an integrated spam-reporting add-in
- Troubleshoot event-based and spam-reporting add-ins
- Debug your add-in with runtime logging