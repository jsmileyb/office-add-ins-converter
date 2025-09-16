{0}------------------------------------------------

# **Build your first Word task pane add-in**

Article • 12/19/2024

In this article, you'll walk through the process of building a Word task pane add-in. You'll use either the Office Add-ins Development Kit or the Yeoman generator to create your Office Add-in. Select the tab for the one you'd like to use and then follow the instructions to create your add-in and test it locally. If you'd like to create the add-in project within Visual Studio Code, we recommend the Office Add-ins Development Kit.

Office Add-ins Development Kit

### **Prerequisites**

- Download and install [Visual Studio Code](https://code.visualstudio.com/) .
- Node.js (the latest LTS version). Visit the [Node.js site](https://nodejs.org/)  to download and install the right version for your operating system. To verify if you've already installed these tools, run the commands node -v and npm -v in your terminal.
- Office connected to a Microsoft 365 subscription. You might qualify for a Microsoft 365 E5 developer subscription through the [Microsoft 365 Developer](https://developer.microsoft.com/microsoft-365/dev-program) [Program,](https://developer.microsoft.com/microsoft-365/dev-program) see [FAQ](https://learn.microsoft.com/en-us/office/developer-program/microsoft-365-developer-program-faq#who-qualifies-for-a-microsoft-365-e5-developer-subscription-) for details. Alternatively, you can [sign up for a 1-month free](https://www.microsoft.com/microsoft-365/try?rtc=1) [trial](https://www.microsoft.com/microsoft-365/try?rtc=1) or [purchase a Microsoft 365 plan](https://www.microsoft.com/microsoft-365/buy/compare-all-microsoft-365-products) .

### **Create the add-in project**

Click the following button to create an add-in project using the Office Add-ins Development Kit for Visual Studio Code. You'll be prompted to install the extension if don't already have it. A page that contains the project description will open in Visual Studio Code.

#### **[Create an add-in in Visual Studio Code](vscode://msoffice.microsoft-office-add-in-debugger/open-specific-sample?sample-id=word-get-started-with-dev-kit)**

In the prompted page, select **Create** to create the add-in project. In the **Workspace folder** dialog that opens, select the folder where you want to create the project.

{1}------------------------------------------------

The Office Add-ins Development Kit will create the project. It will then open the project in a *second* Visual Studio Code window. Close the original Visual Studio Code window.

#### 7 **Note**

If you use VSCode Insiders, or you have problems opening the project page in VSCode, install the extension manually by following **[these steps](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/development-kit-overview?tabs=vscode)**, and find the sample in the sample gallery.

### **Explore the project**

The add-in project that you've created with the Office Add-ins Development Kit contains sample code for a basic task pane add-in. If you'd like to explore the components of your add-in project, open the project in your code editor and review the files listed below. When you're ready to try out your add-in, proceed to the next section.

- 1. The **./manifest.xml** or **./manifest.json** file in the root directory of the project defines the settings and capabilities of the add-in.
- 2. The **./src/taskpane/taskpane.html** file contains the HTML markup for the task pane.

{2}------------------------------------------------

- 3. The **./src/taskpane/taskpane.css** file contains the CSS that's applied to content in the task pane.
- 4. The **./src/taskpane/taskpane.js** file contains the Office JavaScript API code that facilitates interaction between the task pane and the Office client application.

## **Try it out**

- 1. Open the extension by selecting the Office Add-ins Development Kit icon in the **Activity Bar**.
- 2. Select **Preview Your Office Add-in (F5)**
- 3. In the Quick Pick menu, select the option **{Office Application} Desktop (Edge Chromium)**, where '{Office Application}' is the appropriate application, such as "Excel" or "Word". This will launch the add-in and debug the code.

The development kit checks that the prerequisites are met before debugging starts. Check the terminal for detailed information if there are issues with your environment. After this process, the Office desktop application launches and sideloads the add-in. Please note that the first time you run a project, it may make take a few minutes to install the dependencies. You'll need to install the certificate when prompted.

## **Stop testing your Office Add-in**

Once you are finished testing and debugging the add-in, *always* close the add-in by following these steps. (Closing the Office application or web server window doesn't reliably deregister the add-in.)

- 1. Open the extension by selecting the Office Add-ins Development Kit icon in the **Activity Bar**.
- 2. Select **Stop Previewing Your Office Add-in**. This closes the web server and removes the add-in from the registry and cache.
- 3. Close the Office application window.

# **Troubleshooting**

If you have problems running the add-in, take these steps.

- Close any open instances of Office.
- Close the previous web server started for the add-in with the **Stop Previewing Your Office Add-in** Office Add-ins Development Kit extension option.

{3}------------------------------------------------

The article Troubleshoot development errors with Office Add-ins contains solutions to common problems. If you're still having issues, [create a GitHub issue](https://aka.ms/officedevkitnewissue) and we'll help you.

For information on running the add-in on Office on the web, see Sideload Office Add-ins to Office on the web.

For information on debugging on older versions of Office, see Debug add-ins using developer tools in Microsoft Edge Legacy.