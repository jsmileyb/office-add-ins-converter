
# **Word add-ins documentation**

With Word add-ins, you can use familiar web technologies such as HTML, CSS, and JavaScript to build a solution that runs in Word across multiple platforms, including on the web, Windows, Mac, and iPad. Learn how to build, test, debug, and publish Word add-ins.

| About Word add-ins                                                         |
|----------------------------------------------------------------------------|
| e<br>OVERVIEW                                                              |
| What are Word add-ins?                                                     |
| f<br>QUICKSTART                                                            |
| Build your first Word add-in                                               |
| Explore APIs with Script Lab                                               |
| c<br>HOW-TO GUIDE                                                          |
| Use the Word JavaScript API to interact with document content and metadata |
| Test and debug a Word add-in                                               |
| Deploy and publish a Word add-in                                           |
| s<br>SAMPLE                                                                |
| Import a Word document template with a Word add-in                         |
| Manage citations through your Word add-in                                  |

### **Key Office Add-ins concepts**

#### e **OVERVIEW**

Office Add-ins platform overview

p **CONCEPT**

Core concepts for Office Add-ins


Design Office Add-ins

Develop Office Add-ins

#### **Resources**

i **REFERENCE**

[Ask questions](https://stackoverflow.com/questions/tagged/office-js)

[Request features](https://aka.ms/m365dev-suggestions)

[Report issues](https://github.com/officedev/office-js/issues)

Office Add-ins additional resources


# **Word add-ins overview**

Article • 05/29/2025

Do you want to create a solution that extends the functionality of Word? For example, one that involves automated document assembly? Or a solution that binds to and accesses data in a Word document from other data sources? You can use the Office Add-ins platform, which includes the Word JavaScript API and the Office JavaScript API, to extend Word clients running on the web, on a Windows desktop, or on a Mac.

Word add-ins are one of the many development options that you have on the Office Add-ins platform. You can use add-in commands to extend the Word UI and launch task panes that run JavaScript that interacts with the content in a Word document. Any code that you can run in a browser can run in a Word add-in. Add-ins that interact with content in a Word document create requests to act on Word objects and synchronize object state.

#### 7 **Note**

If you plan to **publish** your add-in to AppSource and make it available within the Office experience, make sure that you conform to the **Commercial marketplace certification policies**. For example, to pass validation, your add-in must work across all platforms that support the methods that you define (for more information, see **section 1120.3** and the **[Office Add-in application and availability page](https://learn.microsoft.com/en-us/javascript/api/requirement-sets)**).

The following figure shows an example of a Word add-in that runs in a task pane.


The Word add-in can do the following:

- 1. Send requests to the Word document.
- 2. Use JavaScript to access the paragraph object and update, delete, or move the paragraph.

For example, the following code shows how to append a new sentence to the first paragraph.

```
JavaScript
await Word.run(async (context) => {
 const paragraphs = context.document.body.paragraphs;
 paragraphs.load();
 await context.sync();
 paragraphs.items[0].insertText(' New sentence in the paragraph.',
 Word.InsertLocation.end);
 await context.sync();
});
```
You can use any web server technology to host your Word add-in, such as ASP.NET, NodeJS, or Python. Use your favorite client-side framework—Ember, Backbone, Angular, React—or stick with plain JavaScript to develop your solution. You can also use services like Microsoft Entra and Microsoft Azure to authenticate and host your application respectively.

The Word JavaScript APIs give your application access to the objects and metadata found in a Word document. You can use these APIs to create add-ins that target the following clients.


- Word on the web
- Word 2016 or later on Windows
- Word on Mac
- Word on iPad

Write your add-in once, and it will run in all supported versions of Word across multiple platforms. For details, see [Office client application and platform availability for Office Add-ins.](https://learn.microsoft.com/en-us/javascript/api/requirement-sets)

# **JavaScript APIs for Word**

You can use two sets of JavaScript APIs to interact with the objects and metadata in a Word document.

The first is the [Word JavaScript API](https://learn.microsoft.com/en-us/javascript/api/word). This is an application-specific API model that was introduced with Word 2016. It's a strongly-typed object model that you can use to create Word add-ins that target Word 2016 and later on Windows and on Mac. This object model uses promises and provides access to Word-specific objects like [body,](https://learn.microsoft.com/en-us/javascript/api/word/word.body) [content controls,](https://learn.microsoft.com/en-us/javascript/api/word/word.contentcontrol) [inline](https://learn.microsoft.com/en-us/javascript/api/word/word.inlinepicture) [pictures,](https://learn.microsoft.com/en-us/javascript/api/word/word.inlinepicture) and [paragraphs](https://learn.microsoft.com/en-us/javascript/api/word/word.paragraph). The Word JavaScript API includes TypeScript definitions and vsdoc files so that you can get code hints in your IDE.

The second is the [Common API](https://learn.microsoft.com/en-us/javascript/api/office), which was introduced in Office 2013. Many of the objects in the Common API can be used in add-ins hosted by two or more Office clients. This API uses callbacks extensively.

Currently, all Word clients support Word JavaScript API and the shared Office JavaScript API. For details about supported clients, see [Office client application and platform availability for](https://learn.microsoft.com/en-us/javascript/api/requirement-sets) [Office Add-ins](https://learn.microsoft.com/en-us/javascript/api/requirement-sets).

We recommend that you start with the Word JavaScript API because the object model is easier to use. Use the Word JavaScript API if you need to access the objects in a Word document.

Use the shared Office JavaScript API when you need to do any of the following:

- Perform initialize actions for the application.
- Check the supported requirement set.
- Access metadata, settings, and environmental information for the document.
- Bind to sections in a document and capture events.
- Open a dialog box.

# **Next steps**


Ready to create your first Word add-in? See Build your first Word add-in. Use the add-in manifest to describe where your add-in is hosted, how it's displayed, and define permissions and other information.

To learn more about how to design a world-class Word add-in that creates a compelling experience for your users, see Design guidelines and Best practices.

After you develop your add-in, you can publish it to a network share, an app catalog, or AppSource.

# **See also**

- Developing Office Add-ins
- [Learn about the Microsoft 365 Developer Program](https://aka.ms/m365devprogram)
- Office Add-ins platform overview
- Word JavaScript API reference


# **Build your first Word task pane add-in**

Article • 12/19/2024

In this article, you'll walk through the process of building a Word task pane add-in. You'll use either the Office Add-ins Development Kit or the Yeoman generator to create your Office Add-in. Select the tab for the one you'd like to use and then follow the instructions to create your add-in and test it locally. If you'd like to create the add-in project within Visual Studio Code, we recommend the Office Add-ins Development Kit.

Office Add-ins Development Kit

# **Prerequisites**

- Download and install [Visual Studio Code](https://code.visualstudio.com/) .
- Node.js (the latest LTS version). Visit the [Node.js site](https://nodejs.org/)  to download and install the right version for your operating system. To verify if you've already installed these tools, run the commands node -v and npm -v in your terminal.
- Office connected to a Microsoft 365 subscription. You might qualify for a Microsoft 365 E5 developer subscription through the [Microsoft 365 Developer](https://developer.microsoft.com/microsoft-365/dev-program) [Program,](https://developer.microsoft.com/microsoft-365/dev-program) see [FAQ](https://learn.microsoft.com/en-us/office/developer-program/microsoft-365-developer-program-faq#who-qualifies-for-a-microsoft-365-e5-developer-subscription-) for details. Alternatively, you can [sign up for a 1-month free](https://www.microsoft.com/microsoft-365/try?rtc=1) [trial](https://www.microsoft.com/microsoft-365/try?rtc=1) or [purchase a Microsoft 365 plan](https://www.microsoft.com/microsoft-365/buy/compare-all-microsoft-365-products) .

# **Create the add-in project**

Click the following button to create an add-in project using the Office Add-ins Development Kit for Visual Studio Code. You'll be prompted to install the extension if don't already have it. A page that contains the project description will open in Visual Studio Code.

#### **[Create an add-in in Visual Studio Code](vscode://msoffice.microsoft-office-add-in-debugger/open-specific-sample?sample-id=word-get-started-with-dev-kit)**

In the prompted page, select **Create** to create the add-in project. In the **Workspace folder** dialog that opens, select the folder where you want to create the project.


The Office Add-ins Development Kit will create the project. It will then open the project in a *second* Visual Studio Code window. Close the original Visual Studio Code window.

#### 7 **Note**

If you use VSCode Insiders, or you have problems opening the project page in VSCode, install the extension manually by following **[these steps](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/development-kit-overview?tabs=vscode)**, and find the sample in the sample gallery.

# **Explore the project**

The add-in project that you've created with the Office Add-ins Development Kit contains sample code for a basic task pane add-in. If you'd like to explore the components of your add-in project, open the project in your code editor and review the files listed below. When you're ready to try out your add-in, proceed to the next section.

- 1. The **./manifest.xml** or **./manifest.json** file in the root directory of the project defines the settings and capabilities of the add-in.
- 2. The **./src/taskpane/taskpane.html** file contains the HTML markup for the task pane.


- 3. The **./src/taskpane/taskpane.css** file contains the CSS that's applied to content in the task pane.
- 4. The **./src/taskpane/taskpane.js** file contains the Office JavaScript API code that facilitates interaction between the task pane and the Office client application.

# **Try it out**

- 1. Open the extension by selecting the Office Add-ins Development Kit icon in the **Activity Bar**.
- 2. Select **Preview Your Office Add-in (F5)**
- 3. In the Quick Pick menu, select the option **{Office Application} Desktop (Edge Chromium)**, where '{Office Application}' is the appropriate application, such as "Excel" or "Word". This will launch the add-in and debug the code.

The development kit checks that the prerequisites are met before debugging starts. Check the terminal for detailed information if there are issues with your environment. After this process, the Office desktop application launches and sideloads the add-in. Please note that the first time you run a project, it may make take a few minutes to install the dependencies. You'll need to install the certificate when prompted.

# **Stop testing your Office Add-in**

Once you are finished testing and debugging the add-in, *always* close the add-in by following these steps. (Closing the Office application or web server window doesn't reliably deregister the add-in.)

- 1. Open the extension by selecting the Office Add-ins Development Kit icon in the **Activity Bar**.
- 2. Select **Stop Previewing Your Office Add-in**. This closes the web server and removes the add-in from the registry and cache.
- 3. Close the Office application window.

# **Troubleshooting**

If you have problems running the add-in, take these steps.

- Close any open instances of Office.
- Close the previous web server started for the add-in with the **Stop Previewing Your Office Add-in** Office Add-ins Development Kit extension option.


The article Troubleshoot development errors with Office Add-ins contains solutions to common problems. If you're still having issues, [create a GitHub issue](https://aka.ms/officedevkitnewissue) and we'll help you.

For information on running the add-in on Office on the web, see Sideload Office Add-ins to Office on the web.

For information on debugging on older versions of Office, see Debug add-ins using developer tools in Microsoft Edge Legacy.


# **Tutorial: Create a Word task pane add-in**

Article • 01/16/2025

In this tutorial, you'll create a Word task pane add-in that:

- " Inserts a range of text
- " Formats text
- " Replaces text and inserts text in various locations
- " Inserts images, HTML, and tables
- " Creates and updates content controls

#### **Tip**

If you've already completed the **Build your first Word task pane add-in** quick start, and want to use that project as a starting point for this tutorial, go directly to the **Insert a range of text** section to start this tutorial.

If you want a completed version of this tutorial, visit the **[Office Add-ins samples](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/tutorials/word-tutorial) [repo on GitHub](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/tutorials/word-tutorial)** .

### **Prerequisites**

- Node.js (the latest LTS version). Visit the [Node.js site](https://nodejs.org/) to download and install the right version for your operating system.
- The latest version of Yeoman and the Yeoman generator for Office Add-ins. To install these tools globally, run the following command via the command prompt.

```
command line
npm install -g yo generator-office
```
#### 7 **Note**

Even if you've previously installed the Yeoman generator, we recommend you update your package to the latest version from npm.

- Office connected to a Microsoft 365 subscription (including Office on the web).


#### 7 **Note**

If you don't already have Office, you might qualify for a Microsoft 365 E5 developer subscription through the **[Microsoft 365 Developer Program](https://aka.ms/m365devprogram)** ; for details, see the **[FAQ](https://learn.microsoft.com/en-us/office/developer-program/microsoft-365-developer-program-faq#who-qualifies-for-a-microsoft-365-e5-developer-subscription-)**. Alternatively, you can **[sign up for a 1-month free trial](https://www.microsoft.com/microsoft-365/try)** or **[purchase a Microsoft 365 plan](https://www.microsoft.com/microsoft-365/business/compare-all-microsoft-365-business-products-g)** .

### **Create your add-in project**

Run the following command to create an add-in project using the Yeoman generator. A folder that contains the project will be added to the current directory.

command line

yo office

#### 7 **Note**

When you run the yo office command, you may receive prompts about the data collection policies of Yeoman and the Office Add-in CLI tools. Use the information that's provided to respond to the prompts as you see fit.

When prompted, provide the following information to create your add-in project.

- **Choose a project type:** Office Add-in Task Pane project
- **Choose a script type:** JavaScript
- **What do you want to name your add-in?** My Office Add-in
- **Which Office client application would you like to support?** Word


After you complete the wizard, the generator creates the project and installs supporting Node components.

# **Insert a range of text**

In this step of the tutorial, you'll programmatically test that your add-in supports the user's current version of Word, and then insert a paragraph into the document.

### **Code the add-in**

- 1. Open the project in your code editor.
- 2. Open the file **./src/taskpane/taskpane.html**. This file contains the HTML markup for the task pane.
- 3. Locate the <main> element and delete all lines that appear after the opening <main> tag and before the closing </main> tag.
- 4. Add the following markup immediately after the opening <main> tag.

```
HTML
<button class="ms-Button" id="insert-paragraph">Insert
Paragraph</button><br/><br/>
```
- 5. Open the file **./src/taskpane/taskpane.js**. This file contains the Office JavaScript API code that facilitates interaction between the task pane and the Office client application.


- 6. Remove all references to the run button and the run() function by doing the following:
	- Locate and delete the line document.getElementById("run").onclick = run; .
	- Locate and delete the entire run() function.
- 7. Within the Office.onReady function call, locate the line if (info.host === Office.HostType.Word) { and add the following code immediately after that line. Note:
	- This code adds an event handler for the insert-paragraph button.
	- The insertParagraph function is wrapped in a call to tryCatch (both functions will be added in the next step). This allows any errors generated by the Office JavaScript API layer to be handled separately from your service code.

```
JavaScript
```

```
// Assign event handlers and other initialization logic.
document.getElementById("insert-paragraph").onclick = () =>
tryCatch(insertParagraph);
```
8. Add the following functions to the end of the file. Note:

- Your Word.js business logic will be added to the function passed to Word.run . This logic doesn't execute immediately. Instead, it's added to a queue of pending commands.
- The context.sync method sends all queued commands to Word for execution.
- The tryCatch function will be used by all the functions interacting with the workbook from the task pane. Catching Office JavaScript errors in this fashion is a convenient way to generically handle uncaught errors.

```
JavaScript
async function insertParagraph() {
 await Word.run(async (context) => {
 // TODO1: Queue commands to insert a paragraph into the
document.
 await context.sync();
 });
```


```
}
/** Default helper for invoking an action and handling errors. */
async function tryCatch(callback) {
 try {
 await callback();
 } catch (error) {
 // Note: In a production add-in, you'd want to notify the user
through your add-in's UI.
 console.error(error);
 }
}
```
- 9. Within the insertParagraph() function, replace TODO1 with the following code. Note:
	- The first parameter to the insertParagraph method is the text for the new paragraph.
	- The second parameter is the location within the body where the paragraph will be inserted. Other options for insert paragraph, when the parent object is the body, are "End" and "Replace".

```
JavaScript
const docBody = context.document.body;
docBody.insertParagraph("Office has several versions, including Office
2016, Microsoft 365 subscription, and Office on the web.",
 Word.InsertLocation.start);
```
- 10. Save all your changes to the project.
### **Test the add-in**

- 1. Complete the following steps to start the local web server and sideload your addin.
#### 7 **Note**

- Office Add-ins should use HTTPS, not HTTP, even while you're developing. If you're prompted to install a certificate after you run one of the following commands, accept the prompt to install the certificate that the Yeoman generator provides. You may also have to run your


command prompt or terminal as an administrator for the changes to be made.

- If this is your first time developing an Office Add-in on your machine, you may be prompted in the command line to grant Microsoft Edge WebView a loopback exemption ("Allow localhost loopback for Microsoft Edge WebView?"). When prompted, enter Y to allow the exemption. Note that you'll need administrator privileges to allow the exemption. Once allowed, you shouldn't be prompted for an exemption when you sideload Office Add-ins in the future (unless you remove the exemption from your machine). To learn more, see **["We can't open this](https://learn.microsoft.com/en-us/office/troubleshoot/office-suite-issues/cannot-open-add-in-from-localhost) [add-in from localhost" when loading an Office Add-in or using Fiddler](https://learn.microsoft.com/en-us/office/troubleshoot/office-suite-issues/cannot-open-add-in-from-localhost)**.
#### **Tip**

If you're testing your add-in on Mac, run the following command in the root directory of your project before proceeding. When you run this command, the local web server starts.

command line

npm run dev-server

- To test your add-in in Word, run the following command in the root directory of your project. This starts the local web server (if it isn't already running) and opens Word with your add-in loaded.
command line npm start

- To test your add-in in Word on the web, run the following command in the root directory of your project. When you run this command, the local web


server starts. Replace "{url}" with the URL of a Word document on your OneDrive or a SharePoint library to which you have permissions.

#### 7 **Note**

If you are developing on a Mac, enclose the {url} in single quotation marks. Do *not* do this on Windows.

command line

```
npm run start -- web --document {url}
```
The following are examples.

- npm run start -- web --document https://contoso.sharepoint.com/:t:/g/EZGxP7ksiE5DuxvY638G798BpuhwluxCM fF1WZQj3VYhYQ?e=F4QM1R
- npm run start -- web --document https://1drv.ms/x/s!jkcH7spkM4EGgcZUgqthk4IK3NOypVw?e=Z6G1qp
- npm run start -- web --document https://contoso-my.sharepointdf.com/:t:/p/user/EQda453DNTpFnl1bFPhOVR0BwlrzetbXvnaRYii2lDr_oQ? e=RSccmNP

If your add-in doesn't sideload in the document, manually sideload it by following the instructions in Manually sideload add-ins to Office on the web.

- 2. In Word, if the "My Office Add-in" task pane isn't already open, choose the **Home** tab, and then choose the **Show Taskpane** button on the ribbon to open the add-in task pane.


- 3. In the task pane, choose the **Insert Paragraph** button.
- 4. Make a change in the paragraph.
- 5. Choose the **Insert Paragraph** button again. Note that the new paragraph appears above the previous one because the insertParagraph method is inserting at the start of the document's body.

| Office has several versions, including Office 2016, Microsoft 365 subscription, and Office on the web.             | My Office Add-in | > | × |
|--------------------------------------------------------------------------------------------------------------------|------------------|---|---|
| Office has several versions, لوطِطِلِ including Office 2016, Microsoft 365 subscription, and Office on the<br>web. | Logo             |   |   |
|                                                                                                                    | Welcome          |   |   |
|                                                                                                                    | Insert Paragraph |   |   |

- 6. When you want to stop the local web server and uninstall the add-in, follow the applicable instructions:
	- To stop the server, run the following command. If you used npm start , the following command also uninstalls the add-in.

- If you manually sideloaded the add-in, see Remove a sideloaded add-in.


### **Format text**

In this step of the tutorial, you'll apply a built-in style to text, apply a custom style to text, and change the font of text.

### **Apply a built-in style to text**

- 1. Open the file **./src/taskpane/taskpane.html**.
- 2. Locate the <button> element for the insert-paragraph button, and add the following markup after that line.

```
HTML
<button class="ms-Button" id="apply-style">Apply Style</button><br/>
<br/>
```
- 3. Open the file **./src/taskpane/taskpane.js**.
- 4. Within the Office.onReady function call, locate the line that assigns a click handler to the insert-paragraph button, and add the following code after that line.

```
JavaScript
document.getElementById("apply-style").onclick = () =>
tryCatch(applyStyle);
```
- 5. Add the following function to the end of the file.

```
JavaScript
async function applyStyle() {
 await Word.run(async (context) => {
 // TODO1: Queue commands to style text.
 await context.sync();
 });
}
```
- 6. Within the applyStyle() function, replace TODO1 with the following code. Note that the code applies a style to a paragraph, but styles can also be applied to ranges of text.
JavaScript


### **Apply a custom style to text**

- 1. Open the file **./src/taskpane/taskpane.html**.
- 2. Locate the <button> element for the apply-style button, and add the following markup after that line.

```
HTML
<button class="ms-Button" id="apply-custom-style">Apply Custom
Style</button><br/><br/>
```
- 3. Open the file **./src/taskpane/taskpane.js**.
- 4. Within the Office.onReady function call, locate the line that assigns a click handler to the apply-style button, and add the following code after that line.

```
JavaScript
document.getElementById("apply-custom-style").onclick = () =>
tryCatch(applyCustomStyle);
```
- 5. Add the following function to the end of the file.

```
JavaScript
async function applyCustomStyle() {
 await Word.run(async (context) => {
 // TODO1: Queue commands to apply the custom style.
 await context.sync();
 });
}
```
- 6. Within the applyCustomStyle() function, replace TODO1 with the following code. Note that the code applies a custom style that does not exist yet. You'll create a style with the name **MyCustomStyle** in the Test the add-in step.
JavaScript


```
const lastParagraph = context.document.body.paragraphs.getLast();
lastParagraph.style = "MyCustomStyle";
```
- 7. Save all your changes to the project.
### **Change the font of text**

- 1. Open the file **./src/taskpane/taskpane.html**.
- 2. Locate the <button> element for the apply-custom-style button, and add the following markup after that line.

```
HTML
<button class="ms-Button" id="change-font">Change Font</button><br/>
<br/>
```
- 3. Open the file **./src/taskpane/taskpane.js**.
- 4. Within the Office.onReady function call, locate the line that assigns a click handler to the apply-custom-style button, and add the following code after that line.

```
JavaScript
document.getElementById("change-font").onclick = () =>
tryCatch(changeFont);
```
- 5. Add the following function to the end of the file.

```
JavaScript
```

```
async function changeFont() {
 await Word.run(async (context) => {
 // TODO1: Queue commands to apply a different font.
 await context.sync();
 });
}
```
- 6. Within the changeFont() function, replace TODO1 with the following code. Note that the code gets a reference to the second paragraph by using the ParagraphCollection.getFirst method chained to the Paragraph.getNext method.


```
JavaScript
```

```
const secondParagraph =
context.document.body.paragraphs.getFirst().getNext();
secondParagraph.font.set({
 name: "Courier New",
 bold: true,
 size: 18
 });
```
7. Save all your changes to the project.

### **Test the add-in**

- 1. If the local web server is already running and your add-in is already loaded in Word, proceed to step 2. Otherwise, start the local web server and sideload your add-in.
	- To test your add-in in Word, run the following command in the root directory of your project. This starts the local web server (if it isn't already running) and opens Word with your add-in loaded.

- To test your add-in in Word on the web, run the following command in the root directory of your project. When you run this command, the local web server starts. Replace "{url}" with the URL of a Word document on your OneDrive or a SharePoint library to which you have permissions.
#### 7 **Note**

If you are developing on a Mac, enclose the {url} in single quotation marks. Do *not* do this on Windows.

```
command line
```

```
npm run start -- web --document {url}
```
The following are examples.


- npm run start -- web --document https://contoso.sharepoint.com/:t:/g/EZGxP7ksiE5DuxvY638G798BpuhwluxCM fF1WZQj3VYhYQ?e=F4QM1R
- npm run start -- web --document https://1drv.ms/x/s!jkcH7spkM4EGgcZUgqthk4IK3NOypVw?e=Z6G1qp
- npm run start -- web --document https://contoso-my.sharepointdf.com/:t:/p/user/EQda453DNTpFnl1bFPhOVR0BwlrzetbXvnaRYii2lDr_oQ? e=RSccmNP

If your add-in doesn't sideload in the document, manually sideload it by following the instructions in Manually sideload add-ins to Office on the web.

- 2. If the add-in task pane isn't already open in Word, go to the **Home** tab and choose the **Show Taskpane** button on the ribbon to open it.
- 3. Be sure there are at least three paragraphs in the document. You can choose the **Insert Paragraph** button three times. *Check carefully that there's no blank paragraph at the end of the document. If there is, delete it.*
- 4. In Word, create a [custom style](https://support.microsoft.com/office/d38d6e47-f6fc-48eb-a607-1eb120dec563) named "MyCustomStyle". It can have any formatting that you want.
- 5. Choose the **Apply Style** button. The first paragraph will be styled with the built-in style **Intense Reference**.
- 6. Choose the **Apply Custom Style** button. The last paragraph will be styled with your custom style. (If nothing seems to happen, the last paragraph might be blank. If so, add some text to it.)
- 7. Choose the **Change Font** button. The font of the second paragraph changes to 18 pt., bold, Courier New.


### **Replace text and insert text**

In this step of the tutorial, you'll add text inside and outside of selected ranges of text, and replace the text of a selected range.

### **Add text inside a range**

- 1. Open the file **./src/taskpane/taskpane.html**.
- 2. Locate the <button> element for the change-font button, and add the following markup after that line.

```
HTML
<button class="ms-Button" id="insert-text-into-range">Insert
Abbreviation</button><br/><br/>
```
- 3. Open the file **./src/taskpane/taskpane.js**.
- 4. Within the Office.onReady function call, locate the line that assigns a click handler to the change-font button, and add the following code after that line.

```
JavaScript
document.getElementById("insert-text-into-range").onclick = () =>
tryCatch(insertTextIntoRange);
```
- 5. Add the following function to the end of the file.

```
JavaScript
async function insertTextIntoRange() {
 await Word.run(async (context) => {
 // TODO1: Queue commands to insert text into a selected range.
 // TODO2: Load the text of the range and sync so that the
 // current range text can be read.
 // TODO3: Queue commands to repeat the text of the original
 // range at the end of the document.
 await context.sync();
 });
}
```


- 6. Within the insertTextIntoRange() function, replace TODO1 with the following code. Note:
	- The function is intended to insert the abbreviation ["(M365)"] into the end of the Range whose text is "Microsoft 365". It makes a simplifying assumption that the string is present and the user has selected it.
	- The first parameter of the Range.insertText method is the string to insert into the Range object.
	- The second parameter specifies where in the range the additional text should be inserted. Besides "End", the other possible options are "Start", "Before", "After", and "Replace".
	- The difference between "End" and "After" is that "End" inserts the new text inside the end of the existing range, but "After" creates a new range with the string and inserts the new range after the existing range. Similarly, "Start" inserts text inside the beginning of the existing range and "Before" inserts a new range. "Replace" replaces the text of the existing range with the string in the first parameter.
	- You saw in an earlier stage of the tutorial that the insert* methods of the body object don't have the "Before" and "After" options. This is because you can't put content outside of the document's body.

```
JavaScript
```

```
const doc = context.document;
const originalRange = doc.getSelection();
originalRange.insertText(" (M365)", Word.InsertLocation.end);
```
- 7. We'll skip over TODO2 until the next section. Within the insertTextIntoRange() function, replace TODO3 with the following code. This code is similar to the code you created in the first stage of the tutorial, except that now you are inserting a new paragraph at the end of the document instead of at the start. This new paragraph will demonstrate that the new text is now part of the original range.
#### JavaScript

doc.body.insertParagraph("Original range: " + originalRange.text, Word.InsertLocation.end);


### **Add code to fetch document properties into the task pane's script objects**

In all previous functions in this tutorial, you queued commands to *write* to the Office document. Each function ended with a call to the context.sync() method which sends the queued commands to the document to be executed. But the code you added in the last step calls the originalRange.text property, and this is a significant difference from the earlier functions you wrote, because the originalRange object is only a proxy object that exists in your task pane's script. It doesn't know what the actual text of the range in the document is, so its text property can't have a real value. It's necessary to first fetch the text value of the range from the document and use it to set the value of originalRange.text . Only then can originalRange.text be called without causing an exception to be thrown. This fetching process has three steps.

- 1. Queue a command to load (that is, fetch) the properties that your code needs to read.
- 2. Call the context object's sync method to send the queued command to the document for execution and return the requested information.
- 3. Because the sync method is asynchronous, ensure that it has completed before your code calls the properties that were fetched.

The following step must be completed whenever your code needs to *read* information from the Office document.

- 1. Within the insertTextIntoRange() function, replace TODO2 with the following code.

```
JavaScript
originalRange.load("text");
await context.sync();
```
When you're done, the entire function should look like the following:

```
JavaScript
async function insertTextIntoRange() {
 await Word.run(async (context) => {
 const doc = context.document;
 const originalRange = doc.getSelection();
 originalRange.insertText(" (M365)", Word.InsertLocation.end);
 originalRange.load("text");
```


```
 await context.sync();
 doc.body.insertParagraph("Original range: " + originalRange.text,
Word.InsertLocation.end);
 await context.sync();
 });
}
```
#### **Add text between ranges**

- 1. Open the file **./src/taskpane/taskpane.html**.
- 2. Locate the <button> element for the insert-text-into-range button, and add the following markup after that line.

HTML

<button class="ms-Button" id="insert-text-outside-range">Add Version Info</button><br/><br/>

- 3. Open the file **./src/taskpane/taskpane.js**.
- 4. Within the Office.onReady function call, locate the line that assigns a click handler to the insert-text-into-range button, and add the following code after that line.

```
JavaScript
document.getElementById("insert-text-outside-range").onclick = () =>
tryCatch(insertTextBeforeRange);
```
- 5. Add the following function to the end of the file.

```
JavaScript
async function insertTextBeforeRange() {
 await Word.run(async (context) => {
 // TODO1: Queue commands to insert a new range before the
 // selected range.
 // TODO2: Load the text of the original range and sync so that
the
 // range text can be read and inserted.
```


- });
}

- 6. Within the insertTextBeforeRange() function, replace TODO1 with the following code. Note:
	- The function is intended to add a range whose text is "Office 2019, " before the range with text "Microsoft 365". It makes an assumption that the string is present and the user has selected it.
	- The first parameter of the Range.insertText method is the string to add.
	- The second parameter specifies where in the range the additional text should be inserted. For more details about the location options, see the previous discussion of the insertTextIntoRange function.

```
JavaScript
const doc = context.document;
const originalRange = doc.getSelection();
originalRange.insertText("Office 2019, ", Word.InsertLocation.before);
```
- 7. Within the insertTextBeforeRange() function, replace TODO2 with the following code.

```
JavaScript
originalRange.load("text");
await context.sync();
// TODO3: Queue commands to insert the original range as a
// paragraph at the end of the document.
// TODO4: Make a final call of context.sync here and ensure
// that it runs after the insertParagraph has been queued.
```
- 8. Replace TODO3 with the following code. This new paragraph will demonstrate the fact that the new text is *not* part of the original selected range. The original range still has only the text it had when it was selected.

```
JavaScript
doc.body.insertParagraph("Current text of original range: " +
originalRange.text, Word.InsertLocation.end);
```


- 9. Replace TODO4 with the following code.

```
JavaScript
await context.sync();
```
### **Replace the text of a range**

- 1. Open the file **./src/taskpane/taskpane.html**.
- 2. Locate the <button> element for the insert-text-outside-range button, and add the following markup after that line.

```
HTML
<button class="ms-Button" id="replace-text">Change Quantity
Term</button><br/><br/>
```
- 3. Open the file **./src/taskpane/taskpane.js**.
JavaScript

- 4. Within the Office.onReady function call, locate the line that assigns a click handler to the insert-text-outside-range button, and add the following code after that line.

```
document.getElementById("replace-text").onclick = () =>
tryCatch(replaceText);
```
- 5. Add the following function to the end of the file.

```
JavaScript
async function replaceText() {
 await Word.run(async (context) => {
 // TODO1: Queue commands to replace the text.
 await context.sync();
 });
}
```
- 6. Within the replaceText() function, replace TODO1 with the following code. Note that the function is intended to replace the string "several" with the string "many".


It makes a simplifying assumption that the string is present and the user has selected it.

JavaScript const doc = context.document; const originalRange = doc.getSelection(); originalRange.insertText("many", Word.InsertLocation.replace);

- 7. Save all your changes to the project.
### **Test the add-in**

- 1. If the local web server is already running and your add-in is already loaded in Word, proceed to step 2. Otherwise, start the local web server and sideload your add-in.
	- To test your add-in in Word, run the following command in the root directory of your project. This starts the local web server (if it isn't already running) and opens Word with your add-in loaded.

command line npm start

- To test your add-in in Word on the web, run the following command in the root directory of your project. When you run this command, the local web server starts. Replace "{url}" with the URL of a Word document on your OneDrive or a SharePoint library to which you have permissions.
#### 7 **Note**

If you are developing on a Mac, enclose the {url} in single quotation marks. Do *not* do this on Windows.

command line

npm run start -- web --document {url}

The following are examples.

```
npm run start -- web --document
```
https://contoso.sharepoint.com/:t:/g/EZGxP7ksiE5DuxvY638G798BpuhwluxCM


fF1WZQj3VYhYQ?e=F4QM1R

- npm run start -- web --document https://1drv.ms/x/s!jkcH7spkM4EGgcZUgqthk4IK3NOypVw?e=Z6G1qp
- npm run start -- web --document https://contoso-my.sharepointdf.com/:t:/p/user/EQda453DNTpFnl1bFPhOVR0BwlrzetbXvnaRYii2lDr_oQ? e=RSccmNP

If your add-in doesn't sideload in the document, manually sideload it by following the instructions in Manually sideload add-ins to Office on the web.

- 2. If the add-in task pane isn't already open in Word, go to the **Home** tab and choose the **Show Taskpane** button on the ribbon to open it.
- 3. In the task pane, choose the **Insert Paragraph** button to ensure that there's a paragraph at the start of the document.
- 4. Within the document, select the phrase "Microsoft 365 subscription". *Be careful not to include the preceding space or following comma in the selection.*
- 5. Choose the **Insert Abbreviation** button. Note that " (M365)" is added. Note also that at the bottom of the document a new paragraph is added with the entire expanded text because the new string was added to the existing range.
- 6. Within the document, select the phrase "Microsoft 365". *Be careful not to include the preceding or following space in the selection.*
- 7. Choose the **Add Version Info** button. Note that "Office 2019, " is inserted between "Office 2016" and "Microsoft 365". Note also that at the bottom of the document a new paragraph is added but it contains only the originally selected text because the new string became a new range rather than being added to the original range.
- 8. Within the document, select the word "several". *Be careful not to include the preceding or following space in the selection.*
- 9. Choose the **Change Quantity Term** button. Note that "many" replaces the selected text.


# **Insert images, HTML, and tables**

In this step of the tutorial, you'll learn how to insert images, HTML, and tables into the document.

### **Define an image**

Complete the following steps to define the image that you'll insert into the document in the next part of this tutorial.

- 1. In the root of the project, create a new file named **base64Image.js**.
- 2. Open the file **base64Image.js** and add the following code to specify the Base64 encoded string that represents an image.

```
JavaScript
export const base64Image =

"iVBORw0KGgoAAAANSUhEUgAAAZAAAAEFCAIAAABCdiZrAAAACXBIWXMAAAsSAAALEgHS3X
78AAAgAElEQVR42u2dzW9bV3rGn0w5wLBTRpSACAUDmDRowGoj1DdAtBA6suksZmtmV3Qj+
i8w3XUB00X3pv8CX68Gswq96aKLhI5bCKiM+gpVphIa1qQBcQbyQB/hTJlpOHUXlyEvD885
vLxfvCSfH7KIJVuUrnif+z7nPOd933v37h0IIWQe+BEvASGEgkUIIRQsQggFixBCKFiEEEL
BIoRQsAghhIJFCCEULEIIBYsQQihYhBBCwSKEULAIIYSCRQghFCxCCAWLEEIoWIQQQsEihC
wQCV4CEgDdJvYM9C77f9x8gkyJV4UEznvs6U780rvAfgGdg5EPbr9CyuC1IbSEJGa8KopqB
WC/gI7Fa0MoWCROHJZw/lxWdl3isITeBa8QoWCRyOk2JR9sVdF+qvwnnQPsF+SaRSEjFCwS
Cr0LNCo4rYkfb5s4vj/h33YOcFSWy59VlIsgIRQs4pHTGvYMdJvIjupOx5Ir0Tjtp5K/mTK
wXsSLq2hUWG0R93CXkKg9oL0+ldnFpil+yhlicIM06NA2cXgXySyuV7Fe5CUnFCziyQO2qm
g8BIDUDWzVkUiPfHY8xOCGT77EWkH84FEZbx4DwOotbJpI5nj5CQWLTOMBj8votuRqBWDP8
KJWABIr2KpLwlmHpeHKff4BsmXxFQmhYBGlBxzoy7YlljxOcfFAMottS6JH+4Xh69IhEgoW
cesBNdVQozLyd7whrdrGbSYdIqFgkQkecMD4epO9QB4I46v4tmbtGeK3QYdIKFhE7gEHjO/
odSzsfRzkS1+5h42q+MGOhf2CuPlIh0goWPSAogcccP2RJHI1riP+kQYdVK9Fh0goWPSAk8
2a5xCDG4zPJaWTxnvSIVKwKFj0gEq1go8QgxtUQQeNZtEhUrB4FZbaA9pIN+98hhhcatbNp
qRoGgRKpdAhUrDIMnpAjVrpJSNApK/uRi7pEClYZIk84KDGGQ+IBhhicMP6HRg1ycedgVI6
RELBWl4POFCr8VWkszpe3o76G1aFs9ws+dMhUrDIInvAAeMB0ZBCDG6QBh2kgVI6RAoWWRY
PqBEI9+oQEtKgg3sNpUOkYJGF8oADxgOioUauXKIKOkxV99EhUrDIgnhAG+mCUQQhBpeaNb
```


4JgOn3AegQKVhkvj2gjXRLLrIQgxtUQYdpNYsOkYJF5tUDarQg4hCDS1u3VZd83IOw0iFSs MiceUCNWp3WYH0Wx59R6ls9W1c6RAoWmQ8PaCNdz55hiMEN4zsDNhMDpXSIFCwylx5Qo1a9 C3yVi69a2ajCWZ43NOkQKVgkph5wwHi+KQ4hBs9SC9+RMTpEChaJlwfUFylWEafP5uMKqII OPv0sHSIFi8TFAzpLiXxF/KCbdetEGutFUSa6TXQsdKypv42UgZQhfrWOhbO6q8nPqqCD/z U4OkQKFpm9B7SRbrTpQwzJHNaL/VHyiRVF0dfC2xpOzMnKlUgjW0amhGRW/ZM+w5sqzuqTN Wtb9nKBZDLoEClYZGYe0EYaENWHGDaquHJv5CPnz/H9BToWkjmsFkTdOX0GS22p1ovYNEdU r9vCeR3dJlIG1gojn2o8RKPiRX+D0iw6RAoWmYEH1HioiQZqq47VW32dalUlfi1fQf7ByEd UQpMpYfOJ46UPcFweKaMSaWyaWL8z/Mibxzgqe3G4CC6pT4dIwSLReUCNWrkJMdjh8sMSuk 1d3bReRGb3hy97iS/SEl+5bQ0LqM4B9gvytaptC6kbwz++vD3ZG0r3EBDoWUg6RAoWCd0D9 isXReTKTYghZbhdUB/UYlKV2TSHitZtYc9QrqynDGy/GnGg+4XJr779ShJ0gNdAKR3i/PAj XoIZe8BGBS+uhqtWAF4VXUWu3G//ORVqdVRiEumhWgFoVHT7gB1LnFAvVaJxYZJ+qx/XRuo 1X0+RFqzPsF/QFZuEgrVcHnDPCGbFylnajN/wAZZvqgpR8IzO275tTvjnwl/4sORC6C9xWJ LoYCKNrbpuR3Jazp/jxdUJmksoWIvvAfcLsD4LuLfn5hOJhWlVQ+lyNZDFcUl636GY5/Wpy zo3FRZ+WBeT1JhpGDVlIMMbjYfYM3Ba4zuXgkUPGBD5B5Kl6LaJ4/uh/CCDTvDjW4ROxZm4 gj7+dwZLY24067AkF9OtesCaRYdIwaIHDIzMrmSzv2NNTgl4fLlSXw6kjs8pWN+FfHu3n8p /xpSBjWrwL0eHSMGiB/TL+h1JnNJ+xTA6MawXh1ogTWA5S5tvLS8vMVUM6s1j+TKZEASjQ6 RgkVl6wH4pcUM+zs8qBq9WyRyMGozP+5J0/nzygrrLSkS4ONPmNg/vyr1npiQG9+kQKVhkB h5woFbSI8EuQwxTkS1j2xoG0zsHeBVcRsl/RNMqyoMOG9WRjAUd4pzD4GhoHjDsMIEqchX4 8JuUgU1zJN+kSa4D+LnjHfXiqqsa5Oejb8J/fs9TAZjFtiXXvgADpaqXZsqUFRY94NRq1ag ErFbrRWzVR9Tq9JlOrWy75NncCf982n+o+sYCDJTSIVKw6AGnRhoQbZsBv3S+MlyxAtC7xP F9WMUJDsi5M+gmVCWImpvolorOgXzTMPBAKR0iBWvuPWB4+4CiWj2Rz3MPcFSXHb90Nmawb WDLRVZAc2pHZTkF2fWDKugQRqBUCvcQKVj0gI6qRxYQtfvGBIUdvHQ2fmk/VR7fk5Q5jr+2 fmfygrpTfM+fu8qa6lEFHcIIlGocolWkQwwcLrr79oBB9YRxg7SDXbDjJISue71LHJWnrno +vRh+BX2Xq2QOO6+Hf3TTXsYl43M3BhVcZFNjEyvIluUNvAgrrIX1gINqRdpvM0C1EhatbB vowaM5neOVe/L2VX176/jip88CUysAhyV5SRheoFRSfV+i8RAvckH+XKyweBW8qNWeEelEP 1XkKqgQw3j/T3sxyNv6cSKNm02xA3KrOvLV1gq4Xh1u3vUusWcE7KESK7jZlHvSoDqU+q/4 CAUrItomWtUoRvup1KpRCWxb0KiNqFXvcoreWCem/ETh+ILRYJnvJzlxz+7wrt/l9qkuHUI IrMk9bxaZEjIltl2mYMWDjoVWFae1sAouVeQq2LUYZwfRaVG1dR9PnKp802EpxG016TCOgZ sOb6tk9RayZVZVFKwZ8cff4b/+Htcq8sd17wInJt5UA17SUqnVWR0vbwf5Qn5KgPO6bo0mU 0K2LJetbgtvqjgxQw8uqcbthDH+OrHS/5FV19MuJDXreoSCFQC9C3yxisQK8hVk1dteZ3W8 qQY2VFm68OF/emj0JNJ430DKQCKN3gU6FrrNSHf9VaMrfI68F+ynXVKpkhxndRyX0TlQzv4 hFKyABWuwMPGROWxiJ6kdmmibaJu+7gTpPRbgDbZsqJa9/T8AMrvIlnWx/m4Tx+XhY4yC5R XGGjzRbeHlbd3ZsWQO+Qp2mth84nFtSBoQtS0M1cobqqCD50BpMovrj/Dpufyk1OBXZueKg yq6KVjEI/bZMf3ef6aErTp2XiOzO8UtIe0gCuCoHMWm5MLWyJfK09HTdihdvwPjc+w0J4wv bJv4KhfF2VIKFnHLm8f4KjfhkF0yh00TN5vYfDJ510wVED0qR7ENv7Sa5SZQmlhB/gF2XsO oTdj+O6tjz8Dh3Tlbaow9XMNy/153rGGpDIJ+Ycv5bm6bcvVR5YaiPFCy8Kze6s+4lj4VpI HS1Vv4sORqa09YrlL5fa5hUbBmLFiDd/am6Soi0LtAqzqyMK9Sq8BDDEQVdMBooDSxgvXih AV14RfqxgBSsChYcREsmyv3lImtcU5raJs4q8sjV/MYYpgLrj9SxlP2C/iuiXxFl1EYL4GP ym5/TRQsCla8BKu/3qFNbLl80a9yVKuwUIWzpmKQrnIPBcsrXHQPT+AucXzf70l91lahclT 2FV7tNmEV8fI2t24jI8FLEC52Ysv9wpbAtsVLGNNy2+VyFWGFNX+4SWyReYHpKgrWUuAmsU XiDNNVFKwlsxJBLGyRGVh7LlfFAq5hzeTd38LL27oo0ABpnykSIG766pzWYH3GS0XBWvJr7 yLg8/1F1J18l4pk1lXuhM1CaQkJPixN/jvXKlGMpVpa8u7CvSkj9CGshIIV92e7tOvxeBXG hGFIrN6Sp0ZPa5Jw1gfsdEzBWmbGb4BuE4d3JbdKtszHe1jllZTjsqTBvJtymFCwFpbxpRM 77nAouzE+MnnBAiazK++rYZ9Flw4B4mODgrWkpG5I1nHf1gDFrPa1gveRNmQc+5jnOL2L/p DqzoGkN2mArpChFgrWXD3eS5J38KDJjDTKsMG4aaDlrXTjr1UdJkJPTLpCChYBAEmzSqcHO X8utySZXV65AFBFGezjgULBS1dIwaIflDzehVVeVZHFiIN/VFEGoZtVtyUxbtwrpGDNDb3f heUH26Z4Nq3bkhw5TKT9dtciqihDtynpWN2mK6RgzS/vemH5QemU9kZF0tohX6Er8VteSTm WPQlOZa5w4gwRQsFaZD/Yu5APLOhdyvs6XOfqu+faVhFlOKsrfwXjRRZHzFOwlumeKbkqr2 xaVUmOdL3IiEPA5ZXmhPn4b2edy1gUrOVh/O2uaY/Vu2TEITi1eiCPMrRNnD9XC9Yz0Zgnc 3SFFKxl9YPd5oT+Su2nkgQjIw7TklhR7ldMbOBzQldIwVpOxu+Z8SWScY7K8iKLEQf3bFTl UYZWdZjXVT4zTLrCGD16eAlm6QfdCJZ9WEdYLbYjDmG3FU/mRqoJD90EV3+Ga//o5aUPS77 m2QiFrbQm6l24+ok6B+g2R0pj2xWy9SgFa6HV6o74kO9Ykx/vNsdlyficfGVkanRIgpV/4E uw3v/E4xZBMheYYKn2VZ0HcfS0quK6YaaE4/t8U9MSLlN55X4aRedAXouxVZab54Q0ytBtT nH933KvkIJFwdIEGsaRVjeZEiMOHsurRmWKyTfdlrj1wb1CCtZy+cHT2nSjorotuWbFvMj6 w6/xhxN81xL/G/zsvY7ks384wfdBDHBURRmkB3EmukIBHpOaBVzDmlF55Wa5ffyeyZZF4Vs


rILM79e0XGb/5JX7zS8nHt+r92rDz79gvhPPWVkcZpF0S9cgTpHf51maFtQSCpTqOo0d1WC fPQRUyVFGGs7ouKaq5+IJmJdJYv8PLTMFaDj/ojcZDyd5ZMkd7IqKKMsDHqEcGsihYS+oHT 0zvX016v3FQhYBqrV1/EGeCKxw7pkPBomAtGokV8W3dbXq/Z6A4rMNpYE5Wb8mjDPA9SZuu cOb3Ey9B6OVVUH5wwFEZW3Xxg5kSTkxfUmjj/MrCdz7+ovpvclxYo2HTVKqVz5xtqyo6zfW il+VIQsGaGz/4xnevBelhHQD5Cl7eDqA88fCpcX6cns0Fv3JPHmUQWrZ7Y/yYDvcKaQkX2Q +6P46j5+uS5IN2xCEO9C7xrTWbC36toiyOpgq+KS25SVfICmtpyqsTM5ivbA/7HN8Iy1emj qQKOGu0lIHrj+SfEhD+5mFJ0t85AlQDJrrNwA6Kt01xuZCukIK1sILlIS+qolGRLJDZEQc/ N6dmxqfmU85dufbTANbpPKCa3wXfa+3Co6JjIWX4coWzWt2jJSRT+EGftc/4nSNdlMmWo86 R5ivDg3XdlryBVwR8ZCrVIdiTACdjrnBaJx7g24CCRcIqrwKvO1pVifNKpCPtoZwyRlrQfD 0jM6iJMgQuoEyQUrAWX7B6F8ELVu8S38jMTqYUXS8BZ4ag8VBnGyP7NgQb6z/qMX7ZhV/le pGnoyhYMeP/vouRHxzw5rG80V0008CcZrBzEORS0VSoogxQDBz0D6fpULAWSrAi8IPDukYm E2uF0LfbBTPooQVCIGiiDG0zrEbG7ac8pkPBWiCEwEG3GeLOd/up3IiFXWQ5Xdjx/ZntfKm iDEC4FR9dIQVrQUhmxQXgsLf5pXem0JE9PDN4/jyAELnnS62JMoTa8P7EpCukYC0EH4QZv5 JiH9YZJ6SIg9MM9i5nZgY1VWQgB3EmXnNh9ZCCRcGaSz4cvYE7VhQjoaSHdUKKODjNYIDzu KZl9ZZSI76pRJF1oiukYC2CH3TGoBHccRw99mGdcQKPODjN4Omz2YTabVRa3G3izeMovoHx c+wssihYc+8H30Z1Szcq8tBmgKvv8TGDmV3xweC8DtEwPk2HgkXBmm8/eFoLd+lXuH+kCzc BRhycZtAqzibUDiCxoiyvzuqRjuQQyuf1Ilu/UrDm2Q9G7Jikh3WCKrKcZvDN41BC7X/+Nz Bq+Nk3yurJZnx6UPTllap8/oBFFgVrfv1gxILVu5QfnUvmcOWe3y8+CBB0DuRHgvyI1F//C p9+i7/6Bdbv4E/zuv5/yayyH3QYB3EmVrXCr/jDEu8DCtZ8+sG2OYNz+e2n8m27a76ngQ3+ eYDtrlZv9UXqp3+BRMrVP9FUi1/PQiwEwUoZdIUULPrBaZAeoAtqUEXj4SzbOWmiDG0zuuV C4bcsyDddIQVrDhCO43iblhrMLfRMmSP1+fCP4ITz//4WHUuZ7dpQJ0VndfR6vHkDXSEFa/ 4E68Sc5Tejuns/Mn3dmVY4tUOvg9//J379C/zbTdQ/wN7HcsHSRBla1dmUV3SFFKy5JHVD7 HAS9nEcPefP5YZ0rTDd8BtBBIMKtf/oJwDwP/+N869w/Hf44n3861/iP/4WFy+U/0QTZfB/ EGe9qOyo5bKkFa4MXWE4sKd7OOVVtxnFcRw9x2X5cs+miRdXXX2Fb62RwRMB5hga/4Df/2o 6+dNEGfwfxLle7ddEnqOwp7WRY9gfliJK27PCIh4f0YJDmTmqwzruIw69C5zVh/8FyG//aT q10nRl8H8QJ1/pq1VmVzKIyCXCpaYrpGDNkx98W4vFN3ZUlucPrlXm7JhueE2vEukRKfS8k do5EDdPPWsfoWBF6gfP6gEvAKcM5Cv9/zIl5a0rKZEu5bVeUBGHaFi9pbz5/R/E2aiOaHcy 611oTkwKVti89+7dO14Fd49QC3sfyz+183qkwjosBXacba2AfEVcJrdlSHUKR9SmFdxsyjX uRW6WO2vu+eRL5USc/YKvaHvKwPYriZV+kfPy1ZJZ7Iz63D1DuZT5c953rLBi4gcDyYsmc9 g08cmXkk29xAryD3CzqbyNBXVTzbnyE3GIrnrdVf6YpzW/B3Gc247dVl++PRdZ3Za40qf5O rM6N07Boh8U7yKfO1a2VO28njCeM7GCT750dWupDuv4iThEQ2JFZ119TsRZL478+F+Xhsth nv2ysPSu6TbzLYc/U7BmgvCm9Bm/ShnYtiRS1TlA4yEaD3H+fEQQN5+46imq2q3fqMb62mb Lyvld/g/iOM8k2mcDBl/Tc5ElFNfJXHQDIilYxIVa3Rm5o3wex0kZ2KqL+3ftp3hxFXsGGh U0Ktgv4Is0Xt4eytaVe5MrAlXT95Qx9Zj1yNBEGXoXk+c5pwydZR5EGWzXPCjWfBZZvUvxi cWldwrWbHjXm1xe+Vy92jRH1KpzgL2P5U3Tz+ojp2TyD5SVyADV9r+wTRYfNFGGVnWC706k YdTwyZfYqktkS4gytKrDKzxw9EEVWexBSsGaDb3fTRYsP3lRofl65wD7BV1fBGFH302RJbW rwt0bEzRRBjcHca79UECt3pLIllOju60RKXd+cW9F1umzkQV1ukIKVoz8oLME8Hkcx6l9vU vsFyZvJDnv29XC5JdQFVlOfxSf8krFUXlCeZXMiWLnlC3BBY+30BqUb56LrBO6QgpWHAUr0 OV2Z49NVUJdoGMNb103iqNq+o7wx0RPV2yqowzd5uSMW7eJPUOymDiQLWc1NL6057/Icr9X SChY8ypYmnUQvWYNcBPLUk3WEfb4Z0ggUYZuE1YR1meSWmxgBp1r7SrF8VZkdQ5Glh2Tubj HRyhYS+cHO5bfXXan9LhPFTrvBDfHiVWHdRCbiIMmynBWn24T9rSGr3LKo9HfXygX9Z11nL ciS7jIbOlHwYpXeeW/PcP3DpHSz4xRlVQu+x84N8WcxCHikFjR7QB4OOdsByBe3pYsLyaz2 H6FTVOuj4PX8lZkveVeIQUrzoI10cQl0hNaxDkrLDfbdon0yMKT+0Mqvcv4Rhw2qsqqx89B nLM69gx5CZzZxc5ryev6LLKEGauJdGCjISlYxK8fnHgcZ72Im01dh1+MtsfL7E7OVW1UR/b LT8wpvn/VYZ3ZRhxSN3S1jM+DOGuF4b6EcFoAwJV7uNkUk1+DqtlbkSUU3SyyKFhzU14Zn/ crF826eO9iZP9r09S1kcmWR+zb6bOpl/xVh3VmGHHQ7FT6b9k+qJJ6l3hVxJ4h7jYOjpQPt KljDWs6D0UWE6QUrFiQWBl53gpCI7d7Pyyg6B/UDUer39Vb2KpLNCuRxkYV1x+NfHEPjX1V h3Uwo4jD+h2lmvufiOM85m235ek2cVjCy9uizUysYPMJdn6QLT8rWcI0HbpCCtZ8lFdOd5C 6oSuy7LvIaZGcD/y1AjIlbFsjDY57l97HmqpM1kwiDvryymcDDLuNcrclbpKe1bFfwOFd8e sns9h80k9s+SmyGMgKGjbwc81ZvT+Rwfh85J3npodcIo2bzb4rPH+O/cIEQRQOFWqe4frjO xPZfCIvHAY/bDTkHyjlwE6BBjVAO5nTLd7lH8i+gdbQIx/endp6f3o+LJN7F/hitf//mq6E hBVWkH7QqVbdpqutK2d4WjO7eFCyfZVD4+GEgz7+1QrqoMBaIbqIw8QoQ1BqBXXyw3adL65 KfpvOFT2fK1l0hRSsOfCD475m05zwdLXvnz0DL66i8VByx3YOsGcEMDJeOPo7UvVENahCE2 VwcxAnQLpN7Bfw8rZygd/DShb3CilYMRKsN67Xp3sXw/Upu1mopn2KfXzXqGHnNfIPROGwT WVQM01VveGTuSgiDvoog+cpgT69/4scju8HU9kJx3TWi3M2ryhmcA1rmvexVcSnjntbM5ZC


xaY5YrXsjaSOhY6FRBopA8kcUoauIUnjod8tM0kxpVhC6l0o85ZBoVnKiXgdTeJV09iojvy +vM2nEC6vPaOEa1gUrNAFq22OpNWPyl5GeAqa5Z7z52hUAh5oOkAY/DOgbeLwbmjl6h0Yak /tcyJOYDWggY1qf9vUw6I7xqbpnNZgfUbBoiWM3A96a89wWJrabpw+w8vb2C+EpVZQr75nS iFGHDRRhrYZC7Wy6+j9AqzPvKRzB3WZc7WRrpAVVhRc/AvSPxOfk37sxnoRawUkc0ikJR6w 28J5HWd1nNYiGgm1/Up+cigka3blnq4/xLzMTPT2wx6WkCmxwqJghcnvj/DTDXElItgVk/c NAPjWms3QOjtbr6oKA/5h1eNdAbSqOL6/UG+exMrI6udpDYk0BYuCFSZ//B3+5M/6/9+7wF e5IPNBMUG1sBJsehPA9Ue6iTgLeW2FvHHHcttEiDjgGpZrBmqFIKalxhPVYZ1gIw6a+V0I4 iBOPBEie1QrCtbM3nwLQ+dAua6cLQfWxeEjU/mpbhONh4t5bdtPOZ6egjULuk1f01Jjjqrp eyLtfYC7k9VburWbwCNmfM5RsFheLbQcqyfrCJMTvaFpu9qxIj2IEz0nJu8eClb0tf2iv+1 Uh3Xgu1XWlXu6TqpH5QW/sOfPAztQRcEiruhYvqalzgW9S3yjsGZrBe/9BhIruKZ2fGf1uC RFWZ5TsFjVzxlvHitrAc9FluawN3y3bGd5TsEiEt4uzRNStf6dzMkb3enRRxna5uLXrf0K/ SCApkAULOK2nl+k8yITaoGnyqOL2fLUp+E+Mr2II4t0QsHyJVhLhUpH7L4r7pkYZViex8BS FekULApWpGgm60wVcdCom7N59JLQbXHp3TMJXgK3vOvBqKF3gY6FbhPdJr5rLn5p8HVppJe Tk+tVV10c9ONjF/UgzshNtoKUgR+nkTKGbRqJJ3j42f8Ds4luEx2rr2XfX6BjLdRNqJqsA8 AqTgj967sydJt4cXWh3gypG8M2DKsFAGzJQMGaE2wzdV7v/3/vYl43wpJZbFty0ZmoOJr5X Qiha02U1+QnOSRz/ZbWdmsgTWiDULDmkt5Fv93VfPlKje40KsrjykJr4HFBn23Lds9ujoaO gkVfGWtfqXF2mvZVQgcogZi0bKebo2CRBfSVmo7G0gahmv6lsy2v6OYoWMuL7ewiftPPyle qJutA1oJd1SFe9fcXz83ZD5vvmlPPXiUUrBBpm8Pooz1gZmAr7LtlYXylZiqXUDFldnVtZA IfHTZbN6e67IkVZMvIllm+UbDiR6uKRkWuDs5HfTI39CPz6Cs10/QGa1L6KIOf4ayzdXNTF baZXWxUKVUUrBhjh7bdJyHt289pW+LvKzUrU4OIgz7KoNlVjJub8ybxmV3kK9xJpGDNj2wd lX3Fi2LuKzV7f0dlvK3pogzjW4rxdHOef3H5CvcWKVhzSLeJ43KQrd/j4yuTOeUqsl21ae7 YjoXT2tyUk1N51Y9MShUFa845q6NRCTdtNFtfGc9rjgiDIMks8hXuA1KwFojTGo7LUcfZZ+ srI3Nz3/3g6aKP2nITkIK1yLRNHJVnHF6fua/06eZsVYrDYaYr93CtQqmiYC00024jRkZMf KUtSQM3B8RxLAU3ASlYSydb31Tw5vEcfKsh+cqZuznPV2OjyhHzFKylpNtEozKXzVXc+8p4 ujkPpG7gepWbgBSspSeCbcRoGA+LzkX3GDdmmZuAsXpc8hLMkrUC1uo4q+Pr0nINYpiLQjJ b1kX2ySzgEIp4yNZOE5tPkMzyYsSlYLzZpFpRsIiaTAnbFvIPph75R4L8Lexi5/WEIdWEgk UAIJFGvoKbTS+jlYlPVm9h5zU2TUYWKFhketnaeY3MLi9GRFL1yZfYqlOqKFjEK8kcNk1sv +qHoUgoFzmLzSfYqjOyQMEiQZAysFXHJ19OMWaZuCpjV3D9EXbYv5iCRQJnrYBti9uIgUmV vYzBIcUAAAIqSURBVAmYLfNiULBIaGRK2GlyG9HfNdzFtsVNQAoWiYrBNiJlayq4CUjBIjM yNWnkK9i2uI3oVqq4CUjBIjPG3kbcec1tRPUlysL4nJuAFCwSJ9mytxEpWyNF6Ao2n2CnqZ yXQShYZGasFbBV5zZiX6rsTUDmFShYJNbY24jXHy3venxmt39omZuAFCwyH2TLy7iNuH6nv wlIqaJgkXmzRcu0jWhvAho1bgJSsMg8M9hGXL+zoD9gtp9X4CYgBYssjmwZtUXbRrQPLe80 KVUULLKI2NuIxudzv41obwJuW9wEpGCRRWe92O/FPKfr8VfucROQgkWWjExp/rYR7c7FG1V KFQWLLB+DXszx30a0NwF5aJlQsChb/W3EeMpW6gY3AQkFi4xipx9itY1obwJuW5QqIj5keQ kIEJuRrhxfSlhhkSlka4YjXTm+lFCwyNREP9KV40sJBYv4sGY/bCNeuRfuC63ewvYrbgISC hYJQrY2qmFtIw46F6cMXmlCwSIBEfhIV44vJRQsEi6BjHTl+FJCwSLR4XmkK8eXEgoWmQ3T jnTl+FJCwSIzZjDSVQPHl5JAee/du3e8CsQX3Sa6Y730pB8khIJFCKElJIQQChYhhFCwCCE ULEIIoWARQggFixBCwSKEEAoWIYRQsAghFCxCCKFgEUIIBYsQQsEihBAKFiGEULAIIRQsQg ihYBFCCAWLEELBIoQQChYhhILFS0AIoWARQkjA/D87uqZQTj7xTgAAAABJRU5ErkJggg==" ;

### **Insert an image**

- 1. Open the file **./src/taskpane/taskpane.html**.
- 2. Locate the <button> element for the replace-text button, and add the following markup after that line.

HTML


- 3. Open the file **./src/taskpane/taskpane.js**.
- 4. Locate the Office.onReady function call near the top of the file and add the following code immediately before that line. This code imports the variable that you defined previously in the file **./base64Image.js**.

```
JavaScript
import { base64Image } from "../../base64Image";
```
- 5. Within the Office.onReady function call, locate the line that assigns a click handler to the replace-text button, and add the following code after that line.

```
JavaScript
document.getElementById("insert-image").onclick = () =>
tryCatch(insertImage);
```
- 6. Add the following function to the end of the file.

```
JavaScript
async function insertImage() {
 await Word.run(async (context) => {
 // TODO1: Queue commands to insert an image.
 await context.sync();
 });
}
```
- 7. Within the insertImage() function, replace TODO1 with the following code. Note that this line inserts the Base64-encoded image at the end of the document. (The Paragraph object also has an insertInlinePictureFromBase64 method and other insert* methods. See the following "Insert HTML" section for an example.)
JavaScript

```
context.document.body.insertInlinePictureFromBase64(base64Image,
Word.InsertLocation.end);
```


### **Insert HTML**

- 1. Open the file **./src/taskpane/taskpane.html**.
- 2. Locate the <button> element for the insert-image button, and add the following markup after that line.

```
HTML
<button class="ms-Button" id="insert-html">Insert HTML</button><br/>
<br/>
```
- 3. Open the file **./src/taskpane/taskpane.js**.
- 4. Within the Office.onReady function call, locate the line that assigns a click handler to the insert-image button, and add the following code after that line.

```
JavaScript
document.getElementById("insert-html").onclick = () =>
tryCatch(insertHTML);
```
- 5. Add the following function to the end of the file.

```
JavaScript
async function insertHTML() {
 await Word.run(async (context) => {
 // TODO1: Queue commands to insert a string of HTML.
 await context.sync();
 });
}
```
6. Within the insertHTML() function, replace TODO1 with the following code. Note:

- The first line adds a blank paragraph to the end of the document.
- The second line inserts a string of HTML at the end of the paragraph; specifically two paragraphs, one formatted with the Verdana font, the other with the default styling of the Word document. (As you saw in the insertImage method earlier, the context.document.body object also has the insert* methods.)


#### JavaScript

```
const blankParagraph =
context.document.body.paragraphs.getLast().insertParagraph("",
Word.InsertLocation.after);
blankParagraph.insertHtml('<p style="font-family: verdana;">Inserted
HTML.</p><p>Another paragraph</p>', Word.InsertLocation.end);
```
### **Insert a table**

- 1. Open the file **./src/taskpane/taskpane.html**.
- 2. Locate the <button> element for the insert-html button, and add the following markup after that line.

```
HTML
<button class="ms-Button" id="insert-table">Insert Table</button><br/>
<br/>
```
- 3. Open the file **./src/taskpane/taskpane.js**.
- 4. Within the Office.onReady function call, locate the line that assigns a click handler to the insert-html button, and add the following code after that line.

```
JavaScript
document.getElementById("insert-table").onclick = () =>
tryCatch(insertTable);
```
- 5. Add the following function to the end of the file.

```
JavaScript
async function insertTable() {
 await Word.run(async (context) => {
 // TODO1: Queue commands to get a reference to the paragraph
 // that will precede the table.
 // TODO2: Queue commands to create a table and populate it with
data.
 await context.sync();
 });
}
```


- 6. Within the insertTable() function, replace TODO1 with the following code. Note that this line uses the ParagraphCollection.getFirst method to get a reference to the first paragraph and then uses the Paragraph.getNext method to get a reference to the second paragraph.

```
JavaScript
const secondParagraph =
context.document.body.paragraphs.getFirst().getNext();
```
7. Within the insertTable() function, replace TODO2 with the following code. Note:

- The first two parameters of the insertTable method specify the number of rows and columns.
- The third parameter specifies where to insert the table, in this case after the paragraph.
- The fourth parameter is a two-dimensional array that sets the values of the table cells.
- The table will have plain default styling, but the insertTable method returns a Table object with many members, some of which are used to style the table.

```
JavaScript
const tableData = [
 ["Name", "ID", "Birth City"],
 ["Bob", "434", "Chicago"],
 ["Sue", "719", "Havana"],
 ];
secondParagraph.insertTable(3, 3, Word.InsertLocation.after,
tableData);
```
- 8. Save all your changes to the project.
### **Test the add-in**

- 1. If the local web server is already running and your add-in is already loaded in Word, proceed to step 2. Otherwise, start the local web server and sideload your add-in.


- To test your add-in in Word, run the following command in the root directory of your project. This starts the local web server (if it isn't already running) and opens Word with your add-in loaded.
command line npm start

- To test your add-in in Word on the web, run the following command in the root directory of your project. When you run this command, the local web server starts. Replace "{url}" with the URL of a Word document on your OneDrive or a SharePoint library to which you have permissions.
#### 7 **Note**

If you are developing on a Mac, enclose the {url} in single quotation marks. Do *not* do this on Windows.

command line

npm run start -- web --document {url}

The following are examples.

- npm run start -- web --document https://contoso.sharepoint.com/:t:/g/EZGxP7ksiE5DuxvY638G798BpuhwluxCM fF1WZQj3VYhYQ?e=F4QM1R
- npm run start -- web --document https://1drv.ms/x/s!jkcH7spkM4EGgcZUgqthk4IK3NOypVw?e=Z6G1qp
- npm run start -- web --document https://contoso-my.sharepointdf.com/:t:/p/user/EQda453DNTpFnl1bFPhOVR0BwlrzetbXvnaRYii2lDr_oQ? e=RSccmNP

If your add-in doesn't sideload in the document, manually sideload it by following the instructions in Manually sideload add-ins to Office on the web.

- 2. If the add-in task pane isn't already open in Word, go to the **Home** tab and choose the **Show Taskpane** button on the ribbon to open it.
- 3. In the task pane, choose the **Insert Paragraph** button at least three times to ensure that there are a few paragraphs in the document.


- 4. Choose the **Insert Image** button and note that an image is inserted at the end of the document.
- 5. Choose the **Insert HTML** button and note that two paragraphs are inserted at the end of the document, and that the first one has the Verdana font.
- 6. Choose the **Insert Table** button and note that a table is inserted after the second paragraph.

|                |     |                                                                                                                                                                                                                  | My Office Add-in<br>V V CICUITIL |
|----------------|-----|------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|----------------------------------|
|                |     |                                                                                                                                                                                                                  |                                  |
|                |     | Office has several versions, including Office 2016, Microsoft 365 subscription, and Office on the web.<br>Office has several versions, including Office 2016, Microsoft 365 subscription, and Office on the web. | Insert Paragraph                 |
| Name           | ID  | Birth City                                                                                                                                                                                                       |                                  |
| Bob            | 434 | Chicago                                                                                                                                                                                                          |                                  |
| Sue            | 719 | Havana                                                                                                                                                                                                           | Apply Style                      |
|                |     |                                                                                                                                                                                                                  | Change Font                      |
|                |     |                                                                                                                                                                                                                  |                                  |
|                |     |                                                                                                                                                                                                                  | Insert Abbreviation              |
|                |     |                                                                                                                                                                                                                  | Add Version Info                 |
|                |     |                                                                                                                                                                                                                  |                                  |
|                |     |                                                                                                                                                                                                                  | Change Quantity Term             |
|                |     |                                                                                                                                                                                                                  | Insert Image                     |
| Inserted HTML. |     |                                                                                                                                                                                                                  | Insert HTML                      |

### **Create and update content controls**

In this step of the tutorial, you'll learn how to create Rich Text content controls in the document, and then how to insert and replace content in the controls.

#### 7 **Note**

Before you start this step of the tutorial, we recommend that you create and manipulate Rich Text content controls through the Word UI, so you can be familiar with the controls and their properties. For details, see **[Create forms that users](https://support.microsoft.com/office/040c5cc1-e309-445b-94ac-542f732c8c8b) [complete or print in Word](https://support.microsoft.com/office/040c5cc1-e309-445b-94ac-542f732c8c8b)** .

### **Create a content control**

- 1. Open the file **./src/taskpane/taskpane.html**.


- 2. Locate the <button> element for the insert-table button, and add the following markup after that line.

```
HTML
<button class="ms-Button" id="create-content-control">Create Content
Control</button><br/><br/>
```
- 3. Open the file **./src/taskpane/taskpane.js**.
- 4. Within the Office.onReady function call, locate the line that assigns a click handler to the insert-table button, and add the following code after that line.

```
JavaScript
document.getElementById("create-content-control").onclick = () =>
tryCatch(createContentControl);
```
- 5. Add the following function to the end of the file.

```
JavaScript
async function createContentControl() {
 await Word.run(async (context) => {
 // TODO1: Queue commands to create a content control.
 await context.sync();
 });
}
```
- 6. Within the createContentControl() function, replace TODO1 with the following code. Note:
	- This code is intended to wrap the phrase "Microsoft 365" in a content control. It makes a simplifying assumption that the string is present and the user has selected it.
	- The ContentControl.title property specifies the visible title of the content control.
	- The ContentControl.tag property specifies an tag that can be used to get a reference to a content control using the ContentControlCollection.getByTag method, which you'll use in a later function.


- The ContentControl.appearance property specifies the visual look of the control. Using the value "Tags" means that the control will be wrapped in opening and closing tags, and the opening tag will have the content control's title. Other possible values are "BoundingBox" and "None".
- The ContentControl.color property specifies the color of the tags or the border of the bounding box.

```
const serviceNameRange = context.document.getSelection();
const serviceNameContentControl =
serviceNameRange.insertContentControl();
serviceNameContentControl.title = "Service Name";
serviceNameContentControl.tag = "serviceName";
serviceNameContentControl.appearance = "Tags";
serviceNameContentControl.color = "blue";
```
### **Replace the content of the content control**

- 1. Open the file **./src/taskpane/taskpane.html**.
JavaScript

- 2. Locate the <button> element for the create-content-control button, and add the following markup after that line.

```
HTML
<button class="ms-Button" id="replace-content-in-control">Rename
Service</button><br/><br/>
```
- 3. Open the file **./src/taskpane/taskpane.js**.
- 4. Within the Office.onReady function call, locate the line that assigns a click handler to the create-content-control button, and add the following code after that line.

```
JavaScript
document.getElementById("replace-content-in-control").onclick = () =>
tryCatch(replaceContentInControl);
```
- 5. Add the following function to the end of the file.
JavaScript


```
async function replaceContentInControl() {
 await Word.run(async (context) => {
 // TODO1: Queue commands to replace the text in the Service
Name
 // content control.
 await context.sync();
 });
}
```
- 6. Within the replaceContentInControl() function, replace TODO1 with the following code. Note:
	- The ContentControlCollection.getByTag method returns a ContentControlCollection of all content controls of the specified tag. We use getFirst to get a reference to the desired control.

#### JavaScript

```
const serviceNameContentControl =
context.document.contentControls.getByTag("serviceName").getFirst();
serviceNameContentControl.insertText("Fabrikam Online Productivity
Suite", Word.InsertLocation.replace);
```
- 7. Save all your changes to the project.
### **Test the add-in**

- 1. If the local web server is already running and your add-in is already loaded in Word, proceed to step 2. Otherwise, start the local web server and sideload your add-in.
	- To test your add-in in Word, run the following command in the root directory of your project. This starts the local web server (if it isn't already running) and opens Word with your add-in loaded.

```
command line
npm start
```
- To test your add-in in Word on the web, run the following command in the root directory of your project. When you run this command, the local web


server starts. Replace "{url}" with the URL of a Word document on your OneDrive or a SharePoint library to which you have permissions.

#### 7 **Note**

If you are developing on a Mac, enclose the {url} in single quotation marks. Do *not* do this on Windows.

command line

```
npm run start -- web --document {url}
```
The following are examples.

- npm run start -- web --document https://contoso.sharepoint.com/:t:/g/EZGxP7ksiE5DuxvY638G798BpuhwluxCM fF1WZQj3VYhYQ?e=F4QM1R
- npm run start -- web --document https://1drv.ms/x/s!jkcH7spkM4EGgcZUgqthk4IK3NOypVw?e=Z6G1qp
- npm run start -- web --document https://contoso-my.sharepointdf.com/:t:/p/user/EQda453DNTpFnl1bFPhOVR0BwlrzetbXvnaRYii2lDr_oQ? e=RSccmNP

If your add-in doesn't sideload in the document, manually sideload it by following the instructions in Manually sideload add-ins to Office on the web.

- 2. If the add-in task pane isn't already open in Word, go to the **Home** tab and choose the **Show Taskpane** button on the ribbon to open it.
- 3. In the task pane, choose the **Insert Paragraph** button to ensure that there's a paragraph with "Microsoft 365" at the top of the document.
- 4. In the document, select the text "Microsoft 365" and then choose the **Create Content Control** button. Note that the phrase is wrapped in tags labelled "Service Name".
- 5. Choose the **Rename Service** button and note that the text of the content control changes to "Fabrikam Online Productivity Suite".


### **Next steps**

In this tutorial, you've created a Word task pane add-in that inserts and replaces text, images, and other content in a Word document. To learn more about building Word add-ins, continue to the following article.

**Word add-ins overview**

### **Code samples**

- [Completed Word add-in tutorial](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/tutorials/word-tutorial) : The result of completing this tutorial.
# **See also**

- Office Add-ins platform overview
- Develop Office Add-ins


# **Sample: Import a Word document template with a Word add-in**

Article • 03/11/2024

Templates enable users to quickly create consistent documents for their organizations. Templates can include company information and other critical details that users need for compliance, legal, or other reasons.

This article features a sample add-in that imports a .docx file to use as a template in a Word document. The add-in replaces the current document's content with the content from the template.

# **Prerequisites**

- Office connected to a Microsoft 365 subscription (including Office on the web).


# **Run the sample code**

The sample code for this article is named [Import templates in a Word document](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/word-import-template) . To run the sample, follow the instructions in the [readme](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/word-import-template) .

|                                                                                                                                                                                                                                             |  | Import template - sample<br>V X                                                      |
|---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|--|--------------------------------------------------------------------------------------|
|                                                                                                                                                                                                                                             |  | Import template                                                                      |
| LOGO                                                                                                                                                                                                                                        |  | Template<br>Select a Word document to import.<br>Choose File   template example.docx |
| Memo                                                                                                                                                                                                                                        |  | Update                                                                               |
| To: Recipient Name                                                                                                                                                                                                                          |  | The template has been applied to your<br>document. Go ahead and update the content.  |
| From: Your Name                                                                                                                                                                                                                             |  |                                                                                      |
| CC: Other recipients                                                                                                                                                                                                                        |  | More code samples?                                                                   |
| ◄ Some of the sample text in this document indicates the name of the style<br>applied, so that you can easily apply the same formatting again. To get started<br>right away, just tap any placeholder text (such as this) and start typing. |  | For additional code samples, see Office Add-<br>in code samples.                     |
| View and edit this document in Word on your computer, tablet, or phone. You can edit text; easily insert                                                                                                                                    |  |                                                                                      |

# **Key steps in the sample**

- 1. The user chooses a .docx file they'd like to use as a template.
- 2. The add-in reads the template .docx file then uses [Document.insertFileFromBase64](https://learn.microsoft.com/en-us/javascript/api/word/word.document#word-word-document-insertfilefrombase64-member(1)) to replace the current document's content with the content from the template file.
- 3. The user can make updates to the content of the current document.

# **Make it yours**

The following are a few suggestions for how you could tailor this sample to your scenario.

### **Manage user settings**

Enable single sign-on (SSO) in an Office Add-in to support persisting user data and settings across multiple documents. If your service provides or hosts a document template library, an authorized user can access and apply a template in their document.

You can also persist add-in state and settings in the user's current document.

U **Caution**


Don't store sensitive information such as authentication tokens or connection strings. Properties in the document aren't encrypted or protected.

### **Provide templates**

Provide personalized or company-approved templates for users. These templates can be made accessible from a shared location as part of an authenticated experience.

You can use [content controls](https://learn.microsoft.com/en-us/javascript/api/word/word.contentcontrol), [fields](https://learn.microsoft.com/en-us/javascript/api/word/word.field), and other components as building blocks in your templates.

### **Personalize templates**

Allow users to personalize or refine templates. For templates that may be useful to others (on their team, in their company, etc.), users can upload to a shared location.

### **See also**

- Office Add-in code samples
#### 6 **Collaborate with us on GitHub**

The source for this content can be found on GitHub, where you can also create and review issues and pull requests. For [more information, see our](https://github.com/OfficeDev/office-js-docs-pr/blob/main/Contributing.md) contributor guide.

#### **Office Add-ins feedback**

Office Add-ins is an open source project. Select a link to provide feedback:

[Open a documentation issue](https://github.com/OfficeDev/office-js-docs-pr/issues/new?template=3-customer-feedback.yml&pageUrl=https%3A%2F%2Flearn.microsoft.com%2Fen-us%2Foffice%2Fdev%2Fadd-ins%2Fword%2Fimport-template&pageQueryParams=&contentSourceUrl=https%3A%2F%2Fgithub.com%2FOfficeDev%2Foffice-js-docs-pr%2Fblob%2Fmain%2Fdocs%2Fword%2Fimport-template.md&documentVersionIndependentId=643cd19d-3a79-00e7-b5c8-9aa0a381e061&feedback=%0A%0A%5BEnter+feedback+here%5D%0A&author=%40o365devx&metadata=*+ID%3A+643cd19d-3a79-00e7-b5c8-9aa0a381e061+%0A*+Service%3A+**word**%0A*+Sub-service%3A+**add-ins**)

- [Provide product feedback](https://aka.ms/office-addins-dev-questions)


# **Sample: Manage citations in a Word document using your Word add-in**

Article • 11/29/2023

Citation management is an important aspect of documents, particularly in academia and education. Each citation style has its own guidelines for how citations should be marked in a document as well as where and how the sources should be noted. Such styles include [APA](http://apastyle.apa.org/) and [MLA](https://www.mla.org/MLA-Style) .

This article features a sample add-in that manages citations in a Word document. The add-in displays the references loaded from a .bib file that the user selects to cite in their document.

| Citation manager - sample                                                              |
|----------------------------------------------------------------------------------------|
| Citation management                                                                    |
| Bibliography sources<br>Select a bibliography file to use.<br>Choose File   sample.bib |
| References                                                                             |
| review<br>@ A Review of Pair-wise Testing<br>Jimi Sanchez                              |
| Current selection: Sanchez2016                                                         |
| Insert citation<br>Clear selection                                                     |
| More code samples?<br>For additional code samples, see Office<br>Add-in code samples.  |


# **Prerequisites**

- [Visual Studio Code](https://code.visualstudio.com/Download) .
- Office connected to a Microsoft 365 subscription (including Office on the web).
- [Node.js](https://nodejs.org/) version 16 or greater.
- npm version 8 or greater.

# **Run the sample code**

The sample code for this article is named [Manage citations in a Word document](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/word-citation-management) . To run the sample, follow the instructions in the [readme](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/word-citation-management) .

# **Key steps in the sample**

- 1. The user chooses a local .bib file that contains the references they'd like to cite.
- 2. The add-in reads the .bib file then displays the bibliography references in the task pane. The sample uses [@orcid/bibtexParseJs](https://github.com/ORCID/bibtexParseJs#readme) to parse the .bib file.
- 3. The user chooses the appropriate reference then inserts it at the cursor's location (or at the end of selected text) in the document.
- 4. The add-in adds a reference mark at that location in the document and adds the reference to an endnote. All endnotes are automatically listed at the end of the document.

# **Make it yours**

The following are a few suggestions for how you could tailor this sample to your scenario.

### **Manage user settings**

Enable single sign-on (SSO) in an Office Add-in to support persisting user data and settings across multiple documents. If your service provides or hosts the bibliography library, an authorized user can access and select from that bibliography in their document.

You can also persist add-in state and settings in the user's current document.


Don't store sensitive information such as authentication tokens or connection strings. Properties in the document aren't encrypted or protected.

### **Use footnotes**

List the references in [footnotes](https://learn.microsoft.com/en-us/javascript/api/word/word.range#word-word-range-insertfootnote-member(1)) at the end of the page instead of endnotes, according to the citation style.

Alternatively, allow the user to choose where they'd like the references to be displayed. If so, you can update the add-in to persist the user's preference using a document property or as part of their authenticated experience.

### **Update citation style**

Update the citation style used to display the references in the endnotes (or footnotes).

Alternatively, provide various style options then allow the user to choose. If so, you can update the add-in to persist the user's preference using a document property or as part of their authenticated experience.

### **Replace bibtexParseJs**

Replace the .bib file parser [@orcid/bibtexParseJs](https://github.com/ORCID/bibtexParseJs#readme) with your own or another available parser, especially if this option doesn't provide the functionality you need for your solution.

# **See also**

- Office Add-in code samples
- [npm](https://docs.npmjs.com/downloading-and-installing-node-js-and-npm)
- [@orcid/bibtexParseJs](https://github.com/ORCID/bibtexParseJs#readme)


# **Add headers when a document opens**

07/14/2025

The following sections walk you through how to develop a Word add-in that automatically changes the document header when a new or existing document opens. While this specific sample is for Word, the manifest configuration is the same for Excel and PowerPoint. For an overview of this style of event-based activation pattern, see Activate add-ins with events.

#### ) **Important**

This sample requires you to have a Microsoft 365 subscription with the supported version of Word.

# **Create a new add-in**

Create a new add-in by following the [Word add-in quick start](https://learn.microsoft.com/en-us/office/dev/add-ins/quickstarts/word-quickstart-yo?tabs=yeoman). This will give you a working Office Add-in to which you can add the event-based activation code.

#### 7 **Note**

For a completed version of the sample described in this walkthrough, see the **[Automatically add labels with an add-in when a Word document opens sample in our](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/word-add-label-on-open) [samples GitHub repo](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/word-add-label-on-open)** .

# **Configure the manifest**

To enable an event-based add-in, you must configure the following elements in the VersionOverridesV1_0 node of the manifest.

- In the [Runtimes](https://learn.microsoft.com/en-us/javascript/api/manifest/runtimes) element, make a new [Override element for Runtime](https://learn.microsoft.com/en-us/javascript/api/manifest/override#override-element-for-runtime). Override the "javascript" type and reference the JavaScript file containing the function you want to trigger with the event.
- In the [DesktopFormFactor](https://learn.microsoft.com/en-us/javascript/api/manifest/desktopformfactor) element, add a [FunctionFile](https://learn.microsoft.com/en-us/javascript/api/manifest/functionfile) element for the JavaScript file with the event handler.
- In the [ExtensionPoint](https://learn.microsoft.com/en-us/javascript/api/manifest/extensionpoint) element, set the xsi:type to LaunchEvent . This enables the eventbased activation feature in your add-in.
- In the [LaunchEvent](https://learn.microsoft.com/en-us/javascript/api/manifest/launchevent) element, set the Type to OnDocumentOpened and specify the JavaScript function name of the event handler in the FunctionName attribute.


Use the following sample manifest code to update your project.

- 1. In your code editor, open the quick start project you created.
- 2. Open the **manifest.xml** file located at the root of your project.
- 3. Select the entire **<VersionOverrides>** node (including the open and close tags) and replace it with the following XML.

```
XML
 <VersionOverrides
xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides"
xsi:type="VersionOverridesV1_0">
 <Hosts>
 <Host xsi:type="Document">
 <Runtimes>
 <Runtime resid="Taskpane.Url" lifetime="long" />
 <Runtime resid="WebViewRuntime.Url">
 <Override type="javascript" resid="JsRuntimeWord.Url"/>
 </Runtime>
 </Runtimes>
 <DesktopFormFactor>
 <GetStarted>
 <Title resid="GetStarted.Title"/>
 <Description resid="GetStarted.Description"/>
 <LearnMoreUrl resid="GetStarted.LearnMoreUrl"/>
 </GetStarted>
 <FunctionFile resid="Commands.Url"/>
 <ExtensionPoint xsi:type="LaunchEvent">
 <LaunchEvents>
 <LaunchEvent Type="OnDocumentOpened"
FunctionName="changeHeader"></LaunchEvent>
 </LaunchEvents>
 <SourceLocation resid="WebViewRuntime.Url"/>
 </ExtensionPoint>
 <ExtensionPoint xsi:type="PrimaryCommandSurface">
 <OfficeTab id="TabHome">
 <Group id="CommandsGroup">
 <Label resid="CommandsGroup.Label"/>
 <Icon>
 <bt:Image size="16" resid="Icon.16x16"/>
 <bt:Image size="32" resid="Icon.32x32"/>
 <bt:Image size="80" resid="Icon.80x80"/>
 </Icon>
 <Control xsi:type="Button" id="TaskpaneButton">
 <Label resid="TaskpaneButton.Label"/>
 <Supertip>
 <Title resid="TaskpaneButton.Label"/>
 <Description resid="TaskpaneButton.Tooltip"/>
 </Supertip>
 <Icon>
 <bt:Image size="16" resid="Icon.16x16"/>
```


```
 <bt:Image size="32" resid="Icon.32x32"/>
 <bt:Image size="80" resid="Icon.80x80"/>
 </Icon>
 <Action xsi:type="ShowTaskpane">
 <TaskpaneId>ButtonId1</TaskpaneId>
 <SourceLocation resid="Taskpane.Url"/>
 </Action>
 </Control>
 </Group>
 </OfficeTab>
 </ExtensionPoint>
 </DesktopFormFactor>
 </Host>
 </Hosts>
 <Resources>
 <bt:Images>
 <bt:Image id="Icon.16x16"
DefaultValue="https://localhost:3000/assets/icon-16.png"/>
 <bt:Image id="Icon.32x32"
DefaultValue="https://localhost:3000/assets/icon-32.png"/>
 <bt:Image id="Icon.80x80"
DefaultValue="https://localhost:3000/assets/icon-80.png"/>
 </bt:Images>
 <bt:Urls>
 <bt:Url id="GetStarted.LearnMoreUrl"
DefaultValue="https://go.microsoft.com/fwlink/?LinkId=276812"/>
 <bt:Url id="Commands.Url"
DefaultValue="https://localhost:3000/commands.html"/>
 <bt:Url id="Taskpane.Url"
DefaultValue="https://localhost:3000/taskpane.html"/>
 <bt:Url id="WebViewRuntime.Url"
DefaultValue="https://localhost:3000/commands.html"/>
 <bt:Url id="JsRuntimeWord.Url"
DefaultValue="https://localhost:3000/commands.js"/>
 </bt:Urls>
 <bt:ShortStrings>
 <bt:String id="GetStarted.Title" DefaultValue="Get started with your
sample add-in!"/>
 <bt:String id="CommandsGroup.Label" DefaultValue="Event-based add-in
activation"/>
 <bt:String id="TaskpaneButton.Label" DefaultValue="My add-in"/>
 </bt:ShortStrings>
 <bt:LongStrings>
 <bt:String id="GetStarted.Description" DefaultValue="Your sample add-
in loaded successfully. Go to the HOME tab and click the 'Show Task Pane'
button to get started."/>
 <bt:String id="TaskpaneButton.Tooltip" DefaultValue="Click to show
the task pane"/>
 </bt:LongStrings>
 </Resources>
 </VersionOverrides>
```
- 4. Save your changes.


# **Implement the event handler**

To enable your add-in to act when the OnDocumentOpened event occurs, you must implement a JavaScript event handler. In this section, you'll create the changeHeader function, which adds a "Public" header to new documents or a "Highly Confidential" header to existing documents that already have content.

- 1. In the **./src/commands** folder, open the file named **commands.js**.
- 2. Replace the entire contents of **commands.js** with the following JavaScript code.

```
JavaScript
 /*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under
the MIT license.
 * See LICENSE in the project root for license information.
 */
 /* global global, Office, self, window */
 Office.onReady(() => {
 // If needed, Office.js is ready to be called.
 });
 async function changeHeader(event) {
 Word.run(async (context) => {
 const body = context.document.body;
 body.load("text");
 await context.sync();
 if (body.text.length === 0) {
 // For new or empty documents, make a "Public" header. 
 const header =
context.document.sections.getFirst().getHeader(Word.HeaderFooterType.primary)
;
 const firstPageHeader =
context.document.sections.getFirst().getHeader(Word.HeaderFooterType.firstPag
e);
 header.clear();
 firstPageHeader.clear();
 header.insertParagraph("Public - The data is for the public and
shareable externally", "Start");
 firstPageHeader.insertParagraph("Public - The data is for the public
and shareable externally", "Start");
 header.font.color = "#07641d";
 firstPageHeader.font.color = "#07641d";
 await context.sync();
 } else {
 // For existing documents, make a "Highly Confidential" header.
 const header =
context.document.sections.getFirst().getHeader(Word.HeaderFooterType.primary)
```


```
;
 const firstPageHeader =
context.document.sections.getFirst().getHeader(Word.HeaderFooterType.firstPag
e);
 header.clear();
 firstPageHeader.clear();
 header.insertParagraph("Highly Confidential - The data must be secret
or in some way highly critical", "Start");
 firstPageHeader.insertParagraph("Highly Confidential - The data must
be secret or in some way highly critical", "Start");
 header.font.color = "#f8334d";
 firstPageHeader.font.color = "#f8334d";
 await context.sync();
 }
 });
 // Calling event.completed is required. event.completed lets the platform
know that processing has completed.
 event.completed();
 }
 async function paragraphChanged() {
 await Word.run(async (context) => {
 const results = context.document.body.search("110");
 results.load("length");
 await context.sync();
 if (results.items.length === 0) {
 const header =
context.document.sections.getFirst().getHeader(Word.HeaderFooterType.primary)
;
 header.clear();
 header.insertParagraph("Public - The data is for the public and
shareable externally", "Start");
 const font = header.font;
 font.color = "#07641d";
 await context.sync();
 } else {
 const header =
context.document.sections.getFirst().getHeader(Word.HeaderFooterType.primary)
;
 header.clear();
 header.insertParagraph("Highly Confidential - The data must be secret
or in some way highly critical", "Start");
 const font = header.font;
 font.color = "#f8334d";
 await context.sync();
 }
 });
 }
 async function registerOnParagraphChanged(event) {
 await Word.run(async (context) => {
 let eventContext =
```


```
context.document.onParagraphChanged.add(paragraphChanged);
 await context.sync();
 });
 // Calling event.completed is required. event.completed lets the platform
know that processing has completed.
 event.completed();
 }
 Office.actions.associate("changeHeader", changeHeader);
 Office.actions.associate("registerOnParagraphChanged",
registerOnParagraphChanged);
```
- 3. Save your changes.
# **Test and validate your add-in**

- 1. Run npm start to build your project and launch the web server. **Ignore the Word document that is opened**.
- 2. Manually sideload your add-in in Word on the web by following the guidance at Sideload Office Add-ins to Office on the web. Use the **manifest.xml** in the root of the project.
- 3. Try opening both new and existing Word documents in Word on the web. Headers should automatically be added when they open.

# **See also**

- Activate add-ins with events
- Debug event-based or spam-reporting add-ins
- Troubleshoot event-based and spam-reporting add-ins


# **Word JavaScript API overview**

Article • 05/29/2025

A Word add-in interacts with objects in Word by using the Office JavaScript API, which includes two JavaScript object models:

- **Word JavaScript API**: These are the application-specific APIs for Word. Introduced with Office 2016, the [Word JavaScript API](https://learn.microsoft.com/en-us/javascript/api/word) provides strongly-typed objects that you can use to access objects and metadata in a Word document.
- **Common APIs**: The [Common API](https://learn.microsoft.com/en-us/javascript/api/office), introduced with Office 2013, can be used to access features such as UI, dialogs, and client settings that are common across multiple Office applications.

This section of the documentation focuses on the Word JavaScript API, which you'll use to develop the majority of functionality in add-ins that target Word on the web, or Word 2016 and later. For information about the Common API, see Common JavaScript API object model.

# **Learn programming concepts**

See Word JavaScript object model in Office Add-ins for information about important programming concepts.

# **Learn about API capabilities**

Use other articles in this section of the documentation to learn how to get the whole document from an add-in, use search options in your Word add-in to find text, and more. See the table of contents for the complete list of available articles.

For hands-on experience using the Word JavaScript API to access objects in Word, complete the Word add-in tutorial.

For detailed information about the Word JavaScript API object model, see the [Word JavaScript](https://learn.microsoft.com/en-us/javascript/api/word) [API reference documentation](https://learn.microsoft.com/en-us/javascript/api/word).

# **Try out code samples in Script Lab**

Use Script Lab to get started quickly with a collection of built-in samples that show how to complete tasks with the API. You can run the samples in Script Lab to instantly see the result in the task pane or document, examine the samples to learn how the API works, and even use samples to prototype your own add-in.


# **See also**

- Word add-ins documentation
- Word add-ins overview
- [Word JavaScript API reference](https://learn.microsoft.com/en-us/javascript/api/word)
- [Office client application and platform availability for Office Add-ins](https://learn.microsoft.com/en-us/javascript/api/requirement-sets)


# **word package**

# **Classes**

ノ **Expand table**

| Word.<br>Annotation                    | Represents an annotation attached to a paragraph.                                                                                                                                                                                          |
|----------------------------------------|--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| Word.<br>Annotation<br>Collection      | Contains a collection of Word.Annotation objects.                                                                                                                                                                                          |
| Word.<br>Application                   | Represents the application object.                                                                                                                                                                                                         |
| Word.                                  | Represents the list of available sources attached to the document (in the current list) or                                                                                                                                                 |
| Bibliography                           | the list of sources available in the application (in the master list).                                                                                                                                                                     |
| Word.Body                              | Represents the body of a document or a section.                                                                                                                                                                                            |
| Word.Bookmark                          | Represents a single bookmark in a document, selection, or range. The Bookmark object<br>is a member of the Bookmark collection. The Word.BookmarkCollection includes all the<br>bookmarks listed in the Bookmark dialog box (Insert menu). |
| Word.Bookmark                          | A collection of Word.Bookmark objects that represent the bookmarks in the specified                                                                                                                                                        |
| Collection                             | selection, range, or document.                                                                                                                                                                                                             |
| Word.Border                            | Represents the Border object for text, a paragraph, or a table.                                                                                                                                                                            |
| Word.Border<br>Collection              | Represents the collection of border styles.                                                                                                                                                                                                |
| Word.Border                            | Represents the BorderUniversal object, which manages borders for a range,                                                                                                                                                                  |
| Universal                              | paragraph, table, or frame.                                                                                                                                                                                                                |
| Word.Border<br>Universal<br>Collection | Represents the collection of Word.BorderUniversal objects.                                                                                                                                                                                 |
| Word.Break                             | Represents a break in a Word document. This could be a page, column, or section<br>break.                                                                                                                                                  |
| Word.Break<br>Collection               | Contains a collection of Word.Break objects.                                                                                                                                                                                               |
| Word.Building                          | Represents a building block in a template. A building block is pre-built content, similar                                                                                                                                                  |
| Block                                  | to autotext, that may contain text, images, and formatting.                                                                                                                                                                                |


| Word.Building<br>BlockCategory                  | Represents a category of building blocks in a Word document.                                     |  |
|-------------------------------------------------|--------------------------------------------------------------------------------------------------|--|
| Word.Building<br>BlockCategory<br>Collection    | Represents a collection of Word.BuildingBlockCategory objects in a Word document.                |  |
| Word.Building                                   | Represents a collection of Word.BuildingBlock objects for a specific building block type         |  |
| BlockCollection                                 | and category in a template.                                                                      |  |
| Word.Building<br>BlockEntry<br>Collection       | Represents a collection of building block entries in a Word template.                            |  |
| Word.Building<br>BlockGallery<br>ContentControl | Represents the BuildingBlockGalleryContentControl object.                                        |  |
| Word.Building<br>BlockTypeItem                  | Represents a type of building block in a Word document.                                          |  |
| Word.Building<br>BlockTypeItem<br>Collection    | Represents a collection of building block types in a Word document.                              |  |
| Word.Canvas                                     | Represents a canvas in the document. To get the corresponding Shape object, use<br>Canvas.shape. |  |
| Word.Checkbox<br>ContentControl                 | The data specific to content controls of type CheckBox.                                          |  |
| Word.Color<br>Format                            | Represents the color formatting of a shape or text in Word.                                      |  |
| Word.Combo<br>BoxContent<br>Control             | The data specific to content controls of type 'ComboBox'.                                        |  |
| Word.Comment                                    | Represents a comment in the document.                                                            |  |
| Word.Comment<br>Collection                      | Contains a collection of Word.Comment objects.                                                   |  |
| Word.CommentContentRange                        |                                                                                                  |  |
| Word.Comment<br>Reply                           | Represents a comment reply in the document.                                                      |  |
| Word.Comment                                    | Contains a collection of Word.CommentReply objects. Represents all comment replies               |  |
| ReplyCollection                                 | in one comment thread.                                                                           |  |


| Word.Content<br>Control                       | Represents a content control. Content controls are bounded and potentially labeled<br>regions in a document that serve as containers for specific types of content. Individual<br>content controls may contain contents such as images, tables, or paragraphs of<br>formatted text. Currently, only rich text, plain text, checkbox, dropdown list, and combo<br>box content controls are supported.                         |
|-----------------------------------------------|------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| Word.Content<br>Control<br>Collection         | Contains a collection of Word.ContentControl objects. Content controls are bounded<br>and potentially labeled regions in a document that serve as containers for specific<br>types of content. Individual content controls may contain contents such as images,<br>tables, or paragraphs of formatted text. Currently, only rich text, plain text, checkbox,<br>dropdown list, and combo box content controls are supported. |
| Word.Content<br>ControlListItem               | Represents a list item in a dropdown list or combo box content control.                                                                                                                                                                                                                                                                                                                                                      |
| Word.Content<br>ControlListItem<br>Collection | Contains a collection of Word.ContentControlListItem objects that represent the items<br>in a dropdown list or combo box content control.                                                                                                                                                                                                                                                                                    |
| Word.Critique<br>Annotation                   | Represents an annotation wrapper around critique displayed in the document.                                                                                                                                                                                                                                                                                                                                                  |
| Word.Custom<br>Property                       | Represents a custom property.                                                                                                                                                                                                                                                                                                                                                                                                |
| Word.Custom<br>Property<br>Collection         | Contains the collection of Word.CustomProperty objects.                                                                                                                                                                                                                                                                                                                                                                      |
| Word.Custom                                   | Represents an XML node in a tree in the document. The CustomXmlNode object is a                                                                                                                                                                                                                                                                                                                                              |
| XmlNode                                       | member of the Word.CustomXmlNodeCollection object.                                                                                                                                                                                                                                                                                                                                                                           |
| Word.Custom<br>XmlNode<br>Collection          | Contains a collection of Word.CustomXmlNode objects representing the XML nodes in<br>a document.                                                                                                                                                                                                                                                                                                                             |
| Word.Custom<br>XmlPart                        | Represents a custom XML part.                                                                                                                                                                                                                                                                                                                                                                                                |
| Word.Custom<br>XmlPart<br>Collection          | Contains the collection of Word.CustomXmlPart objects.                                                                                                                                                                                                                                                                                                                                                                       |
| Word.Custom<br>XmlPartScoped<br>Collection    | Contains the collection of Word.CustomXmlPart objects with a specific namespace.                                                                                                                                                                                                                                                                                                                                             |
| Word.Custom<br>XmlPrefix<br>Mapping           | Represents a CustomXmlPrefixMapping object.                                                                                                                                                                                                                                                                                                                                                                                  |


| Word.Custom<br>XmlPrefix<br>Mapping<br>Collection | Represents a collection of Word.CustomXmlPrefixMapping objects.                                                                                     |
|---------------------------------------------------|-----------------------------------------------------------------------------------------------------------------------------------------------------|
| Word.Custom<br>XmlSchema                          | Represents a schema in a Word.CustomXmlSchemaCollection object.                                                                                     |
| Word.Custom<br>XmlSchema<br>Collection            | Represents a collection of Word.CustomXmlSchema objects attached to a data stream.                                                                  |
| Word.Custom<br>XmlValidation<br>Error             | Represents a single validation error in a Word.CustomXmlValidationErrorCollection<br>object.                                                        |
| Word.Custom<br>XmlValidation<br>ErrorCollection   | Represents a collection of Word.CustomXmlValidationError objects.                                                                                   |
| Word.Date<br>PickerContent<br>Control             | Represents the DatePickerContentControl object.                                                                                                     |
| Word.                                             | The Document object is the top level object. A Document object contains one or more                                                                 |
| Document                                          | sections, content controls, and the body that contains the contents of the document.                                                                |
| Word.<br>Document<br>Created                      | The DocumentCreated object is the top level object created by<br>Application.CreateDocument. A DocumentCreated object is a special Document object. |
| Word.<br>Document<br>LibraryVersion               | Represents a document library version.                                                                                                              |
| Word.<br>Document<br>LibraryVersion<br>Collection | Represents the collection of Word.DocumentLibraryVersion objects.                                                                                   |
| Word.<br>Document<br>Properties                   | Represents document properties.                                                                                                                     |
| Word.DropCap                                      | Represents a dropped capital letter in a Word document.                                                                                             |
| Word.Drop<br>DownList<br>ContentControl           | The data specific to content controls of type DropDownList.                                                                                         |
| Word.Field                                        | Represents a field.                                                                                                                                 |


| Word.Field<br>Collection             | Contains a collection of Word.Field objects.                                                                                                                                                                                          |
|--------------------------------------|---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| Word.FillFormat                      | Represents the fill formatting for a shape or text.                                                                                                                                                                                   |
| Word.Font                            | Represents a font.                                                                                                                                                                                                                    |
| Word.Frame                           | Represents a frame. The Frame object is a member of the Word.FrameCollection object.                                                                                                                                                  |
| Word.Frame<br>Collection             | Represents the collection of Word.Frame objects.                                                                                                                                                                                      |
| Word.Glow<br>Format                  | Represents the glow formatting for the font used by the range of text.                                                                                                                                                                |
| Word.Group<br>ContentControl         | Represents the GroupContentControl object.                                                                                                                                                                                            |
| Word.Hyperlink                       | Represents a hyperlink in a Word document.                                                                                                                                                                                            |
| Word.Hyperlink<br>Collection         | Contains a collection of Word.Hyperlink objects.                                                                                                                                                                                      |
| Word.Index                           | Represents a single index. The Index object is a member of the Word.IndexCollection.<br>The IndexCollection includes all the indexes in the document.                                                                                 |
| Word.Index<br>Collection             | A collection of Word.Index objects that represents all the indexes in the document.                                                                                                                                                   |
| Word.Inline<br>Picture               | Represents an inline picture.                                                                                                                                                                                                         |
| Word.Inline<br>Picture<br>Collection | Contains a collection of Word.InlinePicture objects.                                                                                                                                                                                  |
| Word.Line<br>Format                  | Represents line and arrowhead formatting. For a line, the LineFormat object contains<br>formatting information for the line itself; for a shape with a border, this object contains<br>formatting information for the shape's border. |
| Word.Line                            | Represents line numbers in the left margin or to the left of each newspaper-style                                                                                                                                                     |
| Numbering                            | column.                                                                                                                                                                                                                               |
| Word.Link<br>Format                  | Represents the linking characteristics for an OLE object or picture.                                                                                                                                                                  |
| Word.List                            | Contains a collection of Word.Paragraph objects.                                                                                                                                                                                      |
| Word.List<br>Collection              | Contains a collection of Word.List objects.                                                                                                                                                                                           |


| Word.List<br>Format            | Represents the list formatting characteristics of a range.                                                                                              |
|--------------------------------|---------------------------------------------------------------------------------------------------------------------------------------------------------|
| Word.ListItem                  | Represents the paragraph list item format.                                                                                                              |
| Word.ListLevel                 | Represents a list level.                                                                                                                                |
| Word.ListLevel<br>Collection   | Contains a collection of Word.ListLevel objects.                                                                                                        |
| Word.List<br>Template          | Represents a list template.                                                                                                                             |
| Word.NoteItem                  | Represents a footnote or endnote.                                                                                                                       |
| Word.NoteItem<br>Collection    | Contains a collection of Word.NoteItem objects.                                                                                                         |
| Word.Ole                       | Represents the OLE characteristics (other than linking) for an OLE object, ActiveX                                                                      |
| Format                         | control, or field.                                                                                                                                      |
| Word.Page                      | Represents a page in the document. Page objects manage the page layout and<br>content.                                                                  |
| Word.Page<br>Collection        | Represents the collection of page.                                                                                                                      |
| Word.Page<br>Setup             | Represents the page setup settings for a Word document or section.                                                                                      |
| Word.Pane                      | Represents a window pane. The Pane object is a member of the pane collection. The<br>pane collection includes all the window panes for a single window. |
| Word.Pane<br>Collection        | Represents the collection of pane.                                                                                                                      |
| Word.Paragraph                 | Represents a single paragraph in a selection, range, content control, or document<br>body.                                                              |
| Word.Paragraph<br>Collection   | Contains a collection of Word.Paragraph objects.                                                                                                        |
| Word.Paragraph<br>Format       | Represents a style of paragraph in a document.                                                                                                          |
| Word.Picture<br>ContentControl | Represents the PictureContentControl object.                                                                                                            |
| Word.Range                     | Represents a contiguous area in a document.                                                                                                             |
| Word.Range<br>Collection       | Contains a collection of Word.Range objects.                                                                                                            |


| Word.Reflection<br>Format                   | Represents the reflection formatting for a shape in Word.                                                                                                                                                                                              |
|---------------------------------------------|--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| Word.Repeating<br>SectionContent<br>Control | Represents the RepeatingSectionContentControl object.                                                                                                                                                                                                  |
| Word.Repeating<br>SectionItem               | Represents a single item in a Word.RepeatingSectionContentControl.                                                                                                                                                                                     |
| Word.Repeating<br>SectionItem<br>Collection | Represents a collection of Word.RepeatingSectionItem objects in a Word document.                                                                                                                                                                       |
| Word.Request<br>Context                     | The RequestContext object facilitates requests to the Word application. Since the<br>Office add-in and the Word application run in two different processes, the request<br>context is required to get access to the Word object model from the add-in. |
| Word.Reviewer                               | Represents a single reviewer of a document in which changes have been tracked. The<br>Reviewer object is a member of the Word.ReviewerCollection object.                                                                                               |
| Word.Reviewer<br>Collection                 | A collection of Word.Reviewer objects that represents the reviewers of one or more<br>documents. The ReviewerCollection object contains the names of all reviewers who<br>have reviewed documents opened or edited on a computer.                      |
| Word.Revisions                              | Represents the current settings related to the display of reviewers' comments and                                                                                                                                                                      |
| Filter                                      | revision marks in the document.                                                                                                                                                                                                                        |
| Word.Search<br>Options                      | Specifies the options to be included in a search operation. To learn more about how to<br>use search options in the Word JavaScript APIs, read Use search options to find text in<br>your Word add-in.                                                 |
| Word.Section                                | Represents a section in a Word document.                                                                                                                                                                                                               |
| Word.Section<br>Collection                  | Contains the collection of the document's Word.Section objects.                                                                                                                                                                                        |
| Word.Setting                                | Represents a setting of the add-in.                                                                                                                                                                                                                    |
| Word.Setting<br>Collection                  | Contains the collection of Word.Setting objects.                                                                                                                                                                                                       |
| Word.Shading                                | Represents the shading object.                                                                                                                                                                                                                         |
| Word.Shading                                | Represents the ShadingUniversal object, which manages shading for a range,                                                                                                                                                                             |
| Universal                                   | paragraph, frame, or table.                                                                                                                                                                                                                            |
| Word.Shadow<br>Format                       | Represents the shadow formatting for a shape or text in Word.                                                                                                                                                                                          |
| Word.Shape                                  | Represents a shape in the header, footer, or document body. Currently, only the<br>following shapes are supported: text boxes, geometric shapes, groups, pictures, and                                                                                 |


|                                    | canvases.                                                                             |
|------------------------------------|---------------------------------------------------------------------------------------|
| Word.Shape                         | Contains a collection of Word.Shape objects. Currently, only the following shapes are |
| Collection                         | supported: text boxes, geometric shapes, groups, pictures, and canvases.              |
| Word.ShapeFill                     | Represents the fill formatting of a shape object.                                     |
| Word.Shape                         | Represents a shape group in the document. To get the corresponding Shape object,      |
| Group                              | use ShapeGroup.shape.                                                                 |
| Word.ShapeText<br>Wrap             | Represents all the properties for wrapping text around a shape.                       |
| Word.Source                        | Represents an individual source, such as a book, journal article, or interview.       |
| Word.Source<br>Collection          | Represents a collection of Word.Source objects.                                       |
| Word.Style                         | Represents a style in a Word document.                                                |
| Word.Style<br>Collection           | Contains a collection of Word.Style objects.                                          |
| Word.Table                         | Represents a table in a Word document.                                                |
| Word.Table<br>Border               | Specifies the border style.                                                           |
| Word.TableCell                     | Represents a table cell in a Word document.                                           |
| Word.TableCell<br>Collection       | Contains the collection of the document's TableCell objects.                          |
| Word.Table<br>Collection           | Contains the collection of the document's Table objects.                              |
| Word.Table<br>Column               | Represents a table column in a Word document.                                         |
| Word.Table<br>Column<br>Collection | Represents a collection of Word.TableColumn objects in a Word document.               |
| Word.TableRow                      | Represents a row in a Word document.                                                  |
| Word.TableRow<br>Collection        | Contains the collection of the document's TableRow objects.                           |
| Word.TableStyle                    | Represents the TableStyle object.                                                     |
| Word.TabStop                       | Represents a tab stop in a Word document.                                             |


| Word.TabStop<br>Collection           | Represents a collection of tab stops in a Word document.                                                                                                                                                                                                                                                                                                                         |
|--------------------------------------|----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| Word.Template                        | Represents a document template.                                                                                                                                                                                                                                                                                                                                                  |
| Word.Template<br>Collection          | Contains a collection of Word.Template objects that represent all the templates that<br>are currently available. This collection includes open templates, templates attached to<br>open documents, and global templates loaded in the Templates and Add-ins dialog<br>box. To learn how to access this dialog in the Word UI, see Load or unload a template<br>or add-in program |
| Word.Text<br>Column                  | Represents a single text column in a section.                                                                                                                                                                                                                                                                                                                                    |
| Word.Text<br>Column<br>Collection    | A collection of Word.TextColumn objects that represent all the columns of text in the<br>document or a section of the document.                                                                                                                                                                                                                                                  |
| Word.TextFrame                       | Represents the text frame of a shape object.                                                                                                                                                                                                                                                                                                                                     |
| Word.Three<br>Dimensional<br>Format  | Represents a shape's three-dimensional formatting.                                                                                                                                                                                                                                                                                                                               |
| Word.Tracked<br>Change               | Represents a tracked change in a Word document.                                                                                                                                                                                                                                                                                                                                  |
| Word.Tracked<br>Change<br>Collection | Contains a collection of Word.TrackedChange objects.                                                                                                                                                                                                                                                                                                                             |
| Word.View                            | Contains the view attributes (such as show all, field shading, and table gridlines) for a<br>window or pane.                                                                                                                                                                                                                                                                     |
| Word.Window                          | Represents the window that displays the document. A window can be split to contain<br>multiple reading panes.                                                                                                                                                                                                                                                                    |
| Word.Window<br>Collection            | Represents the collection of window objects.                                                                                                                                                                                                                                                                                                                                     |
| Word.Xml<br>Mapping                  | Represents the XML mapping on a Word.ContentControl object between custom XML<br>and that content control. An XML mapping is a link between the text in a content<br>control and an XML element in the custom XML data store for this document.                                                                                                                                  |

# **Interfaces**


| Word.Annotation<br>ClickedEventArgs                 | Holds annotation information that is passed back on annotation inserted event.                 |
|-----------------------------------------------------|------------------------------------------------------------------------------------------------|
| Word.Annotation<br>HoveredEvent<br>Args             | Holds annotation information that is passed back on annotation hovered event.                  |
| Word.Annotation<br>InsertedEventArgs                | Holds annotation information that is passed back on annotation added event.                    |
| Word.Annotation<br>PopupAction<br>EventArgs         | Represents action information that's passed back on annotation pop-up action<br>event.         |
| Word.Annotation<br>RemovedEvent<br>Args             | Holds annotation information that is passed back on annotation removed event.                  |
| Word.Annotation<br>Set                              | Annotations set produced by the add-in. Currently supporting only critiques.                   |
| Word.Comment<br>Detail                              | A structure for the ID and reply IDs of this comment.                                          |
| Word.Comment<br>EventArgs                           | Provides information about the comments that raised the comment event.                         |
| Word.Content<br>ControlAdded<br>EventArgs           | Provides information about the content control that raised contentControlAdded<br>event.       |
| Word.Content<br>ControlData<br>ChangedEvent<br>Args | Provides information about the content control that raised<br>contentControlDataChanged event. |
| Word.Content<br>ControlDeleted<br>EventArgs         | Provides information about the content control that raised contentControlDeleted<br>event.     |
| Word.Content<br>ControlEntered<br>EventArgs         | Provides information about the content control that raised contentControlEntered<br>event.     |
| Word.Content<br>ControlExited<br>EventArgs          | Provides information about the content control that raised contentControlExited<br>event.      |
| Word.Content<br>ControlOptions                      | Specifies the options that define which content controls are returned.                         |


| Word.Content<br>Control<br>Placeholder<br>Options        | The options that define what placeholder to be used in the content control.                                                        |
|----------------------------------------------------------|------------------------------------------------------------------------------------------------------------------------------------|
| Word.Content<br>ControlSelection<br>ChangedEvent<br>Args | Provides information about the content control that raised<br>contentControlSelectionChanged event.                                |
| Word.Critique                                            | Critique that will be rendered as underline for the specified part of paragraph in the<br>document.                                |
| Word.Critique<br>PopupOptions                            | Properties defining the behavior of the pop-up menu for a given critique.                                                          |
| Word.CustomXml<br>AddNodeOptions                         | The options for adding a node to the XML tree.                                                                                     |
| Word.CustomXml<br>AddSchema<br>Options                   | Adds one or more schemas to a schema collection that can then be added to a<br>stream in the data store and to the schema library. |
| Word.CustomXml<br>AddValidation<br>ErrorOptions          | The options that define the descriptive error text and the state of clearedOnUpdate .                                              |
| Word.CustomXml<br>AppendChild<br>NodeOptions             | The options that define the prefix mapping and the source of the custom XML data.                                                  |
| Word.CustomXml<br>InsertNodeBefore<br>Options            | Inserts a new node just before the context node in the tree.                                                                       |
| Word.CustomXml<br>InsertSubtree<br>BeforeOptions         | Inserts a new node just before the context node in the tree.                                                                       |
| Word.CustomXml<br>ReplaceChild<br>NodeOptions            | Removes the specified child node and replaces it with a different node in the same<br>location.                                    |
| Word.Document<br>CompareOptions                          | Specifies the options to be included in a compare document operation.                                                              |
| Word.GetText<br>Options                                  | Specifies the options to be included in a getText operation.                                                                       |
| Word.Hyperlink<br>AddOptions                             | Specifies the options for adding to a Word.HyperlinkCollection object.                                                             |


| Word.IndexAdd<br>Options                                    | Represents options for creating an index in a Word document.                                                      |
|-------------------------------------------------------------|-------------------------------------------------------------------------------------------------------------------|
| Word.IndexMark<br>AllEntriesOptions                         | Represents options for marking all index entries in a Word document.                                              |
| Word.IndexMark<br>EntryOptions                              | Represents options for marking an index entry in a Word document.                                                 |
| Word.InsertFile<br>Options                                  | Specifies the options to determine what to copy when inserting a file.                                            |
| Word.InsertShape<br>Options                                 | Specifies the options to determine location and size when inserting a shape.                                      |
| Word.Interfaces.<br>Annotation<br>CollectionData            | An interface describing the data returned by calling annotationCollection.toJSON() .                              |
| Word.Interfaces.<br>Annotation<br>CollectionLoad<br>Options | Contains a collection of Word.Annotation objects.                                                                 |
| Word.Interfaces.<br>Annotation<br>CollectionUpdate<br>Data  | An interface for updating data on the AnnotationCollection object, for use in<br>annotationCollection.set({  }) . |
| Word.Interfaces.<br>AnnotationData                          | An interface describing the data returned by calling annotation.toJSON() .                                        |
| Word.Interfaces.<br>AnnotationLoad<br>Options               | Represents an annotation attached to a paragraph.                                                                 |
| Word.Interfaces.<br>ApplicationData                         | An interface describing the data returned by calling application.toJSON() .                                       |
| Word.Interfaces.<br>ApplicationLoad<br>Options              | Represents the application object.                                                                                |
| Word.Interfaces.<br>Application<br>UpdateData               | An interface for updating data on the Application object, for use in<br>application.set({  }) .                   |
| Word.Interfaces.<br>BibliographyData                        | An interface describing the data returned by calling bibliography.toJSON() .                                      |


| Word.Interfaces.<br>BibliographyLoad<br>Options           | Represents the list of available sources attached to the document (in the current list)<br>or the list of sources available in the application (in the master list). |
|-----------------------------------------------------------|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| Word.Interfaces.<br>Bibliography<br>UpdateData            | An interface for updating data on the Bibliography object, for use in<br>bibliography.set({  }) .                                                                    |
| Word.Interfaces.<br>BodyData                              | An interface describing the data returned by calling body.toJSON() .                                                                                                 |
| Word.Interfaces.<br>BodyLoadOptions                       | Represents the body of a document or a section.                                                                                                                      |
| Word.Interfaces.<br>BodyUpdateData                        | An interface for updating data on the Body object, for use in body.set({  }) .                                                                                       |
| Word.Interfaces.<br>Bookmark<br>CollectionData            | An interface describing the data returned by calling bookmarkCollection.toJSON() .                                                                                   |
| Word.Interfaces.<br>Bookmark<br>CollectionLoad<br>Options | A collection of Word.Bookmark objects that represent the bookmarks in the specified<br>selection, range, or document.                                                |
| Word.Interfaces.<br>Bookmark<br>CollectionUpdate<br>Data  | An interface for updating data on the BookmarkCollection object, for use in<br>bookmarkCollection.set({  }) .                                                        |
| Word.Interfaces.<br>BookmarkData                          | An interface describing the data returned by calling bookmark.toJSON() .                                                                                             |
| Word.Interfaces.                                          | Represents a single bookmark in a document, selection, or range. The Bookmark                                                                                        |
| BookmarkLoad                                              | object is a member of the Bookmark collection. The Word.BookmarkCollection                                                                                           |
| Options                                                   | includes all the bookmarks listed in the Bookmark dialog box (Insert menu).                                                                                          |
| Word.Interfaces.                                          | An interface for updating data on the Bookmark object, for use in bookmark.set({                                                                                     |
| BookmarkUpdate                                            |                                                                                                                                                                      |
| Data                                                      | }) .                                                                                                                                                                 |
| Word.Interfaces.<br>BorderCollection<br>Data              | An interface describing the data returned by calling borderCollection.toJSON()                                                                                       |
| Word.Interfaces.<br>BorderCollection<br>LoadOptions       | Represents the collection of border styles.                                                                                                                          |


| Word.Interfaces.<br>BorderCollection<br>UpdateData               | An interface for updating data on the BorderCollection object, for use in<br>borderCollection.set({  }) .                   |
|------------------------------------------------------------------|-----------------------------------------------------------------------------------------------------------------------------|
| Word.Interfaces.<br>BorderData                                   | An interface describing the data returned by calling border.toJSON() .                                                      |
| Word.Interfaces.<br>BorderLoad<br>Options                        | Represents the Border object for text, a paragraph, or a table.                                                             |
| Word.Interfaces.<br>BorderUniversal<br>CollectionData            | An interface describing the data returned by calling<br>borderUniversalCollection.toJSON() .                                |
| Word.Interfaces.<br>BorderUniversal<br>CollectionLoad<br>Options | Represents the collection of Word.BorderUniversal objects.                                                                  |
| Word.Interfaces.<br>BorderUniversal<br>CollectionUpdate<br>Data  | An interface for updating data on the BorderUniversalCollection object, for use in<br>borderUniversalCollection.set({  }) . |
| Word.Interfaces.<br>BorderUniversal<br>Data                      | An interface describing the data returned by calling borderUniversal.toJSON() .                                             |
| Word.Interfaces.<br>BorderUniversal<br>LoadOptions               | Represents the BorderUniversal object, which manages borders for a range,<br>paragraph, table, or frame.                    |
| Word.Interfaces.<br>BorderUniversal<br>UpdateData                | An interface for updating data on the BorderUniversal object, for use in<br>borderUniversal.set({  }) .                     |
| Word.Interfaces.<br>BorderUpdate<br>Data                         | An interface for updating data on the Border object, for use in border.set({  }) .                                          |
| Word.Interfaces.<br>BreakCollection<br>Data                      | An interface describing the data returned by calling breakCollection.toJSON() .                                             |
| Word.Interfaces.<br>BreakCollection<br>LoadOptions               | Contains a collection of Word.Break objects.                                                                                |
| Word.Interfaces.                                                 | An interface for updating data on the BreakCollection object, for use in                                                    |
| BreakCollection                                                  | breakCollection.set({  }) .                                                                                                 |


| UpdateData                                                                    |                                                                                                                                                          |
|-------------------------------------------------------------------------------|----------------------------------------------------------------------------------------------------------------------------------------------------------|
| Word.Interfaces.                                                              | An interface describing the data returned by calling break.toJSON()                                                                                      |
| BreakData                                                                     |                                                                                                                                                          |
| Word.Interfaces.<br>BreakLoad<br>Options                                      | Represents a break in a Word document.                                                                                                                   |
| Word.Interfaces.                                                              | An interface for updating data on the Break object, for use in break.set({                                                                               |
| BreakUpdateData                                                               | }) .                                                                                                                                                     |
| Word.Interfaces.<br>BuildingBlock<br>CategoryData                             | An interface describing the data returned by calling buildingBlockCategory.toJSON()                                                                      |
| Word.Interfaces.<br>BuildingBlock<br>CategoryLoad<br>Options                  | Represents a category of building blocks in a Word document.                                                                                             |
| Word.Interfaces.                                                              | An interface describing the data returned by calling buildingBlock.toJSON()                                                                              |
| BuildingBlockData                                                             |                                                                                                                                                          |
| Word.Interfaces.<br>BuildingBlock<br>GalleryContent<br>ControlData            | An interface describing the data returned by calling<br>buildingBlockGalleryContentControl.toJSON() .                                                    |
| Word.Interfaces.<br>BuildingBlock<br>GalleryContent<br>ControlLoad<br>Options | Represents the BuildingBlockGalleryContentControl object.                                                                                                |
| Word.Interfaces.<br>BuildingBlock<br>GalleryContent<br>ControlUpdate<br>Data  | An interface for updating data on the BuildingBlockGalleryContentControl object,<br>for use in buildingBlockGalleryContentControl.set({<br>}) .          |
| Word.Interfaces.<br>BuildingBlock<br>LoadOptions                              | Represents a building block in a template. A building block is pre-built content,<br>similar to autotext, that may contain text, images, and formatting. |
| Word.Interfaces.<br>BuildingBlockType<br>ItemData                             | An interface describing the data returned by calling buildingBlockTypeItem.toJSON()                                                                      |
| Word.Interfaces.<br>BuildingBlockType                                         | Represents a type of building block in a Word document.                                                                                                  |


| ItemLoadOptions                                               |                                                                                                                       |
|---------------------------------------------------------------|-----------------------------------------------------------------------------------------------------------------------|
| Word.Interfaces.<br>BuildingBlock<br>UpdateData               | An interface for updating data on the BuildingBlock object, for use in<br>buildingBlock.set({  }) .                   |
| Word.Interfaces.                                              | An interface describing the data returned by calling canvas.toJSON()                                                  |
| CanvasData                                                    |                                                                                                                       |
| Word.Interfaces.<br>CanvasLoad<br>Options                     | Represents a canvas in the document. To get the corresponding Shape object, use<br>Canvas.shape.                      |
| Word.Interfaces.<br>CanvasUpdate<br>Data                      | An interface for updating data on the Canvas object, for use in canvas.set({<br>}) .                                  |
| Word.Interfaces.<br>CheckboxContent<br>ControlData            | An interface describing the data returned by calling<br>checkboxContentControl.toJSON() .                             |
| Word.Interfaces.<br>CheckboxContent<br>ControlLoad<br>Options | The data specific to content controls of type CheckBox.                                                               |
| Word.Interfaces.<br>CheckboxContent<br>ControlUpdate<br>Data  | An interface for updating data on the CheckboxContentControl object, for use in<br>checkboxContentControl.set({  }) . |
| Word.Interfaces.<br>CollectionLoad<br>Options                 | Provides ways to load properties of only a subset of members of a collection.                                         |
| Word.Interfaces.                                              | An interface describing the data returned by calling colorFormat.toJSON()                                             |
| ColorFormatData                                               |                                                                                                                       |
| Word.Interfaces.<br>ColorFormatLoad<br>Options                | Represents the color formatting of a shape or text in Word.                                                           |
| Word.Interfaces.<br>ColorFormat<br>UpdateData                 | An interface for updating data on the ColorFormat object, for use in<br>colorFormat.set({  }) .                       |
| Word.Interfaces.<br>ComboBox<br>ContentControl<br>Data        | An interface describing the data returned by calling<br>comboBoxContentControl.toJSON() .                             |


| Word.Interfaces.<br>Comment<br>CollectionData                 | An interface describing the data returned by calling commentCollection.toJSON() .                                     |
|---------------------------------------------------------------|-----------------------------------------------------------------------------------------------------------------------|
| Word.Interfaces.<br>Comment<br>CollectionLoad<br>Options      | Contains a collection of Word.Comment objects.                                                                        |
| Word.Interfaces.<br>Comment<br>CollectionUpdate<br>Data       | An interface for updating data on the CommentCollection object, for use in<br>commentCollection.set({  }) .           |
| Word.Interfaces.<br>CommentContent<br>RangeData               | An interface describing the data returned by calling commentContentRange.toJSON() .                                   |
|                                                               | Word.Interfaces.CommentContentRangeLoadOptions                                                                        |
| Word.Interfaces.<br>CommentContent<br>RangeUpdateData         | An interface for updating data on the CommentContentRange object, for use in<br>commentContentRange.set({  }) .       |
| Word.Interfaces.<br>CommentData                               | An interface describing the data returned by calling comment.toJSON() .                                               |
| Word.Interfaces.<br>CommentLoad<br>Options                    | Represents a comment in the document.                                                                                 |
| Word.Interfaces.<br>CommentReply<br>CollectionData            | An interface describing the data returned by calling<br>commentReplyCollection.toJSON() .                             |
| Word.Interfaces.<br>CommentReply<br>CollectionLoad<br>Options | Contains a collection of Word.CommentReply objects. Represents all comment<br>replies in one comment thread.          |
| Word.Interfaces.<br>CommentReply<br>CollectionUpdate<br>Data  | An interface for updating data on the CommentReplyCollection object, for use in<br>commentReplyCollection.set({  }) . |
| Word.Interfaces.<br>CommentReply<br>Data                      | An interface describing the data returned by calling commentReply.toJSON() .                                          |
| Word.Interfaces.<br>CommentReply                              | Represents a comment reply in the document.                                                                           |


| LoadOptions                                                             |                                                                                                                                                                                                                                                                                                                                                                                                                              |
|-------------------------------------------------------------------------|------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| Word.Interfaces.<br>CommentReply<br>UpdateData                          | An interface for updating data on the CommentReply object, for use in<br>commentReply.set({  }) .                                                                                                                                                                                                                                                                                                                            |
| Word.Interfaces.                                                        | An interface for updating data on the Comment object, for use in comment.set({                                                                                                                                                                                                                                                                                                                                               |
| CommentUpdate                                                           |                                                                                                                                                                                                                                                                                                                                                                                                                              |
| Data                                                                    | }) .                                                                                                                                                                                                                                                                                                                                                                                                                         |
| Word.Interfaces.<br>ContentControl<br>CollectionData                    | An interface describing the data returned by calling<br>contentControlCollection.toJSON() .                                                                                                                                                                                                                                                                                                                                  |
| Word.Interfaces.<br>ContentControl<br>CollectionLoad<br>Options         | Contains a collection of Word.ContentControl objects. Content controls are bounded<br>and potentially labeled regions in a document that serve as containers for specific<br>types of content. Individual content controls may contain contents such as images,<br>tables, or paragraphs of formatted text. Currently, only rich text, plain text, checkbox,<br>dropdown list, and combo box content controls are supported. |
| Word.Interfaces.<br>ContentControl<br>CollectionUpdate<br>Data          | An interface for updating data on the ContentControlCollection object, for use in<br>contentControlCollection.set({  }) .                                                                                                                                                                                                                                                                                                    |
| Word.Interfaces.<br>ContentControl<br>Data                              | An interface describing the data returned by calling contentControl.toJSON()                                                                                                                                                                                                                                                                                                                                                 |
| Word.Interfaces.<br>ContentControl<br>ListItemCollection<br>Data        | An interface describing the data returned by calling<br>contentControlListItemCollection.toJSON() .                                                                                                                                                                                                                                                                                                                          |
| Word.Interfaces.<br>ContentControl<br>ListItemCollection<br>LoadOptions | Contains a collection of Word.ContentControlListItem objects that represent the<br>items in a dropdown list or combo box content control.                                                                                                                                                                                                                                                                                    |
| Word.Interfaces.<br>ContentControl<br>ListItemCollection<br>UpdateData  | An interface for updating data on the ContentControlListItemCollection object, for<br>use in contentControlListItemCollection.set({<br>}) .                                                                                                                                                                                                                                                                                  |
| Word.Interfaces.<br>ContentControl<br>ListItemData                      | An interface describing the data returned by calling<br>contentControlListItem.toJSON() .                                                                                                                                                                                                                                                                                                                                    |
| Word.Interfaces.<br>ContentControl                                      | Represents a list item in a dropdown list or combo box content control.                                                                                                                                                                                                                                                                                                                                                      |


| ListItemLoad<br>Options                                         |                                                                                                                                                                                                                                                                                                                                                                                                      |
|-----------------------------------------------------------------|------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| Word.Interfaces.<br>ContentControl<br>ListItemUpdate<br>Data    | An interface for updating data on the ContentControlListItem object, for use in<br>contentControlListItem.set({  }) .                                                                                                                                                                                                                                                                                |
| Word.Interfaces.<br>ContentControl<br>LoadOptions               | Represents a content control. Content controls are bounded and potentially labeled<br>regions in a document that serve as containers for specific types of content.<br>Individual content controls may contain contents such as images, tables, or<br>paragraphs of formatted text. Currently, only rich text, plain text, checkbox,<br>dropdown list, and combo box content controls are supported. |
| Word.Interfaces.<br>ContentControl<br>UpdateData                | An interface for updating data on the ContentControl object, for use in<br>contentControl.set({  }) .                                                                                                                                                                                                                                                                                                |
| Word.Interfaces.<br>Critique<br>AnnotationData                  | An interface describing the data returned by calling critiqueAnnotation.toJSON()                                                                                                                                                                                                                                                                                                                     |
| Word.Interfaces.<br>Critique<br>AnnotationLoad<br>Options       | Represents an annotation wrapper around critique displayed in the document.                                                                                                                                                                                                                                                                                                                          |
| Word.Interfaces.<br>CustomProperty<br>CollectionData            | An interface describing the data returned by calling<br>customPropertyCollection.toJSON() .                                                                                                                                                                                                                                                                                                          |
| Word.Interfaces.<br>CustomProperty<br>CollectionLoad<br>Options | Contains the collection of Word.CustomProperty objects.                                                                                                                                                                                                                                                                                                                                              |
| Word.Interfaces.<br>CustomProperty<br>CollectionUpdate<br>Data  | An interface for updating data on the CustomPropertyCollection object, for use in<br>customPropertyCollection.set({  }) .                                                                                                                                                                                                                                                                            |
| Word.Interfaces.<br>CustomProperty<br>Data                      | An interface describing the data returned by calling customProperty.toJSON()                                                                                                                                                                                                                                                                                                                         |
| Word.Interfaces.<br>CustomProperty<br>LoadOptions               | Represents a custom property.                                                                                                                                                                                                                                                                                                                                                                        |
| Word.Interfaces.                                                | An interface for updating data on the CustomProperty object, for use in                                                                                                                                                                                                                                                                                                                              |
| CustomProperty                                                  | customProperty.set({  }) .                                                                                                                                                                                                                                                                                                                                                                           |


| UpdateData                                                     |                                                                                                                                       |
|----------------------------------------------------------------|---------------------------------------------------------------------------------------------------------------------------------------|
| Word.Interfaces.<br>CustomXmlNode<br>CollectionData            | An interface describing the data returned by calling<br>customXmlNodeCollection.toJSON() .                                            |
| Word.Interfaces.<br>CustomXmlNode<br>CollectionLoad<br>Options | Contains a collection of Word.CustomXmlNode objects representing the XML nodes<br>in a document.                                      |
| Word.Interfaces.<br>CustomXmlNode<br>CollectionUpdate<br>Data  | An interface for updating data on the CustomXmlNodeCollection object, for use in<br>customXmlNodeCollection.set({  }) .               |
| Word.Interfaces.<br>CustomXmlNode<br>Data                      | An interface describing the data returned by calling customXmlNode.toJSON()                                                           |
| Word.Interfaces.<br>CustomXmlNode<br>LoadOptions               | Represents an XML node in a tree in the document. The CustomXmlNode object is a<br>member of the Word.CustomXmlNodeCollection object. |
| Word.Interfaces.<br>CustomXmlNode<br>UpdateData                | An interface for updating data on the CustomXmlNode object, for use in<br>customXmlNode.set({  }) .                                   |
| Word.Interfaces.<br>CustomXmlPart<br>CollectionData            | An interface describing the data returned by calling<br>customXmlPartCollection.toJSON() .                                            |
| Word.Interfaces.<br>CustomXmlPart<br>CollectionLoad<br>Options | Contains the collection of Word.CustomXmlPart objects.                                                                                |
| Word.Interfaces.<br>CustomXmlPart<br>CollectionUpdate<br>Data  | An interface for updating data on the CustomXmlPartCollection object, for use in<br>customXmlPartCollection.set({  }) .               |
| Word.Interfaces.<br>CustomXmlPart<br>Data                      | An interface describing the data returned by calling customXmlPart.toJSON()                                                           |
| Word.Interfaces.<br>CustomXmlPart<br>LoadOptions               | Represents a custom XML part.                                                                                                         |


| Word.Interfaces.<br>CustomXmlPart<br>ScopedCollection<br>Data               | An interface describing the data returned by calling<br>customXmlPartScopedCollection.toJSON() .                                            |
|-----------------------------------------------------------------------------|---------------------------------------------------------------------------------------------------------------------------------------------|
| Word.Interfaces.<br>CustomXmlPart<br>ScopedCollection<br>LoadOptions        | Contains the collection of Word.CustomXmlPart objects with a specific namespace.                                                            |
| Word.Interfaces.<br>CustomXmlPart<br>ScopedCollection<br>UpdateData         | An interface for updating data on the CustomXmlPartScopedCollection object, for use<br>in customXmlPartScopedCollection.set({<br>}) .       |
| Word.Interfaces.<br>CustomXmlPart<br>UpdateData                             | An interface for updating data on the CustomXmlPart object, for use in<br>customXmlPart.set({  }) .                                         |
| Word.Interfaces.<br>CustomXmlPrefix<br>Mapping<br>CollectionData            | An interface describing the data returned by calling<br>customXmlPrefixMappingCollection.toJSON() .                                         |
| Word.Interfaces.<br>CustomXmlPrefix<br>Mapping<br>CollectionLoad<br>Options | Represents a collection of Word.CustomXmlPrefixMapping objects.                                                                             |
| Word.Interfaces.<br>CustomXmlPrefix<br>Mapping<br>CollectionUpdate<br>Data  | An interface for updating data on the CustomXmlPrefixMappingCollection object, for<br>use in customXmlPrefixMappingCollection.set({<br>}) . |
| Word.Interfaces.<br>CustomXmlPrefix<br>MappingData                          | An interface describing the data returned by calling<br>customXmlPrefixMapping.toJSON() .                                                   |
| Word.Interfaces.<br>CustomXmlPrefix<br>MappingLoad<br>Options               | Represents a CustomXmlPrefixMapping object.                                                                                                 |
| Word.Interfaces.<br>CustomXml<br>SchemaCollection<br>Data                   | An interface describing the data returned by calling<br>customXmlSchemaCollection.toJSON() .                                                |


| Word.Interfaces.<br>CustomXml<br>SchemaCollection<br>LoadOptions              | Represents a collection of Word.CustomXmlSchema objects attached to a data<br>stream.                                                           |
|-------------------------------------------------------------------------------|-------------------------------------------------------------------------------------------------------------------------------------------------|
| Word.Interfaces.<br>CustomXml<br>SchemaCollection<br>UpdateData               | An interface for updating data on the CustomXmlSchemaCollection object, for use in<br>customXmlSchemaCollection.set({  }) .                     |
| Word.Interfaces.<br>CustomXml<br>SchemaData                                   | An interface describing the data returned by calling customXmlSchema.toJSON() .                                                                 |
| Word.Interfaces.<br>CustomXml<br>SchemaLoad<br>Options                        | Represents a schema in a Word.CustomXmlSchemaCollection object.                                                                                 |
| Word.Interfaces.<br>CustomXml<br>ValidationError<br>CollectionData            | An interface describing the data returned by calling<br>customXmlValidationErrorCollection.toJSON() .                                           |
| Word.Interfaces.<br>CustomXml<br>ValidationError<br>CollectionLoad<br>Options | Represents a collection of Word.CustomXmlValidationError objects.                                                                               |
| Word.Interfaces.<br>CustomXml<br>ValidationError<br>CollectionUpdate<br>Data  | An interface for updating data on the CustomXmlValidationErrorCollection object,<br>for use in customXmlValidationErrorCollection.set({<br>}) . |
| Word.Interfaces.<br>CustomXml<br>ValidationError<br>Data                      | An interface describing the data returned by calling<br>customXmlValidationError.toJSON() .                                                     |
| Word.Interfaces.<br>CustomXml<br>ValidationError<br>LoadOptions               | Represents a single validation error in a Word.CustomXmlValidationErrorCollection<br>object.                                                    |
| Word.Interfaces.<br>CustomXml<br>ValidationError<br>UpdateData                | An interface for updating data on the CustomXmlValidationError object, for use in<br>customXmlValidationError.set({  }) .                       |


| Word.Interfaces.<br>DatePicker<br>ContentControl<br>Data                | An interface describing the data returned by calling<br>datePickerContentControl.toJSON() .                                                 |
|-------------------------------------------------------------------------|---------------------------------------------------------------------------------------------------------------------------------------------|
| Word.Interfaces.<br>DatePicker<br>ContentControl<br>LoadOptions         | Represents the DatePickerContentControl object.                                                                                             |
| Word.Interfaces.<br>DatePicker<br>ContentControl<br>UpdateData          | An interface for updating data on the DatePickerContentControl object, for use in<br>datePickerContentControl.set({  }) .                   |
| Word.Interfaces.<br>DocumentCreated<br>Data                             | An interface describing the data returned by calling documentCreated.toJSON() .                                                             |
| Word.Interfaces.                                                        | The DocumentCreated object is the top level object created by                                                                               |
| DocumentCreated                                                         | Application.CreateDocument. A DocumentCreated object is a special Document                                                                  |
| LoadOptions                                                             | object.                                                                                                                                     |
| Word.Interfaces.<br>DocumentCreated<br>UpdateData                       | An interface for updating data on the DocumentCreated object, for use in<br>documentCreated.set({  }) .                                     |
| Word.Interfaces.<br>DocumentData                                        | An interface describing the data returned by calling document.toJSON() .                                                                    |
| Word.Interfaces.<br>DocumentLibrary<br>VersionCollection<br>Data        | An interface describing the data returned by calling<br>documentLibraryVersionCollection.toJSON() .                                         |
| Word.Interfaces.<br>DocumentLibrary<br>VersionCollection<br>LoadOptions | Represents the collection of Word.DocumentLibraryVersion objects.                                                                           |
| Word.Interfaces.<br>DocumentLibrary<br>VersionCollection<br>UpdateData  | An interface for updating data on the DocumentLibraryVersionCollection object, for<br>use in documentLibraryVersionCollection.set({<br>}) . |
| Word.Interfaces.<br>DocumentLibrary<br>VersionData                      | An interface describing the data returned by calling<br>documentLibraryVersion.toJSON() .                                                   |
| Word.Interfaces.<br>DocumentLibrary                                     | Represents a document library version.                                                                                                      |


| VersionLoad<br>Options                                     |                                                                                                               |
|------------------------------------------------------------|---------------------------------------------------------------------------------------------------------------|
| Word.Interfaces.                                           | The Document object is the top level object. A Document object contains one or                                |
| DocumentLoad                                               | more sections, content controls, and the body that contains the contents of the                               |
| Options                                                    | document.                                                                                                     |
| Word.Interfaces.<br>Document<br>PropertiesData             | An interface describing the data returned by calling documentProperties.toJSON()                              |
| Word.Interfaces.<br>Document<br>PropertiesLoad<br>Options  | Represents document properties.                                                                               |
| Word.Interfaces.<br>Document<br>PropertiesUpdate<br>Data   | An interface for updating data on the DocumentProperties object, for use in<br>documentProperties.set({  }) . |
| Word.Interfaces.                                           | An interface for updating data on the Document object, for use in document.set({                              |
| DocumentUpdate                                             |                                                                                                               |
| Data                                                       | }) .                                                                                                          |
| Word.Interfaces.                                           | An interface describing the data returned by calling dropCap.toJSON()                                         |
| DropCapData                                                |                                                                                                               |
| Word.Interfaces.<br>DropCapLoad<br>Options                 | Represents a dropped capital letter in a Word document.                                                       |
| Word.Interfaces.<br>DropDownList<br>ContentControl<br>Data | An interface describing the data returned by calling<br>dropDownListContentControl.toJSON() .                 |
| Word.Interfaces.<br>FieldCollection<br>Data                | An interface describing the data returned by calling fieldCollection.toJSON()                                 |
| Word.Interfaces.<br>FieldCollection<br>LoadOptions         | Contains a collection of Word.Field objects.                                                                  |
| Word.Interfaces.<br>FieldCollection<br>UpdateData          | An interface for updating data on the FieldCollection object, for use in<br>fieldCollection.set({  }) .       |
| Word.Interfaces.                                           | An interface describing the data returned by calling field.toJSON()                                           |
| FieldData                                                  |                                                                                                               |


| Word.Interfaces.<br>FieldLoadOptions               | Represents a field.                                                                                     |
|----------------------------------------------------|---------------------------------------------------------------------------------------------------------|
| Word.Interfaces.<br>FieldUpdateData                | An interface for updating data on the Field object, for use in field.set({  }) .                        |
| Word.Interfaces.<br>FillFormatData                 | An interface describing the data returned by calling fillFormat.toJSON() .                              |
| Word.Interfaces.<br>FillFormatLoad<br>Options      | Represents the fill formatting for a shape or text.                                                     |
| Word.Interfaces.<br>FillFormatUpdate<br>Data       | An interface for updating data on the FillFormat object, for use in fillFormat.set({<br>}) .            |
| Word.Interfaces.<br>FontData                       | An interface describing the data returned by calling font.toJSON() .                                    |
| Word.Interfaces.<br>FontLoadOptions                | Represents a font.                                                                                      |
| Word.Interfaces.<br>FontUpdateData                 | An interface for updating data on the Font object, for use in font.set({  }) .                          |
| Word.Interfaces.<br>FrameCollection<br>Data        | An interface describing the data returned by calling frameCollection.toJSON() .                         |
| Word.Interfaces.<br>FrameCollection<br>LoadOptions | Represents the collection of Word.Frame objects.                                                        |
| Word.Interfaces.<br>FrameCollection<br>UpdateData  | An interface for updating data on the FrameCollection object, for use in<br>frameCollection.set({  }) . |
| Word.Interfaces.<br>FrameData                      | An interface describing the data returned by calling frame.toJSON() .                                   |
| Word.Interfaces.<br>FrameLoad<br>Options           | Represents a frame. The Frame object is a member of the Word.FrameCollection<br>object.                 |
| Word.Interfaces.<br>FrameUpdateData                | An interface for updating data on the Frame object, for use in frame.set({  }) .                        |
| Word.Interfaces.<br>GlowFormatData                 | An interface describing the data returned by calling glowFormat.toJSON() .                              |


| Word.Interfaces.<br>GlowFormatLoad<br>Options              | Represents the glow formatting for the font used by the range of text.                                          |
|------------------------------------------------------------|-----------------------------------------------------------------------------------------------------------------|
| Word.Interfaces.<br>GlowFormat<br>UpdateData               | An interface for updating data on the GlowFormat object, for use in glowFormat.set({<br>}) .                    |
| Word.Interfaces.<br>GroupContent<br>ControlData            | An interface describing the data returned by calling groupContentControl.toJSON() .                             |
| Word.Interfaces.<br>GroupContent<br>ControlLoad<br>Options | Represents the GroupContentControl object.                                                                      |
| Word.Interfaces.<br>GroupContent<br>ControlUpdate<br>Data  | An interface for updating data on the GroupContentControl object, for use in<br>groupContentControl.set({  }) . |
| Word.Interfaces.<br>Hyperlink<br>CollectionData            | An interface describing the data returned by calling hyperlinkCollection.toJSON() .                             |
| Word.Interfaces.<br>Hyperlink<br>CollectionLoad<br>Options | Contains a collection of Word.Hyperlink objects.                                                                |
| Word.Interfaces.<br>Hyperlink<br>CollectionUpdate<br>Data  | An interface for updating data on the HyperlinkCollection object, for use in<br>hyperlinkCollection.set({  }) . |
| Word.Interfaces.<br>HyperlinkData                          | An interface describing the data returned by calling hyperlink.toJSON() .                                       |
| Word.Interfaces.<br>HyperlinkLoad<br>Options               | Represents a hyperlink in a Word document.                                                                      |
| Word.Interfaces.<br>HyperlinkUpdate<br>Data                | An interface for updating data on the Hyperlink object, for use in hyperlink.set({<br>}) .                      |
| Word.Interfaces.<br>IndexCollection<br>Data                | An interface describing the data returned by calling indexCollection.toJSON() .                                 |


| Word.Interfaces.<br>IndexCollection<br>LoadOptions             | A collection of Word.Index objects that represents all the indexes in the document.                                     |
|----------------------------------------------------------------|-------------------------------------------------------------------------------------------------------------------------|
| Word.Interfaces.<br>IndexCollection<br>UpdateData              | An interface for updating data on the IndexCollection object, for use in<br>indexCollection.set({  }) .                 |
| Word.Interfaces.<br>IndexData                                  | An interface describing the data returned by calling index.toJSON() .                                                   |
| Word.Interfaces.                                               | Represents a single index. The Index object is a member of the                                                          |
| IndexLoadOptions                                               | Word.IndexCollection. The IndexCollection includes all the indexes in the document.                                     |
| Word.Interfaces.                                               | An interface for updating data on the Index object, for use in index.set({                                              |
| IndexUpdateData                                                | }) .                                                                                                                    |
| Word.Interfaces.<br>InlinePicture<br>CollectionData            | An interface describing the data returned by calling<br>inlinePictureCollection.toJSON() .                              |
| Word.Interfaces.<br>InlinePicture<br>CollectionLoad<br>Options | Contains a collection of Word.InlinePicture objects.                                                                    |
|                                                                |                                                                                                                         |
| Word.Interfaces.<br>InlinePicture<br>CollectionUpdate<br>Data  | An interface for updating data on the InlinePictureCollection object, for use in<br>inlinePictureCollection.set({  }) . |
| Word.Interfaces.                                               | An interface describing the data returned by calling inlinePicture.toJSON()                                             |
| InlinePictureData                                              |                                                                                                                         |
| Word.Interfaces.<br>InlinePictureLoad<br>Options               | Represents an inline picture.                                                                                           |
| Word.Interfaces.<br>InlinePicture<br>UpdateData                | An interface for updating data on the InlinePicture object, for use in<br>inlinePicture.set({  }) .                     |
| Word.Interfaces.                                               | An interface describing the data returned by calling lineFormat.toJSON()                                                |
| LineFormatData                                                 |                                                                                                                         |
| Word.Interfaces.                                               | Represents line and arrowhead formatting. For a line, the LineFormat object contains                                    |
| LineFormatLoad                                                 | formatting information for the line itself; for a shape with a border, this object                                      |
| Options                                                        | contains formatting information for the shape's border.                                                                 |


| UpdateData                                        |                                                                                                       |
|---------------------------------------------------|-------------------------------------------------------------------------------------------------------|
| Word.Interfaces.<br>LineNumbering<br>Data         | An interface describing the data returned by calling lineNumbering.toJSON()                           |
| Word.Interfaces.<br>LineNumbering<br>LoadOptions  | Represents line numbers in the left margin or to the left of each newspaper-style<br>column.          |
| Word.Interfaces.<br>LineNumbering<br>UpdateData   | An interface for updating data on the LineNumbering object, for use in<br>lineNumbering.set({  }) .   |
| Word.Interfaces.                                  | An interface describing the data returned by calling linkFormat.toJSON()                              |
| LinkFormatData                                    |                                                                                                       |
| Word.Interfaces.<br>LinkFormatLoad<br>Options     | Represents the linking characteristics for an OLE object or picture.                                  |
| Word.Interfaces.<br>LinkFormat<br>UpdateData      | An interface for updating data on the LinkFormat object, for use in linkFormat.set({<br>}) .          |
| Word.Interfaces.                                  | An interface describing the data returned by calling listCollection.toJSON()                          |
| ListCollectionData                                |                                                                                                       |
| Word.Interfaces.<br>ListCollectionLoad<br>Options | Contains a collection of Word.List objects.                                                           |
| Word.Interfaces.<br>ListCollection<br>UpdateData  | An interface for updating data on the ListCollection object, for use in<br>listCollection.set({  }) . |
| Word.Interfaces.                                  | An interface describing the data returned by calling list.toJSON()                                    |
| ListData                                          |                                                                                                       |
| Word.Interfaces.                                  | An interface describing the data returned by calling listFormat.toJSON()                              |
| ListFormatData                                    |                                                                                                       |
| Word.Interfaces.<br>ListFormatLoad<br>Options     | Represents the list formatting characteristics of a range.                                            |
| Word.Interfaces.<br>ListFormatUpdate<br>Data      | An interface for updating data on the ListFormat object, for use in listFormat.set({<br>}) .          |
| Word.Interfaces.                                  | An interface describing the data returned by calling listItem.toJSON()                                |
| ListItemData                                      |                                                                                                       |


| Word.Interfaces.<br>ListItemLoad<br>Options                | Represents the paragraph list item format.                                                                      |
|------------------------------------------------------------|-----------------------------------------------------------------------------------------------------------------|
| Word.Interfaces.<br>ListItemUpdate<br>Data                 | An interface for updating data on the ListItem object, for use in listItem.set({<br>}) .                        |
| Word.Interfaces.<br>ListLevel<br>CollectionData            | An interface describing the data returned by calling listLevelCollection.toJSON() .                             |
| Word.Interfaces.<br>ListLevel<br>CollectionLoad<br>Options | Contains a collection of Word.ListLevel objects.                                                                |
| Word.Interfaces.<br>ListLevel<br>CollectionUpdate<br>Data  | An interface for updating data on the ListLevelCollection object, for use in<br>listLevelCollection.set({  }) . |
| Word.Interfaces.<br>ListLevelData                          | An interface describing the data returned by calling listLevel.toJSON() .                                       |
| Word.Interfaces.<br>ListLevelLoad<br>Options               | Represents a list level.                                                                                        |
| Word.Interfaces.<br>ListLevelUpdate<br>Data                | An interface for updating data on the ListLevel object, for use in listLevel.set({<br>}) .                      |
| Word.Interfaces.<br>ListLoadOptions                        | Contains a collection of Word.Paragraph objects.                                                                |
| Word.Interfaces.<br>ListTemplateData                       | An interface describing the data returned by calling listTemplate.toJSON() .                                    |
| Word.Interfaces.<br>ListTemplateLoad<br>Options            | Represents a list template.                                                                                     |
| Word.Interfaces.<br>ListTemplate<br>UpdateData             | An interface for updating data on the ListTemplate object, for use in<br>listTemplate.set({  }) .               |
| Word.Interfaces.<br>NoteItem<br>CollectionData             | An interface describing the data returned by calling noteItemCollection.toJSON() .                              |


| Word.Interfaces.<br>NoteItem<br>CollectionLoad<br>Options | Contains a collection of Word.NoteItem objects.                                                               |
|-----------------------------------------------------------|---------------------------------------------------------------------------------------------------------------|
| Word.Interfaces.<br>NoteItem<br>CollectionUpdate<br>Data  | An interface for updating data on the NoteItemCollection object, for use in<br>noteItemCollection.set({  }) . |
| Word.Interfaces.<br>NoteItemData                          | An interface describing the data returned by calling noteItem.toJSON() .                                      |
| Word.Interfaces.<br>NoteItemLoad<br>Options               | Represents a footnote or endnote.                                                                             |
| Word.Interfaces.<br>NoteItemUpdate<br>Data                | An interface for updating data on the NoteItem object, for use in noteItem.set({<br>}) .                      |
| Word.Interfaces.<br>OleFormatData                         | An interface describing the data returned by calling oleFormat.toJSON() .                                     |
| Word.Interfaces.<br>OleFormatLoad<br>Options              | Represents the OLE characteristics (other than linking) for an OLE object, ActiveX<br>control, or field.      |
| Word.Interfaces.<br>OleFormatUpdate<br>Data               | An interface for updating data on the OleFormat object, for use in oleFormat.set({<br>}) .                    |
| Word.Interfaces.<br>PageCollection<br>Data                | An interface describing the data returned by calling pageCollection.toJSON() .                                |
| Word.Interfaces.<br>PageCollection<br>LoadOptions         | Represents the collection of page.                                                                            |
| Word.Interfaces.<br>PageCollection<br>UpdateData          | An interface for updating data on the PageCollection object, for use in<br>pageCollection.set({  }) .         |
| Word.Interfaces.<br>PageData                              | An interface describing the data returned by calling page.toJSON() .                                          |
| Word.Interfaces.                                          | Represents a page in the document. Page objects manage the page layout and                                    |
| PageLoadOptions                                           | content.                                                                                                      |
| Word.Interfaces.                                          | An interface describing the data returned by calling pageSetup.toJSON()                                       |


| PageSetupData                                              |                                                                                                                 |
|------------------------------------------------------------|-----------------------------------------------------------------------------------------------------------------|
| Word.Interfaces.<br>PageSetupLoad<br>Options               | Represents the page setup settings for a Word document or section.                                              |
| Word.Interfaces.<br>PageSetupUpdate<br>Data                | An interface for updating data on the PageSetup object, for use in pageSetup.set({<br>}) .                      |
| Word.Interfaces.<br>PaneCollection<br>Data                 | An interface describing the data returned by calling paneCollection.toJSON()                                    |
| Word.Interfaces.<br>PaneCollection<br>UpdateData           | An interface for updating data on the PaneCollection object, for use in<br>paneCollection.set({  }) .           |
| Word.Interfaces.                                           | An interface describing the data returned by calling pane.toJSON()                                              |
| PaneData                                                   |                                                                                                                 |
| Word.Interfaces.<br>Paragraph<br>CollectionData            | An interface describing the data returned by calling paragraphCollection.toJSON()                               |
| Word.Interfaces.<br>Paragraph<br>CollectionLoad<br>Options | Contains a collection of Word.Paragraph objects.                                                                |
| Word.Interfaces.<br>Paragraph<br>CollectionUpdate<br>Data  | An interface for updating data on the ParagraphCollection object, for use in<br>paragraphCollection.set({  }) . |
| Word.Interfaces.                                           | An interface describing the data returned by calling paragraph.toJSON()                                         |
| ParagraphData                                              |                                                                                                                 |
| Word.Interfaces.<br>ParagraphFormat<br>Data                | An interface describing the data returned by calling paragraphFormat.toJSON()                                   |
| Word.Interfaces.<br>ParagraphFormat<br>LoadOptions         | Represents a style of paragraph in a document.                                                                  |
| Word.Interfaces.<br>ParagraphFormat<br>UpdateData          | An interface for updating data on the ParagraphFormat object, for use in<br>paragraphFormat.set({  }) .         |
| Word.Interfaces.                                           | Represents a single paragraph in a selection, range, content control, or document                               |
| ParagraphLoad                                              | body.                                                                                                           |


| Options                                                      |                                                                                                                     |
|--------------------------------------------------------------|---------------------------------------------------------------------------------------------------------------------|
| Word.Interfaces.<br>ParagraphUpdate<br>Data                  | An interface for updating data on the Paragraph object, for use in paragraph.set({<br>}) .                          |
| Word.Interfaces.<br>PictureContent<br>ControlData            | An interface describing the data returned by calling pictureContentControl.toJSON()                                 |
| Word.Interfaces.<br>PictureContent<br>ControlLoad<br>Options | Represents the PictureContentControl object.                                                                        |
| Word.Interfaces.<br>PictureContent<br>ControlUpdate<br>Data  | An interface for updating data on the PictureContentControl object, for use in<br>pictureContentControl.set({  }) . |
| Word.Interfaces.<br>RangeCollection<br>Data                  | An interface describing the data returned by calling rangeCollection.toJSON()                                       |
| Word.Interfaces.<br>RangeCollection<br>LoadOptions           | Contains a collection of Word.Range objects.                                                                        |
| Word.Interfaces.<br>RangeCollection<br>UpdateData            | An interface for updating data on the RangeCollection object, for use in<br>rangeCollection.set({  }) .             |
| Word.Interfaces.                                             | An interface describing the data returned by calling range.toJSON()                                                 |
| RangeData                                                    |                                                                                                                     |
| Word.Interfaces.<br>RangeLoad<br>Options                     | Represents a contiguous area in a document.                                                                         |
| Word.Interfaces.                                             | An interface for updating data on the Range object, for use in range.set({                                          |
| RangeUpdateData                                              | }) .                                                                                                                |
| Word.Interfaces.<br>ReflectionFormat<br>Data                 | An interface describing the data returned by calling reflectionFormat.toJSON()                                      |
| Word.Interfaces.<br>ReflectionFormat<br>LoadOptions          | Represents the reflection formatting for a shape in Word.                                                           |


| Word.Interfaces.<br>ReflectionFormat<br>UpdateData                    | An interface for updating data on the ReflectionFormat object, for use in<br>reflectionFormat.set({  }) .                                                                                                                         |
|-----------------------------------------------------------------------|-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| Word.Interfaces.<br>RepeatingSection<br>ContentControl<br>Data        | An interface describing the data returned by calling<br>repeatingSectionContentControl.toJSON() .                                                                                                                                 |
| Word.Interfaces.<br>RepeatingSection<br>ContentControl<br>LoadOptions | Represents the RepeatingSectionContentControl object.                                                                                                                                                                             |
| Word.Interfaces.<br>RepeatingSection<br>ContentControl<br>UpdateData  | An interface for updating data on the RepeatingSectionContentControl object, for<br>use in repeatingSectionContentControl.set({<br>}) .                                                                                           |
| Word.Interfaces.<br>RepeatingSection<br>ItemData                      | An interface describing the data returned by calling repeatingSectionItem.toJSON() .                                                                                                                                              |
| Word.Interfaces.<br>RepeatingSection<br>ItemLoadOptions               | Represents a single item in a Word.RepeatingSectionContentControl.                                                                                                                                                                |
| Word.Interfaces.<br>RepeatingSection<br>ItemUpdateData                | An interface for updating data on the RepeatingSectionItem object, for use in<br>repeatingSectionItem.set({  }) .                                                                                                                 |
| Word.Interfaces.<br>Reviewer<br>CollectionData                        | An interface describing the data returned by calling reviewerCollection.toJSON() .                                                                                                                                                |
| Word.Interfaces.<br>Reviewer<br>CollectionLoad<br>Options             | A collection of Word.Reviewer objects that represents the reviewers of one or more<br>documents. The ReviewerCollection object contains the names of all reviewers who<br>have reviewed documents opened or edited on a computer. |
| Word.Interfaces.<br>Reviewer<br>CollectionUpdate<br>Data              | An interface for updating data on the ReviewerCollection object, for use in<br>reviewerCollection.set({  }) .                                                                                                                     |
| Word.Interfaces.<br>ReviewerData                                      | An interface describing the data returned by calling reviewer.toJSON() .                                                                                                                                                          |
| Word.Interfaces.<br>ReviewerLoad<br>Options                           | Represents a single reviewer of a document in which changes have been tracked. The<br>Reviewer object is a member of the Word.ReviewerCollection object.                                                                          |


| Word.Interfaces.<br>ReviewerUpdate<br>Data           | An interface for updating data on the Reviewer object, for use in reviewer.set({<br>}) .                             |
|------------------------------------------------------|----------------------------------------------------------------------------------------------------------------------|
| Word.Interfaces.<br>RevisionsFilter<br>Data          | An interface describing the data returned by calling revisionsFilter.toJSON() .                                      |
| Word.Interfaces.<br>RevisionsFilter<br>LoadOptions   | Represents the current settings related to the display of reviewers' comments and<br>revision marks in the document. |
| Word.Interfaces.<br>RevisionsFilter<br>UpdateData    | An interface for updating data on the RevisionsFilter object, for use in<br>revisionsFilter.set({  }) .              |
| Word.Interfaces.<br>SearchOptions<br>Data            | An interface describing the data returned by calling searchOptions.toJSON() .                                        |
| Word.Interfaces.                                     | Specifies the options to be included in a search operation. To learn more about how                                  |
| SearchOptions                                        | to use search options in the Word JavaScript APIs, read Use search options to find                                   |
| LoadOptions                                          | text in your Word add-in.                                                                                            |
| Word.Interfaces.<br>SearchOptions<br>UpdateData      | An interface for updating data on the SearchOptions object, for use in<br>searchOptions.set({  }) .                  |
| Word.Interfaces.<br>SectionCollection<br>Data        | An interface describing the data returned by calling sectionCollection.toJSON() .                                    |
| Word.Interfaces.<br>SectionCollection<br>LoadOptions | Contains the collection of the document's Word.Section objects.                                                      |
| Word.Interfaces.<br>SectionCollection<br>UpdateData  | An interface for updating data on the SectionCollection object, for use in<br>sectionCollection.set({  }) .          |
| Word.Interfaces.<br>SectionData                      | An interface describing the data returned by calling section.toJSON() .                                              |
| Word.Interfaces.<br>SectionLoad<br>Options           | Represents a section in a Word document.                                                                             |
| Word.Interfaces.<br>SectionUpdate<br>Data            | An interface for updating data on the Section object, for use in section.set({<br>}) .                               |


| Word.Interfaces.<br>SettingCollection<br>Data        | An interface describing the data returned by calling settingCollection.toJSON() .                           |
|------------------------------------------------------|-------------------------------------------------------------------------------------------------------------|
| Word.Interfaces.<br>SettingCollection<br>LoadOptions | Contains the collection of Word.Setting objects.                                                            |
| Word.Interfaces.<br>SettingCollection<br>UpdateData  | An interface for updating data on the SettingCollection object, for use in<br>settingCollection.set({  }) . |
| Word.Interfaces.<br>SettingData                      | An interface describing the data returned by calling setting.toJSON() .                                     |
| Word.Interfaces.<br>SettingLoad<br>Options           | Represents a setting of the add-in.                                                                         |
| Word.Interfaces.<br>SettingUpdate<br>Data            | An interface for updating data on the Setting object, for use in setting.set({<br>}) .                      |
| Word.Interfaces.<br>ShadingData                      | An interface describing the data returned by calling shading.toJSON() .                                     |
| Word.Interfaces.<br>ShadingLoad<br>Options           | Represents the shading object.                                                                              |
| Word.Interfaces.<br>ShadingUniversal<br>Data         | An interface describing the data returned by calling shadingUniversal.toJSON() .                            |
| Word.Interfaces.<br>ShadingUniversal<br>LoadOptions  | Represents the ShadingUniversal object, which manages shading for a range,<br>paragraph, frame, or table.   |
| Word.Interfaces.<br>ShadingUniversal<br>UpdateData   | An interface for updating data on the ShadingUniversal object, for use in<br>shadingUniversal.set({  }) .   |
| Word.Interfaces.<br>ShadingUpdate<br>Data            | An interface for updating data on the Shading object, for use in shading.set({<br>}) .                      |
| Word.Interfaces.<br>ShadowFormat<br>Data             | An interface describing the data returned by calling shadowFormat.toJSON() .                                |
| Word.Interfaces.<br>ShadowFormat                     | Represents the shadow formatting for a shape or text in Word.                                               |


| LoadOptions                                        |                                                                                                                                                                   |
|----------------------------------------------------|-------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| Word.Interfaces.<br>ShadowFormat<br>UpdateData     | An interface for updating data on the ShadowFormat object, for use in<br>shadowFormat.set({  }) .                                                                 |
| Word.Interfaces.<br>ShapeCollection<br>Data        | An interface describing the data returned by calling shapeCollection.toJSON()                                                                                     |
| Word.Interfaces.<br>ShapeCollection<br>LoadOptions | Contains a collection of Word.Shape objects. Currently, only the following shapes are<br>supported: text boxes, geometric shapes, groups, pictures, and canvases. |
| Word.Interfaces.<br>ShapeCollection<br>UpdateData  | An interface for updating data on the ShapeCollection object, for use in<br>shapeCollection.set({  }) .                                                           |
| Word.Interfaces.                                   | An interface describing the data returned by calling shape.toJSON()                                                                                               |
| ShapeData                                          |                                                                                                                                                                   |
| Word.Interfaces.                                   | An interface describing the data returned by calling shapeFill.toJSON()                                                                                           |
| ShapeFillData                                      |                                                                                                                                                                   |
| Word.Interfaces.<br>ShapeFillLoad<br>Options       | Represents the fill formatting of a shape object.                                                                                                                 |
| Word.Interfaces.<br>ShapeFillUpdate<br>Data        | An interface for updating data on the ShapeFill object, for use in shapeFill.set({<br>}) .                                                                        |
| Word.Interfaces.                                   | An interface describing the data returned by calling shapeGroup.toJSON()                                                                                          |
| ShapeGroupData                                     |                                                                                                                                                                   |
| Word.Interfaces.<br>ShapeGroupLoad<br>Options      | Represents a shape group in the document. To get the corresponding Shape object,<br>use ShapeGroup.shape.                                                         |
| Word.Interfaces.<br>ShapeGroup<br>UpdateData       | An interface for updating data on the ShapeGroup object, for use in shapeGroup.set({<br>}) .                                                                      |
| Word.Interfaces.                                   | Represents a shape in the header, footer, or document body. Currently, only the                                                                                   |
| ShapeLoad                                          | following shapes are supported: text boxes, geometric shapes, groups, pictures, and                                                                               |
| Options                                            | canvases.                                                                                                                                                         |
| Word.Interfaces.<br>ShapeTextWrap<br>Data          | An interface describing the data returned by calling shapeTextWrap.toJSON()                                                                                       |


| Word.Interfaces.<br>ShapeTextWrap<br>LoadOptions    | Represents all the properties for wrapping text around a shape.                                           |
|-----------------------------------------------------|-----------------------------------------------------------------------------------------------------------|
| Word.Interfaces.<br>ShapeTextWrap<br>UpdateData     | An interface for updating data on the ShapeTextWrap object, for use in<br>shapeTextWrap.set({  }) .       |
| Word.Interfaces.<br>ShapeUpdateData                 | An interface for updating data on the Shape object, for use in shape.set({  }) .                          |
| Word.Interfaces.<br>SourceCollection<br>Data        | An interface describing the data returned by calling sourceCollection.toJSON() .                          |
| Word.Interfaces.<br>SourceCollection<br>LoadOptions | Represents a collection of Word.Source objects.                                                           |
| Word.Interfaces.<br>SourceCollection<br>UpdateData  | An interface for updating data on the SourceCollection object, for use in<br>sourceCollection.set({  }) . |
| Word.Interfaces.<br>SourceData                      | An interface describing the data returned by calling source.toJSON() .                                    |
| Word.Interfaces.<br>SourceLoad<br>Options           | Represents an individual source, such as a book, journal article, or interview.                           |
| Word.Interfaces.<br>StyleCollection<br>Data         | An interface describing the data returned by calling styleCollection.toJSON() .                           |
| Word.Interfaces.<br>StyleCollection<br>LoadOptions  | Contains a collection of Word.Style objects.                                                              |
| Word.Interfaces.<br>StyleCollection<br>UpdateData   | An interface for updating data on the StyleCollection object, for use in<br>styleCollection.set({  }) .   |
| Word.Interfaces.<br>StyleData                       | An interface describing the data returned by calling style.toJSON() .                                     |
| Word.Interfaces.<br>StyleLoadOptions                | Represents a style in a Word document.                                                                    |
| Word.Interfaces.<br>StyleUpdateData                 | An interface for updating data on the Style object, for use in style.set({  }) .                          |


| Word.Interfaces.<br>TableBorderData                        | An interface describing the data returned by calling tableBorder.toJSON() .                                     |
|------------------------------------------------------------|-----------------------------------------------------------------------------------------------------------------|
| Word.Interfaces.<br>TableBorderLoad<br>Options             | Specifies the border style.                                                                                     |
| Word.Interfaces.<br>TableBorder<br>UpdateData              | An interface for updating data on the TableBorder object, for use in<br>tableBorder.set({  }) .                 |
| Word.Interfaces.<br>TableCell<br>CollectionData            | An interface describing the data returned by calling tableCellCollection.toJSON() .                             |
| Word.Interfaces.<br>TableCell<br>CollectionLoad<br>Options | Contains the collection of the document's TableCell objects.                                                    |
| Word.Interfaces.<br>TableCell<br>CollectionUpdate<br>Data  | An interface for updating data on the TableCellCollection object, for use in<br>tableCellCollection.set({  }) . |
| Word.Interfaces.<br>TableCellData                          | An interface describing the data returned by calling tableCell.toJSON() .                                       |
| Word.Interfaces.<br>TableCellLoad<br>Options               | Represents a table cell in a Word document.                                                                     |
| Word.Interfaces.<br>TableCellUpdate<br>Data                | An interface for updating data on the TableCell object, for use in tableCell.set({<br>}) .                      |
| Word.Interfaces.<br>TableCollection<br>Data                | An interface describing the data returned by calling tableCollection.toJSON() .                                 |
| Word.Interfaces.<br>TableCollection<br>LoadOptions         | Contains the collection of the document's Table objects.                                                        |
| Word.Interfaces.<br>TableCollection<br>UpdateData          | An interface for updating data on the TableCollection object, for use in<br>tableCollection.set({  }) .         |
| Word.Interfaces.<br>TableColumn<br>CollectionData          | An interface describing the data returned by calling tableColumnCollection.toJSON() .                           |


| Word.Interfaces.<br>TableColumn<br>CollectionLoad<br>Options | Represents a collection of Word.TableColumn objects in a Word document.                                             |
|--------------------------------------------------------------|---------------------------------------------------------------------------------------------------------------------|
| Word.Interfaces.<br>TableColumn<br>CollectionUpdate<br>Data  | An interface for updating data on the TableColumnCollection object, for use in<br>tableColumnCollection.set({  }) . |
| Word.Interfaces.<br>TableColumnData                          | An interface describing the data returned by calling tableColumn.toJSON() .                                         |
| Word.Interfaces.<br>TableColumnLoad<br>Options               | Represents a table column in a Word document.                                                                       |
| Word.Interfaces.<br>TableColumn<br>UpdateData                | An interface for updating data on the TableColumn object, for use in<br>tableColumn.set({  }) .                     |
| Word.Interfaces.<br>TableData                                | An interface describing the data returned by calling table.toJSON() .                                               |
| Word.Interfaces.<br>TableLoadOptions                         | Represents a table in a Word document.                                                                              |
| Word.Interfaces.<br>TableRow<br>CollectionData               | An interface describing the data returned by calling tableRowCollection.toJSON() .                                  |
| Word.Interfaces.<br>TableRow<br>CollectionLoad<br>Options    | Contains the collection of the document's TableRow objects.                                                         |
| Word.Interfaces.<br>TableRow<br>CollectionUpdate<br>Data     | An interface for updating data on the TableRowCollection object, for use in<br>tableRowCollection.set({  }) .       |
| Word.Interfaces.<br>TableRowData                             | An interface describing the data returned by calling tableRow.toJSON() .                                            |
| Word.Interfaces.<br>TableRowLoad<br>Options                  | Represents a row in a Word document.                                                                                |
| Word.Interfaces.<br>TableRowUpdate<br>Data                   | An interface for updating data on the TableRow object, for use in tableRow.set({<br>}) .                            |


| Word.Interfaces.<br>TableStyleData                        | An interface describing the data returned by calling tableStyle.toJSON() .                                                                                                                                                                                                                                                                                                       |
|-----------------------------------------------------------|----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| Word.Interfaces.<br>TableStyleLoad<br>Options             | Represents the TableStyle object.                                                                                                                                                                                                                                                                                                                                                |
| Word.Interfaces.<br>TableStyleUpdate<br>Data              | An interface for updating data on the TableStyle object, for use in tableStyle.set({<br>}) .                                                                                                                                                                                                                                                                                     |
| Word.Interfaces.<br>TableUpdateData                       | An interface for updating data on the Table object, for use in table.set({  }) .                                                                                                                                                                                                                                                                                                 |
| Word.Interfaces.<br>TabStopCollection<br>Data             | An interface describing the data returned by calling tabStopCollection.toJSON() .                                                                                                                                                                                                                                                                                                |
| Word.Interfaces.<br>TabStopCollection<br>LoadOptions      | Represents a collection of tab stops in a Word document.                                                                                                                                                                                                                                                                                                                         |
| Word.Interfaces.<br>TabStopCollection<br>UpdateData       | An interface for updating data on the TabStopCollection object, for use in<br>tabStopCollection.set({  }) .                                                                                                                                                                                                                                                                      |
| Word.Interfaces.<br>TabStopData                           | An interface describing the data returned by calling tabStop.toJSON() .                                                                                                                                                                                                                                                                                                          |
| Word.Interfaces.<br>TabStopLoad<br>Options                | Represents a tab stop in a Word document.                                                                                                                                                                                                                                                                                                                                        |
| Word.Interfaces.<br>Template<br>CollectionData            | An interface describing the data returned by calling templateCollection.toJSON() .                                                                                                                                                                                                                                                                                               |
| Word.Interfaces.<br>Template<br>CollectionLoad<br>Options | Contains a collection of Word.Template objects that represent all the templates that<br>are currently available. This collection includes open templates, templates attached<br>to open documents, and global templates loaded in the Templates and Add-ins<br>dialog box. To learn how to access this dialog in the Word UI, see Load or unload a<br>template or add-in program |
| Word.Interfaces.<br>Template<br>CollectionUpdate<br>Data  | An interface for updating data on the TemplateCollection object, for use in<br>templateCollection.set({  }) .                                                                                                                                                                                                                                                                    |
| Word.Interfaces.<br>TemplateData                          | An interface describing the data returned by calling template.toJSON() .                                                                                                                                                                                                                                                                                                         |


| Word.Interfaces.<br>TemplateLoad<br>Options                 | Represents a document template.                                                                                                 |
|-------------------------------------------------------------|---------------------------------------------------------------------------------------------------------------------------------|
| Word.Interfaces.<br>TemplateUpdate<br>Data                  | An interface for updating data on the Template object, for use in template.set({<br>}) .                                        |
| Word.Interfaces.<br>TextColumn<br>CollectionData            | An interface describing the data returned by calling textColumnCollection.toJSON() .                                            |
| Word.Interfaces.<br>TextColumn<br>CollectionLoad<br>Options | A collection of Word.TextColumn objects that represent all the columns of text in the<br>document or a section of the document. |
| Word.Interfaces.<br>TextColumn<br>CollectionUpdate<br>Data  | An interface for updating data on the TextColumnCollection object, for use in<br>textColumnCollection.set({  }) .               |
| Word.Interfaces.<br>TextColumnData                          | An interface describing the data returned by calling textColumn.toJSON() .                                                      |
| Word.Interfaces.<br>TextColumnLoad<br>Options               | Represents a single text column in a section.                                                                                   |
| Word.Interfaces.<br>TextColumn<br>UpdateData                | An interface for updating data on the TextColumn object, for use in textColumn.set({<br>}) .                                    |
| Word.Interfaces.<br>TextFrameData                           | An interface describing the data returned by calling textFrame.toJSON() .                                                       |
| Word.Interfaces.<br>TextFrameLoad<br>Options                | Represents the text frame of a shape object.                                                                                    |
| Word.Interfaces.<br>TextFrameUpdate<br>Data                 | An interface for updating data on the TextFrame object, for use in textFrame.set({<br>}) .                                      |
| Word.Interfaces.<br>ThreeDimensional<br>FormatData          | An interface describing the data returned by calling<br>threeDimensionalFormat.toJSON() .                                       |
| Word.Interfaces.<br>ThreeDimensional                        | Represents a shape's three-dimensional formatting.                                                                              |


| FormatLoad<br>Options                                          |                                                                                                                         |
|----------------------------------------------------------------|-------------------------------------------------------------------------------------------------------------------------|
| Word.Interfaces.<br>ThreeDimensional<br>FormatUpdate<br>Data   | An interface for updating data on the ThreeDimensionalFormat object, for use in<br>threeDimensionalFormat.set({  }) .   |
| Word.Interfaces.<br>TrackedChange<br>CollectionData            | An interface describing the data returned by calling<br>trackedChangeCollection.toJSON() .                              |
| Word.Interfaces.<br>TrackedChange<br>CollectionLoad<br>Options | Contains a collection of Word.TrackedChange objects.                                                                    |
| Word.Interfaces.<br>TrackedChange<br>CollectionUpdate<br>Data  | An interface for updating data on the TrackedChangeCollection object, for use in<br>trackedChangeCollection.set({  }) . |
| Word.Interfaces.<br>TrackedChange<br>Data                      | An interface describing the data returned by calling trackedChange.toJSON()                                             |
| Word.Interfaces.<br>TrackedChange<br>LoadOptions               | Represents a tracked change in a Word document.                                                                         |
| Word.Interfaces.                                               | An interface describing the data returned by calling view.toJSON()                                                      |
| ViewData                                                       |                                                                                                                         |
| Word.Interfaces.                                               | Contains the view attributes (such as show all, field shading, and table gridlines) for a                               |
| ViewLoadOptions                                                | window or pane.                                                                                                         |
| Word.Interfaces.                                               | An interface for updating data on the View object, for use in view.set({                                                |
| ViewUpdateData                                                 | }) .                                                                                                                    |
| Word.Interfaces.<br>WindowCollection<br>Data                   | An interface describing the data returned by calling windowCollection.toJSON()                                          |
| Word.Interfaces.<br>WindowCollection<br>LoadOptions            | Represents the collection of window objects.                                                                            |
| Word.Interfaces.<br>WindowCollection<br>UpdateData             | An interface for updating data on the WindowCollection object, for use in<br>windowCollection.set({  }) .               |


| Word.Interfaces.<br>WindowData                   | An interface describing the data returned by calling window.toJSON() .                                                                                                                                                                             |
|--------------------------------------------------|----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| Word.Interfaces.<br>WindowLoad<br>Options        | Represents the window that displays the document. A window can be split to contain<br>multiple reading panes.                                                                                                                                      |
| Word.Interfaces.<br>WindowUpdate<br>Data         | An interface for updating data on the Window object, for use in window.set({  }) .                                                                                                                                                                 |
| Word.Interfaces.<br>XmlMappingData               | An interface describing the data returned by calling xmlMapping.toJSON() .                                                                                                                                                                         |
| Word.Interfaces.<br>XmlMappingLoad<br>Options    | Represents the XML mapping on a Word.ContentControl object between custom<br>XML and that content control. An XML mapping is a link between the text in a<br>content control and an XML element in the custom XML data store for this<br>document. |
| Word.Interfaces.<br>XmlMapping<br>UpdateData     | An interface for updating data on the XmlMapping object, for use in xmlMapping.set({<br>}) .                                                                                                                                                       |
| Word.ListFormat<br>CountNumbered<br>ItemsOptions | Represents options for counting numbered items in a range.                                                                                                                                                                                         |
| Word.List<br>TemplateApply<br>Options            | Represents options for applying a list template to a range.                                                                                                                                                                                        |
| Word.Paragraph<br>AddedEventArgs                 | Provides information about the paragraphs that raised the paragraphAdded event.                                                                                                                                                                    |
| Word.Paragraph<br>ChangedEvent<br>Args           | Provides information about the paragraphs that raised the paragraphChanged event.                                                                                                                                                                  |
| Word.Paragraph<br>DeletedEventArgs               | Provides information about the paragraphs that raised the paragraphDeleted event.                                                                                                                                                                  |
| Word.TabStopAdd<br>Options                       | Specifies the options for adding to a Word.TabStopCollection object.                                                                                                                                                                               |
| Word.TextColumn<br>AddOptions                    | Represents options for a new text column in a document or section of a document.                                                                                                                                                                   |
| Word.Window                                      | The options that define whether to save changes before closing and whether to                                                                                                                                                                      |
| CloseOptions                                     | route the document.                                                                                                                                                                                                                                |
| Word.Window                                      | The options for scrolling through the specified pane or window page by page.                                                                                                                                                                       |


| PageScrollOptions             |                                                                                    |
|-------------------------------|------------------------------------------------------------------------------------|
| Word.Window                   | The options that scrolls a window or pane by the specified number of units defined |
| ScrollOptions                 | by the calling method.                                                             |
| Word.XmlSet<br>MappingOptions | The options that define the prefix mapping and the source of the custom XML data.  |

### **Enums**

#### ノ **Expand table**

| Word.Alignment             |                                                                   |
|----------------------------|-------------------------------------------------------------------|
| Word.Annotation<br>State   | Represents the state of the annotation.                           |
| Word.Arrowhead<br>Length   | Specifies the length of the arrowhead at the end of a line.       |
| Word.Arrowhead<br>Style    | Specifies the style of the arrowhead at the end of a line.        |
| Word.Arrowhead<br>Width    | Specifies the width of the arrowhead at the end of a line.        |
| Word.Baseline<br>Alignment | Represents the type of baseline alignment.                        |
| Word.BevelType             | Indicates the bevel type of a Word.ThreeDimensionalFormat object. |
| Word.BodyType              | Represents the types of body objects.                             |
| Word.BorderLine<br>Style   | Specifies the border style for an object.                         |
| Word.BorderLocation        |                                                                   |
| Word.BorderType            |                                                                   |
| Word.BorderWidth           | Represents the width of a style's border.                         |
| Word.BreakType             | Specifies the form of a break.                                    |
| Word.BuildingBlock<br>Type | Specifies the type of building block.                             |
| Word.BuiltInStyle<br>Name  | Represents the built-in style in a Word document.                 |


| Word.CalendarType                            | Calendar types.                                                                |
|----------------------------------------------|--------------------------------------------------------------------------------|
| Word.CellPaddingLocation                     |                                                                                |
| Word.Change<br>TrackingMode                  | Represents the possible change tracking modes.                                 |
| Word.Change<br>TrackingState                 | Specify the track state when ChangeTracking is on.                             |
| Word.Change<br>TrackingVersion               | Specify the current version or the original version of the text.               |
| Word.CharacterCase                           | Specifies the case of the text in the specified range.                         |
| Word.Character<br>Width                      | Specifies the character width of the text in the specified range.              |
| Word.CloseBehavior                           | Specifies the close behavior for Document.close .                              |
| Word.ColorIndex                              | Represents color index values in a Word document.                              |
| Word.ColorType                               | Specifies the color type.                                                      |
| Word.ColumnWidth                             | Specifies the column width options in a Word document.                         |
| Word.Comment<br>ChangeType                   | Represents how the comments in the event were changed.                         |
| Word.Compare<br>Target                       | Specifies the target document for displaying document comparison differences.  |
| Word.Content<br>ControlAppearance            | ContentControl appearance.                                                     |
| Word.Content<br>ControlDateStorage<br>Format | Date storage formats for Word.DatePickerContentControl.                        |
| Word.Content<br>ControlLevel                 | Content control level types.                                                   |
| Word.Content<br>ControlState                 | Represents the state of the content control.                                   |
| Word.Content<br>ControlType                  | Specifies supported content control types and subtypes.                        |
| Word.Continue                                | Specifies whether the formatting from the previous list can be continued.      |
| Word.CritiqueColor                           | Represents the color scheme of a critique in the document, affecting underline |
| Scheme                                       | and highlight.                                                                 |


| Word.CustomXml<br>NodeType            | Represents the type of a Word.CustomXmlNode.                                                                                                                                                                                                    |
|---------------------------------------|-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| Word.CustomXml<br>ValidationErrorType | Represents the type of a Word.CustomXmlValidationError.                                                                                                                                                                                         |
| Word.DefaultList<br>Behavior          | Specifies the default list behavior for a list.                                                                                                                                                                                                 |
| Word.DocPartInsert<br>Type            | Specifies how a building block is inserted into a document.                                                                                                                                                                                     |
| Word.DocumentPropertyType             |                                                                                                                                                                                                                                                 |
| Word.DropPosition                     | Represents the position of a dropped capital letter.                                                                                                                                                                                            |
| Word.EmphasisMark                     | Specifies the type of emphasis mark to use for a character or designated character<br>string.                                                                                                                                                   |
| Word.ErrorCodes                       |                                                                                                                                                                                                                                                 |
| Word.EventSource                      | An enum that specifies an event's source. It can be local or remote (through<br>coauthoring).                                                                                                                                                   |
| Word.EventType                        | Provides information about the type of a raised event.                                                                                                                                                                                          |
| Word.ExtrusionColor<br>Type           | Specifies whether the extrusion color is based on the extruded shape's fill (the<br>front face of the extrusion) and automatically changes when the shape's fill<br>changes, or whether the extrusion color is independent of the shape's fill. |
| Word.FarEastLine                      | Represents the East Asian language to use when breaking lines of text in the                                                                                                                                                                    |
| BreakLanguageId                       | specified document or template.                                                                                                                                                                                                                 |
| Word.FarEastLine                      | Represents the level of line breaking to use for East Asian languages in the                                                                                                                                                                    |
| BreakLevel                            | specified document or template.                                                                                                                                                                                                                 |
| Word.FieldKind                        | Represents the kind of field. Indicates how the field works in relation to updating.                                                                                                                                                            |
| Word.FieldShading                     | Specifies the field shading options in a Word document.                                                                                                                                                                                         |
| Word.FieldType                        | Represents the type of Field.                                                                                                                                                                                                                   |
| Word.FillType                         | Specifies a shape's fill type.                                                                                                                                                                                                                  |
| Word.FlowDirection                    | Specifies the direction in which text flows from one text column to the next.                                                                                                                                                                   |
| Word.FrameSizeRule                    | Represents how Word interprets the rule used to determine the height or width of<br>a Word.Frame.                                                                                                                                               |
| Word.Geometric<br>ShapeType           | Specifies the shape type for a GeometricShape object.                                                                                                                                                                                           |


| Word.GradientColor<br>Type                  | Specifies the type of gradient used in a shape's fill.                                                                                 |
|---------------------------------------------|----------------------------------------------------------------------------------------------------------------------------------------|
| Word.GradientStyle                          | Specifies the style for a gradient fill.                                                                                               |
| Word.GutterPosition                         | Specifies where the gutter appears in the document.                                                                                    |
| Word.GutterStyle                            | Specifies whether the gutter style should conform to left-to-right text flow or<br>right-to-left text flow.                            |
| Word.HeaderFooterType                       |                                                                                                                                        |
| Word.Heading<br>Separator                   | Specifies the type of separator to use for headings.                                                                                   |
| Word.Horizontal<br>InVerticalType           | Specifies the format for horizontal text set within vertical text.                                                                     |
| Word.HyperlinkType                          | Specifies the hyperlink type.                                                                                                          |
| Word.ImageFormat                            |                                                                                                                                        |
| Word.ImeMode                                | Specifies the IME (Input Method Editor) mode.                                                                                          |
| Word.Imported<br>StylesConflict<br>Behavior | Specifies how to handle any conflicts, that is, when imported styles have the same<br>name as existing styles in the current document. |
| Word.IndexFilter                            | Specifies the filter type for an index.                                                                                                |
| Word.IndexFormat                            | Specifies the format for an index.                                                                                                     |
| Word.IndexSortBy                            | Specifies how an index is sorted.                                                                                                      |
| Word.IndexType                              | Specifies the type of index to create.                                                                                                 |
| Word.InsertLocation                         | The insertion location types.                                                                                                          |
| Word.Justification<br>Mode                  | Specifies the character spacing adjustment for a document.                                                                             |
| Word.Kana                                   | Specifies the Kana type.                                                                                                               |
| Word.LanguageId                             | Represents the language ID of a Word document.                                                                                         |
| Word.LayoutMode                             | Specifies how text is laid out in the layout mode for the current document.                                                            |
| Word.Ligature                               | Specifies the type of ligature applied to a font.                                                                                      |
| Word.LightRigType                           | Indicates the effects lighting for an object.                                                                                          |
| Word.LineDashStyle                          | Specifies the dash style for a line.                                                                                                   |


| Word.LineFormat<br>Style            | Specifies the style for a line.                                                   |  |
|-------------------------------------|-----------------------------------------------------------------------------------|--|
| Word.LineSpacing                    | Represents the type of line spacing.                                              |  |
| Word.LineWidth                      | Specifies the width of an object's border.                                        |  |
| Word.LinkType                       | Specifies the type of link.                                                       |  |
| Word.ListApplyTo                    | Specifies the portion of a list to which to apply a list template.                |  |
| Word.ListBuiltInNumberStyle         |                                                                                   |  |
| Word.ListBullet                     |                                                                                   |  |
| Word.ListLevelType                  |                                                                                   |  |
| Word.ListNumbering                  |                                                                                   |  |
| Word.ListType                       | Represents the list type.                                                         |  |
| Word.LocationRelation               |                                                                                   |  |
| Word.NoteItemType                   | Note item type                                                                    |  |
| Word.NumberForm                     | Specifies the number form setting for an OpenType font.                           |  |
| Word.Numbering<br>Rule              | Specifies the numbering rule to apply.                                            |  |
| Word.Number<br>Spacing              | Specifies the number spacing setting for an OpenType font.                        |  |
| Word.NumberType                     | Specifies the type of numbers in a list.                                          |  |
| Word.OleVerb                        | Specifies the action associated with the verb that the OLE object should perform. |  |
| Word.OutlineLevel                   | Represents the outline levels.                                                    |  |
| Word.PageBorderArt                  | Specifies the graphical page border setting of a page.                            |  |
| Word.PageColor                      | Specifies the page color options in a Word document.                              |  |
| Word.Page<br>MovementType           | Specifies the page movement type in a Word document.                              |  |
| Word.Page<br>Orientation            | Specifies a page layout orientation.                                              |  |
| Word.PageSetup<br>VerticalAlignment | Specifies the type of vertical alignment to apply.                                |  |
| Word.PaperSize                      | Specifies a paper size.                                                           |  |


| Word.PatternType                      | Specifies the fill pattern used in a shape.                                                                                                                    |
|---------------------------------------|----------------------------------------------------------------------------------------------------------------------------------------------------------------|
| Word.Preferred                        | Specifies the preferred unit of measure to use when measuring the width of an                                                                                  |
| WidthType                             | item.                                                                                                                                                          |
| Word.PresetCamera                     | Indicates the effects camera type used by the specified object.                                                                                                |
| Word.Preset                           | Specifies the direction that the extrusion's sweep path takes away from the                                                                                    |
| ExtrusionDirection                    | extruded shape (the front face of the extrusion).                                                                                                              |
| Word.Preset<br>GradientType           | Specifies which predefined gradient to use to fill a shape.                                                                                                    |
| Word.PresetLighting                   | Specifies the location of lighting on an extruded (three-dimensional) shape                                                                                    |
| Direction                             | relative to the shape.                                                                                                                                         |
| Word.PresetLighting<br>Softness       | Specifies the intensity of light used on a shape.                                                                                                              |
| Word.PresetMaterial                   | Specifies the extrusion surface material.                                                                                                                      |
| Word.PresetTexture                    | Specifies texture to be used to fill a shape.                                                                                                                  |
| Word.PresetThree<br>DimensionalFormat | Specifies an extrusion (three-dimensional) format.                                                                                                             |
| Word.Range                            | Represents the location of a range. You can get range by calling getRange on                                                                                   |
| Location                              | different objects such as Word.Paragraph and Word.ContentControl.                                                                                              |
| Word.Reading<br>LayoutMargin          | Specifies the margin options in reading layout view in a Word document.                                                                                        |
| Word.ReadingOrder                     | Represents the reading order of text.                                                                                                                          |
| Word.ReflectionType                   | Specifies the type of the Word.ReflectionFormat object.                                                                                                        |
| Word.Relative<br>HorizontalPosition   | Represents what the horizontal position of a shape is relative to. For more<br>information about margins, see Change the margins in your Word document         |
| Word.RelativeSize                     | Represents what the horizontal or vertical size of a shape is relative to. For more<br>information about margins, see Change the margins in your Word document |
| Word.Relative<br>VerticalPosition     | Represents what the vertical position of a shape is relative to. For more<br>information about margins, see Change the margins in your Word document           |
| Word.Revisions<br>BalloonMargin       | Specifies the margin for revision balloons in a Word document.                                                                                                 |
| Word.Revisions<br>BalloonWidthType    | Specifies the width type for revision balloons in a Word document.                                                                                             |
| Word.Revisions                        | Specifies the extent of markup visible in the document.                                                                                                        |


| Markup                              |                                                                                                                                                        |
|-------------------------------------|--------------------------------------------------------------------------------------------------------------------------------------------------------|
| Word.Revisions<br>Mode              | Specifies the display mode for tracked changes in a Word document.                                                                                     |
| Word.RevisionsView                  | Specifies whether Word displays the original version of a document or a version<br>with revisions and formatting changes applied.                      |
| Word.RevisionType                   | Specifies the revision type.                                                                                                                           |
| Word.RulerStyle                     | Specifies the way Word adjusts the table when the left indent is changed.                                                                              |
| Word.SaveBehavior                   | Specifies the save behavior for Document.save                                                                                                          |
| Word.Save<br>Configuration          | Specifies the save options.                                                                                                                            |
| Word.Section                        | Specifies how Word displays the reading order and alignment for the specified                                                                          |
| Direction                           | sections.                                                                                                                                              |
| Word.SectionStart                   | Specifies the type of section break for the specified item.                                                                                            |
| Word.SeekView                       | Specifies the seek view options in a Word document.                                                                                                    |
| Word.Selection                      | This enum sets where the cursor (insertion point) in the document is after a                                                                           |
| Mode                                | selection.                                                                                                                                             |
| Word.Shading<br>TextureType         | Represents the shading texture. To learn more about how to apply backgrounds<br>like textures, see Add, change, or delete the background color in Word |
| Word.ShadowStyle                    | Specifies the type of shadowing effect.                                                                                                                |
| Word.ShadowType                     | Specifies the type of shadow displayed with a shape.                                                                                                   |
| Word.ShapeAuto<br>Size              | Determines the type of automatic sizing allowed.                                                                                                       |
| Word.ShapeFillType                  | Specifies a shape's fill type.                                                                                                                         |
| Word.ShapeScale<br>From             | Specifies which part of the shape retains its position when the shape is scaled.                                                                       |
| Word.ShapeScale<br>Type             | Specifies the scale size type of a shape.                                                                                                              |
| Word.ShapeText<br>Orientation       | Specifies the orientation for the text frame in a shape.                                                                                               |
| Word.ShapeText<br>VerticalAlignment | Specifies the vertical alignment for the text frame in a shape.                                                                                        |
| Word.ShapeText                      | Specifies whether the document text should wrap on both sides of the specified                                                                         |
| WrapSide                            | shape, on either the left or right side only, or on the side of the shape that's                                                                       |


|                              | farther from the respective page margin.                                                                               |  |  |  |  |  |
|------------------------------|------------------------------------------------------------------------------------------------------------------------|--|--|--|--|--|
| Word.ShapeText<br>WrapType   | Specifies how to wrap document text around a shape. For more details, see the<br>"Text Wrapping" tab of Layout options |  |  |  |  |  |
| Word.ShapeType               | Represents the shape type.                                                                                             |  |  |  |  |  |
| Word.ShowSource<br>Documents | Specifies the source documents to show.                                                                                |  |  |  |  |  |
| Word.SpecialPane             | Specifies the special pane options in a Word document.                                                                 |  |  |  |  |  |
| Word.StoryType               | Specifies the type of story in a Word document.                                                                        |  |  |  |  |  |
| Word.StyleType               | Represents the type of style.                                                                                          |  |  |  |  |  |
| Word.StylisticSet            | Specifies the stylistic set to apply to the font.                                                                      |  |  |  |  |  |
| Word.TabAlignment            | Represents the alignment of a tab stop.                                                                                |  |  |  |  |  |
| Word.TabLeader               | Specifies the tab leader style.                                                                                        |  |  |  |  |  |
| Word.TemplateType            | Specifies the type of template.                                                                                        |  |  |  |  |  |
| Word.TextboxTight<br>Wrap    | Represents the type of tight wrap for a text box.                                                                      |  |  |  |  |  |
|                              |                                                                                                                        |  |  |  |  |  |
| Word.Texture                 | Specifies the alignment (the origin of the coordinate grid) for the tiling of the                                      |  |  |  |  |  |
| Alignment                    | texture fill.                                                                                                          |  |  |  |  |  |
| Word.TextureType             | Specifies the texture type for the selected fill.                                                                      |  |  |  |  |  |
| Word.ThemeColor<br>Index     | Specifies the theme colors for document themes.                                                                        |  |  |  |  |  |
| Word.Tracked<br>ChangeType   | TrackedChange type.                                                                                                    |  |  |  |  |  |
| Word.Trailing<br>Character   | Represents the character inserted after the list item mark.                                                            |  |  |  |  |  |
| Word.TwoLines<br>InOneType   | Specifies the two lines in one type.                                                                                   |  |  |  |  |  |
| Word.Underline               | Specifies the underline type.                                                                                          |  |  |  |  |  |
| Word.UnderlineType           | The supported styles for underline format.                                                                             |  |  |  |  |  |
| Word.VerticalAlignment       |                                                                                                                        |  |  |  |  |  |
| Word.ViewType                | Specifies the view type in a Word document.                                                                            |  |  |  |  |  |


### **Functions**

#### ノ **Expand table**

| Word.<br>run(objects,<br>batch) | Executes a batch script that performs actions on the Word object model, using the<br>RequestContext of previously created API objects.                                                                                           |
|---------------------------------|----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| Word.                           | Executes a batch script that performs actions on the Word object model, using the                                                                                                                                                |
| run(object,                     | RequestContext of a previously created API object. When the promise is resolved, any                                                                                                                                             |
| batch)                          | tracked objects that were automatically allocated during execution will be released.                                                                                                                                             |
| Word.<br>run(batch)             | Executes a batch script that performs actions on the Word object model, using a new<br>RequestContext. When the promise is resolved, any tracked objects that were<br>automatically allocated during execution will be released. |

### **Function Details**

### **Word.run(objects, batch)**

Executes a batch script that performs actions on the Word object model, using the RequestContext of previously created API objects.

TypeScript

export function run<T>(objects: OfficeExtension.ClientObject[], batch: (context: Word.RequestContext) => Promise<T>): Promise<T>;

#### **Parameters**

#### **objects** [OfficeExtension.ClientObject](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientobject?view=word-js-preview)[]

An array of previously created API objects. The array will be validated to make sure that all of the objects share the same context. The batch will use this shared RequestContext, which means that any changes applied to these objects will be picked up by context.sync() .

**batch** (context: [Word.RequestContext)](https://learn.microsoft.com/en-us/javascript/api/word/word.requestcontext?view=word-js-preview) => Promise<T>

A function that takes in a RequestContext and returns a promise (typically, just the result of context.sync() ). The context parameter facilitates requests to the Word application. Since


the Office add-in and the Word application run in two different processes, the RequestContext is required to get access to the Word object model from the add-in.

#### **Returns**

Promise<T>

### **Word.run(object, batch)**

Executes a batch script that performs actions on the Word object model, using the RequestContext of a previously created API object. When the promise is resolved, any tracked objects that were automatically allocated during execution will be released.

TypeScript

export function run<T>(object: OfficeExtension.ClientObject, batch: (context: Word.RequestContext) => Promise<T>): Promise<T>;

#### **Parameters**

#### **object** [OfficeExtension.ClientObject](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientobject?view=word-js-preview)

A previously created API object. The batch will use the same RequestContext as the passedin object, which means that any changes applied to the object will be picked up by context.sync() .

#### **batch** (context: [Word.RequestContext)](https://learn.microsoft.com/en-us/javascript/api/word/word.requestcontext?view=word-js-preview) => Promise<T>

A function that takes in a RequestContext and returns a promise (typically, just the result of context.sync() ). The context parameter facilitates requests to the Word application. Since the Office add-in and the Word application run in two different processes, the RequestContext is required to get access to the Word object model from the add-in.

#### **Returns**

Promise<T>

### **Word.run(batch)**

Executes a batch script that performs actions on the Word object model, using a new RequestContext. When the promise is resolved, any tracked objects that were automatically allocated during execution will be released.


```
TypeScript
```

```
export function run<T>(batch: (context: Word.RequestContext) => Promise<T>): 
Promise<T>;
```
#### **Parameters**

**batch** (context: [Word.RequestContext)](https://learn.microsoft.com/en-us/javascript/api/word/word.requestcontext?view=word-js-preview) => Promise<T>

A function that takes in a RequestContext and returns a promise (typically, just the result of context.sync() ). The context parameter facilitates requests to the Word application. Since the Office add-in and the Word application run in two different processes, the RequestContext is required to get access to the Word object model from the add-in.

### **Returns**

Promise<T>


# **Word JavaScript object model in Office Add-ins**

Article • 05/30/2025

This article describes concepts that are fundamental to using the Word JavaScript API to build add-ins.

# **Office.js APIs for Word**

A Word add-in interacts with objects in Word by using the Office JavaScript API. This includes two JavaScript object models:

- **Word JavaScript API**: The [Word JavaScript API](https://learn.microsoft.com/en-us/javascript/api/word) provides strongly-typed objects that work with the document, ranges, tables, lists, formatting, and more. To learn about the asynchronous nature of the Word APIs and how they work with the document, see Using the application-specific API model.
- **Common APIs**: The [Common API](https://learn.microsoft.com/en-us/javascript/api/office) give access to features such as UI, dialogs, and client settings that are common across multiple Office applications. To learn more about using the Common API, see Common JavaScript API object model.

While you'll likely use the Word JavaScript API to develop the majority of functionality in addins that target Word, you'll also use objects in the Common API. For example:

- [Office.Context](https://learn.microsoft.com/en-us/javascript/api/office/office.context): The Context object represents the runtime environment of the add-in and provides access to key objects of the API. It consists of document configuration details such as contentLanguage and officeTheme and also provides information about the addin's runtime environment such as host and platform . Additionally, it provides the requirements.isSetSupported() method, which you can use to check whether a specified requirement set is supported by the Word application where the add-in is running.
- [Office.Document:](https://learn.microsoft.com/en-us/javascript/api/office/office.document) The Office.Document object provides the getFileAsync() method, which you can use to download the Word file where the add-in is running. This is separate from the [Word.Document](https://learn.microsoft.com/en-us/javascript/api/word/word.document) object.


# **Word-specific object model**

To understand the Word APIs, you must understand how key components of a document are related to one another.

- The document contains sections, pages, and document-level entities such as settings and custom XML parts.
- A section contains a body.
- A body has paragraphs, content controls, and range objects, among others.
- A range is a contiguous area of content, including text, whitespace, tables, and images. The [Word.Range](https://learn.microsoft.com/en-us/javascript/api/word/word.range) object contains most of the text manipulation methods.
- A list contains numbered or bulleted paragraphs.
- The document is contained in a window.
- A window has panes. A pane surrounds the visible area of the document.

For the full set of objects supported by the Word JavaScript API, see [Word JavaScript API.](https://learn.microsoft.com/en-us/javascript/api/word)

# **See also**

- Word JavaScript API overview
- Build your first Word add-in
- Word add-in tutorial
- [Word JavaScript API reference](https://learn.microsoft.com/en-us/javascript/api/word)
- [Learn about the Microsoft 365 Developer Program](https://aka.ms/m365devprogram)


# **Create a dictionary task pane add-in**

Article • 02/12/2025

This article shows you an example of a task pane add-in with an accompanying web service that provides dictionary definitions or thesaurus synonyms for the user's current selection in a Word document.

A dictionary Office Add-in is based on the standard task pane add-in with additional features to support querying and displaying definitions from a dictionary XML web service in additional places in the Office application's UI.

In a typical dictionary task pane add-in, a user selects a word or phrase in their document, and the JavaScript logic behind the add-in passes this selection to the dictionary provider's XML web service. The dictionary provider's webpage then updates to show the definitions for the selection to the user.

The XML web service component returns up to three definitions in the format defined by the example OfficeDefinitions XML schema, which are then displayed to the user in other places in the hosting Office application's UI.

Figure 1 shows the selection and display experience for a Bing-branded dictionary addin that's running in Word.

|  | Figure 1. Dictionary add-in displaying definitions for the selected word |  |  |  |  |  |  |
|--|--------------------------------------------------------------------------|--|--|--|--|--|--|
|  |                                                                          |  |  |  |  |  |  |
|  |                                                                          |  |  |  |  |  |  |
|  |                                                                          |  |  |  |  |  |  |

| File        |             |  |                                                    |  |                     |  |            | 그<br>A Editina<br>Home Insert Draw Design Layout Referenc Mailings Review View Develop Help Script La Callout :     | 9 ×<br>ar |
|-------------|-------------|--|----------------------------------------------------|--|---------------------|--|------------|---------------------------------------------------------------------------------------------------------------------|-----------|
|             |             |  |                                                    |  |                     |  | Search     |                                                                                                                     | ><br>×    |
|             |             |  |                                                    |  |                     |  | All v      | Definition                                                                                                          | Q         |
|             | Definition  |  |                                                    |  |                     |  | 1          | O This file<br>Files<br>ും Media                                                                                    |           |
|             |             |  |                                                    |  | Top Results         |  |            |                                                                                                                     |           |
|             |             |  |                                                    |  |                     |  |            | definition<br>defiini tion [ defa niSHan]                                                                           |           |
|             |             |  |                                                    |  |                     |  | NOUN       | definition (noun); definitions (plural noun)                                                                        |           |
|             |             |  |                                                    |  |                     |  | 1.         | a statement of the exact meaning of a word, especially<br>in a dictionary:<br>"a dictionary definition of the verb" |           |
|             |             |  |                                                    |  |                     |  |            | Synonyms: sense, explanation, denotation,<br>connotation, interpretation, elucidation, explication                  |           |
|             |             |  |                                                    |  |                     |  | 2.         | the degree of distinctness in outline of an object,<br>connacially of an imman in ,                                 |           |
|             |             |  |                                                    |  |                     |  |            | Was this useful? Yes No                                                                                             |           |
|             |             |  |                                                    |  |                     |  | Powered by |                                                                                                                     | Bing      |
| Page 1 of 1 | 1 of 1 word |  | OWe're starting the add-ins runtime, just a moment |  | La Display Settings |  | Focus      | les<br>를<br>ie                                                                                                      | +<br>100% |

It's up to you to determine if selecting the **See More** link in the dictionary add-in's HTML UI displays more information within the task pane or opens a separate window to 


the full webpage for the selected word or phrase.

Figure 2 shows the **Define** command in the context menu that enables users to quickly launch installed dictionaries. Figures 3 through 5 show the places in the Office UI where the dictionary XML services are used to provide definitions in Word.

*Figure 3. Definitions in the Spelling and Grammar panes*


| Spelling                                 |  | Jrammar                                           |
|------------------------------------------|--|---------------------------------------------------|
| Fanntistic                               |  | right                                             |
| IGNORE<br>IGNORE ALL<br>ADD              |  | IGNORE                                            |
| Fantastic                                |  | rite                                              |
| CHANGE<br>CHANGE ALL                     |  | CHANGE                                            |
| Fantastic 0                              |  | right @                                           |
| 1. extremely good or pleasant            |  | 1. exactly                                        |
| 2. extremely large                       |  | See more                                          |
| 3. not practical or sensible<br>See more |  | rite @                                            |
| Results by: Bing Dictionary              |  | 1. a traditional ceremony, especially a religious |
|                                          |  | one<br>See more                                   |
| English (United States)                  |  | Results by: Bing Dictionary                       |
|                                          |  | English (United States)                           |
|                                          |  |                                                   |
|                                          |  |                                                   |
|                                          |  |                                                   |
|                                          |  |                                                   |
|                                          |  |                                                   |
|                                          |  |                                                   |

*Figure 4. Definitions in the Thesaurus pane*


| 4 | Bizarre (adj.)    |
|---|-------------------|
|   | Bizarre           |
|   | Eccentric         |
|   | Strange           |
|   | Fanciful          |
|   | Weird             |
|   | Imaginary         |
|   | Whimsical         |
|   | Grotesque         |
|   | Odd               |
|   | Wild              |
|   | Crazy             |
|   | Normal (Antonym)  |
| 4 | Incredible (adj.) |
|   | Incredible        |
|   | Unbelievable      |
|   | Implausible       |
|   | Improbable        |
|   | Unlikely          |
|   | Farfetched        |
|   | Far-fetched       |
|   | Extraordinary     |
|   |                   |

| 1. extremely good or pleasant |
|-------------------------------|
| See more                      |
| Results by: Bing Dictionary   |
|                               |
| English (United States)       |

*Figure 5. Definitions in Reading Mode*


To create a task pane add-in that provides a dictionary lookup, create two main components.

- An XML web service that looks up definitions from a dictionary service, and then returns those values in an XML format that can be consumed and displayed by the dictionary add-in.
- A task pane add-in that submits the user's current selection to the dictionary web service, displays definitions, and can optionally insert those values into the document.

The following sections provide examples of how to create these components.

# **Prerequisites**

- [Visual Studio 2019 or later](https://www.visualstudio.com/vs/) with the **Office/SharePoint development** workload installed.
7 **Note**

If you've previously installed Visual Studio, use the Visual Studio Installer to ensure that the **Office/SharePoint development** workload is installed.

- Office connected to a Microsoft 365 subscription (including Office on the web).
Next, create a Word add-in project in Visual Studio.

- 1. In Visual Studio, choose **Create a new project**.
- 2. Using the search box, enter **add-in**. Choose **Word Web Add-in**, then select **Next**.


- 3. Name your project and select **Create**.
- 4. Visual Studio creates a solution and its two projects appear in **Solution Explorer**. The **Home.html** file opens in Visual Studio.

To learn more about the projects in a Word add-in solution, see the quick start.

# **Create a dictionary XML web service**

The XML web service must return queries to the web service as XML that conforms to the OfficeDefinitions XML schema. The following two sections describe the OfficeDefinitions XML schema, and provide an example of how to code an XML web service that returns queries in that XML format.

### **OfficeDefinitions XML schema**

The following code shows sample XSD for the OfficeDefinitions XML schema example.

```
XML
<?xml version="1.0" encoding="utf-8"?>
<xs:schema
 xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
 xmlns:xs="https://www.w3.org/2001/XMLSchema"
 targetNamespace="http://schemas.microsoft.com/contoso/OfficeDefinitions"
 xmlns="http://schemas.microsoft.com/contoso/OfficeDefinitions">
 <xs:element name="Result">
 <xs:complexType>
 <xs:sequence>
 <xs:element name="SeeMoreURL" type="xs:anyURI"/>
 <xs:element name="Definitions" type="DefinitionListType"/>
 </xs:sequence>
 </xs:complexType>
 </xs:element>
 <xs:complexType name="DefinitionListType">
 <xs:sequence>
 <xs:element name="Definition" maxOccurs="3">
 <xs:simpleType>
 <xs:restriction base="xs:normalizedString">
 <xs:maxLength value="400"/>
 </xs:restriction>
 </xs:simpleType>
 </xs:element>
 </xs:sequence>
 </xs:complexType>
</xs:schema>
```


Returned XML consists of a root **<Result>** element that contains a **<Definitions>** element with zero to three **<Definition>** child elements. Each child element contains definitions that are at most 400 characters in length. Additionally, the URL to the full page on the dictionary site must be provided in the **<SeeMoreURL>** element. The following example shows the structure of returned XML that conforms to the OfficeDefinitions schema.

```
XML
<?xml version="1.0" encoding="utf-8"?>
<Result xmlns="http://schemas.microsoft.com/contoso/OfficeDefinitions">
 <SeeMoreURL xmlns="">https://www.bing.com/search?q=example</SeeMoreURL>
 <Definitions xmlns="">
 <Definition>Definition1</Definition>
 <Definition>Definition2</Definition>
 <Definition>Definition3</Definition>
 </Definitions>
 </Result>
```
### **Sample dictionary XML web service**

The following C# code provides an example of how to write code for an XML web service that returns the result of a dictionary query in the OfficeDefinitions XML format.

```
cs
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Services;
using System.Xml;
using System.Text;
using System.IO;
using System.Net;
using System.Web.Script.Services;
/// <summary>
/// Summary description for _Default.
/// </summary>
[WebService(Namespace = "http://tempuri.org/")]
[WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
// To allow this web service to be called from script, using ASP.NET AJAX,
include the following line. 
[ScriptService]
public class WebService : System.Web.Services.WebService {
 public WebService () {
```


```
 // Uncomment the following line if using designed components.
 // InitializeComponent(); 
 }
 // You can replace this method entirely with your own method that gets
definitions
 // from your data source and then formats it into the example
OfficeDefinitions XML format. 
 // If you need a reference for constructing the returned XML, you can
use this example as a basis.
 [WebMethod]
 public XmlDocument Define(string word)
 {
 StringBuilder sb = new StringBuilder();
 XmlWriter writer = XmlWriter.Create(sb);
 {
 writer.WriteStartDocument();

 writer.WriteStartElement("Result", 
"http://schemas.microsoft.com/contoso/OfficeDefinitions");
 // See More URL should be changed to the dictionary
publisher's page for that word on
 // their website.
 writer.WriteElementString("SeeMoreURL", 
"https://www.bing.com/search?q=" + word);
 writer.WriteStartElement("Definitions");

 writer.WriteElementString("Definition", "Definition
1 of " + word);
 writer.WriteElementString("Definition", "Definition
2 of " + word);
 writer.WriteElementString("Definition", "Definition
3 of " + word);

 writer.WriteEndElement(); // End of Definitions element.
 writer.WriteEndElement(); // End of Result element.

 writer.WriteEndDocument();
 }
 writer.Close();
 XmlDocument doc = new XmlDocument();
 doc.LoadXml(sb.ToString());
 return doc;
 }
}
```
To get started with development, you can do the following.


#### **Create the web service**

- 1. Add a **Web Service (ASMX)** to the add-in's web application project in Visual Studio and name it **DictionaryWebService**.
- 2. Replace the entire content of the associated .asmx.cs file with the preceding C# code sample.

#### **Update the web service markup**

- 1. In the **Solution Explorer**, select the **DictionaryWebService.asmx** file then open its context menu and choose **View Markup**.
- 2. Replace the contents of DictionaryWebService.asmx with the following code.

```
XML
<%@ WebService Language="C#" CodeBehind="DictionaryWebService.asmx.cs"
Class="WebService" %>
```
### **Update the web.config**

- 1. In the **Web.config** of the add-in's web application project, add the following to the **<system.web>** node.

```
XML
<webServices>
 <protocols>
 <add name="HttpGet" />
 <add name="HttpPost" />
 </protocols>
</webServices>
```
- 2. Save your changes.
# **Components of a dictionary add-in**

A dictionary add-in consists of three main component files:

- An XML-formatted add-in only manifest file that describes the add-in.


The JSON-formatted **unified manifest for Microsoft 365** doesn't currently support dictionary add-ins.

- An HTML file that provides the add-in's UI.
- A JavaScript file that provides logic to get the user's selection from the document, sends the selection as a query to the web service, and then displays returned results in the add-in's UI.

### **Example of a dictionary add-in's manifest file**

The following is an example manifest file for a dictionary add-in.

```
XML
```

```
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.0"
 xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
 xsi:type="TaskPaneApp">
 <Id>7164e750-dc86-49c0-b548-1bac57abdc7c</Id>
 <Version>15.0</Version>
 <ProviderName>Microsoft Office Demo Dictionary</ProviderName>
 <DefaultLocale>en-us</DefaultLocale>
 <!--DisplayName is the name that will appear in the user's list of
applications.-->
 <DisplayName DefaultValue="Microsoft Office Demo Dictionary" />
 <!--Description is a 2-3 sentence description of this dictionary. -->
 <Description DefaultValue="The Microsoft Office Demo Dictionary is an
example built to demonstrate how a
 publisher can create a dictionary that integrates with Office. It
doesn't return real definitions." />
 <!--IconUrl is the URI for the icon that will appear in the user's list of
applications.-->
 <IconUrl
DefaultValue="http://contoso/_layouts/images/general/office_logo.jpg" />
 <SupportUrl DefaultValue="[Insert the URL of a page that provides support
information for the app]" />
 <!--Hosts specifies the kind of Office application your dictionary add-in
will support.
 You shouldn't have to modify this area.-->
 <Hosts>
 <Host Name="Document"/>
 </Hosts>
 <DefaultSettings>
 <!--SourceLocation is the URL for your dictionary.-->
 <SourceLocation
DefaultValue="http://contoso/ExampleDictionary/DictionaryHome.html" />
 </DefaultSettings>
 <!--Permissions is the set of permissions a user will have to give your
dictionary.
 If you need write access, such as to allow a user to replace the
highlighted word with a synonym,
```


```
 use ReadWriteDocument. -->
 <Permissions>ReadDocument</Permissions>
 <Dictionary>
 <!--TargetDialects is the set of regional languages your dictionary
contains. For example, if your
 dictionary applies to Spanish (Mexico) and Spanish (Peru), but not
Spanish (Spain), you can specify
 that here. Do not put more than one language (for example, Spanish
and English) here. Publish
 separate languages as separate dictionaries. -->
 <TargetDialects>
 <TargetDialect>EN-AU</TargetDialect>
 <TargetDialect>EN-BZ</TargetDialect>
 <TargetDialect>EN-CA</TargetDialect>
 <TargetDialect>EN-029</TargetDialect>
 <TargetDialect>EN-HK</TargetDialect>
 <TargetDialect>EN-IN</TargetDialect>
 <TargetDialect>EN-ID</TargetDialect>
 <TargetDialect>EN-IE</TargetDialect>
 <TargetDialect>EN-JM</TargetDialect>
 <TargetDialect>EN-MY</TargetDialect>
 <TargetDialect>EN-NZ</TargetDialect>
 <TargetDialect>EN-PH</TargetDialect>
 <TargetDialect>EN-SG</TargetDialect>
 <TargetDialect>EN-ZA</TargetDialect>
 <TargetDialect>EN-TT</TargetDialect>
 <TargetDialect>EN-GB</TargetDialect>
 <TargetDialect>EN-US</TargetDialect>
 <TargetDialect>EN-ZW</TargetDialect>
 </TargetDialects>
 <!--QueryUri is the address of this dictionary's XML web service (which
is used to put definitions in
 additional contexts, such as the spelling checker.)-->
 <QueryUri
DefaultValue="http://contoso/ExampleDictionary/WebService.asmx/Define?
word="/>
 <!--Citation Text, Dictionary Name, and Dictionary Home Page will be
combined to form the citation line
 (for example, this would produce "Examples by: Contoso",
 where "Contoso" is a hyperlink to http://www.contoso.com).-->
 <CitationText DefaultValue="Examples by: " />
 <DictionaryName DefaultValue="Contoso" />
 <DictionaryHomePage DefaultValue="http://www.contoso.com" />
 </Dictionary>
</OfficeApp>
```
The **<Dictionary>** element and its child elements specific to creating a dictionary addin's manifest file are described in the following sections. For information about the other elements in the manifest file, see Office Add-ins with the add-in only manifest.

### **Dictionary element**


Specifies settings for dictionary add-ins.

**Parent element**

**<OfficeApp>**

**Child elements**

**<TargetDialects>**, **<QueryUri>**, **<CitationText>**, **<Name>**, **<DictionaryHomePage>**

#### **Remarks**

The **<Dictionary>** element and its child elements are added to the manifest of a task pane add-in when you create a dictionary add-in.

### **TargetDialects element**

Specifies the regional languages that this dictionary supports. Required for dictionary add-ins.

**Parent element**

**<Dictionary>**

**Child element**

**<TargetDialect>**

#### **Remarks**

The **<TargetDialects>** element and its child elements specify the set of regional languages your dictionary contains. For example, if your dictionary applies to both Spanish (Mexico) and Spanish (Peru), but not Spanish (Spain), you can specify that in this element. Do not specify more than one language (e.g., Spanish and English) in this manifest. Publish separate languages as separate dictionaries.

#### **Example**

| XML                                                                                                                                                                                                                                                                                                                               |
|-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| <targetdialects><br/><targetdialect>EN-AU</targetdialect><br/><targetdialect>EN-BZ</targetdialect><br/><targetdialect>EN-CA</targetdialect><br/><targetdialect>EN-029</targetdialect><br/><targetdialect>EN-HK</targetdialect><br/><targetdialect>EN-IN</targetdialect><br/><targetdialect>EN-ID</targetdialect></targetdialects> |


#### **TargetDialect element**

Specifies a regional language that this dictionary supports. Required for dictionary addins.

#### **Parent element**

#### **<TargetDialects>**

#### **Remarks**

Specify the value for a regional language in the RFC1766 language tag format, such as EN-US.

#### **Example**

XML

<TargetDialect>EN-US</TargetDialect>

### **QueryUri element**

Specifies the endpoint for the dictionary query service. Required for dictionary add-ins.

**Parent element**

**<Dictionary>**

#### **Remarks**

This is the URI of the XML web service for the dictionary provider. The properly escaped query will be appended to this URI.


#### **Example**

XML

<QueryUri DefaultValue="http://msranlc-lingo1/proof.aspx?q="/>

#### **CitationText element**

Specifies the text to use in citations. Required for dictionary add-ins.

**Parent element**

**<Dictionary>**

#### **Remarks**

This element specifies the beginning of the citation text that will be displayed on a line below the content that is returned from the web service (for example, "Results by: " or "Powered by: ").

For this element, you can specify values for additional locales by using the **<Override>** element. For example, if a user is running the Spanish SKU of Office, but using an English dictionary, this allows the citation line to read "Resultados por: Bing" rather than "Results by: Bing". For more information about how to specify values for additional locales, see Localization.

#### **Example**

XML

<CitationText DefaultValue="Results by: " />

### **DictionaryName element**

Specifies the name of this dictionary. Required for dictionary add-ins.

**Parent element**

**<Dictionary>**

#### **Remarks**

This element specifies the link text in the citation text. Citation text is displayed on a line below the content that is returned from the web service.


For this element, you can specify values for additional locales.

#### **Example**

XML

<DictionaryName DefaultValue="Bing Dictionary" />

### **DictionaryHomePage element**

Specifies the URL of the home page for the dictionary. Required for dictionary add-ins.

**Parent element**

**<Dictionary>**

#### **Remarks**

This element specifies the link URL in the citation text. Citation text is displayed on a line below the content that is returned from the web service.

For this element, you can specify values for additional locales.

#### **Example**

XML

<DictionaryHomePage DefaultValue="https://www.bing.com" />

### **Update your dictionary add-in's manifest file**

- 1. Open the manifest file in the add-in project.
- 2. Update the value of the **<ProviderName>** element with your name.
- 3. Replace the value of the **<DisplayName>** element's **<DefaultValue>** attribute with an appropriate name, for example, "Microsoft Office Demo Dictionary".
- 4. Replace the value of the **<Description>** element's **<DefaultValue>** attribute with an appropriate description, for example, "The Microsoft Office Demo Dictionary is an example built to demonstrate how a publisher could create a dictionary that integrates with Office. It doesn't return real definitions.".
- 5. Add the following code after the **<Permissions>** node, replacing "contoso" references with your own company name, then save your changes.


```
XML
<Dictionary>
 <!--TargetDialects is the set of regional languages your dictionary
contains. For example, if your
 dictionary applies to Spanish (Mexico) and Spanish (Peru), but
not Spanish (Spain), you can
 specify that here. Do not put more than one language (for
example, Spanish and English) here.
 Publish separate languages as separate dictionaries. -->
 <TargetDialects>
 <TargetDialect>EN-AU</TargetDialect>
 <TargetDialect>EN-BZ</TargetDialect>
 <TargetDialect>EN-CA</TargetDialect>
 <TargetDialect>EN-029</TargetDialect>
 <TargetDialect>EN-HK</TargetDialect>
 <TargetDialect>EN-IN</TargetDialect>
 <TargetDialect>EN-ID</TargetDialect>
 <TargetDialect>EN-IE</TargetDialect>
 <TargetDialect>EN-JM</TargetDialect>
 <TargetDialect>EN-MY</TargetDialect>
 <TargetDialect>EN-NZ</TargetDialect>
 <TargetDialect>EN-PH</TargetDialect>
 <TargetDialect>EN-SG</TargetDialect>
 <TargetDialect>EN-ZA</TargetDialect>
 <TargetDialect>EN-TT</TargetDialect>
 <TargetDialect>EN-GB</TargetDialect>
 <TargetDialect>EN-US</TargetDialect>
 <TargetDialect>EN-ZW</TargetDialect>
 </TargetDialects>
 <!--QueryUri is the address of this dictionary's XML web service
(which is used to put definitions in
 additional contexts, such as the spelling checker.)-->
 <QueryUri DefaultValue="~remoteAppUrl/DictionaryWebService.asmx"/>
 <!--Citation Text, Dictionary Name, and Dictionary Home Page will be
combined to form the citation
 line (for example, this would produce "Examples by: Contoso",
where "Contoso" is a hyperlink to
 http://www.contoso.com).-->
 <CitationText DefaultValue="Examples by: " />
 <DictionaryName DefaultValue="Contoso" />
 <DictionaryHomePage DefaultValue="http://www.contoso.com" />
</Dictionary>
```
### **Create a dictionary add-in's HTML user interface**

The following two examples show the HTML and CSS files for the UI of the Demo Dictionary add-in. To view how the UI is displayed in the add-in's task pane, see Figure 6 following the code. To see how the implementation of the JavaScript provides


programming logic for this HTML UI, see Write the JavaScript implementation immediately following this section.

In the add-in's web application project in Visual Studio, you can replace the contents of the **./Home.html** file with the following sample HTML.

```
HTML
<!DOCTYPE html>
<html>
<head>
 <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
 <!--The title will not be shown but is supplied to ensure valid HTML.-->
 <title>Example Dictionary</title>
 <!--Required library includes.-->
 <script type="text/javascript"
src="https://ajax.microsoft.com/ajax/4.0/1/MicrosoftAjax.js"></script>
 <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"
type="text/javascript"></script>
 <!--Optional library includes.-->
 <script type="text/javascript"
src="https://ajax.aspnetcdn.com/ajax/jQuery/jquery-1.5.1.js"></script>
 <!--App-specific CSS and JS.-->
 <link rel="Stylesheet" type="text/css" href="Home.css" />
 <script type="text/javascript" src="Home.js"></script>
</head>
<body>
 <div id="mainContainer">
 <div>INSTRUCTIONS</div>
 <ol>
 <li>Ensure there's text in the document.</li>
 <li>Select text.</li>
 </ol>
 <div id="header">
 <span id="headword"></span>
 </div>
 <div>DEFINITIONS</div>
 <ol id="definitions">
 </ol>
 <div id="SeeMore">
 <a id="SeeMoreLink" target="_blank">See More...</a>
 </div>
 <div id="message"></div>
 </div>
</body>
```


```
</html>
```
The following example shows the contents of the .css file.

In the add-in's web application project in Visual Studio, you can replace the contents of the **./Home.css** file with the following sample CSS.

```
CSS
#mainContainer
{
 font-family: Segoe UI;
 font-size: 11pt;
}
#headword
{
 font-family: Segoe UI Semibold;
 color: #262626;
}
#definitions
{
 font-size: 8.5pt;
}
a
{
 font-size: 8pt;
 color: #336699;
 text-decoration: none;
}
a:visited
{
 color: #993366;
}
a:hover, a:active
{ 
 text-decoration: underline;
}
```
*Figure 6. Demo dictionary UI*


|                                         | INSTRUCTIONS                                                                                 |  |  |  |  |  |  |
|-----------------------------------------|----------------------------------------------------------------------------------------------|--|--|--|--|--|--|
|                                         | 1. Ensure there's text in the document.<br>2. Select text.                                   |  |  |  |  |  |  |
| Selected text: fantastic<br>DEFINITIONS |                                                                                              |  |  |  |  |  |  |
|                                         | 1. Definition 1 of fantastic<br>2. Definition 2 of fantastic<br>3. Definition 3 of fantastic |  |  |  |  |  |  |
| See More                                |                                                                                              |  |  |  |  |  |  |

### **Write the JavaScript implementation**

The following example shows the JavaScript implementation in the .js file that's called from the add-in's HTML page to provide the programming logic for the Demo Dictionary add-in. This script uses the XML web service described previously. When placed in the same directory as the example web service, the script will get definitions from that service. It can be used with a public OfficeDefinitions-conforming XML web service by modifying the xmlServiceURL variable at the top of the file.

The primary members of the Office JavaScript API (Office.js) that are called from this implementation are shown in the following list.

- The [initialize](https://learn.microsoft.com/en-us/javascript/api/office) event of the Office object, which is raised when the add-in context is initialized, and provides access to a [Document](https://learn.microsoft.com/en-us/javascript/api/office/office.document) object instance that represents the document the add-in is interacting with.
- The [addHandlerAsync](https://learn.microsoft.com/en-us/javascript/api/office/office.document#office-office-document-addhandlerasync-member(1)) method of the Document object, which is called in the initialize function to add an event handler for the [SelectionChanged](https://learn.microsoft.com/en-us/javascript/api/office/office.documentselectionchangedeventargs) event of the document to listen for user selection changes.
- The [getSelectedDataAsync](https://learn.microsoft.com/en-us/javascript/api/office/office.document#office-office-document-getselecteddataasync-member(1)) method of the Document object, which is called in the tryUpdatingSelectedWord() function when the SelectionChanged event handler is raised to get the word or phrase the user selected, coerce it to plain text, and then execute the selectedTextCallback asynchronous callback function.
- When the selectTextCallback asynchronous callback function that's passed as the *callback* argument of the getSelectedDataAsync method executes, it gets the value of the selected text when the callback returns. It gets that value from the callback's *selectedText* argument (which is of type [AsyncResult)](https://learn.microsoft.com/en-us/javascript/api/office/office.asyncresult) by using the [value](https://learn.microsoft.com/en-us/javascript/api/office/office.asyncresult#office-office-asyncresult-status-member) property of the returned AsyncResult object.


- The rest of the code in the selectedTextCallback function queries the XML web service for definitions.
- The remaining code in the .js file displays the list of definitions in the add-in's HTML UI.

In the add-in's web application project in Visual Studio, you can replace the contents of the **./Home.js** file with the following sample JavaScript.

```
JavaScript
// The document the dictionary add-in is interacting with.
let _doc;
// The last looked-up word, which is also the currently displayed word.
let lastLookup;
// The base URL for the OfficeDefinitions-conforming XML web service to
query for definitions.
const xmlServiceUrl = "DictionaryWebService.asmx/Define";
// Initialize the add-in.
// Office.initialize or Office.onReady is required for all add-ins.
Office.initialize = function (reason) {
 // Checks for the DOM to load using the jQuery ready method.
 $(document).ready(function () {
 // After the DOM is loaded, app-specific code can run.
 // Store a reference to the current document.
 _doc = Office.context.document;
 // Check whether text is already selected.
 tryUpdatingSelectedWord();
 // Add a handler to refresh when the user changes selection.
 _doc.addHandlerAsync("documentSelectionChanged",
tryUpdatingSelectedWord);
 });
}
// Executes when event is raised on the user's selection changes, and at
initialization time.
// Gets the current selection and passes that to asynchronous callback
function.
function tryUpdatingSelectedWord() {
 _doc.getSelectedDataAsync(Office.CoercionType.Text,
selectedTextCallback);
}
// Async callback that executes when the add-in gets the user's selection.
Determines whether anything should
// be done. If so, it makes requests that will be passed to various
functions.
function selectedTextCallback(selectedText) {
 selectedText = $.trim(selectedText.value);
 // Be sure user has selected text. The SelectionChanged event is raised
every time the user moves
```


```
 // the cursor, even if no selection.
 if (selectedText != "") {
 // Check whether the user selected the same word the pane is
currently displaying to
 // avoid unnecessary web calls.
 if (selectedText != lastLookup) {
 // Update the lastLookup variable.
 lastLookup = selectedText;
 // Set the "headword" span to the word you looked up.
 $("#headword").text("Selected text: " + selectedText);
 // AJAX request to get definitions for the selected word; pass
that to refreshDefinitions.
 $.ajax(xmlServiceUrl,
 {
 data: { word: selectedText },
 dataType: 'xml',
 success: refreshDefinitions,
 error: errorHandler
 });
 }
}
// This function is called when the add-in gets back the definitions target
word.
// It removes the old definitions and replaces them with the definitions for
the current word.
// It also sets the "See More" link.
function refreshDefinitions(data, textStatus, jqXHR) {
 $(".definition").remove();
 // Make a new list item for each returned definition that was returned,
set the CSS class,
 // and append it to the definitions div.
 $(data).find("Definition").each(function () {
 $(document.createElement("li"))
 .text($(this).text())
 .addClass("definition")
 .appendTo($("#definitions"));
 });
 // Change the "See More" link to direct to the correct URL.
 $("#SeeMoreLink").attr("href", $(data).find("SeeMoreURL").text());
}
// Basic error handler that writes to a div with id='message'.
function errorHandler(jqXHR, textStatus, errorThrown) {
 document.getElementById('message').innerText
 += ("textStatus:- " + textStatus
 + "\nerrorThrown:- " + errorThrown
 + "\njqXHR:- " + JSON.stringify(jqXHR));
}
```


# **Try it out**

- 1. Using Visual Studio, test the newly created Word add-in by pressing F5 or choosing **Debug** > **Start Debugging** to launch Word with the **Show Taskpane** addin button displayed on the ribbon. The add-in will be hosted locally on IIS.
- 2. In Word, if the add-in task pane isn't already open, choose the **Home** tab, and then choose the **Show Taskpane** button to open the add-in task pane. (If you're using the volume-licensed perpetual version of Office, instead of the Microsoft 365 version or a retail perpetual version, then custom buttons aren't supported. Instead, the task pane will open immediately.)

| File        | Home         | Insert<br>Draw       | Design | Layout | References<br>Mailings                                   | Review                  | View<br>Developer | Help                                  | Script Lab       |        | Comments       | Editing -                                 | ದ್<br>ಕ<br>V |
|-------------|--------------|----------------------|--------|--------|----------------------------------------------------------|-------------------------|-------------------|---------------------------------------|------------------|--------|----------------|-------------------------------------------|--------------|
|             | 13<br>િ<br>2 | Calibri (Body)<br>22 | Aa 4   | V      | !!!<br>==<br>14<br>V<br>V<br>=<br>=<br>NO<br>d<br>0<br>3 | +=<br>+=<br>Styles<br>4 | Editing<br>V      | Dictate<br>Transcribe<br>A Read Aloud | Sensitivity<br>4 | Editor | Reuse<br>Files | Show<br>Taskpane                          |              |
| Clipboard   | 5            | Fort                 |        | ি      | Paragraph                                                | Styles Fa               |                   | Vaice                                 | Sansitivity      | Editor |                | Reuse Files Commands Group                | V            |
|             |              |                      |        |        |                                                          |                         |                   |                                       |                  |        |                | Show Taskpane<br>Click to Show a Taskpane |              |
| Page 1 of 1 | 0 words      | Text Predictions: On | ਦ      |        | Accessibility: Good to go                                |                         |                   | Las Display Settings                  | Facus            |        | 그              | 0                                         | 100%         |

- 3. In Word, add text to the document then select any or all of that text.


# **Get the whole document from an add-in for PowerPoint or Word**

Article • 02/12/2025

You can create an Office Add-in to send or publish a PowerPoint presentation or Word document to a remote location. This article demonstrates how to build a simple task pane add-in for PowerPoint or Word that gets all of the presentation or document as a data object and sends that data to a web server via an HTTP request.

# **Prerequisites for creating an add-in for PowerPoint or Word**

This article assumes that you are using a text editor to create the task pane add-in for PowerPoint or Word. To create the task pane add-in, you must create the following files.

- On a shared network folder or on a web server, you need the following files.
	- An HTML file (**GetDoc_App.html**) that contains the user interface plus links to the JavaScript files (including Office.js and application-specific .js files) and Cascading Style Sheet (CSS) files.
	- A JavaScript file (**GetDoc_App.js**) to contain the programming logic of the addin.
	- A CSS file (**Program.css**) to contain the styles and formatting for the add-in.
- A manifest file (**GetDoc_App.xml** or **GetDoc_App.json**) for the add-in, available on a shared network folder or add-in catalog. The manifest file must point to the location of the HTML file mentioned previously.

Alternatively, you can create an add-in for your Office application using one of the following options. You won't have to create new files as the equivalent of each required file will be available for you to update. For example, the Yeoman generator options include **./src/taskpane/taskpane.html**, **./src/taskpane/taskpane.js**, **./src/taskpane/taskpane.css**, and **./manifest.xml**.

- PowerPoint
	- Visual Studio
	- Yeoman generator for Office Add-ins
- Word
	- Visual Studio


### **Core concepts to know for creating a task pane add-in**

Before you begin creating this add-in for PowerPoint or Word, you should be familiar with building Office Add-ins and working with HTTP requests. This article doesn't discuss how to decode Base64-encoded text from an HTTP request on a web server.

# **Create the manifest for the add-in**

The manifest file for an Office Add-in provides important information about the add-in: what applications can host it, the location of the HTML file, the add-in title and description, and many other characteristics.

In a text editor, add the following code to the manifest file. If you're using a Visual Studio project, select the "Add-in only manifest" option.

```
JSON
Unified manifest for Microsoft 365
  7 Note
  The unified manifest is generally available for production Outlook add-ins. It's
  available only for preview in Excel, PowerPoint, and Word add-ins.
  {
   "$schema": "https://developer.microsoft.com/json-
  schemas/teams/vDevPreview/MicrosoftTeams.schema.json#",
   "manifestVersion": "devPreview",
   "version": "1.0.0.0",
   "id": "[Replace_With_Your_GUID]",
   "localizationInfo": {
   "defaultLanguageTag": "en-us"
   },
   "developer": {
   "name": "[Provider Name e.g., Contoso]",
   "websiteUrl": "[Insert the URL for the app e.g.,
  https://www.contoso.com]",
   "privacyUrl": "[Insert the URL of a page that provides privacy
  information for the app e.g., https://www.contoso.com/privacy]",
   "termsOfUseUrl": "[Insert the URL of a page that provides terms
  of use for the app e.g., https://www.contoso.com/servicesagreement]"
   },
   "name": {
```


```
 "short": "Get Doc add-in",
 "full": "Get Doc add-in"
 },
 "description": {
 "short": "My get PowerPoint or Word document add-in.",
 "full": "My get PowerPoint or Word document add-in."
 },
 "icons": {
 "outline": "_layouts/images/general/office_logo.jpg",
 "color": "_layouts/images/general/office_logo.jpg"
 },
 "accentColor": "#230201",
 "validDomains": [
 "https://www.contoso.com"
 ],
 "showLoadingIndicator": false,
 "isFullScreen": false,
 "defaultBlockUntilAdminAction": false,
 "authorization": {
 "permissions": {
 "resourceSpecific": [
 {
 "name": "Document.ReadWrite.User",
 "type": "Delegated"
 }
 ]
 }
 },
 "extensions": [
 {
 "requirements": {
 "scopes": [
 "document",
             "presentation"
 ]
 },
 "alternates": [
 {
 "alternateIcons": {
 "icon": {
 "size": 32,
                   "url": 
"http://officeimg.vo.msecnd.net/_layouts/images/general/office_logo.jpg"
 },
                "highResolutionIcon": {
 "size": 64,
                   "url": 
"http://officeimg.vo.msecnd.net/_layouts/images/general/office_logo.jpg"
 }
 }
 }
 ]
 }
 ]
```
}


# **Create the user interface for the add-in**

For the user interface of the add-in, you can use HTML written directly into the **GetDoc_App.html** file. The programming logic and functionality of the add-in must be contained in a JavaScript file (for example, **GetDoc_App.js**).

Use the following procedure to create a simple user interface for the add-in that includes a heading and a single button.

- 1. In a new file in the text editor, add the HTML for your selected Office application.

```
HTML
PowerPoint
  <!DOCTYPE html>
  <html>
   <head>
   <meta charset="UTF-8" />
   <meta http-equiv="X-UA-Compatible" content="IE=Edge"/>
   <title>Publish presentation</title>
   <link rel="stylesheet" type="text/css" href="Program.css"
  />
   <script src="https://ajax.aspnetcdn.com/ajax/jquery/jquery-
  1.9.0.min.js" type="text/javascript"></script>
   <script
  src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"
  type="text/javascript"></script>
   <script src="GetDoc_App.js"></script>
   </head>
   <body>
   <form>
   <h1>Publish presentation</h1>
   <br />
   <div><input id='submit' type="button" value="Submit" />
  </div>
   <br />
   <div><h2>Status</h2>
   <div id="status"></div>
   </div>
   </form>
   </body>
  </html>
```


- 2. Save the file as **GetDoc_App.html** using UTF-8 encoding to a network location or to a web server.
7 **Note**

Be sure that the **head** tags of the add-in contains a **script** tag with a valid link to the Office.js file.

- 3. We'll use some CSS to give the add-in a simple yet modern and professional appearance. Use the following CSS to define the style of the add-in.
In a new file in the text editor, add the following CSS.

```
css
body
{
 font-family: "Segoe UI Light","Segoe UI",Tahoma,sans-serif;
}
h1,h2
{
 text-decoration-color:#4ec724;
}
input [type="submit"], input[type="button"]
{
 height:24px;
 padding-left:1em;
 padding-right:1em;
 background-color:white;
 border:1px solid grey;
 border-color: #dedfe0 #b9b9b9 #b9b9b9 #dedfe0;
 cursor:pointer;
}
```
- 4. Save the file as **Program.css** using UTF-8 encoding to the network location or to the web server where the **GetDoc_App.html** file is located.
# **Add the JavaScript to get the document**

In the code for the add-in, a handler to the [Office.initialize](https://learn.microsoft.com/en-us/javascript/api/office#office-office-initialize-function(1)) event adds a handler to the click event of the **Submit** button on the form and informs the user that the add-in is ready.

The following code example shows the event handler for the Office.initialize event along with a helper function, updateStatus , for writing to the status div.


JavaScript

```
// The initialize or onReady function is required for all add-ins.
Office.initialize = function (reason) {
 // Checks for the DOM to load using the jQuery ready method.
 $(document).ready(function () {
 // Run sendFile when Submit is clicked.
 $('#submit').on("click", function () {
 sendFile();
 });
 // Update status.
 updateStatus("Ready to send file.");
 });
}
// Create a function for writing to the status div.
function updateStatus(message) {
 var statusInfo = $('#status');
 statusInfo[0].innerHTML += message + "<br/>";
}
```
When you choose the **Submit** button in the UI, the add-in calls the sendFile function, which contains a call to the [Document.getFileAsync](https://learn.microsoft.com/en-us/javascript/api/office/office.document#office-office-document-getfileasync-member(1)) method. The getFileAsync method uses the asynchronous pattern, similar to other methods in the Office JavaScript API. It has one required parameter, *fileType*, and two optional parameters, *options* and *callback*.

The *fileType* parameter expects one of three constants from the [FileType](https://learn.microsoft.com/en-us/javascript/api/office/office.filetype) enumeration: Office.FileType.Compressed ("compressed"), Office.FileType.PDF ("pdf"), or Office.FileType.Text ("text"). The current file type support for each platform is listed under the [Document.getFileType](https://learn.microsoft.com/en-us/javascript/api/office/office.document#office-office-document-getfileasync-member(1)) remarks. When you pass in **Compressed** for the *fileType* parameter, the getFileAsync method returns the current document as a PowerPoint presentation file (*.pptx) or Word document file (*.docx) by creating a temporary copy of the file on the local computer.

The getFileAsync method returns a reference to the file as a [File](https://learn.microsoft.com/en-us/javascript/api/office/office.file) object. The File object exposes the following four members.

- [size](https://learn.microsoft.com/en-us/javascript/api/office/office.file#office-office-file-size-member) property
- [sliceCount](https://learn.microsoft.com/en-us/javascript/api/office/office.file#office-office-file-slicecount-member) property
- [getSliceAsync](https://learn.microsoft.com/en-us/javascript/api/office/office.file#office-office-file-getsliceasync-member(1)) method
- [closeAsync](https://learn.microsoft.com/en-us/javascript/api/office/office.file#office-office-file-closeasync-member(1)) method

The size property returns the number of bytes in the file. The sliceCount returns the number of [Slice](https://learn.microsoft.com/en-us/javascript/api/office/office.slice) objects (discussed later in this article) in the file.


Use the following code to get the current PowerPoint or Word document as a File object using the Document.getFileAsync method and then make a call to the locally defined getSlice function. Note that the File object, a counter variable, and the total number of slices in the file are passed along in the call to getSlice in an anonymous object.

```
JavaScript
// Get all of the content from a PowerPoint or Word document in 100-KB
chunks of text.
function sendFile() {
 Office.context.document.getFileAsync("compressed",
 { sliceSize: 100000 },
 function (result) {
 if (result.status === Office.AsyncResultStatus.Succeeded) {
 // Get the File object from the result.
 var myFile = result.value;
 var state = {
 file: myFile,
 counter: 0,
 sliceCount: myFile.sliceCount
 };
 updateStatus("Getting file of " + myFile.size + " bytes");
 getSlice(state);
 } else {
 updateStatus(result.status);
 }
 });
}
```
The local function getSlice makes a call to the File.getSliceAsync method to retrieve a slice from the File object. The getSliceAsync method returns a Slice object from the collection of slices. It has two required parameters, *sliceIndex* and *callback*. The *sliceIndex* parameter takes an integer as an indexer into the collection of slices. Like other methods in the Office JavaScript API, the getSliceAsync method also takes a callback function as a parameter to handle the results from the method call.

The Slice object gives you access to the data contained in the file. Unless otherwise specified in the *options* parameter of the getFileAsync method, the Slice object is 4 MB in size. The Slice object exposes three properties: [size](https://learn.microsoft.com/en-us/javascript/api/office/office.slice#office-office-slice-size-member), [data](https://learn.microsoft.com/en-us/javascript/api/office/office.slice#office-office-slice-data-member), and [index](https://learn.microsoft.com/en-us/javascript/api/office/office.slice#office-office-slice-index-member). The size property gets the size, in bytes, of the slice. The index property gets an integer that represents the slice's position in the collection of slices.

JavaScript


```
// Get a slice from the file and then call sendSlice.
function getSlice(state) {
 state.file.getSliceAsync(state.counter, function (result) {
 if (result.status == Office.AsyncResultStatus.Succeeded) {
 updateStatus("Sending piece " + (state.counter + 1) + " of " +
state.sliceCount);
 sendSlice(result.value, state);
 } else {
 updateStatus(result.status);
 }
 });
}
```
The Slice.data property returns the raw data of the file as a byte array. If the data is in text format (that is, XML or plain text), the slice contains the raw text. If you pass in **Office.FileType.Compressed** for the *fileType* parameter of Document.getFileAsync , the slice contains the binary data of the file as a byte array. In the case of a PowerPoint or Word file, the slices contain byte arrays.

You must implement your own function (or use an available library) to convert byte array data to a Base64-encoded string. For information about Base64 encoding with JavaScript, see [Base64 encoding and decoding](https://developer.mozilla.org/docs/Web/JavaScript/Base64_encoding_and_decoding) .

Once you've converted the data to Base64, you can then transmit it to a web server in several ways, including as the body of an HTTP POST request.

Add the following code to send a slice to a web service.

#### 7 **Note**

This code sends a PowerPoint or Word file to the web server in multiple slices. The web server or service must append each individual slice into a single file, and then save it as a .pptx or .docx file before you can perform any manipulations on it.

#### JavaScript

```
function sendSlice(slice, state) {
 var data = slice.data;
 // If the slice contains data, create an HTTP request.
 if (data) {
 // Encode the slice data, a byte array, as a Base64 string.
 // NOTE: The implementation of myEncodeBase64(input) function isn't
 // included with this example. For information about Base64 encoding
with
```


```
 // JavaScript, see
https://developer.mozilla.org/docs/Web/JavaScript/Base64_encoding_and_decodi
ng.
 var fileData = myEncodeBase64(data);
 // Create a new HTTP request. You need to send the request
 // to a webpage that can receive a post.
 var request = new XMLHttpRequest();
 // Create a handler function to update the status
 // when the request has been sent.
 request.onreadystatechange = function () {
 if (request.readyState == 4) {
 updateStatus("Sent " + slice.size + " bytes.");
 state.counter++;
 if (state.counter < state.sliceCount) {
 getSlice(state);
 } else {
 closeFile(state);
 }
 }
 }
 request.open("POST", "[Your receiving page or service]");
 request.setRequestHeader("Slice-Number", slice.index);
 // Send the file as the body of an HTTP POST
 // request to the web server.
 request.send(fileData);
 }
}
```
As the name implies, the File.closeAsync method closes the connection to the document and frees up resources. Although the Office Add-ins sandbox garbage collects out-of-scope references to files, it's still a best practice to explicitly close files once your code is done with them. The closeAsync method has a single parameter, *callback*, that specifies the function to call on the completion of the call.

```
JavaScript
function closeFile(state) {
 // Close the file when you're done with it.
 state.file.closeAsync(function (result) {
 // If the result returns as a success, the
 // file has been successfully closed.
 if (result.status === Office.AsyncResultStatus.Succeeded) {
 updateStatus("File closed.");
 } else {
 updateStatus("File couldn't be closed.");
```


 } }); }

The final JavaScript file could look like the following:

```
JavaScript
/*
* Copyright (c) Microsoft Corporation. All rights reserved. Licensed under
the MIT license.
* See LICENSE in the project root for license information.
*/
// The initialize or onReady function is required for all add-ins.
Office.initialize = function (reason) {
 // Checks for the DOM to load using the jQuery ready method.
 $(document).ready(function () {
 // Run sendFile when Submit is clicked.
 $('#submit').on("click", function () {
 sendFile();
 });
 // Update status.
 updateStatus("Ready to send file.");
 });
}
// Create a function for writing to the status div.
function updateStatus(message) {
 var statusInfo = $('#status');
 statusInfo[0].innerHTML += message + "<br/>";
}
// Get all of the content from a PowerPoint or Word document in 100-KB
chunks of text.
function sendFile() {
 Office.context.document.getFileAsync("compressed",
 { sliceSize: 100000 },
 function (result) {
 if (result.status === Office.AsyncResultStatus.Succeeded) {
 // Get the File object from the result.
 var myFile = result.value;
 var state = {
 file: myFile,
 counter: 0,
 sliceCount: myFile.sliceCount
 };
```


```
 updateStatus("Getting file of " + myFile.size + " bytes");
 getSlice(state);
 } else {
 updateStatus(result.status);
 }
 });
}
// Get a slice from the file and then call sendSlice.
function getSlice(state) {
 state.file.getSliceAsync(state.counter, function (result) {
 if (result.status == Office.AsyncResultStatus.Succeeded) {
 updateStatus("Sending piece " + (state.counter + 1) + " of " +
state.sliceCount);
 sendSlice(result.value, state);
 } else {
 updateStatus(result.status);
 }
 });
}
function sendSlice(slice, state) {
 var data = slice.data;
 // If the slice contains data, create an HTTP request.
 if (data) {
 // Encode the slice data, a byte array, as a Base64 string.
 // NOTE: The implementation of myEncodeBase64(input) function isn't
 // included with this example. For information about Base64 encoding
with
 // JavaScript, see
https://developer.mozilla.org/docs/Web/JavaScript/Base64_encoding_and_decodi
ng.
 var fileData = myEncodeBase64(data);
 // Create a new HTTP request. You need to send the request
 // to a webpage that can receive a post.
 var request = new XMLHttpRequest();
 // Create a handler function to update the status
 // when the request has been sent.
 request.onreadystatechange = function () {
 if (request.readyState == 4) {
 updateStatus("Sent " + slice.size + " bytes.");
 state.counter++;
 if (state.counter < state.sliceCount) {
 getSlice(state);
 } else {
 closeFile(state);
 }
 }
```
}


```
 request.open("POST", "[Your receiving page or service]");
 request.setRequestHeader("Slice-Number", slice.index);
 // Send the file as the body of an HTTP POST
 // request to the web server.
 request.send(fileData);
 }
}
function closeFile(state) {
 // Close the file when you're done with it.
 state.file.closeAsync(function (result) {
 // If the result returns as a success, the
 // file has been successfully closed.
 if (result.status === Office.AsyncResultStatus.Succeeded) {
 updateStatus("File closed.");
 } else {
 updateStatus("File couldn't be closed.");
 }
 });
}
```

# **Use fields in your Word add-in**

Article • 03/21/2025

A [field](https://support.microsoft.com/office/c429bbb0-8669-48a7-bd24-bab6ba6b06bb) in a Word document is a placeholder. It allows you to provide instructions for the content instead of the content itself. You can use fields to create and format a Word template. Word documents support a number of [field types](https://support.microsoft.com/office/1ad6d91a-55a7-4a8d-b535-cf7888659a51) , many with associated parameters for configuring the field. However, Word on the web generally doesn't support adding or editing fields through the UI. For more information, see [Field codes in](https://support.microsoft.com/office/d8f46094-13c3-4966-98c3-259748f3caf1) [Word for the web](https://support.microsoft.com/office/d8f46094-13c3-4966-98c3-259748f3caf1) .

Starting from the WordApi 1.5 requirement set, Word JavaScript APIs allow you to manage fields in your Word add-in. In all platforms, you can get existing fields. You can insert, update, and delete fields in platforms that support those capabilities.

The following sections discuss several of the most frequently used field types: Addin, Date, Hyperlink, and TOC (Table of Contents).

## **Addin field**

The Addin field is meant to store add-in data that's hidden from the Word user interface, regardless of whether fields in the document are set to show or hide its content. The Addin field isn't available in the Word UI's **Field** dialog box. Use the API to insert the Addin field type and set the field's data.

The following code sample shows how to insert an Addin field before the cursor location or your selection in the Word document.

```
JavaScript
// Inserts an Addin field before selection.
async function rangeInsertAddinField() {
 await Word.run(async (context) => {
 let range = context.document.getSelection().getRange();
 const field = range.insertField(Word.InsertLocation.before,
Word.FieldType.addin);
 field.load("result,code");
 await context.sync();
 if (field.isNullObject) {
 console.log("There are no fields in this document.");
 } else {
 console.log("Code of the field: " + field.code);
 console.log("Result of the field: " + JSON.stringify(field.result));
 }
```


 }); }

The following code sample shows how to get the first Addin field found in a document then set that field's data property.

```
JavaScript
// Gets the first Addin field in the document and sets its data.
async function getFirstAddinFieldAndSetData() {
 await Word.run(async (context) => {
 let myFieldTypes = new Array();
 myFieldTypes[0] = Word.FieldType.addin;
 const addinFields =
context.document.body.fields.getByTypes(myFieldTypes);
 let fields = addinFields.load("items");
 await context.sync();
 if (fields.items.length === 0) {
 console.log("No Addin fields in this document.");
 } else {
 fields.load();
 await context.sync();
 const firstAddinField = fields.items[0];
 firstAddinField.load("code,result,data");
 await context.sync();
 console.log("The data of the Addin field before being set:",
firstAddinField.data);
 const data = "Insert your data here";
 //const data = $("#input-reference").val(); // Or get user data from
your add-in's UI.
 firstAddinField.data = data;
 firstAddinField.load("data");
 await context.sync();
 console.log("The data of the Addin field after being set:",
firstAddinField.data);
 }
 });
}
```
## **Date field**

The Date field inserts the current date according to the format you specify. You can toggle between displaying the date or the field code by setting the showCodes field property to false or true respectively.


The following code sample shows how to insert a Date field before the cursor location or your selection in the Word document.

```
JavaScript
// Inserts a Date field before selection.
async function rangeInsertDateField() {
 await Word.run(async (context) => {
 let range = context.document.getSelection().getRange();
 const field = range.insertField(
 Word.InsertLocation.before,
 Word.FieldType.date,
 '\\@ "M/d/yyyy h:mm am/pm"',
 true
 );
 field.load("result,code");
 await context.sync();
 if (field.isNullObject) {
 console.warn("The field wasn't inserted as expected.");
 } else {
 console.log("Code of the field: " + field.code);
 console.log("Result of the field: " + JSON.stringify(field.result));
 }
 });
}
```
### **Further reading**

- [Manage Fields code sample](https://github.com/OfficeDev/office-js-snippets/blob/prod/samples/word/50-document/manage-fields.yaml)
- [Field codes: Date field](https://support.microsoft.com/office/d0c7e1f1-a66a-4b02-a3f4-1a1c56891306)

## **Hyperlink field**

JavaScript

The Hyperlink field inserts the address of either a location in the same document or an external location such as a webpage. When the user selects it, they're navigated to the specified location. You can toggle between displaying the hyperlink address or the field code by setting the showCodes field property to false or true respectively.

The following code sample shows how to insert a Hyperlink field before the cursor location or your selection in the Word document.

```
// Inserts a Hyperlink field before selection.
async function rangeInsertHyperlinkField() {
 await Word.run(async (context) => {
```


```
 let range = context.document.getSelection().getRange();
 const field = range.insertField(
 Word.InsertLocation.before,
 Word.FieldType.hyperlink,
 "https://bing.com",
 true
 );
 field.load("result,code");
 await context.sync();
 if (field.isNullObject) {
 console.warn("The field wasn't inserted as expected.");
 } else {
 console.log("Code of the field: " + field.code);
 console.log("Result of the field: " + JSON.stringify(field.result));
 }
 });
}
```
### **Further reading**

- [Field codes: Hyperlink field](https://support.microsoft.com/office/864f8577-eb2a-4e55-8c90-40631748ef53)
## **TOC (Table of Contents) field**

The TOC field inserts a table of contents, which lists certain areas of a document, like headings. You can toggle between displaying the table of contents or the field code by setting the showCodes field property to false or true respectively.

The following code sample shows how to insert a TOC field at the cursor location or replace your current selection in the Word document.

```
JavaScript
/**
 1. Run setup.
 1. Select "[To place table of contents]" paragraph.
 1. Run rangeInsertTOCField.
 */
// Inserts a TOC (table of contents) field replacing selection.
async function rangeInsertTOCField() {
 await Word.run(async (context) => {
 let range = context.document.getSelection().getRange();
 const field = range.insertField(
 Word.InsertLocation.replace,
 Word.FieldType.toc
 );
 field.load("result,code");
```


```
 await context.sync();
 if (field.isNullObject) {
 console.warn("The field wasn't inserted as expected.");
 } else {
 console.log("Code of the field: " + field.code);
 console.log("Result of the field: " + JSON.stringify(field.result));
 }
 });
}
// Prep document so there'll be elements that could be included in a table
of contents.
async function setup() {
 await Word.run(async (context) => {
 const body: Word.Body = context.document.body;
 body.clear();
 body.insertParagraph("Document title", "End").styleBuiltIn =
Word.BuiltInStyleName.title;
 body.insertParagraph("[To place table of contents]", "End").styleBuiltIn
= Word.BuiltInStyleName.normal;
 body.insertParagraph("Introduction", "End").styleBuiltIn =
Word.BuiltInStyleName.heading1;
 body.insertParagraph("Paragraph 1", "End").styleBuiltIn =
Word.BuiltInStyleName.normal;
 body.insertParagraph("Topic 1", "End").styleBuiltIn =
Word.BuiltInStyleName.heading1;
 body.insertParagraph("Paragraph 2", "End").styleBuiltIn =
Word.BuiltInStyleName.normal;
 body.insertParagraph("Topic 2", "End").styleBuiltIn =
Word.BuiltInStyleName.heading1;
 body.insertParagraph("Paragraph 3", "End").styleBuiltIn =
Word.BuiltInStyleName.normal;
 });
}
```
### **Further reading**

- [Field codes: TOC (Table of Contents) field](https://support.microsoft.com/office/1f538bc4-60e6-4854-9f64-67754d78d05c)
## **See also**

- [Field codes in Word for the web](https://support.microsoft.com/office/d8f46094-13c3-4966-98c3-259748f3caf1)
- [Insert, edit, and view fields in Word](https://support.microsoft.com/office/c429bbb0-8669-48a7-bd24-bab6ba6b06bb)


# **Use search options in your Word add-in to find text**

Article • 02/09/2024

Add-ins frequently need to act based on the text of a document. A search method is exposed by every content control (this includes [Body](https://learn.microsoft.com/en-us/javascript/api/word/word.body#word-word-body-search-member(1)), [Paragraph,](https://learn.microsoft.com/en-us/javascript/api/word/word.paragraph#word-word-paragraph-search-member(1)) [Range](https://learn.microsoft.com/en-us/javascript/api/word/word.range#word-word-range-search-member(1)), [Table,](https://learn.microsoft.com/en-us/javascript/api/word/word.table#word-word-table-search-member(1)) [TableRow](https://learn.microsoft.com/en-us/javascript/api/word/word.tablerow#word-word-tablerow-search-member(1)), and the base [ContentControl](https://learn.microsoft.com/en-us/javascript/api/word/word.contentcontrol#word-word-contentcontrol-search-member(1)) object). This method takes in a string (or wildcard expression) representing the text you are searching for and a [SearchOptions](https://learn.microsoft.com/en-us/javascript/api/word/word.searchoptions) object. It returns a collection of ranges which match the search text.

#### ) **Important**

The Word client may limit the available search options. For more details about current support, see **[Find and replace text](https://support.microsoft.com/office/c6728c16-469e-43cd-afe4-7708c6c779b7)** .

### **Search options**

The search options are a collection of boolean values defining how the search parameter should be treated.

| Property    | Description                                                                                                                                                                                       |
|-------------|---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| ignorePunct | Gets or sets a value indicating whether to ignore all punctuation characters<br>between words. Corresponds to the "Ignore punctuation characters"<br>checkbox in the Find and Replace dialog box. |
| ignoreSpace | Gets or sets a value indicating whether to ignore all whitespace between<br>words. Corresponds to the "Ignore white-space characters" checkbox in the<br>Find and Replace dialog box.             |
| matchCase   | Gets or sets a value indicating whether to perform a case-sensitive search.<br>Corresponds to the "Match case" checkbox in the Find and Replace dialog<br>box.                                    |
| matchPrefix | Gets or sets a value indicating whether to match words that begin with the<br>search string. Corresponds to the "Match prefix" checkbox in the Find and<br>Replace dialog box.                    |
| matchSuffix | Gets or sets a value indicating whether to match words that end with the<br>search string. Corresponds to the "Match suffix" checkbox in the Find and                                             |


| Property       | Description                                                                                                                                                                                                        |
|----------------|--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|                | Replace dialog box.                                                                                                                                                                                                |
| matchWholeWord | Gets or sets a value indicating whether to find operation only entire words,<br>not text that is part of a larger word. Corresponds to the "Find whole words<br>only" checkbox in the Find and Replace dialog box. |
| matchWildcards | Gets or sets a value indicating whether the search will be performed using<br>special search operators. Corresponds to the "Use wildcards" checkbox in the<br>Find and Replace dialog box.                         |

### **Search for special characters**

The following table lists the search notation for certain special characters.

| To find             | Notation |
|---------------------|----------|
| Paragraph mark      | ^p       |
| Tab mark            | ^t       |
| Any character       | ^?       |
| Any digit           | ^#       |
| Any letter          | ^$       |
| Caret character     | ^^       |
| Section character   | ^%       |
| Paragraph character | ^v       |
| Column break        | ^n       |
| Em dash             | ^+       |
| En dash             | ^=       |
| Endnote mark        | ^e       |
| Field               | ^d       |
| Footnote mark       | ^f       |
| Graphic             | ^g       |
| Manual line break   | ^l       |


| To find            | Notation |
|--------------------|----------|
| Manual page break  | ^m       |
| Nonbreaking hyphen | ^~       |
| Nonbreaking space  | ^s       |
| Optional hyphen    | ^-       |
| Section break      | ^b       |
| White Space        | ^w       |

### **Wildcard guidance**

The following table provides guidance around the Word JavaScript API's search wildcards.

| To find                                                                        | Wildcard | Sample                                                                         |
|--------------------------------------------------------------------------------|----------|--------------------------------------------------------------------------------|
| Any single character                                                           | ?        | s?t finds sat and set.                                                         |
| Any string of characters                                                       | *        | s*d finds sad and started.                                                     |
| The beginning of a word                                                        | <        | <(inter) finds interesting and intercept,<br>but not splintered.               |
| The end of a word                                                              | >        | (in)> finds in and within, but not<br>interesting.                             |
| One of the specified characters                                                | [ ]      | w[io]n finds win and won.                                                      |
| Any single character in this range                                             | [-]      | [r-t]ight finds right, sight, and tight.<br>Ranges must be in ascending order. |
| Any single character except the characters<br>in the range inside the brackets | [!x-z]   | t[!a-m]ck finds tock and tuck, but not<br>tack or tick.                        |
| Exactly n occurrences of the previous<br>character or expression               | {n}      | fe{2}d finds feed but not fed.                                                 |
| At least n occurrences of the previous<br>character or expression              | {n,}     | fe{1,}d finds fed and feed.                                                    |
| From n to m occurrences of the previous<br>character or expression             | {n,m}    | 10{1,3} finds 10, 100, and 1000.                                               |


| To find                                                            | Wildcard | Sample                   |
|--------------------------------------------------------------------|----------|--------------------------|
| One or more occurrences of the previous<br>character or expression | @        | lo@t finds lot and loot. |

### **Escape special characters**

Wildcard search is essentially the same as searching on a regular expression. There are special characters in regular expressions, including '[', ']', '(', ')', '{', '}', '*', '?', '<', '>', '!', and '@'. If one of these characters is part of the literal string the code is searching for, then it needs to be escaped, so that Word knows it should be treated literally and not as part of the logic of the regular expression. To escape a character in the Word UI search, you would precede it with a backslash character ('\'), but to escape it programmatically, put it between '[]' characters. For example, '[*]*' searches for any string that begins with a '*' followed by any number of other characters.

## **Examples**

The following examples demonstrate common scenarios.

### **Ignore punctuation search**

```
JavaScript
// Run a batch operation against the Word object model.
await Word.run(async (context) => {
 // Queue a command to search the document and ignore punctuation.
 const searchResults = context.document.body.search('video you',
{ignorePunct: true});
 // Queue a command to load the font property values.
 searchResults.load('font');
 // Synchronize the document state.
 await context.sync();
 console.log('Found count: ' + searchResults.items.length);
 // Queue a set of commands to change the font for each found item.
 for (let i = 0; i < searchResults.items.length; i++) {
 searchResults.items[i].font.color = 'purple';
 searchResults.items[i].font.highlightColor = '#FFFF00'; //Yellow
 searchResults.items[i].font.bold = true;
 }
 // Synchronize the document state.
```


### **Search based on a prefix**

```
JavaScript
// Run a batch operation against the Word object model.
await Word.run(async (context) => {
 // Queue a command to search the document based on a prefix.
 const searchResults = context.document.body.search('vid', {matchPrefix: 
true});
 // Queue a command to load the font property values.
 searchResults.load('font');
 // Synchronize the document state.
 await context.sync();
 console.log('Found count: ' + searchResults.items.length);
 // Queue a set of commands to change the font for each found item.
 for (let i = 0; i < searchResults.items.length; i++) {
 searchResults.items[i].font.color = 'purple';
 searchResults.items[i].font.highlightColor = '#FFFF00'; //Yellow
 searchResults.items[i].font.bold = true;
 }
 // Synchronize the document state.
 await context.sync();
});
```
### **Search based on a suffix**

```
JavaScript
// Run a batch operation against the Word object model.
await Word.run(async (context) => {
 // Queue a command to search the document for any string of characters
after 'ly'.
 const searchResults = context.document.body.search('ly', {matchSuffix: 
true});
 // Queue a command to load the font property values.
 searchResults.load('font');
 // Synchronize the document state.
 await context.sync();
```


```
 console.log('Found count: ' + searchResults.items.length);
 // Queue a set of commands to change the font for each found item.
 for (let i = 0; i < searchResults.items.length; i++) {
 searchResults.items[i].font.color = 'orange';
 searchResults.items[i].font.highlightColor = 'black';
 searchResults.items[i].font.bold = true;
 }
 // Synchronize the document state.
 await context.sync();
});
```
#### **Search using a wildcard**

```
JavaScript
// Run a batch operation against the Word object model.
await Word.run(async (context) => {
 // Queue a command to search the document with a wildcard
 // for any string of characters that starts with 'to' and ends with 'n'.
 const searchResults = context.document.body.search('to*n',
{matchWildcards: true});
 // Queue a command to load the font property values.
 searchResults.load('font');
 // Synchronize the document state.
 await context.sync();
 console.log('Found count: ' + searchResults.items.length);
 // Queue a set of commands to change the font for each found item.
 for (let i = 0; i < searchResults.items.length; i++) {
 searchResults.items[i].font.color = 'purple';
 searchResults.items[i].font.highlightColor = 'pink';
 searchResults.items[i].font.bold = true;
 }
 // Synchronize the document state.
 await context.sync();
});
```
### **Search for a special character**

JavaScript

```
// Run a batch operation against the Word object model.
await Word.run(async (context) => {
```


```
 // Queue a command to search the document for tabs.
 const searchResults = context.document.body.search('^t');
 // Queue a command to load the font property values.
 searchResults.load('font');
 // Synchronize the document state.
 await context.sync();
 console.log('Found count: ' + searchResults.items.length);
 // Queue a set of commands to change the font for each found item.
 for (let i = 0; i < searchResults.items.length; i++) {
 searchResults.items[i].font.color = 'purple';
 searchResults.items[i].font.highlightColor = 'pink';
 searchResults.items[i].font.bold = true;
 }
 // Synchronize the document state.
 await context.sync();
});
```
### **Search using a wildcard for an escaped special character**

As noted earlier in Escape special characters, there are special characters used by regular expressions. In order for a wildcard search to find one of those special characters programmatically, it'll need to be escaped using '[' and ']'. This example shows how to find the '{' special character using a wildcard search.

```
JavaScript
// Run a batch operation against the Word object model.
await Word.run(async (context) => {
 // Queue a command to search the document with a wildcard for an escaped
opening curly brace.
 const searchResults = context.document.body.search('[{]', { 
matchWildcards: true });
 // Queue a command to load the font property values.
 searchResults.load('font');
 // Synchronize the document state.
 await context.sync();
 console.log('Found count: ' + searchResults.items.length);
 // Queue a set of commands to change the font for each found item.
 for (let i = 0; i < searchResults.items.length; i++) {
 searchResults.items[i].font.color = 'purple';
 searchResults.items[i].font.highlightColor = 'pink';
 searchResults.items[i].font.bold = true;
```


```
 }
 // Synchronize the document state.
 await context.sync();
});
```
## **Try code examples in Script Lab**

Get the [Script Lab add-in](https://appsource.microsoft.com/product/office/wa104380862) and try out the code examples provided in this article. To learn more about Script Lab, see Explore Office JavaScript API using Script Lab.

### **See also**

More information can be found in the following:

- Word JavaScript Reference API
- Related Word code samples available in Script Lab:
	- [Search](https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/25-paragraph/search.yaml)
	- [Get word count](https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/25-paragraph/get-word-count.yaml)
- [Find and replace text in Word](https://support.microsoft.com/office/c6728c16-469e-43cd-afe4-7708c6c779b7)

#### 6 **Collaborate with us on GitHub**

The source for this content can be found on GitHub, where you can also create and review issues and pull requests. For more information, see [our](https://github.com/OfficeDev/office-js-docs-pr/blob/main/Contributing.md) [contributor guide](https://github.com/OfficeDev/office-js-docs-pr/blob/main/Contributing.md).

### **Office Add-ins feedback**

Office Add-ins is an open source project. Select a link to provide feedback:

[Open a documentation issue](https://github.com/OfficeDev/office-js-docs-pr/issues/new?template=3-customer-feedback.yml&pageUrl=https%3A%2F%2Flearn.microsoft.com%2Fen-us%2Foffice%2Fdev%2Fadd-ins%2Fword%2Fsearch-option-guidance&pageQueryParams=&contentSourceUrl=https%3A%2F%2Fgithub.com%2FOfficeDev%2Foffice-js-docs-pr%2Fblob%2Fmain%2Fdocs%2Fword%2Fsearch-option-guidance.md&documentVersionIndependentId=b23febfe-fe71-f17d-f757-edb1816a7fa8&feedback=%0A%0A%5BEnter+feedback+here%5D%0A&author=%40o365devx&metadata=*+ID%3A+6f870ea7-9577-ae36-b103-9f35bf43c2e3+%0A*+Service%3A+**word**%0A*+Sub-service%3A+**add-ins**)

- [Provide product feedback](https://aka.ms/office-addins-dev-questions)


# **Work with events using the Word JavaScript API**

08/05/2025

This article introduces key concepts for working with events in Word using the JavaScript API. You'll find practical code samples for registering, handling, and removing event handlers, along with explanations of event life cycles and coauthoring scenarios. Explore the event tables to discover which triggers and objects are supported.

### **Events in Word**

When certain changes occur in a Word document, event notifications fire. The Word JavaScript APIs let you register event handlers that allow your add-in to automatically run designated functions when those changes occur. The following events are currently supported.

| Event                   | Description                                                              | Supported<br>objects | Triggered<br>during<br>coauthoring? |
|-------------------------|--------------------------------------------------------------------------|----------------------|-------------------------------------|
| onAnnotationClicked     | Occurs when the user selects an<br>annotation.                           | Document             | No                                  |
|                         | Event data object:<br>AnnotationClickedEventArgs                         |                      |                                     |
| onAnnotationHovered     | Occurs when the user hovers the cursor<br>over an annotation.            | Document             | No                                  |
|                         | Event data object:<br>AnnotationHoveredEventArgs                         |                      |                                     |
| onAnnotationInserted    | Occurs when the user adds one or more<br>annotations.                    | Document             | No                                  |
|                         | Event data object:<br>AnnotationInsertedEventArgs                        |                      |                                     |
| onAnnotationPopupAction | Occurs when the user performs an action in<br>an annotation pop-up menu. | Document             | No                                  |
|                         | Event data object:<br>AnnotationPopupActionEventArgs                     |                      |                                     |


| Event                 | Description                                                                                                                                                          | Supported<br>objects | Triggered<br>during<br>coauthoring? |
|-----------------------|----------------------------------------------------------------------------------------------------------------------------------------------------------------------|----------------------|-------------------------------------|
| onAnnotationRemoved   | Occurs when the user deletes one or more<br>annotations.                                                                                                             | Document             | No                                  |
|                       | Event data object:<br>AnnotationRemovedEventArgs                                                                                                                     |                      |                                     |
| onContentControlAdded | Occurs when a content control is added.<br>Run context.sync() in the handler to get<br>the new content control's properties.                                         | Document             | Yes                                 |
|                       | Event data object:<br>ContentControlAddedEventArgs                                                                                                                   |                      |                                     |
| onDataChanged         | Occurs when data within the content<br>control are changed. To get the new text,<br>load this content control in the handler. To<br>get the old text, don't load it. | ContentControl       | Yes                                 |
|                       | Event data object:<br>ContentControlDataChangedEventArgs                                                                                                             |                      |                                     |
| onDeleted             | Occurs when the content control is deleted.<br>Don't load this content control in the<br>handler, otherwise you won't be able to get<br>its original properties.     | ContentControl       | Yes                                 |
|                       | Event data object:<br>ContentControlDeletedEventArgs                                                                                                                 |                      |                                     |
| onEntered             | Occurs when the content control is<br>entered.                                                                                                                       | ContentControl       | Yes                                 |
|                       | Event data object:<br>ContentControlEnteredEventArgs                                                                                                                 |                      |                                     |
| onExited              | Occurs when the content control is exited,<br>for example, when the cursor leaves the<br>content control.                                                            | ContentControl       | Yes                                 |
|                       | Event data object:<br>ContentControlExitedEventArgs                                                                                                                  |                      |                                     |
| onParagraphAdded      | Occurs when the user adds new<br>paragraphs.                                                                                                                         | Document             | Yes                                 |


| Event              | Description                                                     | Supported<br>objects | Triggered<br>during<br>coauthoring? |
|--------------------|-----------------------------------------------------------------|----------------------|-------------------------------------|
|                    | Event data object:<br>ParagraphAddedEventArgs                   |                      |                                     |
| onParagraphChanged | Occurs when the user changes paragraphs.                        | Document             | Yes                                 |
|                    | Event data object:<br>ParagraphChangedEventArgs                 |                      |                                     |
| onParagraphDeleted | Occurs when the user deletes paragraphs.                        | Document             | Yes                                 |
|                    | Event data object:<br>ParagraphDeletedEventArgs                 |                      |                                     |
| onSelectionChanged | Occurs when selection within the content<br>control is changed. | ContentControl       | Yes                                 |
|                    | Event data object:<br>ContentControlSelectionChangedEventArgs   |                      |                                     |

### **Events in preview**

#### 7 **Note**

The following events are currently available only in public preview. To use this feature, you must use the preview version of the Office JavaScript API library from the **[Office.js content](https://appsforoffice.microsoft.com/lib/beta/hosted/office.js) [delivery network (CDN)](https://appsforoffice.microsoft.com/lib/beta/hosted/office.js)** . The **[type definition file](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts)** for TypeScript compilation and IntelliSense is found at the CDN and **[DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts)** . You can install these types with npm install --save-dev @types/office-js-preview .

| ノ | Expand table |  |
|---|--------------|--|

| Event            | Description                                       | Supported objects                   | Triggered during<br>coauthoring? |
|------------------|---------------------------------------------------|-------------------------------------|----------------------------------|
| onCommentAdded   | Occurs when new comments are<br>added.            | Body<br>ContentControl<br>Paragraph | Yes                              |
|                  | Event data object:<br>CommentEventArgs            | Range                               |                                  |
| onCommentChanged | Occurs when a comment or its<br>reply is changed. | Body<br>ContentControl              | Yes                              |


| Event               | Description                             | Supported objects                   | Triggered during<br>coauthoring? |
|---------------------|-----------------------------------------|-------------------------------------|----------------------------------|
|                     | Event data object:<br>CommentEventArgs  | Paragraph<br>Range                  |                                  |
| onCommentDeleted    | Occurs when comments are<br>deleted.    | Body<br>Paragraph                   | Yes                              |
|                     | Event data object:<br>CommentEventArgs  |                                     |                                  |
| onCommentDeselected | Occurs when a comment is<br>deselected. | Body<br>ContentControl<br>Paragraph | Yes                              |
|                     | Event data object:<br>CommentEventArgs  | Range                               |                                  |
| onCommentSelected   | Occurs when a comment is<br>selected.   | Body<br>ContentControl<br>Paragraph | Yes                              |
|                     | Event data object:<br>CommentEventArgs  | Range                               |                                  |

## **Event triggers**

Events within a Word document can be triggered by:

- User interaction via the Word user interface (UI) that changes the document.
- Office Add-in (JavaScript) code that changes the document.
- VBA add-in (macro) code that changes the document.
- A coauthor who remotely changes the document using the Word UI or add-in code. For more information, see Events and coauthoring.

Any change that complies with default behavior of Word will trigger the corresponding events in a document.

## **Life cycle of an event handler**

An event handler is created when an add-in registers the event handler. It's destroyed when the add-in deregisters the event handler or when the add-in is refreshed, reloaded, or closed. Event handlers don't persist as part of the Word file, or across sessions with Word on the web.


## **Events and coauthoring**

With coauthoring, multiple people can work together and edit the same Word document simultaneously. For events that can be triggered by a coauthor, such as onParagraphChanged , the corresponding **Event** object will contain a **source** property that indicates whether the event was triggered locally by the current user ( event.source == Local ) or was triggered by the remote coauthor ( event.source == Remote ).

Events that use the following data objects are triggered during coauthoring.

- CommentEventArgs (preview)
- ContentControlAddedEventArgs
- ContentControlDataChangedEventArgs
- ContentControlDeletedEventArgs
- ContentControlEnteredEventArgs
- ContentControlExitedEventArgs
- ContentControlSelectionChangedEventArgs
- ParagraphAddedEventArgs
- ParagraphChangedEventArgs
- ParagraphDeletedEventArgs

## **Register an event handler**

The following code sample registers an event handler for the onParagraphChanged event in the document. The code specifies that when content changes in the document, the handleChange function runs.

```
JavaScript
await Word.run(async (context) => {
 eventContext = context.document.onParagraphChanged.add(handleChange);
 await context.sync();
 console.log("Event handler successfully registered for onParagraphChanged
event in the document.");
}).catch(errorHandlerFunction);
```
As shown in the previous example, when you register an event handler, you indicate the function that should run when the specified event occurs. You can design that function to perform whatever actions your scenario requires. The following code sample shows an event handler function that simply writes information about the event to the console.


```
JavaScript
```

```
async function handleChange(event) {
 await Word.run(async (context) => {
 await context.sync(); 
 console.log("Type of event: " + event.type);
 console.log("Source of event: " + event.source); 
 }).catch(errorHandlerFunction);
}
```
### **Remove an event handler**

The following code sample registers an event handler for the onParagraphChanged event in the document and defines the handleChange function that will run when the event occurs. It also defines the deregisterEventHandler() function that can subsequently be called to remove that event handler. Note that the RequestContext used to create the event handler is needed to remove it.

JavaScript

```
let eventContext;
async function registerEventHandler() {
 await Word.run(async (context) => {
 eventContext = context.document.onParagraphChanged.add(handleChange);
 await context.sync();
 console.log("Event handler successfully registered for onParagraphChanged
event in the document.");
 });
}
async function handleChange(event: Word.ParagraphChangedEventArgs) {
 await Word.run(async (context) => {
 await context.sync();
 console.log(`${event.type} event was detected.`);
 });
}
async function deregisterEventHandler() {
 // The `RequestContext` used to create the event handler is needed to remove it.
 // In this example, `eventContext` is being used to keep track of that context.
 await Word.run(eventContext.context, async (context) => {
 eventContext.remove();
 await context.sync();

 eventContext = null;
 console.log("Removed event handler that was tracking content changes in
paragraphs.");
```


#### }); }

### **Use .track()**

Certain event types also require you to call track() on the object you're adding the event to.

- Content control events
	- onDataChanged
	- onDeleted
	- onEntered
	- onExited
	- onSelectionChanged
- Comment events (preview)
	- onCommentAdded
	- onCommentChanged
	- onCommentDeleted
	- onCommentDeselected
	- onCommentSelected

The following code sample shows how to register an event handler on each content control. Because you're adding the event to the content controls, track() is called on each content control in the collection.

#### TypeScript

```
let eventContexts = [];
await Word.run(async (context) => {
 const contentControls: Word.ContentControlCollection =
context.document.contentControls;
 contentControls.load("items");
 await context.sync();
 // Register the onDeleted event handler on each content control.
 if (contentControls.items.length === 0) {
 console.log("There aren't any content controls in this document so can't
register event handlers.");
 } else {
 for (let i = 0; i < contentControls.items.length; i++) {
 eventContexts[i] =
contentControls.items[i].onDeleted.add(contentControlDeleted);
 // Call track() on each content control.
 contentControls.items[i].track();
```


```
 }
 await context.sync();
 console.log("Added event handlers for when content controls are deleted.");
 }
});
```
The following code sample shows how to register comment event handlers on the document's body object and includes a body.track() call.

```
TypeScript
let eventContexts = [];
// Registers event handlers.
await Word.run(async (context) => {
 const body: Word.Body = context.document.body;
 // Track the body object since you're adding comment events to it.
 body.track();
 await context.sync();
 eventContexts[0] = body.onCommentAdded.add(onEventHandler);
 eventContexts[1] = body.onCommentChanged.add(onChangedHandler);
 eventContexts[2] = body.onCommentDeleted.add(onEventHandler);
 eventContexts[3] = body.onCommentDeselected.add(onEventHandler);
 eventContexts[4] = body.onCommentSelected.add(onEventHandler);
 await context.sync();
 console.log("Event handlers registered.");
});
```
## **See also**

- Word JavaScript object model in Office Add-ins
- These and other examples are available in our Script Lab tool:
	- [On changing content in paragraphs](https://github.com/OfficeDev/office-js-snippets/blob/prod/samples/word/25-paragraph/onchanged-event.yaml)
	- [On deleting content controls](https://github.com/OfficeDev/office-js-snippets/blob/prod/samples/word/10-content-controls/content-control-ondeleted-event.yaml)
	- [Manage comments](https://github.com/OfficeDev/office-js-snippets/blob/prod/samples/word/99-preview-apis/manage-comments.yaml) (preview)


# **Understand when and how to use Office Open XML in your Word add-in**

Article • 06/07/2023

**Provided by:** Stephanie Krieger, Microsoft Corporation | Juan Balmori Labra, Microsoft Corporation

If you're building Office Add-ins to run in Word, you might already know that the Office JavaScript API (Office.js) offers several formats for reading and writing document content. These are called coercion types, and they include plain text, tables, HTML, and Office Open XML.

## **Options for adding rich content**

So what are your options when you need to add rich content to a document, such as images, formatted tables, charts, or even just formatted text?

- 1. **Word JavaScript APIs.** Start with the APIs available through the [WordApi](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) [requirement sets](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) to see if they provide what you need. For an example, see the [Insert formatted text](https://github.com/OfficeDev/office-js-snippets/blob/prod/samples/word/25-paragraph/insert-formatted-text.yaml) code snippet. You can try this and other snippets in the [Script Lab add-in](https://appsource.microsoft.com/product/office/wa104380862) on Word! To learn more about Script Lab, see Explore Office JavaScript API using Script Lab.
- 2. **HTML coercion.** If APIs aren't yet available, you can use HTML for inserting some types of rich content, such as pictures. Depending on your scenario, there can be drawbacks to HTML coercion, such as limitations in the formatting and positioning options available to your content.
- 3. **Office Open XML.** Because Office Open XML is the language in which Word documents (such as .docx and .dotx) are written, you can insert virtually any type of content that a user can add to a Word document, with virtually any type of formatting the user can apply. Determining the Office Open XML markup you need to get it done is easier than you might think.

#### 7 **Note**

Office Open XML is also the language behind PowerPoint and Excel (and, as of Office 2013, Visio) documents. However, currently, you can coerce content as Office Open XML only in Office Add-ins created for Word. For more information about


Office Open XML, including the complete language reference documentation, see the **See also** section.

### **Download the companion code sample**

Download the code sample [Load and write Open XML in your Word add-in](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/word-add-in-load-and-write-open-xml) , which contains the Office Open XML markup and Office.js code required for inserting any of the following examples into Word.

### **Learn about content types**

To begin, take a look at some of the content types you can insert using Office Open XML coercion.

Throughout this article, the terms **content types** and **rich content** refer to the types of rich content you can insert into a Word document.

*Figure 1. Text with direct formatting*

Use direct formatting to specify exactly what the text will look like regardless of existing formatting in the user's document.

*Figure 2. Text formatted using a style*

Use a style to automatically coordinate the look of text you insert with the user's document.

*Figure 3. A simple image*


Use the same method for inserting any Office-supported image format.

*Figure 4. An image formatted using picture styles and effects*

Adding high quality formatting and effects to your images requires much less markup than you might expect.

#### *Figure 5. A content control*

Use content controls with your add-in to add content at a specified (bound) location rather than at the selection.

#### *Figure 6. A text box with WordArt formatting*

Text effects are available in Word for text inside a text box (as shown here) or for regular body text.

*Figure 7. A shape*


Insert built-in or custom drawing shapes, with or without text and formatting effects.

*Figure 8. A table with direct formatting*

| Region    | Q1      | Q2      |
|-----------|---------|---------|
| Southeast | 123,456 | 234,567 |
| Northwest | 234,567 | 345,678 |

Include text formatting, borders, shading, cell sizing, or any table formatting you need.

*Figure 9. A table formatted using a table style*

| Region   Q1         |  | Q2      |
|---------------------|--|---------|
| Southeast   123,456 |  | 234,567 |
| Northwest 234,567   |  | 345,678 |

Use built-in or custom table styles just as easily as using a paragraph style for text.

*Figure 10. A SmartArt diagram*

Office offers a wide array of SmartArt diagram layouts (and you can use Office Open XML to create your own).

*Figure 11. A chart*


You can insert Excel charts as live charts in Word documents, which also means you can use them in your add-in for Word. As you can see by the preceding examples, you can use Office Open XML coercion to insert essentially any type of content that a user can insert into their own document. There are two simple ways to get the Office Open XML markup you need. Either add your rich content to an otherwise blank Word document and then save the file in Word XML Document format or use a test add-in with the [getSelectedDataAsync](https://learn.microsoft.com/en-us/javascript/api/office/office.document#office-office-document-getselecteddataasync-member(1)) method to grab the markup. Both approaches provide essentially the same result.

#### **Tip**

An Office Open XML document is actually a compressed package of files that represent the document contents. Saving the file in the Word XML Document format gives you the entire Office Open XML package flattened into one XML file, which is also what you get when using getSelectedDataAsync to retrieve the Office Open XML markup.

If you save the file to an XML format from Word, note that there are two options under the Save as Type list in the Save As dialog box for .xml format files. Be sure to choose **Word XML Document**, not the Word 2003 option.

Download the code sample named [Word-Add-in-Get-Set-EditOpen-XML](https://github.com/OfficeDev/Word-Add-in-Get-Set-EditOpen-XML) , which you can use as a tool to retrieve and test your markup.

So is that all there is to it? Well, not quite. Yes, for many scenarios, you could use the full, flattened Office Open XML result you see with either of the preceding methods and it would work. The good news is that you probably don't need most of that markup.

If you're one of the many add-in developers seeing Office Open XML markup for the first time, trying to make sense of the massive amount of markup you get for the simplest piece of content might seem overwhelming, but it doesn't have to be.


In this topic, you'll use some common scenarios we've been hearing from the Office Add-ins developer community to show you techniques for simplifying Office Open XML for use in your add-in. We'll explore the markup for some types of content shown earlier along with the information you need for minimizing the Office Open XML payload. We'll also look at the code you need for inserting rich content into a document at the active selection and how to use Office Open XML with the bindings object to add or replace content at specified locations.

## **Explore the Office Open XML document package**

When you use [getSelectedDataAsync](https://learn.microsoft.com/en-us/javascript/api/office/office.document#office-office-document-getselecteddataasync-member(1)) to retrieve the Office Open XML for a selection of content (or when you save the document in Word XML Document format), what you're getting isn't just the markup that describes your selected content; it's an entire document with many options and settings that you almost certainly don't need. In fact, if you use that method from a document that contains a task pane add-in, the markup you get even includes your task pane.

Even a simple Word document package includes parts for document properties, styles, theme (formatting settings), web settings, fonts, and then some, in addition to parts for the actual content.

For example, say that you want to insert just a paragraph of text with direct formatting, as shown earlier in Figure 1. When you grab the Office Open XML for the formatted text using getSelectedDataAsync , you see a large amount of markup. That markup includes a package element that represents an entire document, which contains several parts (commonly referred to as document parts or, in the Office Open XML, as package parts), as you see listed in Figure 13. Each part represents a separate file within the package.

You can edit Office Open XML markup in a text editor like Notepad. If you open it in Visual Studio, use **Edit** > **Advanced** > **Format Document** ( Ctrl + K , Ctrl + D ) to format the package for easier editing. Then you can collapse or expand document parts or sections of them, as shown in Figure 12, to more easily review and edit the content of the Office Open XML package. Each document part begins with a **pkg:part** tag.

*Figure 12. Collapse and expand package parts for easier editing in Visual Studio*


*Figure 13. The parts included in a basic Word Office Open XML document package*

With all that markup, you might be surprised to discover that the only elements you actually need to insert the formatted text example are pieces of the .rels part and the document.xml part.

#### **Tip**

The two lines of markup above the package tag (the XML declarations for version and Office program ID) are assumed when you use the Office Open XML coercion type, so you don't need to include them. Keep them if you want to open your edited markup as a Word document to test it.

Several of the other types of content shown at the start of this topic require additional parts as well (beyond those shown in Figure 13), and you'll address those later in this topic. Meanwhile, since you'll see most of the parts shown in Figure 13 in the markup for any Word document package, here's a quick summary of what each of these parts is for and when you need it:

- Inside the package tag, the first part is the .rels file, which defines relationships between the top-level parts of the package (these are typically the document properties, thumbnail (if any), and main document body). Some of the content in this part is always required in your markup because you need to define the relationship of the main document part (where your content resides) to the document package.
- The document.xml.rels part defines relationships for additional parts required by the document.xml (main body) part, if any.


The .rels files in your package (such as the top-level .rels, document.xml.rels, and others you may see for specific types of content) are an extremely important tool that you can use as a guide for helping you quickly edit down your Office Open XML package. To learn more about how to do this, see **Create your own markup: best practices** later in this topic.

- The document.xml part is the content in the main body of the document. You need elements of this part, of course, since that's where your content appears. But, you don't need everything you see in this part. We'll look at that in more detail later.
- Many parts are automatically ignored by the Set methods when inserting content into a document using Office Open XML coercion, so you might as well remove them. These include the theme1.xml file (the document's formatting theme), the document properties parts (core, add-in, and thumbnail), and setting files (including settings, webSettings, and fontTable).
- In the Figure 1 example, text formatting is directly applied (that is, each font and paragraph formatting setting applied individually). But, if you use a style (such as if you want your text to automatically take on the formatting of the Heading 1 style in the destination document) as shown earlier in Figure 2, then you would need part of the styles.xml part as well as a relationship definition for it. For more information, see the topic section Add objects that use additional Office Open XML parts.

### **Insert document content at the selection**

Let's take a look at the minimal Office Open XML markup required for the formatted text example shown in Figure 1 and the JavaScript required for inserting it at the active selection in the document.

### **Simplified Office Open XML markup**

The Office Open XML example shown here was edited as described in the preceding section to leave just required document parts and only required elements within each of those parts. You'll walk through how to edit the markup yourself (and we'll explain a bit more about the pieces that remain here) in the next section of the topic.

XML

```
<pkg:package
xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
 <pkg:part pkg:name="/_rels/.rels"
```


```
pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"
pkg:padding="512">
 <pkg:xmlData>
 <Relationships
xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
 <Relationship Id="rId1"
Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/of
ficeDocument" Target="word/document.xml"/>
 </Relationships>
 </pkg:xmlData>
 </pkg:part>
 <pkg:part pkg:name="/word/document.xml"
pkg:contentType="application/vnd.openxmlformats-
officedocument.wordprocessingml.document.main+xml">
 <pkg:xmlData>
 <w:document
xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" >
 <w:body>
 <w:p>
 <w:pPr>
 <w:spacing w:before="360" w:after="0" w:line="480"
w:lineRule="auto"/>
 <w:rPr>
 <w:color w:val="70AD47" w:themeColor="accent6"/>
 <w:sz w:val="28"/>
 </w:rPr>
 </w:pPr>
 <w:r>
 <w:rPr>
 <w:color w:val="70AD47" w:themeColor="accent6"/>
 <w:sz w:val="28"/>
 </w:rPr>
 <w:t>This text has formatting directly applied to achieve its
font size, color, line spacing, and paragraph spacing.</w:t>
 </w:r>
 </w:p>
 </w:body>
 </w:document>
 </pkg:xmlData>
 </pkg:part>
</pkg:package>
```
#### 7 **Note**

If you add the markup shown here to an XML file along with the XML declaration tags for version and mso-application at the top of the file (shown in Figure 13), you can open it in Word as a Word document. Or, without those tags, you can still open it using **File** > **Open** in Word. You'll see **Compatibility Mode** on the title bar in Word, because you removed the settings that tell Word this is a Word document.


Since you're adding this markup to an existing Word document, that won't affect your content at all.

### **JavaScript for using setSelectedDataAsync**

Once you save the preceding Office Open XML as an XML file that's accessible from your solution, use the following function to set the formatted text content in the document using Office Open XML coercion.

In the following function, notice that all but the last line are used to get your saved markup for use in the [setSelectedDataAsync](https://learn.microsoft.com/en-us/javascript/api/office/office.document#office-office-document-setselecteddataasync-member(1)) method call at the end of the function. setSelectedDataASync requires only that you specify the content to be inserted and the coercion type.

Replace *yourXMLfilename* with the name and path of the XML file as you've saved it in your solution. If you aren't sure where to include XML files in your solution or how to reference them in your code, see the [Load and write Open XML in your Word add-in](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/word-add-in-load-and-write-open-xml) code sample for examples of that and a working example of the markup and JavaScript shown here.

```
JavaScript
function writeContent() {
 const myOOXMLRequest = new XMLHttpRequest();
 let myXML;
 myOOXMLRequest.open('GET', 'yourXMLfilename', false);
 myOOXMLRequest.send();
 if (myOOXMLRequest.status === 200) {
 myXML = myOOXMLRequest.responseText;
 }
 Office.context.document.setSelectedDataAsync(myXML, { coercionType: 
'ooxml' });
}
```
### **Create your own markup: best practices**

Let's take a closer look at the markup you need to insert the preceding formatted text example.

For this example, start by simply deleting all document parts from the package other than .rels and document.xml. Then, you'll edit those two required parts to simplify things further.


#### ) **Important**

Use the .rels parts as a map to quickly gauge what's included in the package and determine what parts you can delete completely (that is, any parts not related to or referenced by your content). Remember that every document part must have a relationship defined in the package and those relationships appear in the .rels files. So you should see all of them listed in either .rels, document.xml.rels, or a contentspecific .rels file.

The following markup shows the required .rels part before editing. Since we're deleting the add-in and core document property parts, and the thumbnail part, you need to delete those relationships from .rels as well. Notice that this will leave only the relationship (with the relationship ID "rID1" in the following example) for document.xml.

```
XML
<pkg:part pkg:name="/_rels/.rels"
pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"
pkg:padding="512">
 <pkg:xmlData>
 <Relationships
xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
 <Relationship Id="rId3"
Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/
core-properties" Target="docProps/core.xml"/>
 <Relationship Id="rId2"
Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/
thumbnail" Target="docProps/thumbnail.emf"/>
 <Relationship Id="rId1"
Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/of
ficeDocument" Target="word/document.xml"/>
 <Relationship Id="rId4"
Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/ex
tended-properties" Target="docProps/app.xml"/>
 </Relationships>
 </pkg:xmlData>
</pkg:part>
```
Remove the relationships (that is, the **Relationship** tag) for any parts that you completely remove from the package. Including a part without a corresponding relationship, or excluding a part and leaving its relationship in the package, will result in an error.

The following markup shows the document.xml part, which includes our sample formatted text content before editing.


```
<pkg:part pkg:name="/word/document.xml"
pkg:contentType="application/vnd.openxmlformats-
officedocument.wordprocessingml.document.main+xml">
 <pkg:xmlData>
 <w:document mc:Ignorable="w14 w15 wp14"
xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanva
s" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships
```

```
xmlns:v="urn:schemas-microsoft-com:vml"
xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDraw
ing"
```

```
xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDra
wing" xmlns:w10="urn:schemas-microsoft-com:office:word"
```

```
xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
```
" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"

```
xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
```

```
xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup
"
```

```
xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk"
xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"
xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape
">
```

```
 <w:body>
 <w:p>
 <w:pPr>
 <w:spacing w:before="360" w:after="0" w:line="480"
w:lineRule="auto"/>
 <w:rPr>
 <w:color w:val="70AD47" w:themeColor="accent6"/>
 <w:sz w:val="28"/>
 </w:rPr>
 </w:pPr>
 <w:r>
 <w:rPr>
 <w:color w:val="70AD47" w:themeColor="accent6"/>
 <w:sz w:val="28"/>
 </w:rPr>
 <w:t>This text has formatting directly applied to achieve its
font size, color, line spacing, and paragraph spacing.</w:t>
 </w:r>
 <w:bookmarkStart w:id="0" w:name="_GoBack"/>
 <w:bookmarkEnd w:id="0"/>
 </w:p>
 <w:p/>
 <w:sectPr>
 <w:pgSz w:w="12240" w:h="15840"/>
 <w:pgMar w:top="1440" w:right="1440" w:bottom="1440"
w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
 <w:cols w:space="720"/>
 </w:sectPr>
```


 </w:body> </w:document> </pkg:xmlData> </pkg:part>

Since document.xml is the primary document part where you place your content, take a quick walk through that part. (Figure 14, which follows this list, provides a visual reference to show how some of the core content and formatting tags explained here relate to what you see in a Word document.)

- The opening **w:document** tag includes several namespace ( **xmlns** ) listings. Many of those namespaces refer to specific types of content and you only need them if they're relevant to your content.
Notice that the prefix for the tags throughout a document part refers back to the namespaces. In this example, the only prefix used in the tags throughout the document.xml part is **w:**, so the only namespace that you need to leave in the opening **w:document** tag is **xmlns:w**.

#### **Tip**

If you're editing your markup in Visual Studio, after you delete namespaces in any part, look through all tags of that part. If you've removed a namespace that's required for your markup, you'll see a red squiggly underline on the relevant prefix for affected tags. If you remove the **xmlns:mc** namespace, you must also remove the **mc:Ignorable** attribute that precedes the namespace listings.

- Inside the opening body tag, you see a paragraph tag ( **w:p** ), which includes our sample content for this example.
- The **w:pPr** tag includes properties for directly-applied paragraph formatting, such as space before or after the paragraph, paragraph alignment, or indents. (Direct formatting refers to attributes that you apply individually to content rather than as part of a style.) This tag also includes direct font formatting that's applied to the entire paragraph, in a nested **w:rPr** (run properties) tag, which contains the font color and size set in our sample.

You might notice that font sizes and some other formatting settings in Word Office Open XML markup look like they're double the actual size. That's because paragraph and line spacing, as well some section formatting properties shown in the preceding markup, are specified in twips (one-twentieth of a point). Depending on the types of content you work with in Office Open XML, you may see several


additional units of measure, including English Metric Units (914,400 EMUs to an inch), which are used for some Office Art (drawingML) values and 100,000 times actual value, which is used in both drawingML and PowerPoint markup. PowerPoint also expresses some values as 100 times actual and Excel commonly uses actual values.

- Within a paragraph, any content with like properties is included in a run ( **w:r** ), such as is the case with the sample text. Each time there's a change in formatting or content type, a new run starts. (That is, if just one word in the sample text was bold, it would be separated into its own run.) In this example, the content includes just the one text run.
Notice that, because the formatting included in this sample is font formatting (that is, formatting that can be applied to as little as one character), it also appears in the properties for the individual run.

- Also notice the tags for the hidden "_GoBack" bookmark (**w:bookmarkStart** and **w:bookmarkEnd** ), which appear in Word documents by default. You can always delete the start and end tags for the GoBack bookmark from your markup.
- The last piece of the document body is the **w:sectPr** tag, or section properties. This tag includes settings such as margins and page orientation. The content you insert using **setSelectedDataAsync** will take on the active section properties in the destination document by default. So, unless your content includes a section break (in which case you'll see more than one **w:sectPr** tag), you can delete this tag.

*Figure 14. How common tags in document.xml relate to the content and layout of a Word document*


In markup you create, you might see another attribute in several tags that includes the characters **w:rsid**, which you don't see in the examples used in this topic. These are revision identifiers. They're used in Word for the Combine Documents feature and they're on by default. You'll never need them in markup you're inserting with your add-in and turning them off makes for much cleaner markup. You can easily remove existing RSID tags or disable the feature (as described in the following procedure) so that they aren't added to your markup for new content.

Be aware that if you use the co-authoring capabilities in Word (such as the ability to simultaneously edit documents with others), you should enable the feature again when finished generating the markup for your add-in.

To turn off RSID attributes in Word for documents you create going forward, do the following:

- 1. In Word, choose **File** and then choose **Options**.
- 2. In the Word Options dialog box, choose **Trust Center** and then choose **Trust Center Settings**.
- 3. In the Trust Center dialog box, choose **Privacy Options** and then disable the setting **Store random numbers to improve Combine accuracy**. *Note that this setting may not be available in newer versions of Word.*

To remove RSID tags from an existing document, try the following shortcut with the document open in Office Open XML.

- 1. With your insertion point in the main body of the document, press Ctrl + Home to go to the top of the document.
- 2. On the keyboard, press Space , Delete , Space . Then, save the document.

After removing the majority of the markup from this package, you're left with the minimal markup that needs to be inserted for the sample, as shown in the preceding section.

## **Use the same Office Open XML structure for different content types**

Several types of rich content require only the .rels and document.xml components shown in the preceding example, including content controls, Office drawing shapes and text boxes, and tables (unless a style is applied to the table). In fact, you can reuse the same edited package parts and swap out just the **body** content in document.xml for the markup of your content.


To check out the Office Open XML markup for the examples of each of these content types shown earlier in Figures 5 through 8, explore the [Load and write Open XML in](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/word-add-in-load-and-write-open-xml) [your Word add-in](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/word-add-in-load-and-write-open-xml) code sample referenced in the overview section.

Before you move on, take a look at differences to note for a couple of these content types and how to swap out the pieces you need.

### **Understand drawingML markup (Office graphics) in Word**

If the markup for your shape or text box looks far more complex than you would expect, there's a reason for it. With the release of Office 2007, we saw the introduction of the Office Open XML Formats as well as the introduction of a new Office graphics engine that PowerPoint and Excel fully adopted. In the 2007 release, Word only incorporated part of that graphics engine, adopting the updated Excel charting engine, SmartArt graphics, and advanced picture tools. For shapes and text boxes, Word 2007 continued to use legacy drawing objects (VML). It was in the 2010 release that Word took the additional steps with the graphics engine to incorporate updated shapes and drawing tools.

Typically, as you see for the shape and text box examples included in the [Load and write](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/word-add-in-load-and-write-open-xml) [Open XML in your Word add-in](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/word-add-in-load-and-write-open-xml) code sample, the fallback markup can be removed. Word automatically adds missing fallback markup to shapes when a document is saved. However, if you prefer to keep the fallback markup to ensure that you're supporting all user scenarios, there's no harm in retaining it.

If you have grouped drawing objects included in your content, you'll see additional (and apparently repetitive) markup, but this must be retained. Portions of the markup for drawing shapes are duplicated when the object is included in a group.

#### ) **Important**

When working with text boxes and drawing shapes, be sure to check namespaces carefully before removing them from document.xml. (Or, if you're reusing markup from another object type, be sure to add back any required namespaces you might have previously removed from document.xml.) A substantial portion of the namespaces included by default in document.xml are there for drawing object requirements.

### **About graphic positioning**


In the code samples [Load and write Open XML in your Word add-in](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/word-add-in-load-and-write-open-xml) and [Word-Add](https://github.com/OfficeDev/Word-Add-in-Get-Set-EditOpen-XML)[in-Get-Set-EditOpen-XML](https://github.com/OfficeDev/Word-Add-in-Get-Set-EditOpen-XML) , the text box and shape are set up using different types of text wrapping and positioning settings. (Also be aware that the image examples in those code samples are set up using in line with text formatting, which positions a graphic object on the text baseline.)

The shape in those code samples is positioned relative to the right and bottom page margins. Relative positioning lets you more easily coordinate with a user's unknown document setup because it will adjust to the user's margins and run less risk of looking awkward because of paper size, orientation, or margin settings. To retain relative positioning settings when you insert a graphic object, you must retain the paragraph mark (w:p) in which the positioning (known in Word as an anchor) is stored. If you insert the content into an existing paragraph mark rather than including your own, you may be able to retain the same initial visual, but many types of relative references that enable the positioning to automatically adjust to the user's layout may be lost.

### **Work with content controls**

Content controls are an important feature in Word that can greatly enhance the power of your add-in for Word in multiple ways, including giving you the ability to insert content at designated places in the document rather than only at the selection.

In Word, find content controls on the Developer tab of the ribbon, as shown here in Figure 15.

*Figure 15. The Controls group on the Developer tab in Word*

|                      | Aa Aa 금 覽   盧   盧   盧   L Design Mode |  |  |
|----------------------|---------------------------------------|--|--|
| 团 雪 雪   三 Properties |                                       |  |  |
| 물 홈 -                | Group ▼                               |  |  |
| Controls             |                                       |  |  |

Types of content controls in Word include rich text, plain text, picture, building block gallery, check box, dropdown list, combo box, date picker, and repeating section.

- Use the **Properties** command, shown in Figure 15, to edit the title of the control and to set preferences such as hiding the control container.
- Enable **Design Mode** to edit placeholder content in the control.

If your add-in works with a Word template, you can include controls in that template to enhance the behavior of the content. You can also use XML data binding in a Word document to bind content controls to data, such as document properties, for easy form 


completion or similar tasks. (Find controls that are already bound to built-in document properties in Word on the **Insert** tab, under **Quick Parts**.)

When you use content controls with your add-in, you can also greatly expand the options for what your add-in can do using a different type of binding. You can bind to a content control from within the add-in and then write content to the binding rather than to the active selection.

Don't confuse XML data binding in Word with the ability to bind to a control via your add-in. These are completely separate features. However, you can include named content controls in the content you insert via your add-in using OOXML coercion and then use code in the add-in to bind to those controls.

Also be aware that both XML data binding and Office.js can interact with custom XML parts in your app, so it's possible to integrate these powerful tools. To learn about working with custom XML parts in the Office JavaScript API, see the See also section of this topic.

Working with bindings in your Word add-in is covered in the next section of this topic. First, take a look at an example of the Office Open XML required for inserting a rich text content control that you can bind to using your add-in.

#### ) **Important**

Rich text controls are the only type of content control you can use to bind to a content control from within your add-in.

| XML                                                                                                  |
|------------------------------------------------------------------------------------------------------|
|                                                                                                      |
| <pkg:package<br>xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"&gt;</pkg:package<br> |
| <pkg:part <="" pkg:name="/_rels/.rels" td=""></pkg:part>                                             |
| pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"                           |
| pkg:padding="512">                                                                                   |
| <pkg:xmldata></pkg:xmldata>                                                                          |
| <relationships< td=""></relationships<>                                                              |
| xmlns="http://schemas.openxmlformats.org/package/2006/relationships">                                |
| <relationship <="" id="rId1" td=""></relationship>                                                   |
| Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/of                         |
| ficeDocument" Target="word/document.xml"/>                                                           |
|                                                                                                      |
|                                                                                                      |
|                                                                                                      |
| <pkg:part <="" pkg:name="/word/document.xml" td=""></pkg:part>                                       |
| pkg:contentType="application/vnd.openxmlformats                                                      |
| officedocument.wordprocessingml.document.main+xml">                                                  |
| <pkg:xmldata></pkg:xmldata>                                                                          |


```
 <w:document
xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" >
 <w:body>
 <w:p/>
 <w:sdt>
 <w:sdtPr>
 <w:alias w:val="MyContentControlTitle"/>
 <w:id w:val="1382295294"/>
 <w15:appearance w15:val="hidden"/>
 <w:showingPlcHdr/>
 </w:sdtPr>
 <w:sdtContent>
 <w:p>
 <w:r>
 <w:t>[This text is inside a content control that has its
container hidden. You can bind to a content control to add or interact with
content at a specified location in the document.]</w:t>
 </w:r>
 </w:p>
 </w:sdtContent>
 </w:sdt>
 </w:body>
 </w:document>
 </pkg:xmlData>
 </pkg:part>
</pkg:package>
```
As already mentioned, content controls, like formatted text, don't require additional document parts, so only edited versions of the .rels and document.xml parts are included here.

The **w:sdt** tag that you see within the document.xml body represents the content control. If you generate the Office Open XML markup for a content control, you'll see that several attributes have been removed from this example, including the tag and document part properties. Only essential (and a couple of best practice) elements have been retained, including the following:

- The **alias** is the title property from the Content Control Properties dialog box in Word. This is a required property (representing the name of the item) if you plan to bind to the control from within your add-in.
- The unique **id** is a required property. If you bind to the control from within your add-in, the ID is the property the binding uses in the document to identify the applicable named content control.
- The **appearance** attribute is used to hide the control container, for a cleaner look. This feature was introduced in Word 2013, as you see by the use of the w15


namespace. Because this property is used, the w15 namespace is retained at the start of the document.xml part.

- The **showingPlcHdr** attribute is an optional setting that sets the default content you include inside the control (text in this example) as placeholder content. So, if the user clicks or taps in the control area, the entire content is selected rather than behaving like editable content in which the user can make changes.
- Although the empty paragraph mark ( **w:p/** ) that precedes the **sdt** tag isn't required for adding a content control (and will add vertical space above the control in the Word document), it ensures that the control is placed in its own paragraph. This may be important, depending upon the type and formatting of content that will be added in the control.
- If you intend to bind to the control, the default content for the control (what's inside the **sdtContent** tag) must include at least one complete paragraph (as in this example), in order for your binding to accept multi-paragraph rich content.

The document part attribute that was removed from this sample **w:sdt** tag may appear in a content control to reference a separate part in the package where placeholder content information can be stored (parts located in a glossary directory in the Office Open XML package). Although document part is the term used for XML parts (that is, files) within an Office Open XML package, the term document parts as used in the sdt property refers to the same term in Word that's used to describe some content types including building blocks and document property quick parts (for example, built-in XML data-bound controls). If you see parts under a glossary directory in your Office Open XML package, you may need to retain them if the content you're inserting includes these features. For a typical content control that you intend to use to bind to from your add-in, they aren't required. Just remember that, if you do delete the glossary parts from the package, you must also remove the document part attribute from the w:sdt tag.

The next section will discuss how to create and use bindings in your Word add-in.

### **Insert content at a designated location**

You've already looked at how to insert content at the active selection in a Word document. If you bind to a named content control that's in the document, you can insert any of the same content types into that control.

So when might you want to use this approach?

- When you need to add or replace content at specified locations in a template, such as to populate portions of the document from a database.


- When you want the option to replace content that you're inserting at the active selection, such as to provide design element options to the user.
- When you want the user to add data in the document that you can access for use with your add-in, such as to populate fields in the task pane based upon information the user adds in the document.

Download the code sample [Word-Add-in-JavaScript-AddPopulateBindings](https://github.com/OfficeDev/Word-Add-in-JavaScript-AddPopulateBindings) , which provides a working example of how to insert and bind to a content control, and how to populate the binding.

### **Add and bind to a named content control**

As you examine the JavaScript that follows, consider these requirements:

- As previously mentioned, you must use a rich text content control in order to bind to the control from your Word add-in.
- The content control must have a name (this is the **Title** field in the Content Control Properties dialog box, which corresponds to the **Alias** tag in the Office Open XML markup). This is how the code identifies where to place the binding.
- You can have several named controls and bind to them as needed. Use a unique content control name, unique content control ID, and a unique binding ID.

```
JavaScript
function addAndBindControl() {

Office.context.document.bindings.addFromNamedItemAsync("MyContentControlTitl
e", "text", { id: 'myBinding' }, function (result) {
 if (result.status == "failed") {
 if (result.error.message == "The named item does not exist.")
 const myOOXMLRequest = new XMLHttpRequest();
 let myXML;
 myOOXMLRequest.open('GET', 
'../../Snippets_BindAndPopulate/ContentControl.xml', false);
 myOOXMLRequest.send();
 if (myOOXMLRequest.status === 200) {
 myXML = myOOXMLRequest.responseText;
 }
 Office.context.document.setSelectedDataAsync(myXML, { 
coercionType: 'ooxml' }, function (result) {

Office.context.document.bindings.addFromNamedItemAsync("MyContentControlTitl
e", "text", { id: 'myBinding' });
 });
 }
```


}

The code shown here takes the following steps.

- Attempts to bind to the named content control, using [addFromNamedItemAsync](https://learn.microsoft.com/en-us/javascript/api/office/office.bindings#office-office-bindings-addfromnameditemasync-member(1)).
Take this step first if there's a possible scenario for your add-in where the named control could already exist in the document when the code runs. For example, you'll want to do this if the add-in was inserted into and saved with a template that's been designed to work with the add-in, where the control was placed in advance. You also need to do this if you need to bind to a control that was placed earlier by the add-in.

- The callback in the first call to the addFromNamedItemAsync method checks the status of the result to see if the binding failed because the named item doesn't exist in the document (that is, the content control named MyContentControlTitle in this example). If so, the code adds the control at the active selection point (using setSelectedDataAsync ) and then binds to it.
As mentioned earlier and shown in the preceding code, the name of the content control is used to determine where to create the binding. However, in the Office Open XML markup, the code adds the binding to the document using both the name and the ID attribute of the content control.

After running code, if you examine the markup of the document in which your add-in created bindings, you'll see two parts to each binding. In the markup for the content control where a binding was added (in document.xml), you'll see the attribute **w15:webExtensionLinked/**.

In the document part named webExtensions1.xml, you'll see a list of the bindings you've created. Each is identified using the binding ID and the ID attribute of the applicable control, such as the following, where the **appref** attribute is the content control ID: **we:binding id="myBinding" type="text" appref="1382295294"/**.

#### ) **Important**

You must add the binding at the time you intend to act upon it. Don't include the markup for the binding in the Office Open XML for inserting the content control because the process of inserting that markup will strip the binding.

### **Populate a binding**


The code for writing content to a binding is similar to that for writing content to a selection.

```
JavaScript
function populateBinding(filename) {
 const myOOXMLRequest = new XMLHttpRequest();
 let myXML;
 myOOXMLRequest.open('GET', filename, false);
 myOOXMLRequest.send();
 if (myOOXMLRequest.status === 200) {
 myXML = myOOXMLRequest.responseText;
 }
 Office.select("bindings#myBinding").setDataAsync(myXML, { coercionType: 
'ooxml' });
}
```
As with setSelectedDataAsync , you specify the content to be inserted and the coercion type. The only additional requirement for writing to a binding is to identify the binding by ID. Notice how the binding ID used in this code (bindings#myBinding) corresponds to the binding ID established (myBinding) when the binding was created in the previous function.

### **Tip**

The preceding code is all you need whether you are initially populating or replacing the content in a binding. When you insert a new piece of content at a bound location, the existing content in that binding is automatically replaced. Check out an example of this in the previously-referenced code sample **[Word-Add-in-](https://github.com/OfficeDev/Word-Add-in-JavaScript-AddPopulateBindings)[JavaScript-AddPopulateBindings](https://github.com/OfficeDev/Word-Add-in-JavaScript-AddPopulateBindings)** , which provides two separate content samples that you can use interchangeably to populate the same binding.

## **Add objects that use additional Office Open XML parts**

Many types of content require additional document parts in the Office Open XML package, meaning that they either reference information in another part or the content itself is stored in one or more additional parts and referenced in document.xml.

For example, consider the following:

- Content that uses styles for formatting (such as the styled text shown earlier in Figure 2 or the styled table shown in Figure 9) requires the styles.xml part.


- Images (such as those shown in Figures 3 and 4) include the binary image data in one (and sometimes two) additional parts.
- SmartArt diagrams (such as the one shown in Figure 10) require multiple additional parts to describe the layout and content.
- Charts (such as the one shown in Figure 11) require multiple additional parts, including their own relationship (.rels) part.

You can see edited examples of the markup for all of these content types in the previously-referenced code sample [Load and write Open XML in your Word add-in](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/word-add-in-load-and-write-open-xml) . You can insert all of these content types using the same JavaScript code shown earlier (and provided in the referenced code samples) for inserting content at the active selection and writing content to a specified location using bindings.

Remember, if you're retaining any additional parts referenced in document.xml, you will need to retain document.xml.rels and the relationship definitions for the applicable parts you're keeping, such as styles.xml or an image file.

Before you explore the samples, take a look at a few tips for working with each of these content types.

### **Working with styles**

The same approach to editing the markup that you looked at for the preceding example with directly-formatted text applies when using paragraph styles or table styles to format your content. However, the markup for working with paragraph styles is considerably simpler, so that's the example described here.

#### **Editing the markup for content using paragraph styles**

The following markup represents the body content for the styled text example shown in Figure 2.

```
XML
<w:body>
 <w:p>
 <w:pPr>
 <w:pStyle w:val="Heading1"/>
 </w:pPr>
 <w:r>
 <w:t>This text is formatted using the Heading 1 paragraph style.</w:t>
 </w:r>
```


As you see, the markup for formatted text in document.xml is considerably simpler when you use a style, because the style contains all of the paragraph and font formatting that you otherwise need to reference individually. However, as explained earlier, you might want to use styles or direct formatting for different purposes: use direct formatting to specify the appearance of your text regardless of the formatting in the user's document; use a paragraph style (particularly a built-in paragraph style name, such as Heading 1 shown here) to have the text formatting automatically coordinate with the user's document.

Use of a style is a good example of how important it is to read and understand the markup for the content you're inserting, because it isn't explicit that another document part is referenced here. If you include the style definition in this markup and don't include the styles.xml part, the style information in document.xml will be ignored regardless of whether or not that style is in use in the user's document.

However, if you take a look at the styles.xml part, you'll see that only a small portion of this long piece of markup is required when editing markup for use in your add-in:

- The styles.xml part includes several namespaces by default. If you are only retaining the required style information for your content, in most cases you only need to keep the **xmlns:w** namespace.
- The **w:docDefaults** tag content that falls at the top of the styles part will be ignored when your markup is inserted via the add-in and can be removed.
- The largest piece of markup in a styles.xml part is for the **w:latentStyles** tag that appears after docDefaults, which provides information (such as appearance attributes for the Styles pane and Styles gallery) for every available style. This information is also ignored when inserting content via your add-in and so it can be removed.
- Following the latent styles information, you see a definition for each style in use in the document from which you're markup was generated. This includes some default styles that are in use when you create a new document and may not be relevant to your content. You can delete the definitions for any styles that aren't used by your content.

Each built-in heading style has an associated Char style that's a character style version of the same heading format. Unless you've applied the heading style as a character style, you can remove it. If the style is used as a character style, it appears 


in document.xml in a run properties tag ( **w:rPr** ) rather than a paragraph properties ( **w:pPr** ) tag. This should only be the case if you've applied the style to just part of a paragraph, but it can occur inadvertently if the style was incorrectly applied.

- If you're using a built-in style for your content, you don't have to include a full definition. You only must include the style name, style ID, and at least one formatting attribute in order for the coerced Office Open XML to apply the style to your content upon insertion.
However, it's a best practice to include a complete style definition (even if it's the default for built-in styles). If a style is already in use in the destination document, your content will take on the resident definition for the style, regardless of what you include in styles.xml. If the style isn't yet in use in the destination document, your content will use the style definition you provide in the markup.

So, for example, the only content you needed to retain from the styles.xml part for the sample text shown in Figure 2, which is formatted using Heading 1 style, is the following:

#### 7 **Note**

A complete Word definition for the Heading 1 style has been retained in this example.

XML

```
<pkg:part pkg:name="/word/styles.xml"
pkg:contentType="application/vnd.openxmlformats-
officedocument.wordprocessingml.styles+xml">
 <pkg:xmlData>
 <w:styles
xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" >
 <w:style w:type="paragraph" w:styleId="Heading1">
 <w:name w:val="heading 1"/>
 <w:basedOn w:val="Normal"/>
 <w:next w:val="Normal"/>
 <w:link w:val="Heading1Char"/>
 <w:uiPriority w:val="9"/>
 <w:qFormat/>
 <w:pPr>
 <w:keepNext/>
 <w:keepLines/>
 <w:spacing w:before="240" w:after="0" w:line="259"
w:lineRule="auto"/>
 <w:outlineLvl w:val="0"/>
 </w:pPr>
```


```
 <w:rPr>
 <w:rFonts w:asciiTheme="majorHAnsi"
w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi"
w:cstheme="majorBidi"/>
 <w:color w:val="2E74B5" w:themeColor="accent1" w:themeShade="BF"/>
 <w:sz w:val="32"/>
 <w:szCs w:val="32"/>
 </w:rPr>
 </w:style>
 </w:styles>
 </pkg:xmlData>
</pkg:part>
```
#### **Edit the markup for content using table styles**

When your content uses a table style, you need the same relative part of styles.xml as described for working with paragraph styles. That is, you only need to retain the information for the style you're using in your content, and you must include the name, ID, and at least one formatting attribute, but are better off including a complete style definition to address all potential user scenarios.

However, when you look at the markup both for your table in document.xml and for your table style definition in styles.xml, you see enormously more markup than when working with paragraph styles.

- In document.xml, formatting is applied by cell even if it's included in a style. Using a table style won't reduce the volume of markup. The benefit of using table styles for the content is for easy updating and easily coordinating the look of multiple tables.
- In styles.xml, you'll see a substantial amount of markup for a single table style as well, because table styles include several types of possible formatting attributes for each of several table areas, such as the entire table, heading rows, odd and even banded rows and columns (separately), the first column, etc.

### **Work with images**

The markup for an image includes a reference to at least one part that includes the binary data to describe your image. For a complex image, this can be hundreds of pages of markup and you can't edit it. Since you don't ever have to touch the binary parts, you can simply collapse it if you're using a structured editor such as Visual Studio, so that you can still easily review and edit the rest of the package.


If you check out the example markup for the simple image shown earlier in Figure 3, available in the previously-referenced code sample [Load and write Open XML in your](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/word-add-in-load-and-write-open-xml) [Word add-in](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/word-add-in-load-and-write-open-xml) , you'll see that the markup for the image in document.xml includes size and position information as well as a relationship reference to the part that contains the binary image data. That reference is included in the **a:blip** tag, as follows:

XML

#### <a:blip r:embed="rId4" cstate="print">

Be aware that, because a relationship reference is explicitly used ( **r:embed="rID4"** ) and that related part is required in order to render the image, if you don't include the binary data in your Office Open XML package, you'll get an error. This is different from styles.xml, explained previously, which won't throw an error if omitted since the relationship isn't explicitly referenced and the relationship is to a part that provides attributes to the content (formatting) rather than being part of the content itself.

When you review the markup, notice the additional namespaces used in the a:blip tag. You'll see in document.xml that the **xmlns:a** namespace (the main drawingML namespace) is dynamically placed at the beginning of the use of drawingML references rather than at the top of the document.xml part. However, the relationships namespace (r) must be retained where it appears at the start of document.xml. Check your picture markup for additional namespace requirements. Remember that you don't have to memorize which types of content require what namespaces, you can easily tell by reviewing the prefixes of the tags throughout document.xml.

### **Understanding additional image parts and formatting**

When you use some Office picture formatting effects on your image, such as for the image shown in Figure 4, which uses adjusted brightness and contrast settings (in addition to picture styling), a second binary data part for an HD format copy of the image data may be required. This additional HD format is required for formatting considered a layering effect, and the reference to it appears in document.xml similar to the following:

XML

#### <a14:imgLayer r:embed="rId5">

See the required markup for the formatted image shown in Figure 4 (which uses layering effects among others) in the [Load and write Open XML in your Word add-in](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/word-add-in-load-and-write-open-xml) code sample.


### **Work with SmartArt diagrams**

A SmartArt diagram has four associated parts, but only two are always required. You can examine an example of SmartArt markup in the [Load and write Open XML in your Word](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/word-add-in-load-and-write-open-xml) [add-in](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/word-add-in-load-and-write-open-xml) code sample. First, take a look at a brief description of each of the parts and why they are or aren't required:

#### **Tip**

If your content includes more than one diagram, they will be numbered consecutively, replacing the '1' in the file names listed here.

- layout1.xml: This part is required. It includes the markup definition for the layout appearance and functionality.
- data1.xml: This part is required. It includes the data in use in your instance of the diagram.
- drawing1.xml: This part isn't always required but if you apply custom formatting to elements in your instance of a diagram, such as directly formatting individual shapes, you might need to retain it.
- colors1.xml: This part isn't required. It includes color style information, but the colors of your diagram will coordinate by default with the colors of the active formatting theme in the destination document, based on the SmartArt color style you apply from the SmartArt Tools design tab in Word before saving out your Office Open XML markup.
- quickStyles1.xml: This part isn't required. Similar to the colors part, you can remove this as your diagram will take on the definition of the applied SmartArt style that's available in the destination document (that is, it will automatically coordinate with the formatting theme in the destination document).

The SmartArt layout1.xml file is a good example of places you may be able to further trim your markup but mightn't be worth the extra time to do so (because it removes such a small amount of markup relative to the entire package). If you would like to get rid of every last line you can of markup, you can delete the **dgm:sampData** tag and its contents. This sample data defines how the thumbnail preview for the diagram will appear in the SmartArt styles galleries. However, if it's omitted, default sample data is used.

Be aware that the markup for a SmartArt diagram in document.xml contains relationship ID references to the layout, data, colors, and quick styles parts. You can delete the


references in document.xml to the colors and styles parts when you delete those parts and their relationship definitions (and it's certainly a best practice to do so, since you're deleting those relationships), but you won't get an error if you leave them, since they aren't required for your diagram to be inserted into a document. Find these references in document.xml in the **dgm:relIds** tag. Regardless of whether or not you take this step, retain the relationship ID references for the required layout and data parts.

### **Work with charts**

Similar to SmartArt diagrams, charts contain several additional parts. However, the setup for charts is a bit different from SmartArt, in that a chart has its own relationship file. Following is a description of required and removable document parts for a chart.

#### **Tip**

As with SmartArt diagrams, if your content includes more than one chart, they will be numbered consecutively, replacing the '1' in the file names listed here.

- In document.xml.rels, you'll see a reference to the required part that contains the data that describes the chart (chart1.xml).
- You also see a separate relationship file for each chart in your Office Open XML package, such as chart1.xml.rels.

There are three files referenced in chart1.xml.rels, but only one is required. These include the binary Excel workbook data (required) and the color and style parts (colors1.xml and styles1.xml) that you can remove.

Charts that you can create and edit natively in Word are Excel charts, and their data is maintained on an Excel worksheet that's embedded as binary data in your Office Open XML package. Like the binary data parts for images, this Excel binary data is required, but there's nothing to edit in this part. So you can just collapse the part in the editor to avoid having to manually scroll through it all to examine the rest of your Office Open XML package.

However, similar to SmartArt, you can delete the colors and styles parts. If you've used the chart styles and color styles available in to format your chart, the chart will take on the applicable formatting automatically when it's inserted into the destination document.

See the edited markup for the example chart shown in Figure 11 in the [Load and write](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/word-add-in-load-and-write-open-xml) [Open XML in your Word add-in](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/word-add-in-load-and-write-open-xml) code sample.


## **Edit the Office Open XML for use in your task pane add-in**

You've already seen how to identify and edit the content in your markup. If the task still seems difficult when you take a look at the massive Office Open XML package generated for your document, following is a quick summary of recommended steps to help you edit that package down quickly.

Remember that you can use all .rels parts in the package as a map to quickly check for document parts that you can remove.

- 1. Open the flattened XML file in Visual Studio and press Ctrl + K , Ctrl + D to format the file. Then use the collapse/expand buttons on the left to collapse the parts you know you need to remove. You might also want to collapse long parts you need, but know you won't need to edit (such as the base64 binary data for an image file), making the markup faster and easier to visually scan.
- 2. There are several parts of the document package that you can almost always remove when you are preparing Office Open XML markup for use in your add-in. You might want to start by removing these (and their associated relationship definitions), which will greatly reduce the package right away. These include the theme1, fontTable, settings, webSettings, thumbnail, both the core and add-in properties files, and any taskpane or webExtension parts.
- 3. Remove any parts that don't relate to your content, such as footnotes, headers, or footers that you don't require. Again, remember to also delete their associated relationships.
- 4. Review the document.xml.rels part to see if any files referenced in that part are required for your content, such as an image file, the styles part, or SmartArt diagram parts. Delete the relationships for any parts your content doesn't require and confirm that you have also deleted the associated part. If your content doesn't require any of the document parts referenced in document.xml.rels, you can delete that file also.
- 5. If your content has an additional .rels part (such as chart#.xml.rels), review it to see if there are other parts referenced there that you can remove (such as quick styles for charts) and delete both the relationship from that file as well as the associated part.
- 6. Edit document.xml to remove namespaces not referenced in the part, section properties if your content doesn't include a section break, and any markup that


isn't related to the content that you want to insert. If inserting shapes or text boxes, you might also want to remove extensive fallback markup.

- 7. Edit any additional required parts where you know that you can remove substantial markup without affecting your content, such as the styles part.
After you've taken the preceding seven steps, you've likely cut between about 90 and 100 percent of the markup you can remove, depending on your content. In most cases, this is likely to be as far as you want to trim.

Regardless of whether you leave it here or choose to delve further into your content to find every last line of markup you can cut, remember that you can use the previouslyreferenced code sample [Word-Add-in-Get-Set-EditOpen-XML](https://github.com/OfficeDev/Word-Add-in-Get-Set-EditOpen-XML) as a scratch pad to quickly and easily test your edited markup.

#### **Tip**

If you update an Office Open XML snippet in an existing solution while developing, clear temporary Internet files before you run the solution again to update the Office Open XML used by your code. Markup that's included in your solution in XML files is cached on your computer. You can, of course, clear temporary Internet files from your default web browser. To access Internet options and delete these settings from inside Visual Studio 2019, on the **Debug** menu, choose **Options**. Then, under **Environment**, choose **Web Browser** and then choose **Internet Explorer Options**.

## **Create an add-in for both template and standalone use**

In this topic, you've seen several examples of what you can do with Office Open XML in your add-ins. You've looked at a wide range of rich content type examples that you can insert into documents by using the Office Open XML coercion type, together with the JavaScript methods for inserting that content at the selection or to a specified (bound) location.

So, what else do you need to know if you're creating your add-in both for standalone use (that is, inserted from the Store or a proprietary server location) and for use in a precreated template that's designed to work with your add-in? The answer might be that you already know all you need.


The markup for a given content type and methods for inserting it are the same whether your add-in is designed to standalone or work with a template. If you are using templates designed to work with your add-in, just be sure that your JavaScript includes callbacks that account for scenarios where referenced content might already exist in the document (such as demonstrated in the binding example shown in the section Add and bind to a named content control).

When using templates with your app, whether the add-in will be resident in the template at the time that the user created the document or the add-in will be inserting a template, you might also want to incorporate other elements of the API to help you create a more robust, interactive experience. For example, you may want to include identifying data in a customXML part that you can use to determine the template type in order to provide template-specific options to the user. To learn more about how to work with custom XML in your add-ins, see the additional resources that follow.

### **See also**

- Office JavaScript API
- The complete language reference and related documentation on Open XML: [Standard ECMA-376: Office Open XML File Formats](https://www.ecma-international.org/publications-and-standards/standards/ecma-376/)
- Explore Office JavaScript API using Script Lab
- [Exploring the Office JavaScript API: Data Binding and Custom XML Parts](https://learn.microsoft.com/en-us/archive/msdn-magazine/2013/april/microsoft-office-exploring-the-javascript-api-for-office-data-binding-and-custom-xml-parts)
- Companion code sample: [Load and write Open XML in your Word add-in](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/word-add-in-load-and-write-open-xml)
- Other code samples referenced in this article:
	- [Word-Add-in-Get-Set-EditOpen-XML](https://github.com/OfficeDev/Word-Add-in-Get-Set-EditOpen-XML)
	- [Word-Add-in-JavaScript-AddPopulateBindings](https://github.com/OfficeDev/Word-Add-in-JavaScript-AddPopulateBindings)


# **Troubleshoot Word add-ins**

08/06/2025

This article discusses troubleshooting issues that are unique to Word. Use the feedback tool at the end of the page to suggest other issues that can be added to this article.

## **All selected ranges aren't recognized**

If noncontiguous selections are made, the Word API only operates on the last contiguous range in the selection. An unexpected case of this is when you select a column in a table then call, for example, [Document.getSelection](https://learn.microsoft.com/en-us/javascript/api/word/word.document#word-word-document-getselection-member(1)), only the final cell in the selection is returned by the API. Although the selection of a column seems contiguous, the API recognizes it as a noncontiguous selection (e.g., a cell per row).

To learn more generally about making noncontiguous selections, see [How to select items that](https://support.microsoft.com/topic/8b9c1be9-cca3-935a-7cbf-94403aa48d2e) [are not next to each other](https://support.microsoft.com/topic/8b9c1be9-cca3-935a-7cbf-94403aa48d2e) .

## **Annotations don't work**

If the annotation APIs aren't working, it may be because you're not using a Microsoft 365 subscription. If you're using a one-time purchase license, this could be why these APIs aren't working for you.

The annotation APIs rely on a service that requires a Microsoft 365 subscription. Therefore, verify that you're running the add-in in Word connected to a Microsoft 365 subscription license before debugging further.

For more about this problem, see [GitHub issue 4953](https://github.com/OfficeDev/office-js/issues/4953) .

## **Body.insertFileFromBase64 doesn't insert header or footer**

It's by design that the [Body.insertFileFromBase64](https://learn.microsoft.com/en-us/javascript/api/word/word.body#word-word-body-insertfilefrombase64-member(1)) method excludes any header or footer that was in the source file.

To include any headers or footers from the source file, use [Document.insertFileFromBase64](https://learn.microsoft.com/en-us/javascript/api/word/word.document#word-word-document-insertfilefrombase64-member(1)) instead.

## **Can't use Mixed to set a property**


Several enums in Word offer "Mixed" as a valid value. However, the value can primarily be returned when a get a property or make a get* API call. This is because "Mixed" means that several options are applied to the current selection. If you try to set the option to "Mixed", then it isn't clear which actual value should be applied to the selection.

For example, let's say you're working with the borders around a section of text. Each [border](https://learn.microsoft.com/en-us/javascript/api/word/word.border#word-word-border-width-member) can be set to a different [width.](https://learn.microsoft.com/en-us/javascript/api/word/word.borderwidth) If the top border is "Pt025" (that is, 0.25 points), the bottom border is "None", and the left and right borders are "Pt050" (that is, 0.50 points), then when you get the width of the borders, "Mixed" is returned. If you want to change the width of the borders, call the set API on each border using an enum value other than mixed .

This behavior also applies for enum values like "Unknown".

# **Get a GeneralException when working with styles**

If users are hitting a GeneralException when your add-in calls [Document.insertFileFromBase64](https://learn.microsoft.com/en-us/javascript/api/word/word.document#word-word-document-insertfilefrombase64-member(1)) or Style APIs, it may be that those users are exceeding limits imposed by the Word application. To learn more about these limits, see [Operating parameter limitations and specifications in](https://learn.microsoft.com/en-us/office/troubleshoot/word/operating-parameter-limitation) [Word.](https://learn.microsoft.com/en-us/office/troubleshoot/word/operating-parameter-limitation)

## **Layout breaks when using insertHtml while cursor is in content control in header**

This issue may occur when the following three conditions are met.

- 1. Have at least one content control in the header and at least one in the footer of the Word document.
- 2. Ensure the cursor is inside a content control in the header.
- 3. Call [insertHtml](https://learn.microsoft.com/en-us/javascript/api/word/word.contentcontrol#word-word-contentcontrol-inserthtml-member(1)) to set a content control in the footer.

The footer is then unexpectedly mixed with the header. To avoid this, clear the content control in the footer before setting it, as shown in the following code sample.

```
TypeScript
await Word.run(async (context) => {
 // Credit to https://github.com/barisbikmaz for this version of the
workaround.
 // For more information, see https://github.com/OfficeDev/office-
js/issues/129.
 // Let's say there are 2 content controls in the header and 1 in the footer.
 const contentControls = context.document.contentControls;
 contentControls.load();
```


```
 await context.sync().then(function () {
 // Clear the 2 content controls in the header.
 contentControls.items[0].clear(); 
 contentControls.items[1].clear();
 // Clear the control control in the footer then update it.
 contentControls.items[2].clear();
 contentControls.items[2].insertHtml('<p>New Footer</p>', 'Replace');
 });
});
```
## **Lost formatting of last bullet in a list or last paragraph**

If the formatting of the last bullet in a list or the last paragraph is lost in the specified body or range, check if you're using [Body.insertFileFromBase64](https://learn.microsoft.com/en-us/javascript/api/word/word.body#word-word-body-insertfilefrombase64-member(1)) or [Range.insertFileFromBase64](https://learn.microsoft.com/en-us/javascript/api/word/word.range#word-word-range-insertfilefrombase64-member(1)). If so, update your code to use [Document.insertFileFromBase64](https://learn.microsoft.com/en-us/javascript/api/word/word.document#word-word-document-insertfilefrombase64-member(1)) instead.

# **Meaning of null property values in the response**

null has special implications in the Word JavaScript APIs. It's used to represent default values or no formatting.

Formatting properties such as [color](https://learn.microsoft.com/en-us/javascript/api/word/word.font#word-word-font-color-member) will contain null values in the response when different values exist in the specified [range](https://learn.microsoft.com/en-us/javascript/api/word/word.range). For example, if you retrieve a range and load its range.font.color property:

- If all text in the range has the same font color, range.font.color specifies that color.
- If multiple font colors are present within the range, range.font.color is null .

## **Native JavaScript APIs don't work with Word.Table**

The [Word.Table](https://learn.microsoft.com/en-us/javascript/api/word/word.table) object is different from an [HTML table object](https://developer.mozilla.org/docs/Learn_web_development/Core/Structuring_content/HTML_table_basics) . The native JavaScript APIs used to interact with an HTML table can't be used to manage a Word.Table object. Instead, you must use the [Table APIs](https://learn.microsoft.com/en-us/javascript/api/word/word.table) available in the Word Object Model to interact with Word.Table and related objects.

Similarly, don't use Word JavaScript APIs to interact with HTML table objects.

# **Shape APIs can't find shapes**


You have shapes in your document but for some reason, when you used the API to get shapes e.g., context.document.body.shapes , the result is that 0 shapes were found.

One possibility is that the Word template is outdated. If you created a new document from the default template yet you see "Compatibility Mode" in the Word window's title bar, consider updating your default template.

To change the default template, see [Change the Normal template (Normal.dotm)](https://support.microsoft.com/office/06de294b-d216-47f6-ab77-ccb5166f98ea) .

- 1. Use the instructions to find the location of the Normal template on your machine.
- 2. Ensure that Word is closed.
- 3. Rename Normal.dotm in **File Explorer** or similar application. Or you can move Normal.dotm to another location.

#### ) **Important**

Because you renamed or moved Normal.dotm , Word automatically creates a new version the next time you open Word. Any customizations in your original Normal.dotm won't transfer to the new version so you'll need to add your customizations to the new template.

- 4. Open Word and create a new document using the default template. You should no longer see "Compatibility Mode".
- 5. Retry running your code using the Shape API.

### **See also**

- Troubleshoot development errors with Office Add-ins
- Troubleshoot user errors with Office Add-ins

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


The Office Add-ins Development Kit will create the project. It will then open the project in a *second* Visual Studio Code window. Close the original Visual Studio Code window.

#### 7 **Note**

If you use VSCode Insiders, or you have problems opening the project page in VSCode, install the extension manually by following **[these steps](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/development-kit-overview?tabs=vscode)**, and find the sample in the sample gallery.

### **Explore the project**

The add-in project that you've created with the Office Add-ins Development Kit contains sample code for a basic task pane add-in. If you'd like to explore the components of your add-in project, open the project in your code editor and review the files listed below. When you're ready to try out your add-in, proceed to the next section.

- 1. The **./manifest.xml** or **./manifest.json** file in the root directory of the project defines the settings and capabilities of the add-in.
- 2. The **./src/taskpane/taskpane.html** file contains the HTML markup for the task pane.


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


The article Troubleshoot development errors with Office Add-ins contains solutions to common problems. If you're still having issues, [create a GitHub issue](https://aka.ms/officedevkitnewissue) and we'll help you.

For information on running the add-in on Office on the web, see Sideload Office Add-ins to Office on the web.

For information on debugging on older versions of Office, see Debug add-ins using developer tools in Microsoft Edge Legacy.

# **Build your first Word task pane add-in with Visual Studio**

Article • 08/27/2024

In this article, you'll walk through the process of building a Word task pane add-in.

#### **Prerequisites**

- [Visual Studio 2019 or later](https://www.visualstudio.com/vs/) with the **Office/SharePoint development** workload installed.
7 **Note**

If you've previously installed Visual Studio, use the Visual Studio Installer to ensure that the **Office/SharePoint development** workload is installed.

- Office connected to a Microsoft 365 subscription (including Office on the web).
#### **Create the add-in project**

- 1. In Visual Studio, choose **Create a new project**.
- 2. Using the search box, enter **add-in**. Choose **Word Web Add-in**, then select **Next**.
- 3. Name your project and select **Create**.
- 4. Visual Studio creates a solution and its two projects appear in **Solution Explorer**. The **Home.html** file opens in Visual Studio.

#### **Explore the Visual Studio solution**

When you've completed the wizard, Visual Studio creates a solution that contains two projects.

ノ **Expand table**


| Project                       | Description                                                                                                                                                                                                                                                                                                                                                                                                                                            |
|-------------------------------|--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| Add-in<br>project             | Contains only an XML-formatted add-in only manifest file, which contains all the<br>settings that describe your add-in. These settings help the Office application<br>determine when your add-in should be activated and where the add-in should<br>appear. Visual Studio generates the contents of this file for you so that you can<br>run the project and use your add-in immediately. Change these settings any time<br>by modifying the XML file. |
| Web<br>application<br>project | Contains the content pages of your add-in, including all the files and file<br>references that you need to develop Office-aware HTML and JavaScript pages.<br>While you develop your add-in, Visual Studio hosts the web application on your<br>local IIS server. When you're ready to publish the add-in, you'll need to deploy<br>this web application project to a web server.                                                                      |

#### **Update the code**

- 1. **Home.html** specifies the HTML that will be rendered in the add-in's task pane. In **Home.html**, replace the <body> element with the following markup and save the file.

```
HTML
<body>
 <div id="content-header">
 <div class="padding">
 <h1>Welcome</h1>
 </div>
 </div>
 <div id="content-main">
 <div class="padding">
 <p>Choose the buttons below to add boilerplate text to the
document by using the Word JavaScript API.</p>
 <br />
 <h3>Try it out</h3>
 <button id="emerson">Add quote from Ralph Waldo
Emerson</button>
 <br /><br />
 <button id="checkhov">Add quote from Anton Chekhov</button>
 <br /><br />
 <button id="proverb">Add Chinese proverb</button>
 </div>
 </div>
 <br />
 <div id="supportedVersion"/>
</body>
```


- 2. Open the file **Home.js** in the root of the web application project. This file specifies the script for the add-in. Replace the entire contents with the following code and save the file.

```
JavaScript
'use strict';
(function () {
 Office.onReady(function() {
 // Office is ready.
 $(document).ready(function () {
 // The document is ready.
 // Use this to check whether the API is supported in the
Word client.
 if (Office.context.requirements.isSetSupported('WordApi', 
'1.1')) {
 // Do something that is only available via the new
APIs.
 $('#emerson').on("click",
insertEmersonQuoteAtSelection);
 $('#checkhov').on("click",
insertChekhovQuoteAtTheBeginning);
 $('#proverb').on("click",
insertChineseProverbAtTheEnd);
 $('#supportedVersion').html('This code is using Word
2016 or later.');
 } else {
 // Lets you know that this code will not work with your
version of Word.
 $('#supportedVersion').html('This code requires Word
2016 or later.');
 }
 });
 });
 async function insertEmersonQuoteAtSelection() {
 await Word.run(async (context) => {
 // Create a proxy object for the document.
 const thisDocument = context.document;
 // Queue a command to get the current selection.
 // Create a proxy range object for the selection.
 const range = thisDocument.getSelection();
 // Queue a command to replace the selected text.
 range.insertText('"Hitch your wagon to a star." - Ralph
Waldo Emerson\n', Word.InsertLocation.replace);
 // Synchronize the document state by executing the queued
commands,
```


```
 // and return a promise to indicate task completion.
 await context.sync();
 console.log('Added a quote from Ralph Waldo Emerson.');
 })
 .catch(function (error) {
 console.log('Error: ' + JSON.stringify(error));
 if (error instanceof OfficeExtension.Error) {
 console.log('Debug info: ' + 
JSON.stringify(error.debugInfo));
 }
 });
 }
 async function insertChekhovQuoteAtTheBeginning() {
 await Word.run(async (context) => {
 // Create a proxy object for the document body.
 const body = context.document.body;
 // Queue a command to insert text at the start of the
document body.
 body.insertText('"Knowledge is of no value unless you put
it into practice." - Anton Chekhov\n', Word.InsertLocation.start);
 // Synchronize the document state by executing the queued
commands,
 // and return a promise to indicate task completion.
 await context.sync();
 console.log('Added a quote from Anton Chekhov.');
 })
 .catch(function (error) {
 console.log('Error: ' + JSON.stringify(error));
 if (error instanceof OfficeExtension.Error) {
 console.log('Debug info: ' + 
JSON.stringify(error.debugInfo));
 }
 });
 }
 async function insertChineseProverbAtTheEnd() {
 await Word.run(async (context) => {
 // Create a proxy object for the document body.
 const body = context.document.body;
 // Queue a command to insert text at the end of the
document body.
 body.insertText('"To know the road ahead, ask those coming
back." - Chinese proverb\n', Word.InsertLocation.end);
 // Synchronize the document state by executing the queued
commands,
 // and return a promise to indicate task completion.
 await context.sync();
 console.log('Added a quote from a Chinese proverb.');
```


```
 })
 .catch(function (error) {
 console.log('Error: ' + JSON.stringify(error));
 if (error instanceof OfficeExtension.Error) {
 console.log('Debug info: ' + 
JSON.stringify(error.debugInfo));
 }
 });
 }
})();
```
- 3. Open the file **Home.css** in the root of the web application project. This file specifies the custom styles for the add-in. Replace the entire contents with the following code and save the file.

```
css
#content-header {
 background: #2a8dd4;
 color: #fff;
 position: absolute;
 top: 0;
 left: 0;
 width: 100%;
 height: 80px;
 overflow: hidden;
}
#content-main {
 background: #fff;
 position: fixed;
 top: 80px;
 left: 0;
 right: 0;
 bottom: 0;
 overflow: auto;
}
.padding {
 padding: 15px;
}
```
### **Update the manifest**

- 1. Open the add-in only manifest file in the add-in project. This file defines the addin's settings and capabilities.
- 2. The ProviderName element has a placeholder value. Replace it with your name.


- 3. The DefaultValue attribute of the DisplayName element has a placeholder. Replace it with **My Office Add-in**.
- 4. The DefaultValue attribute of the Description element has a placeholder. Replace it with **A task pane add-in for Word**.
- 5. Save the file.

```
XML
...
<ProviderName>John Doe</ProviderName>
<DefaultLocale>en-US</DefaultLocale>
<!-- The display name of your add-in. Used on the Store and various
places of the Office UI such as the add-in's dialog. -->
<DisplayName DefaultValue="My Office Add-in" />
<Description DefaultValue="A task pane add-in for Word."/>
...
```
#### **Try it out**

- 1. Using Visual Studio, test the newly created Word add-in by pressing F5 or choosing **Debug** > **Start Debugging** to launch Word with the **Show Taskpane** addin button displayed on the ribbon. The add-in will be hosted locally on IIS.
- 2. In Word, if the add-in task pane isn't already open, choose the **Home** tab, and then choose the **Show Taskpane** button on the ribbon to open the add-in task pane. (If you're using a volume-licensed perpetual version of Office 2016 or older, then custom buttons aren't supported. Instead, the task pane will open immediately.)

| File         | Home    | Insert<br>Draw       |    | Design Layout             | References                         | Mailings                 | Review | View        |              | Developer Help                        | Script Lab        |        | Comments       | Editing -                  | B V                      | ಿಗೆ  |
|--------------|---------|----------------------|----|---------------------------|------------------------------------|--------------------------|--------|-------------|--------------|---------------------------------------|-------------------|--------|----------------|----------------------------|--------------------------|------|
|              | િ<br>0  | Calibri (Body)       | V  | 11<br>V<br>Pos<br>ਧ<br>Pa | ~ [= ~<br>=<br>li<br>=<br>=<br>్రా | u s<br>+=<br>=<br>=<br>1 | **     | Styles<br>4 | Editing<br>V | Dictate<br>Transcribe<br>A Read Aloud | Sensitivity<br>ﻬﻪ | Editor | Reuse<br>Files | Show<br>Taskpane           |                          |      |
| Clipboard Fa |         | Fort                 |    | િય                        | Paragraph                          |                          |        | Styles Fa   |              | Vaice                                 | Sansitivity       | Editor |                | Reuse Files Commands Group |                          | V    |
|              |         |                      |    |                           |                                    |                          |        |             |              |                                       |                   |        |                | Show Taskpane              | Click to Show a Taskpane |      |
| Page 1 of 1  | 0 Mords | Text Predictions: On | 13 |                           | Accessibility: Good to go          |                          |        |             |              | Las Display Settings                  | Focus             | ਸੀ।    | 년              | 10                         | t                        | 100% |

- 3. In the task pane, choose any of the buttons to add boilerplate text to the document.


| Dictate<br>Calibri (Body)<br>+=<br>Transcribe<br>Editina<br>Sensitivity<br>Show<br>Styles<br>Editor<br>laskpane<br>A Read Aloud<br>FILES<br>Clipboard<br>2<br>Fort<br>E<br>Paragraph<br>Styles Fa<br>WOICE<br>ﺔ<br>Commands Group<br>Section with<br>My Office Add-in                                                           | V<br>× |
|---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|--------|
|                                                                                                                                                                                                                                                                                                                                 |        |
|                                                                                                                                                                                                                                                                                                                                 |        |
| Welcome<br>"Knowledge is of no value unless you put it into practice." - Anton Chekhov<br>"Hitch your wagon to a star." - Ralph Waldo Emerson<br>Choose the buttons below to add<br>"To know the road ahead, ask those coming back." - Chinese proverb<br>boilerplate text to the document by using<br>the Word JavaScript API. |        |
| Try it out<br>Add quote from Ralph Waldo Emerson<br>Add quote from Anton Chekhov<br>Add Chinese proverb<br>Accessibility: Good to go<br>Add-ins loaded successfully<br>Display Settings<br>Forus<br>Page 1 of 1<br>36 words<br>Text Predictions: On                                                                             |        |

#### 7 **Note**

To see the console.log output, you'll need a separate set of developer tools for a JavaScript console. To learn more about F12 tools and the Microsoft Edge DevTools, visit **Debug add-ins using developer tools for Internet Explorer**, **Debug add-ins using developer tools for Edge Legacy**, or **Debug add-ins using developer tools in Microsoft Edge (Chromium-based)**.

#### **Next steps**

Congratulations, you've successfully created a Word task pane add-in! Next, to learn more about developing Office Add-ins with Visual Studio, continue to the following article.

**Develop Office Add-ins with Visual Studio**

## **Troubleshooting**

- Ensure your environment is ready for Office development by following the instructions in Set up your development environment.
- Some of the sample code uses ES6 JavaScript. This isn't compatible with older versions of Office that use the Trident (Internet Explorer 11) browser engine. For information on how to support those platforms in your add-in, see Support older Microsoft webviews and Office versions. If you don't already have a Microsoft 365 subscription to use for development, you might qualify for a Microsoft 365 E5


developer subscription through the [Microsoft 365 Developer Program](https://aka.ms/m365devprogram) ; for details, see the [FAQ](https://learn.microsoft.com/en-us/office/developer-program/microsoft-365-developer-program-faq#who-qualifies-for-a-microsoft-365-e5-developer-subscription-). Alternatively, you can [sign up for a 1-month free trial](https://www.microsoft.com/microsoft-365/try) or [purchase a Microsoft 365 plan](https://www.microsoft.com/microsoft-365/business/compare-all-microsoft-365-business-products-g) .

- If your add-in shows an error (for example, "This add-in could not be started. Close this dialog to ignore the problem or click "Restart" to try again.") when you press F5 or choose **Debug** > **Start Debugging** in Visual Studio, see Debug Office Addins in Visual Studio for other debugging options.
#### **Code samples**

- [Word "Hello world" add-in](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/hello-world/word-hello-world) : Learn how to build a simple Office Add-in with only a manifest, HTML web page, and a logo.
#### **See also**

- Office Add-ins platform overview
- Develop Office Add-ins
- Word add-ins overview
- [Word add-in code samples](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Word)
- Word JavaScript API reference
- Publish your add-in using Visual Studio