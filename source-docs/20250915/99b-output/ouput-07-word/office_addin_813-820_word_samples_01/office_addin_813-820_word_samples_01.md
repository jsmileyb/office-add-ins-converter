{0}------------------------------------------------

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

{1}------------------------------------------------

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

{2}------------------------------------------------

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

{3}------------------------------------------------

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

{4}------------------------------------------------

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

{5}------------------------------------------------

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

{6}------------------------------------------------

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

{7}------------------------------------------------

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