
# **PowerPoint add-ins documentation**

With PowerPoint add-ins, you can use familiar web technologies such as HTML, CSS, and JavaScript to build a solution that can run in PowerPoint across multiple platforms, including on the web, Windows, Mac, and iPad. Learn how to build, test, debug, and publish PowerPoint add-ins.

| About PowerPoint add-ins                       |  |
|------------------------------------------------|--|
| e<br>OVERVIEW                                  |  |
| What are PowerPoint add-ins?                   |  |
| JavaScript API for PowerPoint                  |  |
| f<br>QUICKSTART                                |  |
| Build your first PowerPoint add-in             |  |
| Explore Office JavaScript API using Script Lab |  |
| c<br>HOW-TO GUIDE                              |  |
| Test and debug a PowerPoint add-ins            |  |
| Deploy and publish a PowerPoint add-ins        |  |

#### **Key Office Add-ins concepts**

e **OVERVIEW**

Office Add-ins platform overview

#### b **GET STARTED**

Core concepts for Office Add-ins

Design Office Add-ins

Develop Office Add-ins


#### **Resources**

i **REFERENCE**

[Ask questions](https://stackoverflow.com/questions/tagged/office-js)

[Request features](https://aka.ms/m365dev-suggestions)

[Report issues](https://github.com/officedev/office-js/issues)

Office Add-ins additional resources


# **PowerPoint add-ins**

06/13/2025

You can use PowerPoint add-ins to build engaging solutions for your users' presentations across platforms including Windows, iPad, Mac, and in a browser. You can create two types of PowerPoint add-ins:

- Use **task pane add-ins** to bring in reference information or insert data into the presentation via a service. For example, see the [Pexels - Free Stock Photos](https://appsource.microsoft.com/product/office/wa104379997) add-in, which you can use to add professional photos to your presentation. To create your own task pane add-in, you can start with Build your first PowerPoint task pane add-in.
- Use **content add-ins** to add dynamic HTML5 content to your presentations. For example, see the [LucidChart Diagrams for PowerPoint](https://appsource.microsoft.com/product/office/wa104380117) add-in, which injects interactive diagrams from LucidChart into your deck. To create your own content add-in, start with Build your first PowerPoint content add-in.

# **PowerPoint add-in scenarios**

The code examples in this article demonstrate some basic tasks that can be useful when developing add-ins for PowerPoint.

# **Add a new slide then navigate to it**

In the following code sample, the addAndNavigateToNewSlide function calls the [SlideCollection.add](https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.slidecollection#powerpoint-powerpoint-slidecollection-add-member(1)) method to add a new slide to the presentation. The function then calls the [Presentation.setSelectedSlides](https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.presentation#powerpoint-powerpoint-presentation-setselectedslides-member(1)) method to navigate to the new slide.

```
JavaScript
async function addAndNavigateToNewSlide() {
 // Adds a new slide then navigates to it.
 await PowerPoint.run(async (context) => {
 const slideCountResult = context.presentation.slides.getCount();
 context.presentation.slides.add();
 await context.sync();
 const newSlide =
context.presentation.slides.getItemAt(slideCountResult.value);
 newSlide.load("id");
 await context.sync();
 console.log(`Added slide - ID: ${newSlide.id}`);
```


```
 // Navigate to the new slide.
 context.presentation.setSelectedSlides([newSlide.id]);
 await context.sync();
 });
}
```
# **Navigate to a particular slide in the presentation**

In the following code sample, the getSelectedSlides function calls the [Presentation.getSelectedSlides](https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.presentation#powerpoint-powerpoint-presentation-getselectedslides-member(1)) method to get the selected slides then logs their IDs. The function can then act on the current slide (or first slide from the selection).

```
JavaScript
async function getSelectedSlides() {
 // Gets the ID of the current slide (or selected slides).
 await PowerPoint.run(async (context) => {
 const selectedSlides = context.presentation.getSelectedSlides();
 selectedSlides.load("items/id");
 await context.sync();
 if (selectedSlides.items.length === 0) {
 console.warn("No slides were selected.");
 return;
 }
 console.log("IDs of selected slides:");
 selectedSlides.items.forEach(item => {
 console.log(item.id);
 });
 // Navigate to first selected slide.
 const currentSlide = selectedSlides.items[0];
 console.log(`Navigating to slide with ID ${currentSlide.id} ...`);
 context.presentation.setSelectedSlides([currentSlide.id]);
 // Perform actions on current slide...
 });
}
```
## **Navigate between slides in the presentation**

In the following code sample, the goToSlideByIndex function calls the Presentation.setSelectedSlides method to navigate to the first slide in the presentation, which has the index 0. The maximum slide index you can navigate to in this sample is slideCountResult.value - 1 .


```
JavaScript
```

```
async function goToSlideByIndex() {
 await PowerPoint.run(async (context) => {
 // Gets the number of slides in the presentation.
 const slideCountResult = context.presentation.slides.getCount();
 await context.sync();
 if (slideCountResult.value === 0) {
 console.warn("There are no slides.");
 return;
 }
 const slide = context.presentation.slides.getItemAt(0); // First slide
 //const slide = context.presentation.slides.getItemAt(slideCountResult.value -
1); // Last slide
 slide.load("id");
 await context.sync();
 console.log(`Slide ID: ${slide.id}`);
 // Navigate to the slide.
 context.presentation.setSelectedSlides([slide.id]);
 await context.sync();
 });
}
```
# **Get the URL of the presentation**

In the following code sample, the getFileUrl function calls the [Document.getFileProperties](https://learn.microsoft.com/en-us/javascript/api/office/office.document#office-office-document-getfilepropertiesasync-member(1)) method to get the URL of the presentation file.

```
JavaScript
function getFileUrl() {
 // Gets the URL of the current file.
 Office.context.document.getFilePropertiesAsync(function (asyncResult) {
 const fileUrl = asyncResult.value.url;
 if (fileUrl === "") {
 console.warn("The file hasn't been saved yet. Save the file and try
again.");
 } else {
 console.log(`File URL: ${fileUrl}`);
 }
 });
}
```
# **Create a presentation**


Your add-in can create a new presentation, separate from the PowerPoint instance in which the add-in is currently running. The PowerPoint namespace has the createPresentation method for this purpose. When this method is called, the new presentation is immediately opened and displayed in a new instance of PowerPoint. Your add-in remains open and running with the previous presentation.

```
JavaScript
PowerPoint.createPresentation();
```
The createPresentation method can also create a copy of an existing presentation. The method accepts a Base64-encoded string representation of an .pptx file as an optional parameter. The resulting presentation will be a copy of that file, assuming the string argument is a valid .pptx file. The [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) class can be used to convert a file into the required Base64-encoded string, as demonstrated in the following example.

JavaScript const myFile = document.getElementById("file") as HTMLInputElement; const reader = new FileReader(); reader.onload = function (event) { // Strip off the metadata before the Base64-encoded string. const startIndex = reader.result.toString().indexOf("base64,"); const copyBase64 = reader.result.toString().substr(startIndex + 7); PowerPoint.createPresentation(copyBase64); }; // Read in the file as a data URL so we can parse the Base64-encoded string. reader.readAsDataURL(myFile.files[0]);

To see a full code sample that includes an HTML implementation, see [Create presentation](https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/powerpoint/document/create-presentation.yaml) .

# **Detect the presentation's active view and handle the ActiveViewChanged event**

If you're building a content add-in, you'll need to get the presentation's active view and handle the [Document.ActiveViewChanged](https://learn.microsoft.com/en-us/javascript/api/office/office.eventtype#fields) event as part of your [Office.onReady](https://learn.microsoft.com/en-us/javascript/api/office#office-office-onready-function(1)) call.

7 **Note**


In PowerPoint on the web, the Document.ActiveViewChanged event will never fire because **Slide Show** mode is treated as a new session. In this case, the add-in must fetch the active view on load, as shown in the following code sample.

Note the following about the code sample:

- The getActiveFileView function calls the [Document.getActiveViewAsync](https://learn.microsoft.com/en-us/javascript/api/office/office.document#office-office-document-getactiveviewasync-member(1)) method to return whether the presentation's current view is "edit" (any of the view where you can edit slides, such as **Normal**, **Slide Sorter**, or **Outline**) or "read" (**Slide Show** or **Reading View**), represented by the [ActiveView](https://learn.microsoft.com/en-us/javascript/api/office/office.activeview) enum.
- The registerActiveViewChanged function calls the [Document.addHandlerAsync](https://learn.microsoft.com/en-us/javascript/api/office/office.document#office-office-document-addhandlerasync-member(1)) method to register a handler for the Document.ActiveViewChanged event.
- To display information, this example uses the showNotification function, which is included in the Visual Studio Office Add-ins project templates. If you aren't using Visual Studio to develop your add-in, you'll need to replace the showNotification function with your own code.

JavaScript

```
// General Office.onReady function. Called after the add-in loads and Office JS is
initialized.
Office.onReady(() => {
 // Get whether the current view is edit or read.
 const currentView = getActiveFileView();
 // Register the active view changed handler.
 registerActiveViewChanged();
 // Render the content based off of the current view.
 if (currentView === Office.ActiveView.Read) {
 // Handle read view.
 console.log('Current view is read.');
 // You can add any specific logic for the read view here.
 } else {
 // Handle edit view.
 console.log('Current view is edit.');
 // You can add any specific logic for the edit view here.
 }
});
// Gets the active file view.
function getActiveFileView() {
 console.log('Getting active file view...');
 Office.context.document.getActiveViewAsync(function (result) {
 if (result.status === Office.AsyncResultStatus.Succeeded) {
 console.log('Active view:', result.value);
 return result.value;
 } else {
```


```
 console.error('Error getting active view:', result.error.message);
 showNotification('Error:', result.error.message);
 return null;
 }
 });
}
// Registers the ActiveViewChanged event.
function registerActiveViewChanged() {
 console.log('Registering ActiveViewChanged event handler...');
 Office.context.document.addHandlerAsync(
 Office.EventType.ActiveViewChanged,
 activeViewHandler,
 function (result) {
 if (result.status === Office.AsyncResultStatus.Failed) {
 console.error('Failed to register active view changed handler:',
result.error.message);
 showNotification('Error:', result.error.message);
 } else {
 console.log('Active view changed handler registered
successfully.');
 }
 });
}
// ActiveViewChanged event handler.
function activeViewHandler(eventArgs) {
 console.log('Active view changed:', JSON.stringify(eventArgs));
 showNotification('Active view changed', `The active view has changed to: 
${eventArgs.activeView}`);
 // You can add logic here based on the new active view.
}
```
# **See also**

- Developing Office Add-ins
- [Learn about the Microsoft 365 Developer Program](https://aka.ms/m365devprogram)
- PowerPoint quick starts
	- Build your first PowerPoint content add-in
	- Build your first PowerPoint task pane add-in
- [PowerPoint Code Samples](https://developer.microsoft.com/microsoft-365/gallery/?filterBy=Samples,PowerPoint)
- How to save add-in state and settings per document for content and task pane add-ins
- Read and write data to the active selection in a document or spreadsheet
- Get the whole document from an add-in for PowerPoint or Word
- Use document themes in your PowerPoint add-ins


# **Build your first PowerPoint task pane add-in**

Article • 09/17/2024

In this article, you'll walk through the process of building a PowerPoint task pane add-in.

## **Prerequisites**

- Node.js (the latest LTS version). Visit the [Node.js site](https://nodejs.org/) to download and install the right version for your operating system.
- The latest version of Yeoman and the Yeoman generator for Office Add-ins. To install these tools globally, run the following command via the command prompt.

command line

npm install -g yo generator-office

7 **Note**

Even if you've previously installed the Yeoman generator, we recommend you update your package to the latest version from npm.

- Office connected to a Microsoft 365 subscription (including Office on the web).
7 **Note**

If you don't already have Office, you might qualify for a Microsoft 365 E5 developer subscription through the **[Microsoft 365 Developer Program](https://aka.ms/m365devprogram)** ; for details, see the **[FAQ](https://learn.microsoft.com/en-us/office/developer-program/microsoft-365-developer-program-faq#who-qualifies-for-a-microsoft-365-e5-developer-subscription-)**. Alternatively, you can **[sign up for a 1-month free trial](https://www.microsoft.com/microsoft-365/try)** or **[purchase a Microsoft 365 plan](https://www.microsoft.com/microsoft-365/business/compare-all-microsoft-365-business-products-g)** .

# **Create the add-in project**

Run the following command to create an add-in project using the Yeoman generator. A folder that contains the project will be added to the current directory.


#### 7 **Note**

When you run the yo office command, you may receive prompts about the data collection policies of Yeoman and the Office Add-in CLI tools. Use the information that's provided to respond to the prompts as you see fit.

When prompted, provide the following information to create your add-in project.

- **Choose a project type:** Office Add-in Task Pane project
- **Choose a script type:** Javascript
- **What do you want to name your add-in?** My Office Add-in
- **Which Office client application would you like to support?** PowerPoint

After you complete the wizard, the generator creates the project and installs supporting Node components.

# **Explore the project**

The add-in project that you've created with the Yeoman generator contains sample code for a basic task pane add-in. If you'd like to explore the components of your add-in project, open the project in your code editor and review the files listed below. When you're ready to try out your add-in, proceed to the next section.

- The **./manifest.xml** or **manifest.json** file in the root directory of the project defines the settings and capabilities of the add-in.


- The **./src/taskpane/taskpane.html** file contains the HTML markup for the task pane.
- The **./src/taskpane/taskpane.css** file contains the CSS that's applied to content in the task pane.
- The **./src/taskpane/taskpane.js** file contains the Office JavaScript API code that facilitates interaction between the task pane and the Office client application.

# **Try it out**

- 1. Navigate to the root folder of the project.
command line cd "My Office Add-in"

- 2. Complete the following steps to start the local web server and sideload your addin.
#### 7 **Note**

- Office Add-ins should use HTTPS, not HTTP, even while you're developing. If you're prompted to install a certificate after you run one of the following commands, accept the prompt to install the certificate that the Yeoman generator provides. You may also have to run your command prompt or terminal as an administrator for the changes to be made.
- If this is your first time developing an Office Add-in on your machine, you may be prompted in the command line to grant Microsoft Edge WebView a loopback exemption ("Allow localhost loopback for Microsoft Edge WebView?"). When prompted, enter Y to allow the exemption. Note that you'll need administrator privileges to allow the exemption. Once allowed, you shouldn't be prompted for an exemption when you sideload Office Add-ins in the future (unless you remove the exemption from your machine). To learn more, see **["We can't open this](https://learn.microsoft.com/en-us/office/troubleshoot/office-suite-issues/cannot-open-add-in-from-localhost) [add-in from localhost" when loading an Office Add-in or using Fiddler](https://learn.microsoft.com/en-us/office/troubleshoot/office-suite-issues/cannot-open-add-in-from-localhost)**.


#### **Tip**

If you're testing your add-in on Mac, run the following command before proceeding. When you run this command, the local web server starts.

command line

npm run dev-server

- To test your add-in in PowerPoint, run the following command in the root directory of your project. This starts the local web server (if it's not already running) and opens PowerPoint with your add-in loaded.
command line npm start

- To test your add-in in PowerPoint on a browser, run the following command in the root directory of your project. When you run this command, the local web server starts. Replace "{url}" with the URL of a PowerPoint document on your OneDrive or a SharePoint library to which you have permissions.
#### 7 **Note**

If you are developing on a Mac, enclose the {url} in single quotation marks. Do *not* do this on Windows.

command line

```
npm run start -- web --document {url}
```
The following are examples.


- npm run start -- web --document https://contoso.sharepoint.com/:t:/g/EZGxP7ksiE5DuxvY638G798BpuhwluxCM fF1WZQj3VYhYQ?e=F4QM1R npm run start -- web --document
https://1drv.ms/x/s!jkcH7spkM4EGgcZUgqthk4IK3NOypVw?e=Z6G1qp

- npm run start -- web --document https://contoso-my.sharepointdf.com/:t:/p/user/EQda453DNTpFnl1bFPhOVR0BwlrzetbXvnaRYii2lDr_oQ? e=RSccmNP
If your add-in doesn't sideload in the document, manually sideload it by following the instructions in Manually sideload add-ins to Office on the web.

- 3. In PowerPoint, insert a new blank slide, choose the **Home** tab, and then choose the **Show Taskpane** button on the ribbon to open the add-in task pane.
- 4. At the bottom of the task pane, choose the **Run** link to insert the text "Hello World" into the current slide.


- 5. When you want to stop the local web server and uninstall the add-in, follow the applicable instructions:
	- To stop the server, run the following command. If you used npm start , the following command also uninstalls the add-in.

| command line |  |  |  |  |  |
|--------------|--|--|--|--|--|
| npm stop     |  |  |  |  |  |

- If you manually sideloaded the add-in, see Remove a sideloaded add-in.
## **Next steps**

Congratulations, you've successfully created a PowerPoint task pane add-in! Next, learn more about the capabilities of a PowerPoint add-in and build a more complex add-in by following along with the PowerPoint add-in tutorial.

# **Troubleshooting**

- Ensure your environment is ready for Office development by following the instructions in Set up your development environment.
- Some of the sample code uses ES6 JavaScript. This isn't compatible with older versions of Office that use the Trident (Internet Explorer 11) browser engine. For information on how to support those platforms in your add-in, see Support older


Microsoft webviews and Office versions. If you don't already have a Microsoft 365 subscription to use for development, you might qualify for a Microsoft 365 E5 developer subscription through the [Microsoft 365 Developer Program](https://aka.ms/m365devprogram) ; for details, see the [FAQ](https://learn.microsoft.com/en-us/office/developer-program/microsoft-365-developer-program-faq#who-qualifies-for-a-microsoft-365-e5-developer-subscription-). Alternatively, you can [sign up for a 1-month free trial](https://www.microsoft.com/microsoft-365/try) or [purchase a Microsoft 365 plan](https://www.microsoft.com/microsoft-365/business/compare-all-microsoft-365-business-products-g) .

- The automatic npm install step Yo Office performs may fail. If you see errors when trying to run npm start , navigate to the newly created project folder in a command prompt and manually run npm install . For more information about Yo Office, see Create Office Add-in projects using the Yeoman Generator.
## **Code samples**

- [PowerPoint "Hello world" add-in](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/hello-world/powerpoint-hello-world) : Learn how to build a simple Office Add-in with only a manifest, HTML web page, and a logo.
## **See also**

- Office Add-ins platform overview
- Develop Office Add-ins
- Using Visual Studio Code to publish


# **Build your first PowerPoint content addin**

Article • 08/27/2024

In this article, you'll walk through the process of building a PowerPoint content add-in using Visual Studio.

## **Prerequisites**

- [Visual Studio 2019 or later](https://www.visualstudio.com/vs/) with the **Office/SharePoint development** workload installed.
7 **Note**

If you've previously installed Visual Studio, use the Visual Studio Installer to ensure that the **Office/SharePoint development** workload is installed.

- Office connected to a Microsoft 365 subscription (including Office on the web).
### **Create the add-in project**

- 1. In Visual Studio, choose **Create a new project**.
- 2. Using the search box, enter **add-in**. Choose **PowerPoint Web Add-in**, then select **Next**.
- 3. Name your project and select **Create**.
- 4. In the **Create Office Add-in** dialog window, choose **Insert content into PowerPoint slides**, and then choose **Finish** to create the project.
- 5. Visual Studio creates a solution and its two projects appear in **Solution Explorer**. The **Home.html** file opens in Visual Studio.

# **Explore the Visual Studio solution**

When you've completed the wizard, Visual Studio creates a solution that contains two projects.


| Project                       | Description                                                                                                                                                                                                                                                                                                                                                                                                                                            |
|-------------------------------|--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| Add-in<br>project             | Contains only an XML-formatted add-in only manifest file, which contains all the<br>settings that describe your add-in. These settings help the Office application<br>determine when your add-in should be activated and where the add-in should<br>appear. Visual Studio generates the contents of this file for you so that you can<br>run the project and use your add-in immediately. Change these settings any time<br>by modifying the XML file. |
| Web<br>application<br>project | Contains the content pages of your add-in, including all the files and file<br>references that you need to develop Office-aware HTML and JavaScript pages.<br>While you develop your add-in, Visual Studio hosts the web application on your<br>local IIS server. When you're ready to publish the add-in, you'll need to deploy<br>this web application project to a web server.                                                                      |

## **Update the code**

- 1. **Home.html** specifies the HTML that will be rendered in the add-in's task pane. In **Home.html**, find the <p> element that contains the text "This example will read the current document selection." and the <button> element where the id is "get-datafrom-selection". Replace these entire elements with the following markup then save the file.

```
HTML
<p class="ms-font-m-plus">This example will get some details about the
current slide.</p>
<button class="Button Button--primary" id="get-data-from-selection">
 <span class="Button-icon"><i class="ms-Icon ms-Icon--plus"></i>
</span>
 <span class="Button-label">Get slide details</span>
 <span class="Button-description">Gets and displays the current
slide's details.</span>
</button>
```
- 2. Open the file **Home.js** in the root of the web application project. This file specifies the script for the add-in. Find the getDataFromSelection function and replace the entire function with the following code then save the file.
JavaScript // Gets some details about the current slide and displays them in a notification.


```
function getDataFromSelection() {
 if (Office.context.document.getSelectedDataAsync) {

Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideR
ange,
 function (result) {
 if (result.status ===
Office.AsyncResultStatus.Succeeded) {
 showNotification('Some slide details are:', '"' + 
JSON.stringify(result.value) + '"');
 } else {
 showNotification('Error:', result.error.message);
 }
 }
 );
 } else {
 app.showNotification('Error:', 'Reading selection data is not
supported by this host application.');
 }
}
```
## **Update the manifest**

- 1. Open the add-in only manifest file in the add-in project. This file defines the addin's settings and capabilities.
- 2. The ProviderName element has a placeholder value. Replace it with your name.
- 3. The DefaultValue attribute of the DisplayName element has a placeholder. Replace it with **My Office Add-in**.
- 4. The DefaultValue attribute of the Description element has a placeholder. Replace it with **A content add-in for PowerPoint.**.
- 5. Save the file. The updated lines should look like the following code sample.

```
XML
...
<ProviderName>John Doe</ProviderName>
<DefaultLocale>en-US</DefaultLocale>
<!-- The display name of your add-in. Used on the store and various
places of the Office UI such as the add-ins dialog. -->
<DisplayName DefaultValue="My Office Add-in" />
<Description DefaultValue="A content add-in for PowerPoint."/>
...
```


# **Try it out**

- 1. Using Visual Studio, test the newly created PowerPoint add-in by pressing F5 or choosing the **Start** button to launch PowerPoint with the content add-in displayed over the slide.
- 2. In PowerPoint, choose the **Get slide details** button in the content add-in to get details about the current slide.

| Welcome<br>This example will get some details about the current |
|-----------------------------------------------------------------|
| slide.<br>Get slide details                                     |
| Find more samples online                                        |
|                                                                 |
| `lick tc                                                        |
|                                                                 |

#### 7 **Note**

To see the console.log output, you'll need a separate set of developer tools for a JavaScript console. To learn more about F12 tools and the Microsoft Edge DevTools, visit **Debug add-ins using developer tools for Internet Explorer**, **Debug add-ins using developer tools for Edge Legacy**, or **Debug add-ins using developer tools in Microsoft Edge (Chromium-based)**.

# **Next steps**

Congratulations, you've successfully created a PowerPoint content add-in! Next, learn more about developing Office Add-ins with Visual Studio.

# **Troubleshooting**


- Ensure your environment is ready for Office development by following the instructions in Set up your development environment.
- Some of the sample code uses ES6 JavaScript. This isn't compatible with older versions of Office that use the Trident (Internet Explorer 11) browser engine. For information on how to support those platforms in your add-in, see Support older Microsoft webviews and Office versions. If you don't already have a Microsoft 365 subscription to use for development, you might qualify for a Microsoft 365 E5 developer subscription through the [Microsoft 365 Developer Program](https://aka.ms/m365devprogram) ; for details, see the [FAQ](https://learn.microsoft.com/en-us/office/developer-program/microsoft-365-developer-program-faq#who-qualifies-for-a-microsoft-365-e5-developer-subscription-). Alternatively, you can [sign up for a 1-month free trial](https://www.microsoft.com/microsoft-365/try) or [purchase a Microsoft 365 plan](https://www.microsoft.com/microsoft-365/business/compare-all-microsoft-365-business-products-g) .
- If your add-in shows an error (for example, "This add-in could not be started. Close this dialog to ignore the problem or click "Restart" to try again.") when you press F5 or choose **Debug** > **Start Debugging** in Visual Studio, see Debug Office Addins in Visual Studio for other debugging options.

## **See also**

- Office Add-ins platform overview
- Develop Office Add-ins
- Using Visual Studio Code to publish


# **Tutorial: Create a PowerPoint task pane add-in**

Article • 11/21/2024

In this tutorial, you'll create a PowerPoint task pane add-in that:

- " Adds an image to a slide
- " Adds text to a slide
- " Gets slide metadata
- " Adds new slides
- " Navigates between slides

## **Create the add-in**

#### **Tip**

If you've already completed the **Build your first PowerPoint task pane add-in** quick start using the Yeoman generator, and want to use that project as a starting point for this tutorial, go directly to the **Insert an image** section to start this tutorial.

If you want a completed version of this tutorial, visit the **[Office Add-ins samples](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/tutorials/powerpoint-tutorial-yo) [repo on GitHub](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/tutorials/powerpoint-tutorial-yo)** .

### **Prerequisites**

- Node.js (the latest LTS version). Visit the [Node.js site](https://nodejs.org/) to download and install the right version for your operating system.
- The latest version of Yeoman and the Yeoman generator for Office Add-ins. To install these tools globally, run the following command via the command prompt.

```
command line
npm install -g yo generator-office
```
7 **Note**


Even if you've previously installed the Yeoman generator, we recommend you update your package to the latest version from npm.

- Office connected to a Microsoft 365 subscription (including Office on the web).
#### 7 **Note**

If you don't already have Office, you might qualify for a Microsoft 365 E5 developer subscription through the **[Microsoft 365 Developer Program](https://aka.ms/m365devprogram)** ; for details, see the **[FAQ](https://learn.microsoft.com/en-us/office/developer-program/microsoft-365-developer-program-faq#who-qualifies-for-a-microsoft-365-e5-developer-subscription-)**. Alternatively, you can **[sign up for a 1-month free trial](https://www.microsoft.com/microsoft-365/try)** or **[purchase a Microsoft 365 plan](https://www.microsoft.com/microsoft-365/business/compare-all-microsoft-365-business-products-g)** .

### **Create the add-in project**

Run the following command to create an add-in project using the Yeoman generator. A folder that contains the project will be added to the current directory.

command line

yo office

#### 7 **Note**

When you run the yo office command, you may receive prompts about the data collection policies of Yeoman and the Office Add-in CLI tools. Use the information that's provided to respond to the prompts as you see fit.

When prompted, provide the following information to create your add-in project.

- **Choose a project type:** Office Add-in Task Pane project
- **Choose a script type:** JavaScript
- **What do you want to name your add-in?** My Office Add-in
- **Which Office client application would you like to support?** PowerPoint


After you complete the wizard, the generator creates the project and installs supporting Node components.

### **Complete setup**

- 1. Navigate to the root directory of the project.
command line

cd "My Office Add-in"

- 2. Open your project in VS Code or your preferred code editor.
On Windows, you can navigate to the root directory of the project via the command line and then enter code . to open that folder in VS Code. On Mac, you'll need to **[add the code command to the path](https://code.visualstudio.com/docs/setup/mac#_launching-from-the-command-line)** before you can use that command to open the project folder in VS Code.

## **Insert an image**

Complete the following steps to add code that inserts an image into a slide.

- 1. Open the project in your code editor.
- 2. In the root of the project, create a new file named **base64Image.js**.


- 3. Open the file **base64Image.js** and add the following code to specify the Base64 encoded string that represents an image.
JavaScript

export const base64Image =

"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" ;

- 4. Open the file **./src/taskpane/taskpane.html**. This file contains the HTML markup for the task pane.
- 5. Locate the <body> element. Replace it with the following markup, then save the file.

```
HTML
<body class="ms-font-m ms-welcome ms-Fabric">
 <!-- TODO2: Update the header node. -->
 <header class="ms-welcome__header ms-bgColor-neutralLighter">
 <img width="90" height="90" src="../../assets/logo-filled.png"
alt="Contoso" title="Contoso" />
 <h1 class="ms-font-su">Welcome</h1>
 </header>
 <section id="sideload-msg" class="ms-welcome__main">
 <h2 class="ms-font-xl">Please <a target="_blank"
href="https://learn.microsoft.com/office/dev/add-ins/testing/test-
debug-office-add-ins#sideload-an-office-add-in-for-
testing">sideload</a> your add-in to see app body.</h2>
 </section>
 <main id="app-body" class="ms-welcome__main" style="display:
none;">
 <div class="padding">
 <!-- TODO1: Create the insert-image button. -->
 <!-- TODO3: Create the insert-text button. -->
 <!-- TODO4: Create the get-slide-metadata button. -->
 <!-- TODO5: Create the add-slides and go-to-slide buttons.
-->
 </div>
 </main>
 <section id="display-msg" class="ms-welcome__main">
 <div class="padding">
 <h3>Message</h3>
 <div id="message"></div>
 </div>
 </section>
</body>
```


- 6. In the **taskpane.html** file, replace TODO1 with the following markup. This markup defines the **Insert Image** button that will appear within the add-in's task pane.

```
HTML
<button class="ms-Button" id="insert-image">Insert Image</button><br/>
<br/>
```
- 7. Open the file **./src/taskpane/taskpane.js**. This file contains the Office JavaScript API code that facilitates interaction between the task pane and the Office client application. Replace the entire contents with the following code and save the file.

```
JavaScript
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed
under the MIT license.
 * See LICENSE in the project root for license information.
 */
/* global document, Office */
// TODO1: Import Base64-encoded string for image.
Office.onReady((info) => {
 if (info.host === Office.HostType.PowerPoint) {
 document.getElementById("sideload-msg").style.display = "none";
 document.getElementById("app-body").style.display = "flex";
 // TODO2: Assign event handler for insert-image button.
 // TODO4: Assign event handler for insert-text button.
 // TODO6: Assign event handler for get-slide-metadata button.
 // TODO8: Assign event handlers for add-slides and the four
navigation buttons.
 }
});
// TODO3: Define the insertImage function.
// TODO5: Define the insertText function.
// TODO7: Define the getSlideMetadata function.
// TODO9: Define the addSlides and navigation functions.
async function clearMessage(callback) {
 document.getElementById("message").innerText = "";
 await callback();
}
function setMessage(message) {
 document.getElementById("message").innerText = message;
}
```


```
// Default helper for invoking an action and handling errors.
async function tryCatch(callback) {
 try {
 document.getElementById("message").innerText = "";
 await callback();
 } catch (error) {
 setMessage("Error: " + error.toString());
 }
}
```
- 8. In the **taskpane.js** file above the Office.onReady function call near the top of the file, replace TODO1 with the following code. This code imports the variable that you defined previously in the file **./base64Image.js**.

```
JavaScript
import { base64Image } from "../../base64Image";
```
- 9. In the **taskpane.js** file, replace TODO2 with the following code to assign the event handler for the **Insert Image** button.

```
JavaScript
document.getElementById("insert-image").onclick = () =>
clearMessage(insertImage);
```
- 10. In the **taskpane.js** file, replace TODO3 with the following code to define the insertImage function. This function uses the Office JavaScript API to insert the image into the document. Note:
	- The coercionType option that's specified as the second parameter of the setSelectedDataAsync request indicates the type of data being inserted.
	- The asyncResult object encapsulates the result of the setSelectedDataAsync request, including status and error information if the request failed.

```
JavaScript
function insertImage() {
 // Call Office.js to insert the image into the document.
 Office.context.document.setSelectedDataAsync(
 base64Image,
 {
 coercionType: Office.CoercionType.Image
 },
```


```
 (asyncResult) => {
 if (asyncResult.status === Office.AsyncResultStatus.Failed) {
 setMessage("Error: " + asyncResult.error.message);
 }
 }
 );
}
```
- 11. Save all your changes to the project.
### **Test the add-in**

- 1. Navigate to the root folder of the project.
command line

cd "My Office Add-in"

- 2. Complete the following steps to start the local web server and sideload your addin.
7 **Note**

- Office Add-ins should use HTTPS, not HTTP, even while you're developing. If you're prompted to install a certificate after you run one of the following commands, accept the prompt to install the certificate that the Yeoman generator provides. You may also have to run your command prompt or terminal as an administrator for the changes to be made.
- If this is your first time developing an Office Add-in on your machine, you may be prompted in the command line to grant Microsoft Edge WebView a loopback exemption ("Allow localhost loopback for Microsoft Edge WebView?"). When prompted, enter Y to allow the exemption. Note that you'll need administrator privileges to allow the exemption. Once allowed, you shouldn't be prompted for an exemption when you sideload Office Add-ins in the future (unless you remove the exemption from your machine). To learn more, see **["We can't open this](https://learn.microsoft.com/en-us/office/troubleshoot/office-suite-issues/cannot-open-add-in-from-localhost) [add-in from localhost" when loading an Office Add-in or using Fiddler](https://learn.microsoft.com/en-us/office/troubleshoot/office-suite-issues/cannot-open-add-in-from-localhost)**.


#### **Tip**

If you're testing your add-in on Mac, run the following command before proceeding. When you run this command, the local web server starts.

command line

npm run dev-server

- To test your add-in in PowerPoint, run the following command in the root directory of your project. This starts the local web server (if it's not already running) and opens PowerPoint with your add-in loaded.
command line npm start

- To test your add-in in PowerPoint on a browser, run the following command in the root directory of your project. When you run this command, the local web server starts. Replace "{url}" with the URL of a PowerPoint document on your OneDrive or a SharePoint library to which you have permissions.
#### 7 **Note**

If you are developing on a Mac, enclose the {url} in single quotation marks. Do *not* do this on Windows.

command line

```
npm run start -- web --document {url}
```
The following are examples.


- npm run start -- web --document https://contoso.sharepoint.com/:t:/g/EZGxP7ksiE5DuxvY638G798BpuhwluxCM fF1WZQj3VYhYQ?e=F4QM1R npm run start -- web --document
	- https://1drv.ms/x/s!jkcH7spkM4EGgcZUgqthk4IK3NOypVw?e=Z6G1qp
- npm run start -- web --document https://contoso-my.sharepointdf.com/:t:/p/user/EQda453DNTpFnl1bFPhOVR0BwlrzetbXvnaRYii2lDr_oQ? e=RSccmNP

If your add-in doesn't sideload in the document, manually sideload it by following the instructions in Manually sideload add-ins to Office on the web.

- 3. If the add-in task pane isn't already open in PowerPoint, choose the **Home** tab, and then choose the **Show Taskpane** button on the ribbon to open the add-in task pane.
- 4. In the task pane, choose the **Insert Image** button to add the image to the current slide.


- 5. When you want to stop the local web server and uninstall the add-in, follow the applicable instructions:
	- To stop the server, run the following command. If you used npm start , the following command also uninstalls the add-in.

command line npm stop

- If you manually sideloaded the add-in, see Remove a sideloaded add-in.
# **Customize user interface (UI) elements**

Complete the following steps to add markup that customizes the task pane UI.

- 1. In the **taskpane.html** file, replace TODO2 and the current header section with the following markup to update the header section and title in the task pane. Note:
	- The styles that begin with ms- are defined by Fabric Core in Office Add-ins, a JavaScript front-end framework for building user experiences for Office. The **taskpane.html** file includes a reference to the Fabric Core stylesheet.

```
HTML
<header id="content-header">
 <div class="ms-Grid ms-bgColor-neutralPrimary">
 <div class="ms-Grid-row">
 <div class="padding ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-
lg12"> <div class="ms-font-xl ms-fontColor-white ms-fontWeight-
semibold">My PowerPoint add-in</div></div>
 </div>
 </div>
</header>
```
- 2. Save all your changes to the project.
## **Test the add-in**

- 1. If the local web server isn't already running, complete the following steps to start the local web server and sideload your add-in.


- Office Add-ins should use HTTPS, not HTTP, even while you're developing. If you're prompted to install a certificate after you run one of the following commands, accept the prompt to install the certificate that the Yeoman generator provides. You may also have to run your command prompt or terminal as an administrator for the changes to be made.
- If this is your first time developing an Office Add-in on your machine, you may be prompted in the command line to grant Microsoft Edge WebView a loopback exemption ("Allow localhost loopback for Microsoft Edge WebView?"). When prompted, enter Y to allow the exemption. Note that you'll need administrator privileges to allow the exemption. Once allowed, you shouldn't be prompted for an exemption when you sideload Office Add-ins in the future (unless you remove the exemption from your machine). To learn more, see **["We can't open this](https://learn.microsoft.com/en-us/office/troubleshoot/office-suite-issues/cannot-open-add-in-from-localhost) [add-in from localhost" when loading an Office Add-in or using Fiddler](https://learn.microsoft.com/en-us/office/troubleshoot/office-suite-issues/cannot-open-add-in-from-localhost)**.

#### **Tip**

If you're testing your add-in on Mac, run the following command before proceeding. When you run this command, the local web server starts.

command line

npm run dev-server

- To test your add-in in PowerPoint, run the following command in the root directory of your project. This starts the local web server (if it's not already running) and opens PowerPoint with your add-in loaded.
command line


```
npm start
```
- To test your add-in in PowerPoint on a browser, run the following command in the root directory of your project. When you run this command, the local web server starts. Replace "{url}" with the URL of a PowerPoint document on your OneDrive or a SharePoint library to which you have permissions.
7 **Note**

If you are developing on a Mac, enclose the {url} in single quotation marks. Do *not* do this on Windows.

command line

npm run start -- web --document {url}

The following are examples.

- npm run start -- web --document
https://contoso.sharepoint.com/:t:/g/EZGxP7ksiE5DuxvY638G798BpuhwluxCM fF1WZQj3VYhYQ?e=F4QM1R

- npm run start -- web --document https://1drv.ms/x/s!jkcH7spkM4EGgcZUgqthk4IK3NOypVw?e=Z6G1qp
- npm run start -- web --document https://contoso-my.sharepointdf.com/:t:/p/user/EQda453DNTpFnl1bFPhOVR0BwlrzetbXvnaRYii2lDr_oQ? e=RSccmNP

If your add-in doesn't sideload in the document, manually sideload it by following the instructions in Manually sideload add-ins to Office on the web.

- 2. If the add-in task pane isn't already open in PowerPoint, select the **Show Taskpane** button on the ribbon to open it.

| இ<br>AutoSave ( Off )                   | 間り・ひ 里<br>O<br>=<br>PowerPoint add-in J G  · Saved to this PC >                                                                                                                                              | સ્ત્ર                                                | 0<br>×        |
|-----------------------------------------|--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|------------------------------------------------------|---------------|
| File<br>Home                            | Insert Draw Design Transitions Animations Slide Show Record Review View Help Script Lab                                                                                                                      | ●<br>ត្រូវបា                                         | ar<br>19<br>> |
| Slides<br>Paste<br>S<br>Clipboard<br>তি | Company<br>AV -<br>U<br>ಿನ<br>Drawing<br>Paragraph<br>Editing<br>Sensitivity<br>Add-ins<br>Dictate<br>A^ A*<br>V<br>><br>V<br>As<br>A<br>Aa v<br>V<br>ン<br>V<br>Font<br>2<br>Voice<br>Sensitivity<br>Add-ins | LOGO<br>Designer<br>Show<br>Taskpane<br>Commands Gro |               |

- 3. Notice that the task pane now contains an updated header section and title.


## **Insert text**

Complete the following steps to add code that inserts text into the title slide which contains an image.

- 1. In the **taskpane.html** file, replace TODO3 with the following markup. This markup defines the **Insert Text** button that will appear within the add-in's task pane.

```
HTML
<button class="ms-Button" id="insert-text">Insert Text</button><br/>
<br/>
```
- 2. In the **taskpane.js** file, replace TODO4 with the following code to assign the event handler for the **Insert Text** button.

```
JavaScript
document.getElementById("insert-text").onclick = () =>
clearMessage(insertText);
```
- 3. In the **taskpane.js** file, replace TODO5 with the following code to define the insertText function. This function inserts text into the current slide.


```
JavaScript
```

```
function insertText() {
 Office.context.document.setSelectedDataAsync("Hello World!",
(asyncResult) => {
 if (asyncResult.status === Office.AsyncResultStatus.Failed) {
 setMessage("Error: " + asyncResult.error.message);
 }
 });
}
```
- 4. Save all your changes to the project.
## **Test the add-in**

- 1. Navigate to the root folder of the project.

```
command line
```
cd "My Office Add-in"

- 2. If the local web server isn't already running, complete the following steps to start the local web server and sideload your add-in.
#### 7 **Note**

- Office Add-ins should use HTTPS, not HTTP, even while you're developing. If you're prompted to install a certificate after you run one of the following commands, accept the prompt to install the certificate that the Yeoman generator provides. You may also have to run your command prompt or terminal as an administrator for the changes to be made.
- If this is your first time developing an Office Add-in on your machine, you may be prompted in the command line to grant Microsoft Edge WebView a loopback exemption ("Allow localhost loopback for Microsoft Edge WebView?"). When prompted, enter Y to allow the exemption. Note that you'll need administrator privileges to allow the exemption. Once allowed, you shouldn't be prompted for an exemption when you sideload Office Add-ins in the future (unless you remove the


exemption from your machine). To learn more, see **["We can't open this](https://learn.microsoft.com/en-us/office/troubleshoot/office-suite-issues/cannot-open-add-in-from-localhost) [add-in from localhost" when loading an Office Add-in or using Fiddler](https://learn.microsoft.com/en-us/office/troubleshoot/office-suite-issues/cannot-open-add-in-from-localhost)**.

#### **Tip**

If you're testing your add-in on Mac, run the following command before proceeding. When you run this command, the local web server starts.

command line

npm run dev-server

- To test your add-in in PowerPoint, run the following command in the root directory of your project. This starts the local web server (if it's not already running) and opens PowerPoint with your add-in loaded.

| command line |  |  |  |  |  |
|--------------|--|--|--|--|--|
| npm start    |  |  |  |  |  |

- To test your add-in in PowerPoint on a browser, run the following command in the root directory of your project. When you run this command, the local web server starts. Replace "{url}" with the URL of a PowerPoint document on your OneDrive or a SharePoint library to which you have permissions.
#### 7 **Note**

If you are developing on a Mac, enclose the {url} in single quotation marks. Do *not* do this on Windows.

command line

npm run start -- web --document {url}


The following are examples.

- npm run start -- web --document https://contoso.sharepoint.com/:t:/g/EZGxP7ksiE5DuxvY638G798BpuhwluxCM fF1WZQj3VYhYQ?e=F4QM1R
- npm run start -- web --document https://1drv.ms/x/s!jkcH7spkM4EGgcZUgqthk4IK3NOypVw?e=Z6G1qp
- npm run start -- web --document https://contoso-my.sharepointdf.com/:t:/p/user/EQda453DNTpFnl1bFPhOVR0BwlrzetbXvnaRYii2lDr_oQ? e=RSccmNP

If your add-in doesn't sideload in the document, manually sideload it by following the instructions in Manually sideload add-ins to Office on the web.

- 3. If the add-in task pane isn't already open in PowerPoint, select the **Show Taskpane** button on the ribbon to open it.
- 4. In the task pane, choose the **Insert Image** button to add the image to the current slide, then choose a design for the slide that contains a text box for the title.


- 5. Put your cursor in the text box on the title slide and then in the task pane, choose the **Insert Text** button to add text to the slide.
# **Get slide metadata**

Complete the following steps to add code that retrieves metadata for the selected slide.


- 1. In the **taskpane.html** file, replace TODO4 with the following markup. This markup defines the **Get Slide Metadata** button that will appear within the add-in's task pane.

```
HTML
<button class="ms-Button" id="get-slide-metadata">Get Slide
Metadata</button><br/><br/>
```
- 2. In the **taskpane.js** file, replace TODO6 with the following code to assign the event handler for the **Get Slide Metadata** button.

```
JavaScript
document.getElementById("get-slide-metadata").onclick = () =>
clearMessage(getSlideMetadata);
```
- 3. In the **taskpane.js** file, replace TODO7 with the following code to define the getSlideMetadata function. This function retrieves metadata for the selected slides and writes it to the Message section in the add-in task pane.

```
JavaScript
function getSlideMetadata() {

Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideR
ange, (asyncResult) => {
 if (asyncResult.status === Office.AsyncResultStatus.Failed) {
 setMessage("Error: " + asyncResult.error.message);
 } else {
 setMessage("Metadata for selected slides: " + 
JSON.stringify(asyncResult.value));
 }
 });
}
```
- 4. Save all your changes to the project.
## **Test the add-in**

- 1. Navigate to the root folder of the project.
command line cd "My Office Add-in"


- 2. If the local web server isn't already running, complete the following steps to start the local web server and sideload your add-in.
#### 7 **Note**

- Office Add-ins should use HTTPS, not HTTP, even while you're developing. If you're prompted to install a certificate after you run one of the following commands, accept the prompt to install the certificate that the Yeoman generator provides. You may also have to run your command prompt or terminal as an administrator for the changes to be made.
- If this is your first time developing an Office Add-in on your machine, you may be prompted in the command line to grant Microsoft Edge WebView a loopback exemption ("Allow localhost loopback for Microsoft Edge WebView?"). When prompted, enter Y to allow the exemption. Note that you'll need administrator privileges to allow the exemption. Once allowed, you shouldn't be prompted for an exemption when you sideload Office Add-ins in the future (unless you remove the exemption from your machine). To learn more, see **["We can't open this](https://learn.microsoft.com/en-us/office/troubleshoot/office-suite-issues/cannot-open-add-in-from-localhost) [add-in from localhost" when loading an Office Add-in or using Fiddler](https://learn.microsoft.com/en-us/office/troubleshoot/office-suite-issues/cannot-open-add-in-from-localhost)**.

#### **Tip**

If you're testing your add-in on Mac, run the following command before proceeding. When you run this command, the local web server starts.

command line

npm run dev-server


- To test your add-in in PowerPoint, run the following command in the root directory of your project. This starts the local web server (if it's not already running) and opens PowerPoint with your add-in loaded.
command line npm start

- To test your add-in in PowerPoint on a browser, run the following command in the root directory of your project. When you run this command, the local web server starts. Replace "{url}" with the URL of a PowerPoint document on your OneDrive or a SharePoint library to which you have permissions.
#### 7 **Note**

If you are developing on a Mac, enclose the {url} in single quotation marks. Do *not* do this on Windows.

command line

npm run start -- web --document {url}

The following are examples.

- npm run start -- web --document https://contoso.sharepoint.com/:t:/g/EZGxP7ksiE5DuxvY638G798BpuhwluxCM fF1WZQj3VYhYQ?e=F4QM1R
- npm run start -- web --document https://1drv.ms/x/s!jkcH7spkM4EGgcZUgqthk4IK3NOypVw?e=Z6G1qp
- npm run start -- web --document https://contoso-my.sharepointdf.com/:t:/p/user/EQda453DNTpFnl1bFPhOVR0BwlrzetbXvnaRYii2lDr_oQ? e=RSccmNP

If your add-in doesn't sideload in the document, manually sideload it by following the instructions in Manually sideload add-ins to Office on the web.

- 3. If the add-in task pane isn't already open in PowerPoint, select the **Show Taskpane** button on the ribbon to open it.


| P<br>AutoSave (1<br>Off )                    | 19 × (1) 4)<br>នា<br>G · Saved to this PC ✓<br>0<br>। ><br>PowerPoint add-in  .                                                                                                            | લ્મુ<br>×                                            |
|----------------------------------------------|--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|------------------------------------------------------|
| File<br>Home                                 | Insert Draw Design Transitions Animations Slide Show Record Review View Help Script Lab                                                                                                    | O<br>瑞<br>a<br>ಲ್ಲ<br>>                              |
| Slides<br>Paste<br>20<br>V<br>Clipboard<br>ি | 00<br>AV -<br>U<br>র্না করে<br>Paragraph<br>Drawing<br>Editing<br>Sensitivity<br>Add-ins<br>Dictate<br>A A A Ao<br>V<br>A × Aa ×<br>V<br>V<br>Font<br>7<br>Voice<br>Sensitivity<br>Add-ins | Logo<br>Designer<br>Show<br>Taskpane<br>Commands Gro |

- 4. In the task pane, choose the **Get Slide Metadata** button to get the metadata for the selected slide. The slide metadata is written in the Message section below the buttons in the task pane. In this case, the slides array within the JSON metadata contains one object that specifies the id , title , and index of the selected slide. If multiple slides had been selected when you retrieved slide metadata, the slides array within the JSON metadata would contain one object for each selected slide.

| E                                | PowerPoint add-in                                    | Q<br>G  · Saved to this PC V          |                                                                                | ્ર<br>0                                  | ×      |
|----------------------------------|------------------------------------------------------|---------------------------------------|--------------------------------------------------------------------------------|------------------------------------------|--------|
| File<br>Home<br>Insert           | Draw Design Transitions Animations Slide Show Record | Review                                | O<br>View Help Script Lab                                                      | 哪<br>್ರ<br>B ×                           | ਕੇ     |
| Slides<br>Clipboard<br>Font<br>2 | Drawing<br>Paragraph<br>><br>N                       | Editing<br>Dictate<br>><br>><br>Voice | Sensitivity<br>Add-ins<br>Designer<br>Sensitivity<br>Add-ins                   | Logo<br>Show<br>Taskpane<br>Commands Gro | >      |
|                                  |                                                      |                                       | My Office Add-in                                                               |                                          | ><br>× |
|                                  |                                                      |                                       | My PowerPoint add-in                                                           |                                          |        |
|                                  |                                                      |                                       | Insert Image                                                                   |                                          |        |
|                                  |                                                      |                                       | Insert Text                                                                    |                                          |        |
|                                  |                                                      |                                       |                                                                                | Get Slide Metadata                       |        |
|                                  | Click to add title                                   |                                       |                                                                                |                                          |        |
|                                  | Click to add subtitle                                |                                       | Message                                                                        |                                          |        |
|                                  |                                                      |                                       | Metadata for selected slides: {"slides":<br>[{"id":256,"title":"","index":1}]} |                                          |        |
|                                  |                                                      |                                       |                                                                                |                                          |        |
|                                  |                                                      |                                       |                                                                                |                                          |        |
|                                  |                                                      |                                       |                                                                                |                                          |        |
|                                  |                                                      |                                       |                                                                                |                                          |        |
|                                  |                                                      |                                       |                                                                                |                                          |        |

# **Navigate between slides**

Complete the following steps to add code that navigates between the slides of a document.

- 1. In the **taskpane.html** file, replace TODO5 with the following markup. This markup defines the four navigation buttons that will appear within the add-in's task pane.
HTML

<button class="ms-Button" id="add-slides">Add Slides</button><br/><br/> <button class="ms-Button" id="go-to-first-slide">Go to First


```
Slide</button><br/><br/>
<button class="ms-Button" id="go-to-next-slide">Go to Next
Slide</button><br/><br/>
<button class="ms-Button" id="go-to-previous-slide">Go to Previous
Slide</button><br/><br/>
<button class="ms-Button" id="go-to-last-slide">Go to Last
Slide</button><br/><br/>
```
- 2. In the **taskpane.js** file, replace TODO8 with the following code to assign the event handlers for the **Add Slides** and four navigation buttons.

```
document.getElementById("add-slides").onclick = () =>
tryCatch(addSlides);
document.getElementById("go-to-first-slide").onclick = () =>
clearMessage(goToFirstSlide);
document.getElementById("go-to-next-slide").onclick = () =>
clearMessage(goToNextSlide);
document.getElementById("go-to-previous-slide").onclick = () =>
clearMessage(goToPreviousSlide);
document.getElementById("go-to-last-slide").onclick = () =>
clearMessage(goToLastSlide);
```
JavaScript

- 3. In the **taskpane.js** file, replace TODO9 with the following code to define the addSlides and navigation functions. Each of these functions uses the goToByIdAsync method to select a slide based upon its position in the document (first, last, previous, and next).

```
JavaScript
async function addSlides() {
 await PowerPoint.run(async function (context) {
 context.presentation.slides.add();
 context.presentation.slides.add();
 await context.sync();
 goToLastSlide();
 setMessage("Success: Slides added.");
 });
}
function goToFirstSlide() {
 Office.context.document.goToByIdAsync(Office.Index.First,
Office.GoToType.Index, (asyncResult) => {
 if (asyncResult.status === Office.AsyncResultStatus.Failed) {
 setMessage("Error: " + asyncResult.error.message);
 }
 });
```


```
}
function goToLastSlide() {
 Office.context.document.goToByIdAsync(Office.Index.Last,
Office.GoToType.Index, (asyncResult) => {
 if (asyncResult.status === Office.AsyncResultStatus.Failed) {
 setMessage("Error: " + asyncResult.error.message);
 }
 });
}
function goToPreviousSlide() {
 Office.context.document.goToByIdAsync(Office.Index.Previous,
Office.GoToType.Index, (asyncResult) => {
 if (asyncResult.status === Office.AsyncResultStatus.Failed) {
 setMessage("Error: " + asyncResult.error.message);
 }
 });
}
function goToNextSlide() {
 Office.context.document.goToByIdAsync(Office.Index.Next,
Office.GoToType.Index, (asyncResult) => {
 if (asyncResult.status === Office.AsyncResultStatus.Failed) {
 setMessage("Error: " + asyncResult.error.message);
 }
 });
}
```
- 4. Save all your changes to the project.
## **Test the add-in**

- 1. Navigate to the root folder of the project.
command line

cd "My Office Add-in"

- 2. If the local web server isn't already running, complete the following steps to start the local web server and sideload your add-in.
#### 7 **Note**

- Office Add-ins should use HTTPS, not HTTP, even while you're developing. If you're prompted to install a certificate after you run one of the following commands, accept the prompt to install the certificate


that the Yeoman generator provides. You may also have to run your command prompt or terminal as an administrator for the changes to be made.

- If this is your first time developing an Office Add-in on your machine, you may be prompted in the command line to grant Microsoft Edge WebView a loopback exemption ("Allow localhost loopback for Microsoft Edge WebView?"). When prompted, enter Y to allow the exemption. Note that you'll need administrator privileges to allow the exemption. Once allowed, you shouldn't be prompted for an exemption when you sideload Office Add-ins in the future (unless you remove the exemption from your machine). To learn more, see **["We can't open this](https://learn.microsoft.com/en-us/office/troubleshoot/office-suite-issues/cannot-open-add-in-from-localhost) [add-in from localhost" when loading an Office Add-in or using Fiddler](https://learn.microsoft.com/en-us/office/troubleshoot/office-suite-issues/cannot-open-add-in-from-localhost)**.
#### **Tip**

If you're testing your add-in on Mac, run the following command before proceeding. When you run this command, the local web server starts.

command line

npm run dev-server

- To test your add-in in PowerPoint, run the following command in the root directory of your project. This starts the local web server (if it's not already running) and opens PowerPoint with your add-in loaded.
command line npm start

- To test your add-in in PowerPoint on a browser, run the following command in the root directory of your project. When you run this command, the local


web server starts. Replace "{url}" with the URL of a PowerPoint document on your OneDrive or a SharePoint library to which you have permissions.

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

- 3. If the add-in task pane isn't already open in PowerPoint, select the **Show Taskpane** button on the ribbon to open it.
- 4. In the task pane, choose the **Add Slides** button. Two new slides are added to the document and the last slide in the document is selected and displayed.


- 5. In the task pane, choose the **Go to First Slide** button. The first slide in the document is selected and displayed.

| 商<br>વે<br>File<br>Insert Draw Design Transitions Animations Slide Show Record<br>Review<br>View Help<br>Script Lab<br>0<br>್ರಿ<br>图 ><br>Home<br>Logo<br>Drawing<br>Slides<br>Editing<br>Paragraph<br>Sensitivity<br>Add-ins<br>Designer<br>Show<br>Dictate<br>Paste<br>V<br>V<br>Taskpane<br><<br>><br>Clipboard<br>Font<br>િ<br>Voice<br>Sensitivity<br>Commands Gro<br>2<br>Add-ins<br>My Office Add-in<br>×<br>><br>My PowerPoint add-in<br><<br>2<br>Insert Image<br>3<br>Insert Text<br>Get Slide Metadata<br>Add Slides<br>Click to add title<br>Go to First Slide<br>Click to add subtitle<br>Go to Next Slide<br>Go to Previous Slide |
|-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 |
|                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 |
|                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 |
|                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 |
|                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 |
|                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 |
|                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 |
|                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 |
|                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 |
|                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 |
|                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 |
|                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 |
| Go to Last Slide                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                |
| Message                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                         |

- 6. In the task pane, choose the **Go to Next Slide** button. The next slide in the document is selected and displayed.


- 7. In the task pane, choose the **Go to Previous Slide** button. The previous slide in the document is selected and displayed.

| P<br>AutoSave     | Q<br>PowerPoint add-in<br>G  · Saved to this PC V                                                                                                         | ્ર                   |                                          | 0           | ×  |
|-------------------|-----------------------------------------------------------------------------------------------------------------------------------------------------------|----------------------|------------------------------------------|-------------|----|
| File<br>Home      | Design Transitions Animations Slide Show Record Review<br>View Help<br>0<br>Insert Draw<br>Script Lab                                                     | தீ                   | ു                                        | 图 >         | વે |
| Clipboard<br>2    | Drawing<br>Editing<br>Paragraph<br>Designer<br>Slides<br>Sensitivity<br>Add-ins<br>Dictate<br>><br>V<br><<br>Font<br>Voice<br>Sensitivity<br>ਨ<br>Add-ins |                      | Logo<br>Show<br>Taskpane<br>Commands Gro |             | >  |
|                   | My Office Add-in                                                                                                                                          |                      |                                          | >           | ×  |
|                   | My PowerPoint add-in                                                                                                                                      |                      |                                          |             | <  |
| 2                 |                                                                                                                                                           | Insert Image         |                                          |             |    |
| 3                 | Insert Text                                                                                                                                               |                      |                                          |             |    |
|                   |                                                                                                                                                           | Get Slide Metadata   |                                          |             |    |
|                   | Add Slides                                                                                                                                                |                      |                                          |             |    |
|                   | Click to add title                                                                                                                                        | Go to First Slide    |                                          |             |    |
|                   | Click to add subtitle                                                                                                                                     | Go to Next Slide     |                                          |             |    |
|                   |                                                                                                                                                           |                      |                                          |             |    |
|                   |                                                                                                                                                           | Go to Previous Slide |                                          |             |    |
|                   |                                                                                                                                                           | Go to Last Slide     |                                          |             |    |
|                   | 1 DD                                                                                                                                                      | Message              |                                          |             |    |
| m<br>Slide 1 of 2 | - Notes - Display Settings<br>402 Accorcibility Invoctianto<br>00<br>日日<br>D<br>ID                                                                        |                      |                                          | + 2006 (A)> |    |

- 8. In the task pane, choose the **Go to Last Slide** button. The last slide in the document is selected and displayed.


- 9. If the web server is running, run the following command when you want to stop the server.

| command line |  |  |  |  |
|--------------|--|--|--|--|
| npm stop     |  |  |  |  |

# **Code samples**

- [Completed PowerPoint add-in tutorial](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/tutorials/powerpoint-tutorial-yo) : The result of completing this tutorial.
# **Next steps**

In this tutorial, you created a PowerPoint add-in that inserts an image, inserts text, gets slide metadata, and navigates between slides. To learn more about building PowerPoint add-ins, continue to the following article.

**PowerPoint add-ins overview**

# **See also**

- Office Add-ins platform overview


- Develop Office Add-ins


# **JavaScript API for PowerPoint**

Article • 05/30/2025

A PowerPoint add-in interacts with objects in PowerPoint by using the Office JavaScript API, which includes two JavaScript object models:

- **PowerPoint JavaScript API**: The [PowerPoint JavaScript API](https://learn.microsoft.com/en-us/javascript/api/powerpoint) provides strongly-typed objects that you can use to access objects in PowerPoint. To learn about the asynchronous nature of the PowerPoint JavaScript APIs and how they work with the presentation, see Using the application-specific API model.
- **Common APIs**: The [Common API](https://learn.microsoft.com/en-us/javascript/api/office) can be used to access features such as UI, dialogs, and client settings that are common across multiple Office applications. To learn more about using the Common API, see Common JavaScript API object model.

# **Learn programming concepts**

See PowerPoint add-ins overview for information about important programming concepts.

# **Learn about API capabilities**

For detailed information about the PowerPoint JavaScript API object model, see the [PowerPoint](https://learn.microsoft.com/en-us/javascript/api/powerpoint) [JavaScript API reference documentation](https://learn.microsoft.com/en-us/javascript/api/powerpoint).

For hands-on experience interacting with content in PowerPoint, complete the PowerPoint addin tutorial.

# **Try out code samples in Script Lab**

Use Script Lab to get started quickly with a collection of built-in samples that show how to complete tasks with the API. You can run the samples in Script Lab to instantly see the result in the task pane or document, examine the samples to learn how the API works, and even use samples to prototype your own add-in.

# **See also**

- PowerPoint add-ins documentation
- PowerPoint add-ins overview
- [PowerPoint JavaScript API reference](https://learn.microsoft.com/en-us/javascript/api/powerpoint)
- [Office client application and platform availability for Office Add-ins](https://learn.microsoft.com/en-us/javascript/api/requirement-sets)


- API Reference documentation


# **powerpoint package**

# **Classes**

ノ **Expand table**

| PowerPoint.Application                  |                                                                                    |
|-----------------------------------------|------------------------------------------------------------------------------------|
| PowerPoint.Binding                      | Represents an Office.js binding that is defined in the presentation.               |
| PowerPoint.Binding                      | Represents the collection of all the binding objects that are part of the          |
| Collection                              | presentation.                                                                      |
| PowerPoint.Border                       | Represents the properties for a table cell border.                                 |
| PowerPoint.Borders                      | Represents the borders for a table cell.                                           |
| PowerPoint.Bullet                       | Represents the bullet formatting properties of a text that is attached to the      |
| Format                                  | PowerPoint.ParagraphFormat.                                                        |
| PowerPoint.Custom<br>Property           | Represents a custom property.                                                      |
| PowerPoint.Custom<br>PropertyCollection | A collection of custom properties.                                                 |
| PowerPoint.Custom<br>XmlPart            | Represents a custom XML part object.                                               |
| PowerPoint.Custom<br>XmlPartCollection  | A collection of custom XML parts.                                                  |
| PowerPoint.Custom                       | A scoped collection of custom XML parts. A scoped collection is the result of some |
| XmlPartScoped                           | operation (such as filtering by namespace). A scoped collection cannot be scoped   |
| Collection                              | any further.                                                                       |
| PowerPoint.<br>Document<br>Properties   | Represents presentation properties.                                                |
| PowerPoint.<br>Hyperlink                | Represents a single hyperlink.                                                     |
| PowerPoint.<br>HyperlinkCollection      | Represents a collection of hyperlinks.                                             |
| PowerPoint.Margins                      | Represents the margins of a table cell.                                            |
| PowerPoint.Page<br>Setup                | Represents the page setup information for the presentation.                        |


| PowerPoint.                                            | Represents the paragraph formatting properties of a text that is attached to the                                                                                                                                                                                            |  |  |
|--------------------------------------------------------|-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|--|--|
| ParagraphFormat                                        | PowerPoint.TextRange.                                                                                                                                                                                                                                                       |  |  |
| PowerPoint.<br>PlaceholderFormat                       | Represents the properties of a placeholder shape.                                                                                                                                                                                                                           |  |  |
| PowerPoint.Presentation                                |                                                                                                                                                                                                                                                                             |  |  |
| PowerPoint.Request<br>Context                          | The RequestContext object facilitates requests to the PowerPoint application. Since<br>the Office add-in and the PowerPoint application run in two different processes,<br>the request context is required to get access to the PowerPoint object model from<br>the add-in. |  |  |
| PowerPoint.Shape                                       | Represents a single shape in the slide.                                                                                                                                                                                                                                     |  |  |
| PowerPoint.Shape<br>Collection                         | Represents the collection of shapes.                                                                                                                                                                                                                                        |  |  |
| PowerPoint.Shape<br>Fill                               | Represents the fill formatting of a shape object.                                                                                                                                                                                                                           |  |  |
| PowerPoint.Shape                                       | Represents the font attributes, such as font name, font size, and color, for a shape's                                                                                                                                                                                      |  |  |
| Font                                                   | TextRange object.                                                                                                                                                                                                                                                           |  |  |
| PowerPoint.Shape                                       | Represents a shape group inside a presentation. To get the corresponding Shape                                                                                                                                                                                              |  |  |
| Group                                                  | object, use ShapeGroup.shape .                                                                                                                                                                                                                                              |  |  |
| PowerPoint.Shape                                       | Represents the line formatting for the shape object. For images and geometric                                                                                                                                                                                               |  |  |
| LineFormat                                             | shapes, line formatting represents the border of the shape.                                                                                                                                                                                                                 |  |  |
| PowerPoint.Shape<br>ScopedCollection                   | Represents a collection of shapes.                                                                                                                                                                                                                                          |  |  |
| PowerPoint.Slide                                       | Represents a single slide of a presentation.                                                                                                                                                                                                                                |  |  |
| PowerPoint.Slide<br>Background                         | Represents a background of a slide.                                                                                                                                                                                                                                         |  |  |
| PowerPoint.Slide<br>BackgroundFill                     | Represents the fill formatting of a slide background object.                                                                                                                                                                                                                |  |  |
| PowerPoint.Slide<br>Background<br>GradientFill         | Represents PowerPoint.SlideBackground gradient fill properties.                                                                                                                                                                                                             |  |  |
| PowerPoint.Slide<br>BackgroundPattern<br>Fill          | Represents PowerPoint.SlideBackground pattern fill properties.                                                                                                                                                                                                              |  |  |
| PowerPoint.Slide<br>BackgroundPicture<br>OrTextureFill | Represents PowerPoint.SlideBackground picture or texture fill properties.                                                                                                                                                                                                   |  |  |


| PowerPoint.Slide<br>BackgroundSolidFill | Represents PowerPoint.SlideBackground solid fill properties.                  |
|-----------------------------------------|-------------------------------------------------------------------------------|
| PowerPoint.Slide<br>Collection          | Represents the collection of slides in the presentation.                      |
| PowerPoint.Slide<br>Layout              | Represents the layout of a slide.                                             |
| PowerPoint.Slide<br>LayoutBackground    | Represents the background of a slide layout.                                  |
| PowerPoint.Slide<br>LayoutCollection    | Represents the collection of layouts provided by the Slide Master for slides. |
| PowerPoint.Slide<br>Master              | Represents the Slide Master of a slide.                                       |
| PowerPoint.Slide<br>MasterBackground    | Represents the background of a slide master.                                  |
| PowerPoint.Slide<br>MasterCollection    | Represents the collection of Slide Masters in the presentation.               |
| PowerPoint.Slide<br>ScopedCollection    | Represents a collection of slides in the presentation.                        |
| PowerPoint.Table                        | Represents a table.                                                           |
| PowerPoint.Table<br>Cell                | Represents a table.                                                           |
| PowerPoint.Table<br>CellCollection      | Represents a collection of table cells.                                       |
| PowerPoint.Table<br>Column              | Represents a column in a table.                                               |
| PowerPoint.Table<br>ColumnCollection    | Represents a collection of table columns.                                     |
| PowerPoint.Table<br>Row                 | Represents a row in a table.                                                  |
| PowerPoint.Table<br>RowCollection       | Represents a collection of table rows.                                        |
| PowerPoint.Table<br>StyleOptions        | Represents the available table style options.                                 |
| PowerPoint.Tag                          | Represents a single tag in the slide.                                         |


| PowerPoint.Tag<br>Collection    | Represents the collection of tags.                                           |
|---------------------------------|------------------------------------------------------------------------------|
| PowerPoint.Text<br>Frame        | Represents the text frame of a shape object.                                 |
| PowerPoint.Text                 | Contains the text that is attached to a shape, in addition to properties and |
| Range                           | methods for manipulating the text.                                           |
| PowerPoint.Theme<br>ColorScheme | Represents a theme color scheme.                                             |

# **Interfaces**

#### ノ **Expand table**

| PowerPoint.Add<br>SlideOptions                                 | Represents the available options when adding a new slide.                                                   |
|----------------------------------------------------------------|-------------------------------------------------------------------------------------------------------------|
| PowerPoint.Border<br>Properties                                | Represents the properties for a table cell border.                                                          |
| PowerPoint.Fill<br>Properties                                  | Represents the fill formatting of a table cell.                                                             |
| PowerPoint.Font<br>Properties                                  | Represents the font attributes, such as font name, size, and color.                                         |
| PowerPoint.Insert<br>SlideOptions                              | Represents the available options when inserting slides.                                                     |
| PowerPoint.<br>Interfaces.Binding<br>CollectionData            | An interface describing the data returned by calling bindingCollection.toJSON() .                           |
| PowerPoint.<br>Interfaces.Binding<br>CollectionLoad<br>Options | Represents the collection of all the binding objects that are part of the presentation.                     |
| PowerPoint.<br>Interfaces.Binding<br>CollectionUpdate<br>Data  | An interface for updating data on the BindingCollection object, for use in<br>bindingCollection.set({  }) . |
| PowerPoint.<br>Interfaces.Binding<br>Data                      | An interface describing the data returned by calling binding.toJSON() .                                     |


| PowerPoint.<br>Interfaces.Binding<br>LoadOptions               | Represents an Office.js binding that is defined in the presentation.                                         |
|----------------------------------------------------------------|--------------------------------------------------------------------------------------------------------------|
| PowerPoint.<br>Interfaces.Border<br>Data                       | An interface describing the data returned by calling border.toJSON() .                                       |
| PowerPoint.<br>Interfaces.Border<br>LoadOptions                | Represents the properties for a table cell border.                                                           |
| PowerPoint.<br>Interfaces.Borders<br>Data                      | An interface describing the data returned by calling borders.toJSON() .                                      |
| PowerPoint.<br>Interfaces.Borders<br>LoadOptions               | Represents the borders for a table cell.                                                                     |
| PowerPoint.<br>Interfaces.Border<br>UpdateData                 | An interface for updating data on the Border object, for use in border.set({  }) .                           |
| PowerPoint.<br>Interfaces.Bullet<br>FormatData                 | An interface describing the data returned by calling bulletFormat.toJSON() .                                 |
| PowerPoint.<br>Interfaces.Bullet<br>FormatLoad<br>Options      | Represents the bullet formatting properties of a text that is attached to the<br>PowerPoint.ParagraphFormat. |
| PowerPoint.<br>Interfaces.Bullet<br>FormatUpdate<br>Data       | An interface for updating data on the BulletFormat object, for use in<br>bulletFormat.set({  }) .            |
| PowerPoint.<br>Interfaces.<br>CollectionLoad<br>Options        | Provides ways to load properties of only a subset of members of a collection.                                |
| PowerPoint.<br>Interfaces.Custom<br>PropertyCollection<br>Data | An interface describing the data returned by calling<br>customPropertyCollection.toJSON() .                  |
| PowerPoint.<br>Interfaces.Custom                               | A collection of custom properties.                                                                           |


| PropertyCollection<br>LoadOptions                                    |                                                                                                                           |
|----------------------------------------------------------------------|---------------------------------------------------------------------------------------------------------------------------|
| PowerPoint.<br>Interfaces.Custom<br>PropertyCollection<br>UpdateData | An interface for updating data on the CustomPropertyCollection object, for use in<br>customPropertyCollection.set({  }) . |
| PowerPoint.<br>Interfaces.Custom<br>PropertyData                     | An interface describing the data returned by calling customProperty.toJSON()                                              |
| PowerPoint.<br>Interfaces.Custom<br>PropertyLoad<br>Options          | Represents a custom property.                                                                                             |
| PowerPoint.<br>Interfaces.Custom<br>PropertyUpdate<br>Data           | An interface for updating data on the CustomProperty object, for use in<br>customProperty.set({  }) .                     |
| PowerPoint.<br>Interfaces.Custom<br>XmlPartCollection<br>Data        | An interface describing the data returned by calling<br>customXmlPartCollection.toJSON() .                                |
| PowerPoint.<br>Interfaces.Custom<br>XmlPartCollection<br>LoadOptions | A collection of custom XML parts.                                                                                         |
| PowerPoint.<br>Interfaces.Custom<br>XmlPartCollection<br>UpdateData  | An interface for updating data on the CustomXmlPartCollection object, for use in<br>customXmlPartCollection.set({  }) .   |
| PowerPoint.<br>Interfaces.Custom<br>XmlPartData                      | An interface describing the data returned by calling customXmlPart.toJSON()                                               |
| PowerPoint.<br>Interfaces.Custom<br>XmlPartLoad<br>Options           | Represents a custom XML part object.                                                                                      |
| PowerPoint.<br>Interfaces.Custom<br>XmlPartScoped<br>CollectionData  | An interface describing the data returned by calling<br>customXmlPartScopedCollection.toJSON() .                          |


| PowerPoint.<br>Interfaces.Custom<br>XmlPartScoped<br>CollectionLoad<br>Options | A scoped collection of custom XML parts. A scoped collection is the result of some<br>operation (such as filtering by namespace). A scoped collection cannot be scoped<br>any further. |
|--------------------------------------------------------------------------------|----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| PowerPoint.<br>Interfaces.Custom<br>XmlPartScoped<br>CollectionUpdate<br>Data  | An interface for updating data on the CustomXmlPartScopedCollection object, for use<br>in customXmlPartScopedCollection.set({<br>}) .                                                  |
| PowerPoint.<br>Interfaces.<br>Document<br>PropertiesData                       | An interface describing the data returned by calling documentProperties.toJSON() .                                                                                                     |
| PowerPoint.<br>Interfaces.<br>Document<br>PropertiesLoad<br>Options            | Represents presentation properties.                                                                                                                                                    |
| PowerPoint.<br>Interfaces.<br>Document<br>PropertiesUpdate<br>Data             | An interface for updating data on the DocumentProperties object, for use in<br>documentProperties.set({  }) .                                                                          |
| PowerPoint.<br>Interfaces.<br>Hyperlink<br>CollectionData                      | An interface describing the data returned by calling hyperlinkCollection.toJSON() .                                                                                                    |
| PowerPoint.<br>Interfaces.<br>Hyperlink<br>CollectionLoad<br>Options           | Represents a collection of hyperlinks.                                                                                                                                                 |
| PowerPoint.<br>Interfaces.<br>Hyperlink<br>CollectionUpdate<br>Data            | An interface for updating data on the HyperlinkCollection object, for use in<br>hyperlinkCollection.set({  }) .                                                                        |
| PowerPoint.<br>Interfaces.<br>HyperlinkData                                    | An interface describing the data returned by calling hyperlink.toJSON() .                                                                                                              |


| PowerPoint.<br>Interfaces.<br>HyperlinkLoad<br>Options       | Represents a single hyperlink.                                                                            |
|--------------------------------------------------------------|-----------------------------------------------------------------------------------------------------------|
| PowerPoint.<br>Interfaces.<br>HyperlinkUpdate<br>Data        | An interface for updating data on the Hyperlink object, for use in hyperlink.set({<br>}) .                |
| PowerPoint.<br>Interfaces.Margins<br>Data                    | An interface describing the data returned by calling margins.toJSON() .                                   |
| PowerPoint.<br>Interfaces.Margins<br>LoadOptions             | Represents the margins of a table cell.                                                                   |
| PowerPoint.<br>Interfaces.Margins<br>UpdateData              | An interface for updating data on the Margins object, for use in margins.set({<br>}) .                    |
| PowerPoint.<br>Interfaces.Page<br>SetupData                  | An interface describing the data returned by calling pageSetup.toJSON() .                                 |
|                                                              |                                                                                                           |
| PowerPoint.<br>Interfaces.Page<br>SetupLoadOptions           | Represents the page setup information for the presentation.                                               |
| PowerPoint.<br>Interfaces.Page<br>SetupUpdateData            | An interface for updating data on the PageSetup object, for use in pageSetup.set({<br>}) .                |
| PowerPoint.<br>Interfaces.<br>ParagraphFormat<br>Data        | An interface describing the data returned by calling paragraphFormat.toJSON() .                           |
| PowerPoint.<br>Interfaces.<br>ParagraphFormat<br>LoadOptions | Represents the paragraph formatting properties of a text that is attached to the<br>PowerPoint.TextRange. |
| PowerPoint.<br>Interfaces.<br>ParagraphFormat<br>UpdateData  | An interface for updating data on the ParagraphFormat object, for use in<br>paragraphFormat.set({  }) .   |


| Placeholder<br>FormatData                                          |                                                                                                         |  |
|--------------------------------------------------------------------|---------------------------------------------------------------------------------------------------------|--|
| PowerPoint.<br>Interfaces.<br>Placeholder<br>FormatLoad<br>Options | Represents the properties of a placeholder shape.                                                       |  |
| PowerPoint.<br>Interfaces.<br>PresentationData                     | An interface describing the data returned by calling presentation.toJSON()                              |  |
| PowerPoint.Interfaces.PresentationLoadOptions                      |                                                                                                         |  |
| PowerPoint.<br>Interfaces.Shape<br>CollectionData                  | An interface describing the data returned by calling shapeCollection.toJSON()                           |  |
| PowerPoint.<br>Interfaces.Shape<br>CollectionLoad<br>Options       | Represents the collection of shapes.                                                                    |  |
| PowerPoint.<br>Interfaces.Shape<br>CollectionUpdate<br>Data        | An interface for updating data on the ShapeCollection object, for use in<br>shapeCollection.set({  }) . |  |
| PowerPoint.<br>Interfaces.Shape<br>Data                            | An interface describing the data returned by calling shape.toJSON()                                     |  |
| PowerPoint.<br>Interfaces.Shape<br>FillData                        | An interface describing the data returned by calling shapeFill.toJSON()                                 |  |
| PowerPoint.<br>Interfaces.Shape<br>FillLoadOptions                 | Represents the fill formatting of a shape object.                                                       |  |
| PowerPoint.<br>Interfaces.Shape<br>FillUpdateData                  | An interface for updating data on the ShapeFill object, for use in shapeFill.set({<br>}) .              |  |
| PowerPoint.<br>Interfaces.Shape<br>FontData                        | An interface describing the data returned by calling shapeFont.toJSON()                                 |  |
| PowerPoint.                                                        | Represents the font attributes, such as font name, font size, and color, for a shape's                  |  |
| Interfaces.Shape                                                   | TextRange object.                                                                                       |  |


| FontLoadOptions                                                    |                                                                                                                                              |
|--------------------------------------------------------------------|----------------------------------------------------------------------------------------------------------------------------------------------|
| PowerPoint.<br>Interfaces.Shape<br>FontUpdateData                  | An interface for updating data on the ShapeFont object, for use in shapeFont.set({<br>}) .                                                   |
| PowerPoint.<br>Interfaces.Shape<br>GroupData                       | An interface describing the data returned by calling shapeGroup.toJSON()                                                                     |
| PowerPoint.<br>Interfaces.Shape<br>GroupLoad<br>Options            | Represents a shape group inside a presentation. To get the corresponding Shape<br>object, use ShapeGroup.shape                               |
| PowerPoint.<br>Interfaces.Shape<br>LineFormatData                  | An interface describing the data returned by calling shapeLineFormat.toJSON()                                                                |
| PowerPoint.<br>Interfaces.Shape<br>LineFormatLoad<br>Options       | Represents the line formatting for the shape object. For images and geometric<br>shapes, line formatting represents the border of the shape. |
| PowerPoint.<br>Interfaces.Shape<br>LineFormatUpdate<br>Data        | An interface for updating data on the ShapeLineFormat object, for use in<br>shapeLineFormat.set({  }) .                                      |
| PowerPoint.<br>Interfaces.Shape<br>LoadOptions                     | Represents a single shape in the slide.                                                                                                      |
| PowerPoint.<br>Interfaces.Shape<br>ScopedCollection<br>Data        | An interface describing the data returned by calling<br>shapeScopedCollection.toJSON() .                                                     |
| PowerPoint.<br>Interfaces.Shape<br>ScopedCollection<br>LoadOptions | Represents a collection of shapes.                                                                                                           |
| PowerPoint.<br>Interfaces.Shape<br>ScopedCollection<br>UpdateData  | An interface for updating data on the ShapeScopedCollection object, for use in<br>shapeScopedCollection.set({  }) .                          |
| PowerPoint.<br>Interfaces.Shape<br>UpdateData                      | An interface for updating data on the Shape object, for use in shape.set({<br>}) .                                                           |


| PowerPoint.<br>Interfaces.Slide<br>BackgroundData                            | An interface describing the data returned by calling slideBackground.toJSON() .                                                 |
|------------------------------------------------------------------------------|---------------------------------------------------------------------------------------------------------------------------------|
| PowerPoint.<br>Interfaces.Slide<br>BackgroundFill<br>Data                    | An interface describing the data returned by calling slideBackgroundFill.toJSON() .                                             |
| PowerPoint.<br>Interfaces.Slide<br>BackgroundFill<br>LoadOptions             | Represents the fill formatting of a slide background object.                                                                    |
| PowerPoint.<br>Interfaces.Slide<br>Background<br>GradientFillData            | An interface describing the data returned by calling<br>slideBackgroundGradientFill.toJSON() .                                  |
| PowerPoint.<br>Interfaces.Slide<br>Background<br>GradientFillLoad<br>Options | Represents PowerPoint.SlideBackground gradient fill properties.                                                                 |
| PowerPoint.<br>Interfaces.Slide<br>Background<br>GradientFillUpdate<br>Data  | An interface for updating data on the SlideBackgroundGradientFill object, for use in<br>slideBackgroundGradientFill.set({  }) . |
| PowerPoint.<br>Interfaces.Slide<br>BackgroundLoad<br>Options                 | Represents a background of a slide.                                                                                             |
| PowerPoint.<br>Interfaces.Slide<br>Background<br>PatternFillData             | An interface describing the data returned by calling<br>slideBackgroundPatternFill.toJSON() .                                   |
| PowerPoint.<br>Interfaces.Slide<br>Background<br>PatternFillLoad<br>Options  | Represents PowerPoint.SlideBackground pattern fill properties.                                                                  |
| PowerPoint.<br>Interfaces.Slide<br>Background                                | An interface for updating data on the SlideBackgroundPatternFill object, for use in<br>slideBackgroundPatternFill.set({  }) .   |


| PatternFillUpdate<br>Data                                                            |                                                                                                                                                   |
|--------------------------------------------------------------------------------------|---------------------------------------------------------------------------------------------------------------------------------------------------|
| PowerPoint.<br>Interfaces.Slide<br>Background<br>PictureOrTexture<br>FillData        | An interface describing the data returned by calling<br>slideBackgroundPictureOrTextureFill.toJSON() .                                            |
| PowerPoint.<br>Interfaces.Slide<br>Background<br>PictureOrTexture<br>FillLoadOptions | Represents PowerPoint.SlideBackground picture or texture fill properties.                                                                         |
| PowerPoint.<br>Interfaces.Slide<br>Background<br>PictureOrTexture<br>FillUpdateData  | An interface for updating data on the SlideBackgroundPictureOrTextureFill object,<br>for use in slideBackgroundPictureOrTextureFill.set({<br>}) . |
| PowerPoint.<br>Interfaces.Slide<br>BackgroundSolid<br>FillData                       | An interface describing the data returned by calling<br>slideBackgroundSolidFill.toJSON() .                                                       |
| PowerPoint.<br>Interfaces.Slide<br>BackgroundSolid<br>FillLoadOptions                | Represents PowerPoint.SlideBackground solid fill properties.                                                                                      |
| PowerPoint.<br>Interfaces.Slide<br>BackgroundSolid<br>FillUpdateData                 | An interface for updating data on the SlideBackgroundSolidFill object, for use in<br>slideBackgroundSolidFill.set({  }) .                         |
| PowerPoint.<br>Interfaces.Slide<br>Background<br>UpdateData                          | An interface for updating data on the SlideBackground object, for use in<br>slideBackground.set({  }) .                                           |
| PowerPoint.<br>Interfaces.Slide<br>CollectionData                                    | An interface describing the data returned by calling slideCollection.toJSON()                                                                     |
| PowerPoint.<br>Interfaces.Slide<br>CollectionLoad<br>Options                         | Represents the collection of slides in the presentation.                                                                                          |


| PowerPoint.<br>Interfaces.Slide<br>CollectionUpdate<br>Data        | An interface for updating data on the SlideCollection object, for use in<br>slideCollection.set({  }) .             |
|--------------------------------------------------------------------|---------------------------------------------------------------------------------------------------------------------|
| PowerPoint.<br>Interfaces.Slide<br>Data                            | An interface describing the data returned by calling slide.toJSON() .                                               |
| PowerPoint.<br>Interfaces.Slide<br>LayoutBackground<br>Data        | An interface describing the data returned by calling<br>slideLayoutBackground.toJSON() .                            |
| PowerPoint.<br>Interfaces.Slide<br>LayoutBackground<br>LoadOptions | Represents the background of a slide layout.                                                                        |
| PowerPoint.<br>Interfaces.Slide<br>LayoutBackground<br>UpdateData  | An interface for updating data on the SlideLayoutBackground object, for use in<br>slideLayoutBackground.set({  }) . |
| PowerPoint.<br>Interfaces.Slide<br>LayoutCollection<br>Data        | An interface describing the data returned by calling<br>slideLayoutCollection.toJSON() .                            |
| PowerPoint.<br>Interfaces.Slide<br>LayoutCollection<br>LoadOptions | Represents the collection of layouts provided by the Slide Master for slides.                                       |
| PowerPoint.<br>Interfaces.Slide<br>LayoutCollection<br>UpdateData  | An interface for updating data on the SlideLayoutCollection object, for use in<br>slideLayoutCollection.set({  }) . |
| PowerPoint.<br>Interfaces.Slide<br>LayoutData                      | An interface describing the data returned by calling slideLayout.toJSON() .                                         |
| PowerPoint.<br>Interfaces.Slide<br>LayoutLoad<br>Options           | Represents the layout of a slide.                                                                                   |
| PowerPoint.<br>Interfaces.Slide<br>LoadOptions                     | Represents a single slide of a presentation.                                                                        |


| PowerPoint.<br>Interfaces.Slide<br>Master<br>BackgroundData            | An interface describing the data returned by calling<br>slideMasterBackground.toJSON() .                            |
|------------------------------------------------------------------------|---------------------------------------------------------------------------------------------------------------------|
| PowerPoint.<br>Interfaces.Slide<br>Master<br>BackgroundLoad<br>Options | Represents the background of a slide master.                                                                        |
| PowerPoint.<br>Interfaces.Slide<br>MasterCollection<br>Data            | An interface describing the data returned by calling<br>slideMasterCollection.toJSON() .                            |
| PowerPoint.<br>Interfaces.Slide<br>MasterCollection<br>LoadOptions     | Represents the collection of Slide Masters in the presentation.                                                     |
| PowerPoint.<br>Interfaces.Slide<br>MasterCollection<br>UpdateData      | An interface for updating data on the SlideMasterCollection object, for use in<br>slideMasterCollection.set({  }) . |
| PowerPoint.<br>Interfaces.Slide<br>MasterData                          | An interface describing the data returned by calling slideMaster.toJSON() .                                         |
| PowerPoint.<br>Interfaces.Slide<br>MasterLoad<br>Options               | Represents the Slide Master of a slide.                                                                             |
| PowerPoint.<br>Interfaces.Slide<br>ScopedCollection<br>Data            | An interface describing the data returned by calling<br>slideScopedCollection.toJSON() .                            |
| PowerPoint.<br>Interfaces.Slide<br>ScopedCollection<br>LoadOptions     | Represents a collection of slides in the presentation.                                                              |
| PowerPoint.<br>Interfaces.Slide<br>ScopedCollection<br>UpdateData      | An interface for updating data on the SlideScopedCollection object, for use in<br>slideScopedCollection.set({  }) . |


| PowerPoint.<br>Interfaces.Table<br>CellCollectionData              | An interface describing the data returned by calling tableCellCollection.toJSON() .                                 |
|--------------------------------------------------------------------|---------------------------------------------------------------------------------------------------------------------|
| PowerPoint.<br>Interfaces.Table<br>CellCollectionLoad<br>Options   | Represents a collection of table cells.                                                                             |
| PowerPoint.<br>Interfaces.Table<br>CellCollection<br>UpdateData    | An interface for updating data on the TableCellCollection object, for use in<br>tableCellCollection.set({  }) .     |
| PowerPoint.<br>Interfaces.Table<br>CellData                        | An interface describing the data returned by calling tableCell.toJSON() .                                           |
| PowerPoint.<br>Interfaces.Table<br>CellLoadOptions                 | Represents a table.                                                                                                 |
| PowerPoint.<br>Interfaces.Table<br>CellUpdateData                  | An interface for updating data on the TableCell object, for use in tableCell.set({<br>}) .                          |
| PowerPoint.<br>Interfaces.Table<br>ColumnCollection<br>Data        | An interface describing the data returned by calling<br>tableColumnCollection.toJSON() .                            |
| PowerPoint.<br>Interfaces.Table<br>ColumnCollection<br>LoadOptions | Represents a collection of table columns.                                                                           |
| PowerPoint.<br>Interfaces.Table<br>ColumnCollection<br>UpdateData  | An interface for updating data on the TableColumnCollection object, for use in<br>tableColumnCollection.set({  }) . |
| PowerPoint.<br>Interfaces.Table<br>ColumnData                      | An interface describing the data returned by calling tableColumn.toJSON() .                                         |
| PowerPoint.<br>Interfaces.Table<br>ColumnLoad<br>Options           | Represents a column in a table.                                                                                     |


| PowerPoint.<br>Interfaces.Table<br>ColumnUpdate<br>Data         | An interface for updating data on the TableColumn object, for use in<br>tableColumn.set({  }) .               |
|-----------------------------------------------------------------|---------------------------------------------------------------------------------------------------------------|
| PowerPoint.<br>Interfaces.Table<br>Data                         | An interface describing the data returned by calling table.toJSON() .                                         |
| PowerPoint.<br>Interfaces.Table<br>LoadOptions                  | Represents a table.                                                                                           |
| PowerPoint.<br>Interfaces.Table<br>RowCollection<br>Data        | An interface describing the data returned by calling tableRowCollection.toJSON() .                            |
| PowerPoint.<br>Interfaces.Table<br>RowCollection<br>LoadOptions | Represents a collection of table rows.                                                                        |
| PowerPoint.<br>Interfaces.Table<br>RowCollection<br>UpdateData  | An interface for updating data on the TableRowCollection object, for use in<br>tableRowCollection.set({  }) . |
| PowerPoint.<br>Interfaces.Table<br>RowData                      | An interface describing the data returned by calling tableRow.toJSON() .                                      |
| PowerPoint.<br>Interfaces.Table<br>RowLoadOptions               | Represents a row in a table.                                                                                  |
| PowerPoint.<br>Interfaces.Table<br>RowUpdateData                | An interface for updating data on the TableRow object, for use in tableRow.set({<br>}) .                      |
| PowerPoint.<br>Interfaces.Table<br>StyleOptionsData             | An interface describing the data returned by calling tableStyleOptions.toJSON() .                             |
| PowerPoint.<br>Interfaces.Table<br>StyleOptionsLoad<br>Options  | Represents the available table style options.                                                                 |
| PowerPoint.                                                     | An interface for updating data on the TableStyleOptions object, for use in                                    |
| Interfaces.Table                                                | tableStyleOptions.set({  }) .                                                                                 |


| StyleOptions<br>UpdateData                                 |                                                                                                                    |
|------------------------------------------------------------|--------------------------------------------------------------------------------------------------------------------|
| PowerPoint.<br>Interfaces.Tag<br>CollectionData            | An interface describing the data returned by calling tagCollection.toJSON()                                        |
| PowerPoint.<br>Interfaces.Tag<br>CollectionLoad<br>Options | Represents the collection of tags.                                                                                 |
| PowerPoint.<br>Interfaces.Tag<br>CollectionUpdate<br>Data  | An interface for updating data on the TagCollection object, for use in<br>tagCollection.set({  }) .                |
| PowerPoint.                                                | An interface describing the data returned by calling tag.toJSON()                                                  |
| Interfaces.TagData                                         |                                                                                                                    |
| PowerPoint.<br>Interfaces.TagLoad<br>Options               | Represents a single tag in the slide.                                                                              |
| PowerPoint.<br>Interfaces.Tag<br>UpdateData                | An interface for updating data on the Tag object, for use in tag.set({<br>}) .                                     |
| PowerPoint.<br>Interfaces.Text<br>FrameData                | An interface describing the data returned by calling textFrame.toJSON()                                            |
| PowerPoint.<br>Interfaces.Text<br>FrameLoad<br>Options     | Represents the text frame of a shape object.                                                                       |
| PowerPoint.<br>Interfaces.Text<br>FrameUpdateData          | An interface for updating data on the TextFrame object, for use in textFrame.set({<br>}) .                         |
| PowerPoint.<br>Interfaces.Text<br>RangeData                | An interface describing the data returned by calling textRange.toJSON()                                            |
| PowerPoint.<br>Interfaces.Text<br>RangeLoad<br>Options     | Contains the text that is attached to a shape, in addition to properties and methods<br>for manipulating the text. |


| PowerPoint.<br>Interfaces.Text<br>RangeUpdateData                 | An interface for updating data on the TextRange object, for use in textRange.set({<br>}) .                                                                                                                                                                                                                                                                                                                                     |
|-------------------------------------------------------------------|--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| PowerPoint.Shape<br>AddOptions                                    | Represents the available options when adding shapes.                                                                                                                                                                                                                                                                                                                                                                           |
| PowerPoint.Shape<br>GetImageOptions                               | Represents the available options when getting an image of a shape. The image is<br>scaled to fit into the desired dimensions. If width and height aren't specified, the<br>true size of the shape is used. If only one of either width or height is specified, the<br>other will be calculated to preserve aspect ratio. The resulting dimensions will<br>automatically be clamped to the maximum supported size if too large. |
| PowerPoint.Slide<br>Background<br>GradientFill<br>Options         | Represents the available options for setting a PowerPoint.SlideBackground gradient<br>fill.                                                                                                                                                                                                                                                                                                                                    |
| PowerPoint.Slide<br>Background<br>PatternFillOptions              | Represents the available options for setting a PowerPoint.SlideBackground pattern<br>fill.                                                                                                                                                                                                                                                                                                                                     |
| PowerPoint.Slide<br>Background<br>PictureOrTexture<br>FillOptions | Represents PowerPoint.SlideBackground picture or texture fill options.                                                                                                                                                                                                                                                                                                                                                         |
| PowerPoint.Slide<br>BackgroundSolid<br>FillOptions                | Represents the available options for setting a PowerPoint.SlideBackground solid fill.                                                                                                                                                                                                                                                                                                                                          |
| PowerPoint.Slide<br>GetImageOptions                               | Represents the available options when getting an image of a slide.                                                                                                                                                                                                                                                                                                                                                             |
| PowerPoint.Table<br>AddOptions                                    | Represents the available options when adding a table.                                                                                                                                                                                                                                                                                                                                                                          |
| PowerPoint.Table<br>CellBorders                                   | Represents the borders of a table cell.                                                                                                                                                                                                                                                                                                                                                                                        |
| PowerPoint.Table<br>CellMargins                                   | Represents the margins of a table cell.                                                                                                                                                                                                                                                                                                                                                                                        |
| PowerPoint.Table<br>CellProperties                                | Represents the table cell properties to update.                                                                                                                                                                                                                                                                                                                                                                                |
| PowerPoint.Table<br>ClearOptions                                  | Represents the available options when clearing a table.                                                                                                                                                                                                                                                                                                                                                                        |
| PowerPoint.Table<br>ColumnProperties                              | Provides the table column properties.                                                                                                                                                                                                                                                                                                                                                                                          |


| PowerPoint.Table<br>MergedArea<br>Properties | Represents the properties of a merged area of cells in a table.                |
|----------------------------------------------|--------------------------------------------------------------------------------|
| PowerPoint.Table<br>RowProperties            | Provides the table row properties.                                             |
| PowerPoint.Text<br>Run                       | Represents a sequence of one or more characters with the same font attributes. |

## **Enums**

#### ノ **Expand table**

| PowerPoint.<br>BindingType                          | Represents the possible binding types.                                                    |  |
|-----------------------------------------------------|-------------------------------------------------------------------------------------------|--|
| PowerPoint.<br>BulletStyle                          | Specifies the style of a bullet.                                                          |  |
| PowerPoint.<br>BulletType                           | Specifies the type of a bullet.                                                           |  |
| PowerPoint.<br>ConnectorType                        | Specifies the connector type for line shapes.                                             |  |
| PowerPoint.<br>Document<br>PropertyType             | Specifies the document property type for custom properties.                               |  |
| PowerPoint.ErrorCodes                               |                                                                                           |  |
| PowerPoint.<br>GeometricShape<br>Type               | Specifies the shape type for a GeometricShape object.                                     |  |
| PowerPoint.<br>InsertSlide<br>Formatting            | Specifies the formatting options for when slides are inserted.                            |  |
| PowerPoint.<br>Paragraph<br>Horizontal<br>Alignment | Represents the horizontal alignment of the PowerPoint.TextFrame in a<br>PowerPoint.Shape. |  |
| PowerPoint.<br>PlaceholderType                      | Specifies the type of a placeholder.                                                      |  |


| PowerPoint.<br>ShapeAutoSize                       | Determines the type of automatic sizing allowed.                                                                                                                                                                                                                                                                                                           |  |
|----------------------------------------------------|------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|--|
| PowerPoint.<br>ShapeFillType                       | Specifies a shape's fill type.                                                                                                                                                                                                                                                                                                                             |  |
| PowerPoint.<br>ShapeFont<br>UnderlineStyle         | The type of underline applied to a font.                                                                                                                                                                                                                                                                                                                   |  |
| PowerPoint.<br>ShapeGetImage<br>FormatType         | Represents the format of an image.                                                                                                                                                                                                                                                                                                                         |  |
| PowerPoint.<br>ShapeLineDash<br>Style              | Specifies the dash style for a line.                                                                                                                                                                                                                                                                                                                       |  |
| PowerPoint.<br>ShapeLineStyle                      | Specifies the style for a line.                                                                                                                                                                                                                                                                                                                            |  |
| PowerPoint.<br>ShapeType                           | Specifies the type of a shape.                                                                                                                                                                                                                                                                                                                             |  |
| PowerPoint.                                        | Use with setZOrder to move the specified shape up or down the collection's z-order,                                                                                                                                                                                                                                                                        |  |
| ShapeZOrder                                        | which shifts it in front of or behind other shapes.                                                                                                                                                                                                                                                                                                        |  |
| PowerPoint.Slide<br>BackgroundFill<br>Type         | Specifies the fill type for a PowerPoint.SlideBackground.                                                                                                                                                                                                                                                                                                  |  |
| PowerPoint.Slide<br>Background<br>GradientFillType | Specifies the gradient fill type for a PowerPoint.SlideBackgroundGradientFill.                                                                                                                                                                                                                                                                             |  |
| PowerPoint.Slide<br>Background<br>PatternFillType  | Specifies the pattern fill type for a PowerPoint.SlideBackgroundPatternFill.                                                                                                                                                                                                                                                                               |  |
| PowerPoint.Slide<br>LayoutType                     | Specifies the type of a slide layout.                                                                                                                                                                                                                                                                                                                      |  |
| PowerPoint.<br>TableStyle                          | Represents the available built-in table styles.                                                                                                                                                                                                                                                                                                            |  |
| PowerPoint.Text<br>Vertical<br>Alignment           | Represents the vertical alignment of a PowerPoint.TextFrame in a PowerPoint.Shape. If<br>one of the centered options is selected, the contents of the TextFrame will be centered<br>horizontally within the Shape as a group. To change the horizontal alignment of a text,<br>see PowerPoint.ParagraphFormat and PowerPoint.ParagraphHorizontalAlignment. |  |
| PowerPoint.                                        | Specifies the theme colors used in PowerPoint.                                                                                                                                                                                                                                                                                                             |  |


# **Functions**

#### ノ **Expand table**

| PowerPoint.create<br>Presentation(base64File) | Creates and opens a new presentation. Optionally, the presentation can be<br>prepopulated with a Base64-encoded .pptx file.<br>[ API set: PowerPointApi 1.1 ]                                                                                                              |
|-----------------------------------------------|----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| PowerPoint.run(batch)                         | Executes a batch script that performs actions on the PowerPoint object<br>model, using a new RequestContext. When the promise is resolved, any<br>tracked objects that were automatically allocated during execution will be<br>released.                                  |
| PowerPoint.run(object,<br>batch)              | Executes a batch script that performs actions on the PowerPoint object<br>model, using the RequestContext of a previously-created API object. When<br>the promise is resolved, any tracked objects that were automatically<br>allocated during execution will be released. |
| PowerPoint.run(objects,                       | Executes a batch script that performs actions on the PowerPoint object                                                                                                                                                                                                     |
| batch)                                        | model, using the RequestContext of previously-created API objects.                                                                                                                                                                                                         |

# **Function Details**

## **PowerPoint.createPresentation(base64File)**

Creates and opens a new presentation. Optionally, the presentation can be prepopulated with a Base64-encoded .pptx file.

#### [ [API set: PowerPointApi 1.1](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/powerpoint/powerpoint-api-requirement-sets) ]

TypeScript

export function createPresentation(base64File?: string): Promise<void>;

#### **Parameters**

**base64File** string

Optional. The Base64-encoded .pptx file. The default value is null. The maximum length of the string is 71,680,000 characters.


#### **Returns**

Promise<void>

### **Examples**

```
TypeScript
const myFile = <HTMLInputElement>document.getElementById("file");
const reader = new FileReader();
reader.onload = (event) => {
 // Remove the metadata before the base64-encoded string.
 const startIndex = reader.result.toString().indexOf("base64,");
 const copyBase64 = reader.result.toString().substr(startIndex + 7);
 PowerPoint.createPresentation(copyBase64);
};
// Read in the file as a data URL so we can parse the base64-encoded string.
reader.readAsDataURL(myFile.files[0]);
```
## **PowerPoint.run(batch)**

Executes a batch script that performs actions on the PowerPoint object model, using a new RequestContext. When the promise is resolved, any tracked objects that were automatically allocated during execution will be released.

TypeScript

```
export function run<T>(batch: (context: PowerPoint.RequestContext) =>
OfficeExtension.IPromise<T>): OfficeExtension.IPromise<T>;
```
#### **Parameters**

**batch** (context: [PowerPoint.RequestContext](https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.requestcontext?view=powerpoint-js-preview)) => [OfficeExtension.IPromise](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.ipromise?view=powerpoint-js-preview)<T>

A function that takes in a RequestContext and returns a promise (typically, just the result of "context.sync()"). The context parameter facilitates requests to the PowerPoint application. Since the Office add-in and the PowerPoint application run in two different processes, the RequestContext is required to get access to the PowerPoint object model from the add-in.

### **Returns**

[OfficeExtension.IPromise<](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.ipromise?view=powerpoint-js-preview)T>


### **PowerPoint.run(object, batch)**

Executes a batch script that performs actions on the PowerPoint object model, using the RequestContext of a previously-created API object. When the promise is resolved, any tracked objects that were automatically allocated during execution will be released.

TypeScript

```
export function run<T>(object: OfficeExtension.ClientObject, batch: (context:
PowerPoint.RequestContext) => OfficeExtension.IPromise<T>): 
OfficeExtension.IPromise<T>;
```
#### **Parameters**

#### **object** [OfficeExtension.ClientObject](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientobject?view=powerpoint-js-preview)

A previously-created API object. The batch will use the same RequestContext as the passedin object, which means that any changes applied to the object will be picked up by "context.sync()".

#### **batch** (context: [PowerPoint.RequestContext](https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.requestcontext?view=powerpoint-js-preview)) => [OfficeExtension.IPromise](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.ipromise?view=powerpoint-js-preview)<T>

A function that takes in a RequestContext and returns a promise (typically, just the result of "context.sync()"). The context parameter facilitates requests to the PowerPoint application. Since the Office add-in and the PowerPoint application run in two different processes, the RequestContext is required to get access to the PowerPoint object model from the add-in.

#### **Returns**

[OfficeExtension.IPromise<](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.ipromise?view=powerpoint-js-preview)T>

## **PowerPoint.run(objects, batch)**

Executes a batch script that performs actions on the PowerPoint object model, using the RequestContext of previously-created API objects.

```
TypeScript
```

```
export function run<T>(objects: OfficeExtension.ClientObject[], batch:
(context: PowerPoint.RequestContext) => OfficeExtension.IPromise<T>): 
OfficeExtension.IPromise<T>;
```
#### **Parameters**


#### **objects** [OfficeExtension.ClientObject](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientobject?view=powerpoint-js-preview)[]

An array of previously-created API objects. The array will be validated to make sure that all of the objects share the same context. The batch will use this shared RequestContext, which means that any changes applied to these objects will be picked up by "context.sync()".

#### **batch** (context: [PowerPoint.RequestContext](https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.requestcontext?view=powerpoint-js-preview)) => [OfficeExtension.IPromise](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.ipromise?view=powerpoint-js-preview)<T>

A function that takes in a RequestContext and returns a promise (typically, just the result of "context.sync()"). The context parameter facilitates requests to the PowerPoint application. Since the Office add-in and the PowerPoint application run in two different processes, the RequestContext is required to get access to the PowerPoint object model from the add-in.

#### **Returns**

[OfficeExtension.IPromise<](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.ipromise?view=powerpoint-js-preview)T>


# **PowerPoint JavaScript object model in Office Add-ins**

06/20/2025

This article describes concepts that are fundamental to using the PowerPoint JavaScript API to build add-ins.

# **Office.js APIs for PowerPoint**

A PowerPoint add-in interacts with objects in PowerPoint by using the Office JavaScript API. This includes two JavaScript object models:

- **PowerPoint JavaScript API**: The [PowerPoint JavaScript API](https://learn.microsoft.com/en-us/javascript/api/powerpoint) provides strongly-typed objects that work with the presentation, slides, tables, shapes, formatting, and more. To learn about the asynchronous nature of the PowerPoint APIs and how they work with the presentation, see Using the application-specific API model.
- **Common APIs**: The [Common API](https://learn.microsoft.com/en-us/javascript/api/office) give access to features such as UI, dialogs, and client settings that are common across multiple Office applications. To learn more about using the Common API, see Common JavaScript API object model.

While you'll likely use the PowerPoint JavaScript API to develop the majority of functionality in add-ins that target PowerPoint, you'll also use objects in the Common API. For example:

- [Office.Context](https://learn.microsoft.com/en-us/javascript/api/office/office.context): The Office.Context object represents the runtime environment of the add-in and provides access to key objects of the API. It consists of presentation configuration details such as contentLanguage and officeTheme and also provides information about the add-in's runtime environment such as host and platform . Additionally, it provides the requirements.isSetSupported() method, which you can use to check whether a specified requirement set is supported by the PowerPoint application where the add-in is running.
- [Office.Document:](https://learn.microsoft.com/en-us/javascript/api/office/office.document) The Office.Document object provides the getFileAsync() method, which you can use to download the PowerPoint file where the add-in is running. It also provides the getActiveViewAsync() method, which you can use to check whether the presentation is in a "read" or "edit" view. "edit" corresponds to any of the views in which you can edit slides: Normal, Slide Sorter, or Outline View. "read" corresponds to either Slide Show or Reading View.

# **PowerPoint-specific object model**


To understand the PowerPoint APIs, you must understand how key components of a presentation are related to one another.

- The presentation contains slides and presentation-level entities such as settings and custom XML parts.
- A slide contains content like shapes, text, and tables.
- A layout determines how a slide's content is organized and displayed.

For the full set of objects supported by the PowerPoint JavaScript API, see [PowerPoint](https://learn.microsoft.com/en-us/javascript/api/powerpoint) [JavaScript API.](https://learn.microsoft.com/en-us/javascript/api/powerpoint)

# **See also**

- PowerPoint JavaScript API overview
- Build your first PowerPoint add-in
- PowerPoint add-in tutorial
- [PowerPoint JavaScript API reference](https://learn.microsoft.com/en-us/javascript/api/powerpoint)
- [Learn about the Microsoft 365 Developer Program](https://aka.ms/m365devprogram)


# **Add and delete slides in PowerPoint**

Article • 07/21/2022

A PowerPoint add-in can add slides to the presentation and optionally specify which slide master, and which layout of the master, is used for the new slide. The add-in can also delete slides.

The APIs for adding slides are primarily used in scenarios where the IDs of the slide masters and layouts in the presentation are known at coding time or can be found in a data source at runtime. In such a scenario, either you or the customer must create and maintain a data source that correlates the selection criterion (such as the names or images of slide masters and layouts) with the IDs of the slide masters and layouts. The APIs can also be used in scenarios where the user can insert slides that use the default slide master and the master's default layout, and in scenarios where the user can select an existing slide and create a new one with the same slide master and layout (but not the same content). See Selecting which slide master and layout to use for more information about this.

# **Add a slide with SlideCollection.add**

Add slides with the [SlideCollection.add](https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.slidecollection#powerpoint-powerpoint-slidecollection-add-member(1)) method. The following is a simple example in which a slide that uses the presentation's default slide master and the first layout of that master is added. The method always adds new slides to the end of the presentation. The following is an example.

```
JavaScript
async function addSlide() {
 await PowerPoint.run(async function(context) {
 context.presentation.slides.add();
 await context.sync();
 });
}
```
### **Select which slide master and layout to use**

Use the [AddSlideOptions](https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.addslideoptions) parameter to control which slide master is used for the new slide and which layout within the master is used. The following is an example. About this code, note:


- You can include either or both the properties of the AddSlideOptions object.
- If both properties are used, then the specified layout must belong to the specified master or an error is thrown.
- If the masterId property isn't present (or its value is an empty string), then the default slide master is used and the layoutId must be a layout of that slide master.
- The default slide master is the slide master used by the last slide in the presentation. (In the unusual case where there are currently no slides in the presentation, then the default slide master is the first slide master in the presentation.)
- If the layoutId property isn't present (or its value is an empty string), then the first layout of the master that is specified by the masterId is used.
- Both properties are strings of one of three possible forms: *nnnnnnnnnn***#**, **#***mmmmmmmmm*, or *nnnnnnnnnn***#***mmmmmmmmm*, where *nnnnnnnnnn* is the master's or layout's ID (typically 10 digits) and *mmmmmmmmm* is the master's or layout's creation ID (typically 6 - 10 digits). Some examples are 2147483690#2908289500 , 2147483690# , and #2908289500 .

```
JavaScript
```

```
async function addSlide() {
 await PowerPoint.run(async function(context) {
 context.presentation.slides.add({
 slideMasterId: "2147483690#2908289500",
 layoutId: "2147483691#2499880"
 });

 await context.sync();
 });
}
```
There is no practical way that users can discover the ID or creation ID of a slide master or layout. For this reason, you can really only use the AddSlideOptions parameter when either you know the IDs at coding time or your add-in can discover them at runtime. Because users can't be expected to memorize the IDs, you also need a way to enable the user to select slides, perhaps by name or by an image, and then correlate each title or image with the slide's ID.

Accordingly, the AddSlideOptions parameter is primarily used in scenarios in which the add-in is designed to work with a specific set of slide masters and layouts whose IDs are known. In such a scenario, either you or the customer must create and maintain a data source that correlates a selection criterion (such as slide master and layout names or images) with the corresponding IDs or creation IDs.


### **Have the user choose a matching slide**

If your add-in can be used in scenarios where the new slide should use the same combination of slide master and layout that is used by an *existing* slide, then your addin can (1) prompt the user to select a slide and (2) read the IDs of the slide master and layout. The following steps show how to read the IDs and add a slide with a matching master and layout.

- 1. Create a function to get the index of the selected slide. The following is an example. About this code, note:
	- It uses the [Office.context.document.getSelectedDataAsync](https://learn.microsoft.com/en-us/javascript/api/office/office.document#office-office-document-getselecteddataasync-member(1)) method of the Common JavaScript APIs.
	- The call to getSelectedDataAsync is embedded in a Promise-returning function. For more information about why and how to do this, see Wrap Common APIs in promise-returning functions.
	- getSelectedDataAsync returns an array because multiple slides can be selected. In this scenario, the user has selected just one, so the code gets the first (0th) slide, which is the only one selected.
	- The index value of the slide is the 1-based value the user sees beside the slide in the thumbnails pane.

```
JavaScript
function getSelectedSlideIndex() {
 return new OfficeExtension.Promise<number>(function(resolve,
reject) {

Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideR
ange, function(asyncResult) {
 try {
 if (asyncResult.status ===
Office.AsyncResultStatus.Failed) {
 reject(console.error(asyncResult.error.message));
 } else {
 resolve(asyncResult.value.slides[0].index);
 }
 } 
 catch (error) {
 reject(console.log(error));
 }
 });
 });
}
```
- 2. Call your new function inside the [PowerPoint.run()](https://learn.microsoft.com/en-us/javascript/api/powerpoint#PowerPoint_run_batch_) of the main function that adds the slide. The following is an example.


```
JavaScript
async function addSlideWithMatchingLayout() {
 await PowerPoint.run(async function(context) {
 let selectedSlideIndex = await getSelectedSlideIndex();
 // Decrement the index because the value returned by
getSelectedSlideIndex()
 // is 1-based, but SlideCollection.getItemAt() is 0-based.
 const realSlideIndex = selectedSlideIndex - 1;
 const selectedSlide =
context.presentation.slides.getItemAt(realSlideIndex).load("slideMaster
/id, layout/id");
 await context.sync();
 context.presentation.slides.add({
 slideMasterId: selectedSlide.slideMaster.id,
 layoutId: selectedSlide.layout.id
 });
 await context.sync();
 });
}
```
# **Delete slides**

Delete a slide by getting a reference to the [Slide](https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.slide) object that represents the slide and call the Slide.delete method. The following is an example in which the 4th slide is deleted.

```
JavaScript
async function deleteSlide() {
 await PowerPoint.run(async function(context) {
 // The slide index is zero-based. 
 const slide = context.presentation.slides.getItemAt(3);
 slide.delete();
 await context.sync();
 });
}
```


# **Insert slides in a PowerPoint presentation**

Article • 09/20/2022

A PowerPoint add-in can insert slides from one presentation into the current presentation by using PowerPoint's application-specific JavaScript library. You can control whether the inserted slides keep the formatting of the source presentation or the formatting of the target presentation.

The slide insertion APIs are primarily used in presentation template scenarios: There are a small number of known presentations which serve as pools of slides that can be inserted by the add-in. In such a scenario, either you or the customer must create and maintain a data source that correlates the selection criterion (such as slide titles or images) with slide IDs. The APIs can also be used in scenarios where the user can insert slides from any arbitrary presentation, but in that scenario the user is effectively limited to inserting *all* the slides from the source presentation. See Selecting which slides to insert for more information about this.

There are two steps to inserting slides from one presentation into another.

- 1. Convert the source presentation file (.pptx) into a base64-formatted string.
- 2. Use the insertSlidesFromBase64 method to insert one or more slides from the base64 file into the current presentation.

# **Convert the source presentation to base64**

There are many ways to convert a file to base64. Which programming language and library you use, and whether to convert on the server-side of your add-in or the clientside is determined by your scenario. Most commonly, you'll do the conversion in JavaScript on the client-side by using a [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) object. The following example shows this practice.

- 1. Begin by getting a reference to the source PowerPoint file. In this example, we will use an <input> control of type file to prompt the user to choose a file. Add the following markup to the add-in page.

```
HTML
<section>
 <p>Select a PowerPoint presentation from which to insert slides</p>
 <form>
```


```
 <input type="file" id="file" />
 </form>
</section>
```
This markup adds the UI in the following screenshot to the page.

There are many other ways to get a PowerPoint file. For example, if the file is stored on OneDrive or SharePoint, you can use Microsoft Graph to download it. For more information, see **[Working with files in Microsoft Graph](https://learn.microsoft.com/en-us/graph/api/resources/onedrive)** and **[Access Files with Microsoft Graph](https://learn.microsoft.com/en-us/training/modules/msgraph-access-file-data/)**.

- 2. Add the following code to the add-in's JavaScript to assign a function to the input control's change event. (You create the storeFileAsBase64 function in the next step.)

```
JavaScript
$("#file").on("change", storeFileAsBase64);
```
- 3. Add the following code. Note the following about this code.
	- The reader.readAsDataURL method converts the file to base64 and stores it in the reader.result property. When the method completes, it triggers the onload event handler.
	- The onload event handler trims metadata off of the encoded file and stores the encoded string in a global variable.
	- The base64-encoded string is stored globally because it will be read by another function that you create in a later step.

#### JavaScript

```
let chosenFileBase64;
async function storeFileAsBase64() {
 const reader = new FileReader();
```


```
 reader.onload = async (event) => {
 const startIndex = reader.result.toString().indexOf("base64,");
 const copyBase64 = reader.result.toString().substr(startIndex +
7);
 chosenFileBase64 = copyBase64;
 };
 const myFile = document.getElementById("file") as HTMLInputElement;
 reader.readAsDataURL(myFile.files[0]);
}
```
# **Insert slides with insertSlidesFromBase64**

Your add-in inserts slides from another PowerPoint presentation into the current presentation with the [Presentation.insertSlidesFromBase64](https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.presentation#powerpoint-powerpoint-presentation-insertslidesfrombase64-member(1)) method. The following is a simple example in which all of the slides from the source presentation are inserted at the beginning of the current presentation and the inserted slides keep the formatting of the source file. Note that chosenFileBase64 is a global variable that holds a base64 encoded version of a PowerPoint presentation file.

```
JavaScript
async function insertAllSlides() {
 await PowerPoint.run(async function(context) {
 context.presentation.insertSlidesFromBase64(chosenFileBase64);
 await context.sync();
 });
}
```
You can control some aspects of the insertion result, including where the slides are inserted and whether they get the source or target formatting , by passing an [InsertSlideOptions](https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.insertslideoptions) object as a second parameter to insertSlidesFromBase64 . The following is an example. About this code, note:

- There are two possible values for the formatting property: "UseDestinationTheme" and "KeepSourceFormatting". Optionally, you can use the InsertSlideFormatting enum, (e.g., PowerPoint.InsertSlideFormatting.useDestinationTheme ).
- The function will insert the slides from the source presentation immediately after the slide specified by the targetSlideId property. The value of this property is a string of one of three possible forms: *nnn***#**, **#***mmmmmmmmm*, or *nnn***#***mmmmmmmmm*, where *nnn* is the slide's ID (typically 3 digits) and *mmmmmmmmm* is the slide's creation ID (typically 9 digits). Some examples are 267#763315295 , 267# , and #763315295 .


```
JavaScript
async function insertSlidesDestinationFormatting() {
 await PowerPoint.run(async function(context) {
 context.presentation
 .insertSlidesFromBase64(chosenFileBase64,
 {
 formatting: "UseDestinationTheme",
                      targetSlideId: "267#"
 }
 );
 await context.sync();
 });
}
```
Of course, you typically won't know at coding time the ID or creation ID of the target slide. More commonly, an add-in will ask users to select the target slide. The following steps show how to get the *nnn***#** ID of the currently selected slide and use it as the target slide.

- 1. Create a function that gets the ID of the currently selected slide by using the [Office.context.document.getSelectedDataAsync](https://learn.microsoft.com/en-us/javascript/api/office/office.document#office-office-document-getselecteddataasync-member(1)) method of the Common JavaScript APIs. The following is an example. Note that the call to getSelectedDataAsync is embedded in a Promise-returning function. For more information about why and how to do this, see Wrap Common-APIs in promise-returning functions.

```
JavaScript
function getSelectedSlideID() {
 return new OfficeExtension.Promise<string>(function (resolve, reject)
{

Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideR
ange, function (asyncResult) {
 try {
 if (asyncResult.status === Office.AsyncResultStatus.Failed) {
 reject(console.error(asyncResult.error.message));
 } else {
 resolve(asyncResult.value.slides[0].id);
 }
 }
 catch (error) {
 reject(console.log(error));
 }
 });
 })
}
```


- 2. Call your new function inside the [PowerPoint.run()](https://learn.microsoft.com/en-us/javascript/api/powerpoint#PowerPoint_run_batch_) of the main function and pass the ID that it returns (concatenated with the "#" symbol) as the value of the targetSlideId property of the InsertSlideOptions parameter. The following is an example.

```
JavaScript
async function insertAfterSelectedSlide() {
 await PowerPoint.run(async function(context) {
 const selectedSlideID = await getSelectedSlideID();
 context.presentation.insertSlidesFromBase64(chosenFileBase64, {
 formatting: "UseDestinationTheme",
 targetSlideId: selectedSlideID + "#"
 });
 await context.sync();
 });
}
```
### **Selecting which slides to insert**

You can also use the [InsertSlideOptions](https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.insertslideoptions) parameter to control which slides from the source presentation are inserted. You do this by assigning an array of the source presentation's slide IDs to the sourceSlideIds property. The following is an example that inserts four slides. Note that each string in the array must follow one or another of the patterns used for the targetSlideId property.

```
JavaScript
async function insertAfterSelectedSlide() {
 await PowerPoint.run(async function(context) {
 const selectedSlideID = await getSelectedSlideID();
 context.presentation.insertSlidesFromBase64(chosenFileBase64, {
 formatting: "UseDestinationTheme",
 targetSlideId: selectedSlideID + "#",
 sourceSlideIds: ["267#763315295", "256#", "#926310875", "1270#"]
 });
 await context.sync();
 });
}
```


The slides will be inserted in the same relative order in which they appear in the source presentation, regardless of the order in which they appear in the array.

There is no practical way that users can discover the ID or creation ID of a slide in the source presentation. For this reason, you can really only use the sourceSlideIds property when either you know the source IDs at coding time or your add-in can retrieve them at runtime from some data source. Because users cannot be expected to memorize slide IDs, you also need a way to enable the user to select slides, perhaps by title or by an image, and then correlate each title or image with the slide's ID.

Accordingly, the sourceSlideIds property is primarily used in presentation template scenarios: The add-in is designed to work with a specific set of presentations that serve as pools of slides that can be inserted. In such a scenario, either you or the customer must create and maintain a data source that correlates a selection criterion (such as titles or images) with slide IDs or slide creation IDs that has been constructed from the set of possible source presentations.


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


## **Core concepts to know for creating a task pane add-in**

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


# **Use custom tags for presentations, slides, and shapes in PowerPoint**

Article • 07/21/2022

An add-in can attach custom metadata, in the form of key-value pairs, called "tags", to presentations, specific slides, and specific shapes on a slide.

There are two main scenarios for using tags:

- When applied to a slide or a shape, a tag enables the object to be categorized for batch processing. For example, suppose a presentation has some slides that should be included in presentations to the East region but not the West region. Similarly, there are alternative slides that should be shown only to the West. Your add-in can create a tag with the key REGION and the value East and apply it to the slides that should only be used in the East. The tag's value is set to West for the slides that should only be shown to the West region. Just before a presentation to the East, a button in the add-in runs code that loops through all the slides checking the value of the REGION tag. Slides where the region is West are deleted. The user then closes the add-in and starts the slide show.
- When applied to a presentation, a tag is effectively a custom property in the presentation document (similar to a [CustomProperty](https://learn.microsoft.com/en-us/javascript/api/word/word.customproperty) in Word).

# **Tag slides and shapes**

A tag is a key-value pair, where the value is always of type string and is represented by a [Tag](https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.tag) object. Each type of parent object, such as a [Presentation,](https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.presentation) [Slide,](https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.slide) or [Shape](https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.shape) object, has a tags property of type [TagsCollection.](https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.tagcollection)

## **Add, update, and delete tags**

To add a tag to an object, call the [TagCollection.add](https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.tagcollection#powerpoint-powerpoint-tagcollection-add-member(1)) method of the parent object's tags property. The following code adds two tags to the first slide of a presentation. About this code, note:

- The first parameter of the add method is the key in the key-value pair.
- The second parameter is the value.
- The key is in uppercase letters. This isn't strictly mandatory for the add method; however, the key is always stored by PowerPoint as uppercase, and *some tagrelated methods do require that the key be expressed in uppercase*, so we


recommend as a best practice that you always use uppercase in your code for a tag key.

JavaScript

```
async function addMultipleSlideTags() {
 await PowerPoint.run(async function(context) {
 const slide = context.presentation.slides.getItemAt(0);
 slide.tags.add("OCEAN", "Arctic");
 slide.tags.add("PLANET", "Jupiter");
 await context.sync();
 });
}
```
The add method is also used to update a tag. The following code changes the value of the PLANET tag.

```
JavaScript
async function updateTag() {
 await PowerPoint.run(async function(context) {
 const slide = context.presentation.slides.getItemAt(0);
 slide.tags.add("PLANET", "Mars");
 await context.sync();
 });
}
```
To delete a tag, call the delete method on it's parent TagsCollection object and pass the key of the tag as the parameter. For an example, see Set custom metadata on the presentation.

### **Use tags to selectively process slides and shapes**

Consider the following scenario: Contoso Consulting has a presentation they show to all new customers. But some slides should only be shown to customers that have paid for "premium" status. Before showing the presentation to non-premium customers, they make a copy of it and delete the slides that only premium customers should see. An add-in enables Contoso to tag which slides are for premium customers and to delete these slides when needed. The following list outlines the major coding steps to create this functionality.

- 1. Create a function that tags the currently selected slide as intended for Premium customers. About this code, note:


- The getSelectedSlideIndex function is defined in the next step. It returns the 1-based index of the currently selected slide.
- The value returned by the getSelectedSlideIndex function has to be decremented because the [SlideCollection.getItemAt](https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.slidecollection#powerpoint-powerpoint-slidecollection-getitemat-member(1)) method is 0-based.

```
JavaScript
async function addTagToSelectedSlide() {
 await PowerPoint.run(async function(context) {
 let selectedSlideIndex = await getSelectedSlideIndex();
 selectedSlideIndex = selectedSlideIndex - 1;
 const slide =
context.presentation.slides.getItemAt(selectedSlideIndex);
 slide.tags.add("CUSTOMER_TYPE", "Premium");
 await context.sync();
 });
}
```
- 2. The following code creates a method to get the index of the selected slide. About this code, note:
	- It uses the [Office.context.document.getSelectedDataAsync](https://learn.microsoft.com/en-us/javascript/api/office/office.document#office-office-document-getselecteddataasync-member(1)) method of the Common JavaScript APIs.
	- The call to getSelectedDataAsync is embedded in a promise-returning function. For more information about why and how to do this, see Wrap Common APIs in promise-returning functions.
	- getSelectedDataAsync returns an array because multiple slides can be selected. In this scenario, the user has selected just one, so the code gets the first (0th) slide, which is the only one selected.
	- The index value of the slide is the 1-based value the user sees beside the slide in the PowerPoint UI thumbnails pane.

```
JavaScript
function getSelectedSlideIndex() {
 return new OfficeExtension.Promise<number>(function(resolve,
reject) {

Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideR
ange, function(asyncResult) {
 try {
 if (asyncResult.status ===
Office.AsyncResultStatus.Failed) {
 reject(console.error(asyncResult.error.message));
 } else {
 resolve(asyncResult.value.slides[0].index);
```


```
 }
 } 
 catch (error) {
 reject(console.log(error));
 }
 });
 });
}
```
- 3. The following code creates a function to delete slides that are tagged for premium customers. About this code, note:
	- Because the key and value properties of the tags are going to be read after the context.sync , they must be loaded first.

```
JavaScript
async function deleteSlidesByAudience() {
 await PowerPoint.run(async function(context) {
 const slides = context.presentation.slides;
 slides.load("tags/key, tags/value");
 await context.sync();
 for (let i = 0; i < slides.items.length; i++) {
 let currentSlide = slides.items[i];
 for (let j = 0; j < currentSlide.tags.items.length; j++) {
 let currentTag = currentSlide.tags.items[j];
 if (currentTag.key === "CUSTOMER_TYPE" && currentTag.value ===
"Premium") {
 currentSlide.delete();
 }
 }
 }
 await context.sync();
 });
}
```
## **Set custom metadata on the presentation**

Add-ins can also apply tags to the presentation as a whole. This enables you to use tags for document-level metadata similar to how the [CustomProperty](https://learn.microsoft.com/en-us/javascript/api/word/word.customproperty)class is used in Word. But unlike the Word CustomProperty class, the value of a PowerPoint tag can only be of type string .

The following code is an example of adding a tag to a presentation.


```
JavaScript
async function addPresentationTag() {
 await PowerPoint.run(async function (context) {
 let presentationTags = context.presentation.tags;
 presentationTags.add("SECURITY", "Internal-Audience-Only");
 await context.sync();
 });
}
```
The following code is an example of deleting a tag from a presentation. Note that the key of the tag is passed to the delete method of the parent TagsCollection object.

```
JavaScript
async function deletePresentationTag() {
 await PowerPoint.run(async function (context) {
 let presentationTags = context.presentation.tags;
 presentationTags.delete("SECURITY");
 await context.sync();
 });
}
```


# **Use document themes in your PowerPoint add-ins**

Article • 06/18/2024

An [Office theme](https://support.microsoft.com/office/83e68627-2c17-454a-9fd8-62deb81951a6) consists, in part, of a visually coordinated set of fonts and colors that you can apply to presentations, documents, worksheets, and emails. To apply or customize the theme of a presentation in PowerPoint, you use the **Themes** and **Variants** groups on **Design** tab of the ribbon. PowerPoint assigns a new blank presentation with the default **Office Theme**, but you can choose other themes available on the **Design** tab, download additional themes from Office.com, or create and customize your own theme.

Using **OfficeThemes.css**, design add-ins that are coordinated with PowerPoint in two ways.

- **In content add-ins for PowerPoint**. Use the document theme classes of **OfficeThemes.css** to specify fonts and colors that match the theme of the presentation your content add-in is inserted into - and those fonts and colors will dynamically update if a user changes or customizes the presentation's theme.
- **In task pane add-ins for PowerPoint**. Use the Office UI theme classes of **OfficeThemes.css** to specify the same fonts and background colors used in the UI so that your task pane add-ins will match the colors of built-in task panes - and those colors will dynamically update if a user changes the Office UI theme.

## **Document theme colors**

Every Office document theme defines 12 colors. Ten of these colors are available when you set font, background, and other color settings in a presentation with the color picker.


To view or customize the full set of 12 theme colors in PowerPoint, in the **Variants** group on the **Design** tab, click the **More** drop-down - then select **Colors** > **Customize Colors** to display the **Create New Theme Colors** dialog box.

| ?<br>×<br>Create New Theme Colors |  |           |           |  |
|-----------------------------------|--|-----------|-----------|--|
| Theme colors                      |  | Sample    |           |  |
| Text/Background - Dark 1          |  | Text      | Text      |  |
| Text/Background - Light 1         |  |           |           |  |
| Text/Background - Dark 2          |  |           |           |  |
| Text/Background - Light 2         |  |           |           |  |
| Accent 1                          |  |           | Hyperlink |  |
| Accent 2                          |  | Hyperlink | Hyperlink |  |
| Accent 3                          |  |           |           |  |
| Accent 4                          |  |           |           |  |
| Accent 5                          |  |           |           |  |
| Accent 6                          |  |           |           |  |
| Hyperlink                         |  |           |           |  |
| Eollowed Hyperlink                |  |           |           |  |
| Custom 1<br>Name:                 |  |           |           |  |
| Reset                             |  | Save      | Cancel    |  |

The first four colors are for text and backgrounds. Text that is created with the light colors will always be legible over the dark colors, and text that is created with dark colors will always be legible over the light colors. The next six are accent colors that are always visible over the four potential background colors. The last two colors are for hyperlinks and followed hyperlinks.

# **Document theme fonts**

Every Office document theme also defines two fonts -- one for headings and one for body text. PowerPoint uses these fonts to construct automatic text styles. In addition, **Quick Styles** galleries for text and **WordArt** use these same theme fonts. These two fonts are available as the first two selections when you select fonts with the font picker.


| Theme Fonts     |            |
|-----------------|------------|
| ™ Calibri Light | (Headings) |
| Tr Calibri      | (Body)     |
| All Fonts       |            |
| TT Agency B     |            |
| Tr Aharoni      | אבגד הוז   |
| TT Aldhabi      | ,8 ,51     |
| T ALGERIAN      |            |

To view or customize theme fonts in PowerPoint, in the **Variants** group on the **Design** tab, click the **More** drop-down - then select **Fonts** > **Customize Fonts** to display the **Create New Theme Fonts** dialog box.

|                                                         | ?<br>×<br>Create New Theme Fonts                                                      |
|---------------------------------------------------------|---------------------------------------------------------------------------------------|
| Heading font:<br>Calibri Light<br>Body font:<br>Calibri | Sample<br>><br>Heading<br>Body text body text body text.<br>Body text body text.<br>> |
| Name:                                                   | Save                                                                                  |
| Custom 1                                                | Cancel                                                                                |

### **Office UI theme fonts and colors**

Office also lets you choose between several predefined themes that specify some of the colors and fonts used in the UI of all Office applications. To do that, you use the **File** > **Account** > **Office Theme** drop-down (from any Office application).

| White      |  |
|------------|--|
| White      |  |
| Light Gray |  |
| Dark Gray  |  |

**OfficeThemes.css** includes classes that you can use in your task pane add-ins for PowerPoint so they will use these same fonts and colors. This lets you design your task pane add-ins that match the appearance of built-in task panes.

# **Use OfficeThemes.css**

Using the **OfficeThemes.css** file with your content add-ins for PowerPoint lets you coordinate the appearance of your add-in with the theme applied to the presentation 


it's running with. Using the **OfficeThemes.css** file with your task pane add-ins for PowerPoint lets you coordinate the appearance of your add-in with the fonts and colors of the Office UI.

# **Add the OfficeThemes.css file to your project**

Use the following steps to add and reference the **OfficeThemes.css** file to your add-in project.

#### 7 **Note**

The steps in this procedure only apply to Visual Studio 2015. If you are using Visual Studio 2019, the **OfficeThemes.css** file is created automatically for any new PowerPoint add-in projects that you create.

- 1. In **Solution Explorer**, right-click (or select and hold) the **Content** folder in the *project_name***Web** project, choose **Add**, and then select **Style Sheet**.
- 2. Name the new style sheet **OfficeThemes**.

#### ) **Important**

The style sheet must be named OfficeThemes, or the feature that dynamically updates add-in fonts and colors when a user changes the theme won't work.

- 3. Delete the default **body** class ( body {} ) in the file, and copy and paste the following CSS code into the file.

```
css
/* The following classes describe the common theme information for
office documents */
/* Basic Font and Background Colors for text */
.office-docTheme-primary-fontColor { color:#000000; } 
.office-docTheme-primary-bgColor { background-color:#ffffff; } 
.office-docTheme-secondary-fontColor { color: #000000; } 
.office-docTheme-secondary-bgColor { background-color: #ffffff; } 
/* Accent color definitions for fonts */
.office-contentAccent1-color { color:#5b9bd5; } 
.office-contentAccent2-color { color:#ed7d31; } 
.office-contentAccent3-color { color:#a5a5a5; } 
.office-contentAccent4-color { color:#ffc000; }
```


```
.office-contentAccent5-color { color:#4472c4; } 
.office-contentAccent6-color { color:#70ad47; } 
/* Accent color for backgrounds */
.office-contentAccent1-bgColor { background-color:#5b9bd5; } 
.office-contentAccent2-bgColor { background-color:#ed7d31; } 
.office-contentAccent3-bgColor { background-color:#a5a5a5; } 
.office-contentAccent4-bgColor { background-color:#ffc000; } 
.office-contentAccent5-bgColor { background-color:#4472c4; } 
.office-contentAccent6-bgColor { background-color:#70ad47; } 
/* Accent color for borders */
.office-contentAccent1-borderColor { border-color:#5b9bd5; } 
.office-contentAccent2-borderColor { border-color:#ed7d31; } 
.office-contentAccent3-borderColor { border-color:#a5a5a5; } 
.office-contentAccent4-borderColor { border-color:#ffc000; } 
.office-contentAccent5-borderColor { border-color:#4472c4; } 
.office-contentAccent6-borderColor { border-color:#70ad47; } 
/* links */
.office-a { color: #0563c1; } 
.office-a:visited { color: #954f72; } 
/* Body Fonts */
.office-bodyFont-eastAsian { } /* East Asian name of the Font */
.office-bodyFont-latin { font-family:"Calibri"; } /* Latin name of the
Font */
.office-bodyFont-script { } /* Script name of the Font */
.office-bodyFont-localized { font-family:"Calibri"; } /* Localized name
of the Font. Corresponds to the default font of the culture currently
used in Office.*/
/* Headers Font */
.office-headerFont-eastAsian { } 
.office-headerFont-latin { font-family:"Calibri Light"; } 
.office-headerFont-script { } 
.office-headerFont-localized { font-family:"Calibri Light"; } 
/* The following classes define font and background colors for Office
UI themes. These classes should only be used in task pane add-ins */
/* Basic Font and Background Colors for PPT */
.office-officeTheme-primary-fontColor { color:#b83b1d; } 
.office-officeTheme-primary-bgColor { background-color:#dedede; } 
.office-officeTheme-secondary-fontColor { color:#262626; } 
.office-officeTheme-secondary-bgColor { background-color:#ffffff; }
```
- 4. If you are using a tool other than Visual Studio to create your add-in, copy the CSS code from the previous step into a text file. Then, save the file as **OfficeThemes.css**.


# **Reference OfficeThemes.css in your add-in's HTML pages**

To use the **OfficeThemes.css** file in your add-in project, add a <link> tag that references the **OfficeThemes.css** file inside the <head> tag of the web pages (such as an .html, .aspx, or .php file) that implement the UI of your add-in in this format.

```
HTML
```

```
<link href="<local_path_to_OfficeThemes.css>" rel="stylesheet"
type="text/css" />
```
To do this in Visual Studio, follow these steps.

- 1. Choose **Create a new project**.
- 2. Using the search box, enter **add-in**. Choose **PowerPoint Web Add-in**, then select **Next**.
- 3. Name your project and select **Create**.
- 4. In the **Create Office Add-in** dialog window, choose **Add new functionalities to PowerPoint**, and then choose **Finish** to create the project.
- 5. Visual Studio creates a solution and its two projects appear in **Solution Explorer**. The **Home.html** file opens in Visual Studio.
- 6. In the HTML pages that implement the UI of your add-in, such as Home.html in the default template, add the following <link> tag inside the <head> tag that references the **OfficeThemes.css** file.

```
HTML
<link href="../../Content/OfficeThemes.css" rel="stylesheet"
type="text/css" />
```
If you are creating your add-in with a tool other than Visual Studio, add a <link> tag with the same format specifying a relative path to the copy of **OfficeThemes.css** that will be deployed with your add-in.

## **Use OfficeThemes.css document theme classes in your content add-in's HTML page**


The following shows a simple example of HTML in a content add-in that uses the OfficeTheme.css document theme classes. For details about the **OfficeThemes.css** classes that correspond to the 12 colors and 2 fonts used in a document theme, see Theme classes for content add-ins.

At runtime, when inserted into a presentation that uses the default **Office Theme**, the content add-in is rendered like this.

| Hello world! |  |  |  |  |
|--------------|--|--|--|--|
| Hello world! |  |  |  |  |
| Hello world! |  |  |  |  |
| Hello world! |  |  |  |  |
| Hello world! |  |  |  |  |
| Hello world! |  |  |  |  |

If you change the presentation to use another theme or customize the presentation's theme, the fonts and colors specified with **OfficeThemes.css** classes will dynamically update to correspond to the fonts and colors of the presentation's theme. Using the


same HTML example as above, if the presentation the add-in is inserted into uses the **Facet** theme, the add-in rendering will look like this.

| Hello world!    |  |
|-----------------|--|
| Hello world!    |  |
| Hello world!    |  |
| Hello world!    |  |
| Hello world!    |  |
| Hello world!    |  |
| الواسعند مالمال |  |

## **Use OfficeThemes.css Office UI theme classes in your task pane add-in's HTML page**

In addition to the document theme, users can customize the color scheme of the Office user interface for all Office applications using the **File** > **Account** > **Office Theme** dropdown box.

The following shows a simple example of HTML in a task pane add-in that uses OfficeTheme.css classes to specify font color and background color. For details about the **OfficeThemes.css** classes that correspond to fonts and colors of the Office UI theme, see Theme classes for task pane add-ins.

```
HTML
<body>
 <div id="content-header" class="office-officeTheme-primary-fontColor
office-officeTheme-primary-bgColor">
 <div class="padding">
 <h1>Welcome</h1>
 </div>
 </div>
 <div id="content-main" class="office-officeTheme-secondary-fontColor
office-officeTheme-secondary-bgColor">
 <div class="padding">
 <p>Add home screen content here.</p>
 <p>For example:</p>
 <button id="get-data-from-selection">Get data from
selection</button>
 <p><a target="_blank" class="office-a"
```


```
href="https://go.microsoft.com/fwlink/?LinkId=276812">Find more samples
online...</a></p>
 </div>
 </div>
</body>
```
When running in PowerPoint with **File** > **Account** > **Office Theme** set to **White**, the task pane add-in is rendered like this.

| TaskPaneTheme                                 |  |
|-----------------------------------------------|--|
| Welcome                                       |  |
| Add home screen content here.<br>For example: |  |
| Get data from selection                       |  |
| Find more samples online                      |  |
|                                               |  |
|                                               |  |

If you change **OfficeTheme** to **Dark Gray**, the fonts and colors specified with **OfficeThemes.css** classes will dynamically update to render like this.

**OfficeTheme.css classes**


The **OfficeThemes.css** file contains two sets of classes you can use with your content and task pane add-ins for PowerPoint.

## **Theme classes for content add-ins**

The **OfficeThemes.css** file provides classes that correspond to the 2 fonts and 12 colors used in a document theme. These classes are appropriate to use with content add-ins for PowerPoint so that your add-in's fonts and colors will be coordinated with the presentation it's inserted into.

### **Theme fonts for content add-ins**

ノ **Expand table**

| Class                          | Description                                                         |
|--------------------------------|---------------------------------------------------------------------|
| office-bodyFont<br>eastAsian   | East Asian name of the body font.                                   |
| office-bodyFont<br>latin       | Latin name of the body font. Default "Calabri"                      |
| office-bodyFont<br>script      | Script name of the body font.                                       |
| office-bodyFont                | Localized name of the body font. Specifies the default font name    |
| localized                      | according to the culture currently used in Office.                  |
| office-headerFont<br>eastAsian | East Asian name of the headers font.                                |
| office-headerFont<br>latin     | Latin name of the headers font. Default "Calabri Light"             |
| office-headerFont<br>script    | Script name of the headers font.                                    |
| office-headerFont              | Localized name of the headers font. Specifies the default font name |
| localized                      | according to the culture currently used in Office.                  |

## **Theme colors for content add-ins**


| Class                               | Description                                      |  |  |
|-------------------------------------|--------------------------------------------------|--|--|
| office-docTheme-primary-fontColor   | Primary font color. Default #000000              |  |  |
| office-docTheme-primary-bgColor     | Primary font background color. Default #FFFFFF   |  |  |
| office-docTheme-secondary-fontColor | Secondary font color. Default #000000            |  |  |
| office-docTheme-secondary-bgColor   | Secondary font background color. Default #FFFFFF |  |  |
| office-contentAccent1-color         | Font accent color 1. Default #5B9BD5             |  |  |
| office-contentAccent2-color         | Font accent color 2. Default #ED7D31             |  |  |
| office-contentAccent3-color         | Font accent color 3. Default #A5A5A5             |  |  |
| office-contentAccent4-color         | Font accent color 4. Default #FFC000             |  |  |
| office-contentAccent5-color         | Font accent color 5. Default #4472C4             |  |  |
| office-contentAccent6-color         | Font accent color 6. Default #70AD47             |  |  |
| office-contentAccent1-bgColor       | Background accent color 1. Default #5B9BD5       |  |  |
| office-contentAccent2-bgColor       | Background accent color 2. Default #ED7D31       |  |  |
| office-contentAccent3-bgColor       | Background accent color 3. Default #A5A5A5       |  |  |
| office-contentAccent4-bgColor       | Background accent color 4. Default #FFC000       |  |  |
| office-contentAccent5-bgColor       | Background accent color 5. Default #4472C4       |  |  |
| office-contentAccent6-bgColor       | Background accent color 6. Default #70AD47       |  |  |
| office-contentAccent1-borderColor   | Border accent color 1. Default #5B9BD5           |  |  |
| office-contentAccent2-borderColor   | Border accent color 2. Default #ED7D31           |  |  |
| office-contentAccent3-borderColor   | Border accent color 3. Default #A5A5A5           |  |  |
| office-contentAccent4-borderColor   | Border accent color 4. Default #FFC000           |  |  |
| office-contentAccent5-borderColor   | Border accent color 5. Default #4472C4           |  |  |
| office-contentAccent6-borderColor   | Border accent color 6. Default #70AD47           |  |  |
| office-a                            | Hyperlink color. Default #0563C1                 |  |  |
| office-a:visited                    | Followed hyperlink color. Default #954F72        |  |  |

The following screenshot shows examples of all of the theme color classes (except for the two hyperlink colors) assigned to add-in text when using the default Office theme.


| office-docTheme-primary-fontColor   |  |  |  |  |
|-------------------------------------|--|--|--|--|
| office-docTheme-primary-bgColor     |  |  |  |  |
| office-docTheme-secondary-fontColor |  |  |  |  |
| office-docTheme-secondary-bgColor   |  |  |  |  |
| office-contentAccent1-color         |  |  |  |  |
| office-contentAccent2-color         |  |  |  |  |
| office-contentAccent3-color         |  |  |  |  |
| office-contentAccent4-color         |  |  |  |  |
| office-contentAccent5-color         |  |  |  |  |
| office-contentAccent6-color         |  |  |  |  |
| office-contentAccent1-bgColor       |  |  |  |  |
| office-contentAccent2-bgColor       |  |  |  |  |
| office-contentAccent3-bgColor       |  |  |  |  |
| office-contentAccent4-bgColor       |  |  |  |  |
| office-contentAccent5-bgColor       |  |  |  |  |
| office-contentAccent6-bgColor       |  |  |  |  |
| office-contentAccent1-borderColor   |  |  |  |  |
| office-contentAccent2-borderColor   |  |  |  |  |
| office-contentAccent3-borderColor   |  |  |  |  |
| office-contentAccent4-borderColor   |  |  |  |  |
| office-contentAccent5-borderColor   |  |  |  |  |
| office-contentAccent6-borderColor   |  |  |  |  |

### **Theme classes for task pane add-ins**

The **OfficeThemes.css** file provides classes that correspond to the four colors assigned to fonts and backgrounds used by the Office application UI theme. These classes are appropriate to use with task add-ins for PowerPoint, so that your add-in's colors are coordinated with the other built-in task panes in Office.

### **Theme font and background colors for task pane add-ins**

ノ **Expand table**

| Class                                  | Description                                 |  |
|----------------------------------------|---------------------------------------------|--|
| office-officeTheme-primary-fontColor   | Primary font color. Default #B83B1D         |  |
| office-officeTheme-primary-bgColor     | Primary background color. Default #DEDEDE   |  |
| office-officeTheme-secondary-fontColor | Secondary font color. Default #262626       |  |
| office-officeTheme-secondary-bgColor   | Secondary background color. Default #FFFFFF |  |

## **See also**


#### 6 **Collaborate with us on GitHub**

The source for this content can be found on GitHub, where you can also create and review issues and pull requests. For more information, see [our](https://github.com/OfficeDev/office-js-docs-pr/blob/main/Contributing.md) [contributor guide](https://github.com/OfficeDev/office-js-docs-pr/blob/main/Contributing.md).

#### **Office Add-ins feedback**

Office Add-ins is an open source project. Select a link to provide feedback:

- [Open a documentation issue](https://github.com/OfficeDev/office-js-docs-pr/issues/new?template=3-customer-feedback.yml&pageUrl=https%3A%2F%2Flearn.microsoft.com%2Fen-us%2Foffice%2Fdev%2Fadd-ins%2Fpowerpoint%2Fuse-document-themes-in-your-powerpoint-add-ins&pageQueryParams=&contentSourceUrl=https%3A%2F%2Fgithub.com%2FOfficeDev%2Foffice-js-docs-pr%2Fblob%2Fmain%2Fdocs%2Fpowerpoint%2Fuse-document-themes-in-your-powerpoint-add-ins.md&documentVersionIndependentId=612a326c-6cb6-fafb-31ad-3b58e27d76e2&feedback=%0A%0A%5BEnter+feedback+here%5D%0A&author=%40o365devx&metadata=*+ID%3A+27a2e981-f6ee-2ca1-57a5-402a2ea4e08f+%0A*+Service%3A+**powerpoint**%0A*+Sub-service%3A+**add-ins**)
- [Provide product feedback](https://aka.ms/office-addins-dev-questions)


# **Work with shapes using the PowerPoint JavaScript API**

Article • 05/07/2025

This article describes how to use geometric shapes, lines, and text boxes in conjunction with the [Shape](https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.shape) and [ShapeCollection](https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.shapecollection) APIs.

# **Create shapes**

Shapes are created through and stored in a slide's shape collection ( slide.shapes ). ShapeCollection has several .add* methods for this purpose. All shapes have names and IDs generated for them when they are added to the collection. These are the name and id properties, respectively. name can be set by your add-in.

### **Geometric shapes**

A geometric shape is created with one of the overloads of ShapeCollection.addGeometricShape . The first parameter is either a [GeometricShapeType](https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.geometricshapetype) enum or the string equivalent of one of the enum's values. There is an optional second parameter of type [ShapeAddOptions](https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.shapeaddoptions) that can specify the initial size of the shape and its position relative to the top and left sides of the slide, measured in points. Or these properties can be set after the shape is created.

The following code sample creates a rectangle named **"Square"** that is positioned 100 points from the top and left sides of the slide. The method returns a Shape object.

```
JavaScript
```

```
// This sample creates a rectangle positioned 100 points from the top and left
sides
// of the slide and is 150x150 points. The shape is put on the first slide.
await PowerPoint.run(async (context) => {
 const shapes = context.presentation.slides.getItemAt(0).shapes;
 const rectangle =
shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle);
 rectangle.left = 100;
 rectangle.top = 100;
 rectangle.height = 150;
 rectangle.width = 150;
 rectangle.name = "Square";
 await context.sync();
});
```


### **Lines**

A line is created with one of the overloads of ShapeCollection.addLine . The first parameter is either a [ConnectorType](https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.connectortype) enum or the string equivalent of one of the enum's values to specify how the line contorts between endpoints. There is an optional second parameter of type [ShapeAddOptions](https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.shapeaddoptions) that can specify the start and end points of the line. Or these properties can be set after the shape is created. The method returns a Shape object.

#### 7 **Note**

When the shape is a line, the top and left properties of the Shape and ShapeAddOptions objects specify the starting point of the line relative to the top and left edges of the slide. The height and width properties specify the endpoint of the line *relative to the start point*. So, the end point relative to the top and left edges of the slide is ( top + height ) by ( left + width ). The unit of measure for all properties is points and negative values are allowed.

The following code sample creates a straight line on the slide.

```
JavaScript
// This sample creates a straight line on the first slide.
await PowerPoint.run(async (context) => {
 const shapes = context.presentation.slides.getItemAt(0).shapes;
 const line = shapes.addLine(PowerPoint.ConnectorType.straight, {left: 200, 
top: 50, height: 300, width: 150});
 line.name = "StraightLine";
 await context.sync();
});
```
## **Text boxes**

A text box is created with the [addTextBox](https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.shapecollection#powerpoint-powerpoint-shapecollection-addtextbox-member(1)) method. The first parameter is the text that should appear in the box initially. There is an optional second parameter of type [ShapeAddOptions](https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.shapeaddoptions) that can specify the initial size of the text box and its position relative to the top and left sides of the slide. Or these properties can be set after the shape is created.

The following code sample shows how to create a text box on the first slide.

```
JavaScript
// This sample creates a text box with the text "Hello!" and sizes it
appropriately.
```


```
await PowerPoint.run(async (context) => {
 const shapes = context.presentation.slides.getItemAt(0).shapes;
 const textbox = shapes.addTextBox("Hello!");
 textbox.left = 100;
 textbox.top = 100;
 textbox.height = 300;
 textbox.width = 450;
 textbox.name = "Textbox";
 await context.sync();
});
```
## **Move and resize shapes**

Shapes sit on top of the slide. Their placement is defined by the left and top properties. These act as margins from slide's respective edges, measured in points, with left: 0 and top: 0 being the upper-left corner. The shape size is specified by the height and width properties. Your code can move or resize the shape by resetting these properties. (These properties have a slightly different meaning when the shape is a line. See Lines.)

## **Text in shapes**

Geometric shapes can contain text. Shapes have a textFrame property of type [TextFrame](https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.textframe). The TextFrame object manages the text display options (such as margins and text overflow). TextFrame.textRange is a [TextRange](https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.textrange) object with the text content and font settings.

The following code sample creates a geometric shape named **"Braces"** with the text **"Shape text"**. It also adjusts the shape and text colors, as well as sets the text's vertical alignment to the center.

```
JavaScript
```

```
// This sample creates a light blue rectangle with braces ("{}") on the left and
right ends
// and adds the purple text "Shape text" to the center.
await PowerPoint.run(async (context) => {
 const shapes = context.presentation.slides.getItemAt(0).shapes;
 const braces =
shapes.addGeometricShape(PowerPoint.GeometricShapeType.bracePair);
 braces.left = 100;
 braces.top = 400;
 braces.height = 50;
 braces.width = 150;
 braces.name = "Braces";
 braces.fill.setSolidColor("lightblue");
 braces.textFrame.textRange.text = "Shape text";
 braces.textFrame.textRange.font.color = "purple";
```


```
 braces.textFrame.verticalAlignment =
PowerPoint.TextVerticalAlignment.middleCentered;
 await context.sync();
});
```
## **Group and ungroup shapes**

In PowerPoint, you can group several shapes and treat them like a single shape. You can subsequently ungroup grouped shapes. To learn more about grouping objects in the PowerPoint UI, see [Group or ungroup shapes, pictures, or other objects](https://support.microsoft.com/office/a7374c35-20fe-4e0a-9637-7de7d844724b) .

### **Group shapes**

To group shapes with the JavaScript API, use [ShapeCollection.addGroup](https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.shapecollection#powerpoint-powerpoint-shapecollection-addgroup-member(1)).

The following code sample shows how to group existing shapes of type [GeometricShape](https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.shapetype) found on the current slide.

```
TypeScript
// Groups the geometric shapes on the current slide.
await PowerPoint.run(async (context) => {
 // Get the shapes on the current slide.
 context.presentation.load("slides");
 const slide = context.presentation.getSelectedSlides().getItemAt(0);
 slide.load("shapes/items/type,shapes/items/id");
 await context.sync();
 const shapes = slide.shapes;
 const shapesToGroup = shapes.items.filter((item) => item.type ===
PowerPoint.ShapeType.geometricShape);
 if (shapesToGroup.length === 0) {
 console.warn("No shapes on the current slide, so nothing to group.");
 return;
 }
 // Group the geometric shapes.
 console.log(`Number of shapes to group: ${shapesToGroup.length}`);
 const group = shapes.addGroup(shapesToGroup);
 group.load("id");
 await context.sync();
 console.log(`Grouped shapes. Group ID: ${group.id}`);
});
```
**Ungroup shapes**


To ungroup shapes with the JavaScript API, get the [group](https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.shape#powerpoint-powerpoint-shape-group-member) property from the group's Shape object then call [ShapeGroup.ungroup](https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.shapegroup#powerpoint-powerpoint-shapegroup-ungroup-member(1)).

The following code sample shows how to ungroup the first shape group found on the current slide.

```
JavaScript
// Ungroups the first shape group on the current slide.
await PowerPoint.run(async (context) => {
 // Get the shapes on the current slide.
 context.presentation.load("slides");
 const slide = context.presentation.getSelectedSlides().getItemAt(0);
 slide.load("shapes/items/type,shapes/items/id");
 await context.sync();
 const shapes = slide.shapes;
 const shapeGroups = shapes.items.filter((item) => item.type ===
PowerPoint.ShapeType.group);
 if (shapeGroups.length === 0) {
 console.warn("No shape groups on the current slide, so nothing to
ungroup.");
 return;
 }
 // Ungroup the first grouped shapes.
 const firstGroupId = shapeGroups[0].id;
 const shapeGroupToUngroup = shapes.getItem(firstGroupId);
 shapeGroupToUngroup.group.ungroup();
 await context.sync();
 console.log(`Ungrouped shapes with group ID: ${firstGroupId}`);
});
```
## **Delete shapes**

Shapes are removed from the slide with the Shape object's delete method.

The following code sample shows how to delete shapes.

```
JavaScript
await PowerPoint.run(async (context) => {
 // Delete all shapes from the first slide.
 const shapes = context.presentation.slides.getItemAt(0).shapes;
 // Load all the shapes in the collection without loading their properties.
 shapes.load("items/$none");
 await context.sync();
```


```
 shapes.items.forEach(function (shape) {
 shape.delete();
 });
 await context.sync();
});
```
# **See also**

- Work with tables using the PowerPoint JavaScript API
- Bind to shapes in a PowerPoint presentation
- [Group or ungroup shapes, pictures, or other objects](https://support.microsoft.com/office/a7374c35-20fe-4e0a-9637-7de7d844724b)


# **Bind to shapes in a PowerPoint presentation**

Article • 04/30/2025

Your PowerPoint add-in can bind to shapes to consistently access them through an identifier. The add-in establishes a binding by calling [BindingCollection.add](https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.bindingcollection#powerpoint-powerpoint-bindingcollection-add-member(1)) and assigning a unique identifier. Use the identifier at any time to reference the shape and access its properties. Creating bindings provides the following value to your add-in.

- Establishes a relationship between the add-in and the shape in the document. Bindings are persisted in the document and can be accessed at a later time.
- Enables access to shape properties to read or update, without requiring the user to select any shapes.

The following image shows how an add-in might bind to two shapes on a slide. Each shape has a binding ID created by the add-in: star and pie . Using the binding ID, the add-in can access the desired shape to update properties.

## **Scenario: Use bindings to sync with a data source**

A common scenario for using bindings is to keep shapes up to date with a data source. Often when creating a presentation, users copy and paste images from the data source into the presentation. Over time, to keep the images up to date, they will manually copy and paste the latest images from the data source. An add-in can help automate this process by retrieving upto-date images from the data source on the user's behalf. When a shape fill needs updating,


the add-in uses the binding to find the correct shape and update the shape fill with the newer image.

In a general implementation, there are two components to consider for binding a shape in PowerPoint and updating it with a new image from a data source.

- 1. **The data source**. This is any source of data or asset library such as Microsoft SharePoint or Microsoft OneDrive.
- 2. **The PowerPoint add-in**. The add-in gets data from the data source based on what the user needs. It converts the data to a Base64-encoded image. This is the only fill type the bound shape can accept. It inserts a shape upon the user's request and binds it with a unique identifier. Then it fills the shape with the Base64 image based on the original data source. Shapes are updated upon the user's request and the add-in uses the binding identifier to find the shape and update the image with the last saved Base64 image.

#### 7 **Note**

You decide the implementation details of how to sync updates from the data source and how to get or create images. This article only describes how to use the Office JS APIs in your add-in to bind a shape and update it with latest images.

## **Create a bound shape in PowerPoint**

Use the PowerPoint.BindingCollection.add() method for the presentation to create a binding which refers to a particular shape.

The following sample shows how to create a shape on the first selected slide.


```
JavaScript
```

```
await PowerPoint.run(async (context) => {
 const slides = context.presentation.getSelectedSlides();
 // Insert new shape on first selected slide. 
 const myShape = slides
 .getItemAt(0)
 .shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, {
 top: 100,
 left: 30,
 width: 200,
 height: 200
 });
 // Fill shape with a Base64-encoded image. 
 // Note: The image is typically created from a data source request. 
 const productsImage = "...base64 image data...";
 myShape.fill.setImage(productsImage);
});
```
Call BindingCollection.add to add the binding to the bindings collection in PowerPoint. The following sample shows how to add a new binding for a shape to the bindings collection.

JavaScript // Create a binding ID to track the shape for later updates. const bindingId = "productChart"; // Create binding by adding the new shape to the bindings collection. context.presentation.bindings.add(myShape, PowerPoint.BindingType.shape, bindingId);

# **Refresh a bound shape with updated data**

After there's an update to the image data, refresh the shape image by finding it via the binding identifier. The following code sample shows how to find a bound shape with the identifier and fill it with an updated image. The image is updated by the add-in based on the data source request or provided by the data source directly.

```
JavaScript
async function updateBinding(bindingId, image) {
 await PowerPoint.run(async (context) => {
 try {
 // Get the shape based on binding ID. 
 const myShape = context.presentation.bindings
 .getItem(bindingId)
 .getShape();
```


```
 // Update the shape to latest image. 
 myShape.fill.setImage(image);
 await context.sync();
 } catch (err) {
 console.error(err);
 }
 });
}
```
# **Delete a binding**

The following sample shows how to delete a binding by deleting it from the bindings collection.

```
JavaScript
async function deleteBinding(bindingId) {
 await PowerPoint.run(async (context) => {
 context.presentation.bindings.getItemAt(bindingId).delete();
 await context.sync();
 });
}
```
# **Load bindings**

When a user opens a presentation and your add-in first loads, you can load all the bindings to continue working with them. The following code shows how to load all bindings in a presentation and display them in the console.

```
JavaScript
async function loadBindings() {
 await PowerPoint.run(async (context) => {
 try {
 let myBindings = context.presentation.bindings;
 myBindings.load("items");
 await context.sync();
 // Log all binding IDs to console.
 if (myBindings.items.length > 0) {
 myBindings.items.forEach(async (binding) => {
 console.log(binding.id);
 });
 }
 } catch (err) {
```


```
 console.error(err);
 }
 });
}
```
# **Error handling when a binding or shape is deleted**

When a shape is deleted, its associated binding is also removed from the PowerPoint binding collection. Any object references you have to the binding, or shape, will return errors if you access any properties or methods on those objects. Be sure to handle potential error scenarios for a deleted shape if your add-in keeps Binding or Shape objects.

The following code shows one approach to error handling when a binding object references a deleted binding. Use a try/catch statement and then call a function to reload all binding and shape references when an error occurs.

```
JavaScript
async function getShapeFromBindingID(id) {
 await PowerPoint.run(async (context) => {
 try {
 const binding = context.presentation.bindings.getItemAt(id);
 const shape = binding.getShape();
 await context.sync();
 return shape;
 } catch (err) {
 console.log(err);
 return undefined;
 }
 });
}
```
# **See also**

When maintaining freshness on shapes, you may also want to check the zOrder. See the [zOrderPosition](https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.shape) property for more information.

- Work with shapes using the PowerPoint JavaScript API
- Bind to regions in a document or spreadsheet


# **Work with tables using the PowerPoint JavaScript API**

Article • 04/30/2025

This article provides code samples that show how to create tables and control formatting by using the PowerPoint JavaScript API.

# **Create an empty table**

To create an empty table, call the [ShapeCollection.addTable()](https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.shapecollection#powerpoint-powerpoint-shapecollection-addtable-member(1)) method and specify how many rows and columns the table needs. The following code sample shows how to create a table with 3 rows and 4 columns.

```
JavaScript
await PowerPoint.run(async (context) => {
 const shapes = context.presentation.getSelectedSlides().getItemAt(0).shapes;
 // Add a table (which is a type of Shape).
 const shape = shapes.addTable(3, 4);
 await context.sync();
});
```
The previous sample doesn't specify any options, so the table defaults to formatting provided by PowerPoint. The following image shows an example of an empty table created with default formatting in PowerPoint.

# **Specify values**

You can populate the table with string values when you create it. To do this provide a 2 dimensional array of values in the [TableAddOptions](https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.tableaddoptions) object. The following code sample creates a table with string values from "1" to "12". Note the following:

- An empty cell must be specified as an empty string "". If a value is undefined or missing, addTable throws an error.
- The outer array contains a list of rows. Each row is an inner array containing a list of string cell values.
- The function named insertTableOnCurrentSlide is used in other samples in this article.


```
JavaScript
```

```
async function run() {
 const options: PowerPoint.TableAddOptions = {
 values: [
 ["1", "2", "", "4"], // Cell 3 is blank.
 ["5", "6", "7", "8"],
 ["9", "10", "11", "12"]
 ],
 };
 await insertTableOnCurrentSlide(3, 4, options);
}
async function insertTableOnCurrentSlide(rowCount: number, columnCount: number,
options: PowerPoint.TableAddOptions) {
 await PowerPoint.run(async (context) => {
 const shapes =
context.presentation.getSelectedSlides().getItemAt(0).shapes;
 // Add a table (which is a type of Shape).
 const shape = shapes.addTable(rowCount, columnCount, options);
 await context.sync();
 });
}
```
The previous sample creates a table with values as shown in the following image.

| 5 | 6  |    | 8  |
|---|----|----|----|
| 9 | 10 | 11 | 12 |

# **Specify cell formatting**

You can specify cell formatting when you create a table, including border style, fill style, font style, horizontal alignment, indent level, and vertical alignment. These formats are specified by the [TableCellProperties](https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.tablecellproperties) object.

# **Uniform cell formatting**

Uniform cell formatting applies to the entire table. For example, if you set the uniform font color to white, all table cells will use the white font. Uniform cell formatting is useful for controlling the default formatting you want on the entire table.

Specify uniform cell formatting for the entire table using the [TableAddOptions.uniformCellProperties](https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.tableaddoptions#powerpoint-powerpoint-tableaddoptions-uniformcellproperties-member) property. The following code sample shows how to set all table cells to dark slate blue fill color and bold white font.


```
JavaScript
```

```
const rowCount = 3;
const columnCount = 4;
const options: PowerPoint.TableAddOptions = {
 values: [
 ["1", "2", "", "4"],
 ["5", "6", "7", "8"],
 ["9", "10", "11", "12"]
 ],
 uniformCellProperties: {
 fill: { color: "darkslateblue" },
 font: { bold: true, color: "white" }
 }
};
await insertTableOnCurrentSlide(rowCount, columnCount, options);
```
The previous sample creates a table as shown in the following image.

| 1   | 2  |    |    |
|-----|----|----|----|
| P   | 6  | 7  | 8  |
| 197 | 10 | 11 | 12 |

## **Specific cell formatting**

Specific cell formatting applies to individual cells and overrides the uniform cell formatting, if any. Set individual cell formatting by using the [TableAddOptions.specificCellProperties](https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.tableaddoptions#powerpoint-powerpoint-tableaddoptions-specificcellproperties-member) property. The following code sample shows how to set the fill color to black for the cell at row 1, column 1.

Note the specificCellProperties must be a 2D array that matches the 2D size of the table exactly. The sample first creates the entire empty 2D array of objects. Then it sets the specific cell format at row 1, column 1, after the options object is created.

```
JavaScript
const rowCount = 3;
const columnCount = 4;
// Compact syntax to create a 2D array filled with empty and distinct objects.
const specificCellProperties = Array(rowCount).fill("").map(_ =>
Array(columnCount).fill("").map(_ => ({})));
const options: PowerPoint.TableAddOptions = {
 values: [
 ["1", "2", "", "4"],
 ["5", "6", "7", "8"],
 ["9", "10", "11", "12"]
 ],
```


```
 uniformCellProperties: {
 fill: { color: "darkslateblue" },
 font: { bold: true, color: "white" }
 },
 specificCellProperties // Array values are empty objects at this point.
};
// Set fill color for specific cell at row 1, column 1.
options.specificCellProperties[1][1] = {
 fill: { color: "black" }
};
await insertTableOnCurrentSlide(rowCount, columnCount, options);
```
The previous sample creates a table with a specific format applied to the cell in row 1, column 1 as shown in the following image.

|   | 2  |    |    |
|---|----|----|----|
| 5 | 6  |    | 8  |
| 9 | 10 | 11 | 12 |

The previous sample uses the [font](https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.tablecellproperties#powerpoint-powerpoint-tablecellproperties-font-member) property which is of type [FontProperties](https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.fontproperties). The font property allows you to specify many properties, such as bold, italic, name, color, and more. The following code sample shows how to specify multiple properties for a font for a cell.

```
JavaScript
options.specificCellProperties[1][1] = {
 font: {
 color: "orange",
 name: "Arial",
 size: 50,
 allCaps: true,
 italic: true
 }
};
```
You can also specify a [fill](https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.tablecellproperties#powerpoint-powerpoint-tablecellproperties-fill-member) property which is of type [FillProperties](https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.fillproperties). The fill property can specify a color and the transparency percentage. The following code sample shows how to create a fill for all table cells using the color "light red" and a 50% transparency.

JavaScript

```
uniformCellProperties: {
 fill: {
 color: "lightred",
 transparency: 0.5
```


# **Borders**

Use the [TableCellProperties.borders](https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.tablecellproperties#powerpoint-powerpoint-tablecellproperties-borders-member) object to define borders for cells in the table. The following code sample shows how to set the borders of a cell in row 1 by column 1 to a red border with weight 3.

```
JavaScript
```

```
const columnCount = 3;
const rowCount = 3;
// Compact syntax to create a 2D array filled with empty and distinct objects.
const specificCellProperties = Array(rowCount).fill(undefined).map(_ =>
Array(columnCount).fill(undefined).map(_ => ({})));
const options: PowerPoint.TableAddOptions = {
 values: [
 ["1", "2", "3"],
 ["4", "5", "6"],
 ["7", "8", "9"]
 ],
 uniformCellProperties: {
 fill: {
 color: "lightcyan",
 transparency: 0.5
 },
 },
 specificCellProperties
};
options.specificCellProperties[1][1] = {
 font: {
 color: "red",
 name: "Arial",
 size: 50,
 allCaps: true,
 italic: true
 },
 borders: {
 bottom: {
 color: "red",
 weight: 3
 },
 left: {
 color: "red",
 weight: 3
 },
 right: {
 color: "red",
 weight: 3
 },
```


```
 top: {
 color: "red",
 weight: 3
 }
 }
};
await insertTableOnCurrentSlide(rowCount, columnCount, options);
```
## **Horizontal and vertical alignment**

Use the [TableCellProperties.horizontalAlignment](https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.tablecellproperties#powerpoint-powerpoint-tablecellproperties-horizontalalignment-member) property to control text alignment in a cell. The following example shows how to set horizontal alignment to left, right, and center for three cells in a table. For a list of all alignment options, see the [ParagraphHorizontalAlignment](https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.paragraphhorizontalalignment) enum.

```
JavaScript
const rowCount = 3;
const columnCount = 3;
// Compact syntax to create a 2D array filled with empty and distinct objects.
const specificCellProperties = Array(rowCount).fill("").map(_ =>
Array(columnCount).fill("").map(_ => ({})));
const options: PowerPoint.TableAddOptions = {
 values: [
 ["Left aligned, top", "\n\n", ""],
 ["Centered", "\n\n", ""],
 ["Right aligned, bottom", "\n\n", ""]
 ],
 uniformCellProperties: {
 fill: { color: "lightblue" },
 borders: {
 bottom: {
 color: "black",
 weight: 3
 },
 left: {
 color: "black",
 weight: 3
 },
 right: {
 color: "black",
 weight: 3
 },
 top: {
 color: "black",
 weight: 3
 }
 }
 },
 specificCellProperties // Array values are empty objects at this point.
};
```


```
options.specificCellProperties[0][0] = {
 horizontalAlignment: PowerPoint.ParagraphHorizontalAlignment.left,
 verticalAlignment: 0 //PowerPoint.TextVerticalAlignment.top
};
options.specificCellProperties[1][0] = {
 horizontalAlignment: PowerPoint.ParagraphHorizontalAlignment.center,
 verticalAlignment: 1 //PowerPoint.TextVerticalAlignment.middle
};
options.specificCellProperties[2][0] = {
 horizontalAlignment: PowerPoint.ParagraphHorizontalAlignment.right,
 verticalAlignment: 2 //PowerPoint.TextVerticalAlignment.bottom
};
await insertTableOnCurrentSlide(3, 3, options);
```
The previous sample creates a table with left/top, centered, and right/bottom text alignment as shown in the following image.

| Left aligned, top     |  |  |
|-----------------------|--|--|
| Centered              |  |  |
| Right aligned, bottom |  |  |

# **Specify row and column widths**

Specify row and column widths using the [TableAddOptions.rows](https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.tableaddoptions#powerpoint-powerpoint-tableaddoptions-rows-member) and [TableAddOptions.columns](https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.tableaddoptions#powerpoint-powerpoint-tableaddoptions-columns-member) properties. The rows property is an array of [TableRowProperties](https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.tablerowproperties) that you use to set each row's [rowHeight](https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.tablerowproperties#powerpoint-powerpoint-tablerowproperties-rowheight-member) property. Similarly, the columns property is an array of [TableColumnProperties](https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.tablecolumnproperties) you use to set each column's [columnWidth](https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.tablecolumnproperties#powerpoint-powerpoint-tablecolumnproperties-columnwidth-member) property. The width or height is set in points.

The height or width that you set may not be honored by PowerPoint if it needs to fit the text. For example, if the text is too wide for a column, PowerPoint will increase the row height so that it can wrap the text to the next line. Similarly, the column width will increase if the specified size is smaller than a single character in the specified font size.

The following code example shows how to set row height and column width for a new table. Note that the rows and columns properties must be set to an array of objects equal to their count.


```
const columnCount = 3;
const rowCount = 3;
const options: PowerPoint.TableAddOptions = {
 values: [
 ["Width 72pt", "Width 244pt", "Width 100pt"],
 ["", "", ""],
 ["", "^\n\nHeight 200 pt\n\nv", ""]
 ],
 // Initialize columns with an array of empty objects for each column.
 columns: Array(columnCount).fill("").map(_ => ({})),
 rows: Array(columnCount).fill("").map(_ => ({})),
 uniformCellProperties: {
 fill: { color: "lightcyan" },
 horizontalAlignment: PowerPoint.ParagraphHorizontalAlignment.center,
 verticalAlignment: 1, //PowerPoint.TextVerticalAlignment.middle
 borders: {
 bottom: {
 color: "black",
 weight: 3
 },
 left: {
 color: "black",
 weight: 3
 },
 right: {
 color: "black",
 weight: 3
 },
 top: {
 color: "black",
 weight: 3
 }
 }
 }
};
options.columns[0].columnWidth = 72;
options.columns[1].columnWidth = 244;
options.columns[2].columnWidth = 100;
options.rows[2].rowHeight = 200;
await insertTableOnCurrentSlide(rowCount, columnCount, options);
```
The previous sample creates a table with three custom column widths, and one custom row height, as shown in the following image.


| Width<br>72pt | Width 244pt   | Width<br>100pt |
|---------------|---------------|----------------|
|               |               |                |
|               |               |                |
|               | へ             |                |
|               | Height 200 pt |                |
|               | V             |                |
|               |               |                |

# **Specify merged areas**

A merged area is two or more cells combined so that they share a single value and format. In appearance the merged area spans multiple rows or columns. A merged area is indexed by its upper left table cell location (row, column) when setting its value or format. The upper left cell of the merged area is always used to set the value and formatting. All other cells in the merged area must be empty strings with no formatting applied.

To specify a merged area, provide the upper left location where the area starts (row, column) and the length of the area in rows and columns. The following diagram shows an example of these values for a merged area that is 3 rows by 2 columns in size. Note that merged areas can't overlap with each other.

Use the [TableAddOptions.mergedAreas](https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.tableaddoptions#powerpoint-powerpoint-tableaddoptions-mergedareas-member) property to specify one or more merged areas. The following code sample shows how to create a table with two merged areas. About the code sample, note the following:

- The values property must only specify the value for the upper left corner of the merged area. All other cell values in the merged area must specify empty strings ("").
- Each merged area must specify the upper left corner location (row, column) and the length in cells of the merged area in terms of row count and column count.

JavaScript

```
const rowCount = 3;
const columnCount = 4;
```


```
// Compact syntax to create a 2D array filled with empty and distinct objects.
const specificCellProperties = Array(rowCount).fill("").map(_ =>
Array(columnCount).fill("").map(_ => ({})));
const options: PowerPoint.TableAddOptions = {
 values: [
 ["1", "This is a merged cell", "", "4"],
 ["5", "6", "This is also a merged cell", "8"],
 ["9", "10", "", "12"]
 ],
 uniformCellProperties: {
 fill: { color: "darkslateblue" },
 font: { bold: true, color: "white" },
 borders: {
 bottom: {
 color: "black",
 weight: 3
 },
 left: {
 color: "black",
 weight: 3
 },
 right: {
 color: "black",
 weight: 3
 },
 top: {
 color: "black",
 weight: 3
 }
 }
 },
 mergedAreas: [{ rowIndex: 0, columnIndex: 1, rowCount: 1, columnCount: 2 },
 { rowIndex: 1, columnIndex: 2, rowCount: 2, columnCount: 1 }
 ],
 specificCellProperties // Array values are empty objects at this point.
};
// Set fill color for specific cell at row 1, column 1.
options.specificCellProperties[1][1] = {
 fill: { color: "black" }
};
await insertTableOnCurrentSlide(rowCount, columnCount, options);
```
The previous sample creates a table with two merged areas as shown in the following image.

|   | This is a merged cell |                               |    |
|---|-----------------------|-------------------------------|----|
| 5 | (5                    | This is also a<br>merged cell |    |
| 9 | 10                    |                               | 12 |

## **Get and set table cell values**


After a table is created you can get or set string values in the cells. Note that this is the only part of a table you can change. You can't change borders, fonts, widths, or other cell properties. If you need to update a table, delete it and recreate it. The following code sample shows how to find an existing table and set a new value for a cell in the table.

```
JavaScript
await PowerPoint.run(async (context) => {
 // Load shapes.
 const shapes = context.presentation.getSelectedSlides().getItemAt(0).shapes;
 shapes.load("items");
 await context.sync();
 // Find the first shape of type table.
 const shape = shapes.items.find((shape) => shape.type ===
PowerPoint.ShapeType.table)
 const table = shape.getTable();
 table.load("values");
 await context.sync();
 // Set the value of the specified table cell.
 let values = table.values;
 values[1][1] = "A new value";
 table.values = values;
 await context.sync();
});
```
You can also get the following read-only properties from the table.

- **rowCount**
- **columnCount**

The following sample shows how to get the table properties and log them to the console. The sample also shows how to get the merged areas in the table.

```
JavaScript
await PowerPoint.run(async (context) => {
 // Load shapes.
 const shapes = context.presentation.getSelectedSlides().getItemAt(0).shapes;
 shapes.load("items");
 await context.sync();
 // Find the first shape of type table.
 const shape = shapes.items.find((shape) => shape.type ===
PowerPoint.ShapeType.table)
 const table = shape.getTable();
 // Load row and column counts.
 table.load("rowCount, columnCount");
 // Load the merged areas.
 const mergedAreas = table.getMergedAreas();
 mergedAreas.load("items");
 await context.sync();
```


```
 // Log the table properties.
 console.log(mergedAreas);
 console.log(table.rowCount);
 console.log(table.columnCount);
```
});


# **Project add-ins documentation**

With Project add-ins, you can use familiar web technologies such as HTML, CSS, and JavaScript to build a solution that can run in Project on Windows. Learn how to build, test, debug, and publish Project add-ins.

| About Project add-ins               |  |
|-------------------------------------|--|
| e<br>OVERVIEW                       |  |
| What are Project add-ins?           |  |
| f<br>QUICKSTART                     |  |
| Build your first Project add-in     |  |
| c<br>HOW-TO GUIDE                   |  |
| Test and debug a Project add-in     |  |
| Deploy and publish a Project add-in |  |
| Key Office Add-ins concepts         |  |
| e<br>OVERVIEW                       |  |
| Office Add-ins platform overview    |  |
| b<br>GET STARTED                    |  |
| Core concepts for Office Add-ins    |  |

Design Office Add-ins

Develop Office Add-ins

#### **Resources**

i **REFERENCE**


[Ask questions](https://stackoverflow.com/questions/tagged/office-js)

[Request features](https://feedbackportal.microsoft.com/feedback/forum/40792262-301c-ec11-b6e7-0022481f8472)

[Report issues](https://github.com/officedev/office-js/issues)

Join Office Add-ins community call

Office Add-ins additional resources

[Download samples](https://developer.microsoft.com/microsoft-365/gallery/?filterBy=Project,Samples)

# **Build your first PowerPoint task pane add-in**

Article • 09/17/2024

In this article, you'll walk through the process of building a PowerPoint task pane add-in.

### **Prerequisites**

- Node.js (the latest LTS version). Visit the [Node.js site](https://nodejs.org/) to download and install the right version for your operating system.
- The latest version of Yeoman and the Yeoman generator for Office Add-ins. To install these tools globally, run the following command via the command prompt.

command line

npm install -g yo generator-office

7 **Note**

Even if you've previously installed the Yeoman generator, we recommend you update your package to the latest version from npm.

- Office connected to a Microsoft 365 subscription (including Office on the web).
7 **Note**

If you don't already have Office, you might qualify for a Microsoft 365 E5 developer subscription through the **[Microsoft 365 Developer Program](https://aka.ms/m365devprogram)** ; for details, see the **[FAQ](https://learn.microsoft.com/en-us/office/developer-program/microsoft-365-developer-program-faq#who-qualifies-for-a-microsoft-365-e5-developer-subscription-)**. Alternatively, you can **[sign up for a 1-month free trial](https://www.microsoft.com/microsoft-365/try)** or **[purchase a Microsoft 365 plan](https://www.microsoft.com/microsoft-365/business/compare-all-microsoft-365-business-products-g)** .

## **Create the add-in project**

Run the following command to create an add-in project using the Yeoman generator. A folder that contains the project will be added to the current directory.


#### 7 **Note**

When you run the yo office command, you may receive prompts about the data collection policies of Yeoman and the Office Add-in CLI tools. Use the information that's provided to respond to the prompts as you see fit.

When prompted, provide the following information to create your add-in project.

- **Choose a project type:** Office Add-in Task Pane project
- **Choose a script type:** Javascript
- **What do you want to name your add-in?** My Office Add-in
- **Which Office client application would you like to support?** PowerPoint

After you complete the wizard, the generator creates the project and installs supporting Node components.

## **Explore the project**

The add-in project that you've created with the Yeoman generator contains sample code for a basic task pane add-in. If you'd like to explore the components of your add-in project, open the project in your code editor and review the files listed below. When you're ready to try out your add-in, proceed to the next section.

- The **./manifest.xml** or **manifest.json** file in the root directory of the project defines the settings and capabilities of the add-in.


- The **./src/taskpane/taskpane.html** file contains the HTML markup for the task pane.
- The **./src/taskpane/taskpane.css** file contains the CSS that's applied to content in the task pane.
- The **./src/taskpane/taskpane.js** file contains the Office JavaScript API code that facilitates interaction between the task pane and the Office client application.

## **Try it out**

- 1. Navigate to the root folder of the project.
command line cd "My Office Add-in"

- 2. Complete the following steps to start the local web server and sideload your addin.
#### 7 **Note**

- Office Add-ins should use HTTPS, not HTTP, even while you're developing. If you're prompted to install a certificate after you run one of the following commands, accept the prompt to install the certificate that the Yeoman generator provides. You may also have to run your command prompt or terminal as an administrator for the changes to be made.
- If this is your first time developing an Office Add-in on your machine, you may be prompted in the command line to grant Microsoft Edge WebView a loopback exemption ("Allow localhost loopback for Microsoft Edge WebView?"). When prompted, enter Y to allow the exemption. Note that you'll need administrator privileges to allow the exemption. Once allowed, you shouldn't be prompted for an exemption when you sideload Office Add-ins in the future (unless you remove the exemption from your machine). To learn more, see **["We can't open this](https://learn.microsoft.com/en-us/office/troubleshoot/office-suite-issues/cannot-open-add-in-from-localhost) [add-in from localhost" when loading an Office Add-in or using Fiddler](https://learn.microsoft.com/en-us/office/troubleshoot/office-suite-issues/cannot-open-add-in-from-localhost)**.


#### **Tip**

If you're testing your add-in on Mac, run the following command before proceeding. When you run this command, the local web server starts.

command line

npm run dev-server

- To test your add-in in PowerPoint, run the following command in the root directory of your project. This starts the local web server (if it's not already running) and opens PowerPoint with your add-in loaded.
command line npm start

- To test your add-in in PowerPoint on a browser, run the following command in the root directory of your project. When you run this command, the local web server starts. Replace "{url}" with the URL of a PowerPoint document on your OneDrive or a SharePoint library to which you have permissions.
#### 7 **Note**

If you are developing on a Mac, enclose the {url} in single quotation marks. Do *not* do this on Windows.

command line

```
npm run start -- web --document {url}
```
The following are examples.


- npm run start -- web --document https://contoso.sharepoint.com/:t:/g/EZGxP7ksiE5DuxvY638G798BpuhwluxCM fF1WZQj3VYhYQ?e=F4QM1R npm run start -- web --document
https://1drv.ms/x/s!jkcH7spkM4EGgcZUgqthk4IK3NOypVw?e=Z6G1qp

- npm run start -- web --document https://contoso-my.sharepointdf.com/:t:/p/user/EQda453DNTpFnl1bFPhOVR0BwlrzetbXvnaRYii2lDr_oQ? e=RSccmNP
If your add-in doesn't sideload in the document, manually sideload it by following the instructions in Manually sideload add-ins to Office on the web.

- 3. In PowerPoint, insert a new blank slide, choose the **Home** tab, and then choose the **Show Taskpane** button on the ribbon to open the add-in task pane.
- 4. At the bottom of the task pane, choose the **Run** link to insert the text "Hello World" into the current slide.


- 5. When you want to stop the local web server and uninstall the add-in, follow the applicable instructions:
	- To stop the server, run the following command. If you used npm start , the following command also uninstalls the add-in.

| command line |  |  |  |  |  |
|--------------|--|--|--|--|--|
| npm stop     |  |  |  |  |  |

- If you manually sideloaded the add-in, see Remove a sideloaded add-in.
### **Next steps**

Congratulations, you've successfully created a PowerPoint task pane add-in! Next, learn more about the capabilities of a PowerPoint add-in and build a more complex add-in by following along with the PowerPoint add-in tutorial.

## **Troubleshooting**

- Ensure your environment is ready for Office development by following the instructions in Set up your development environment.
- Some of the sample code uses ES6 JavaScript. This isn't compatible with older versions of Office that use the Trident (Internet Explorer 11) browser engine. For information on how to support those platforms in your add-in, see Support older


Microsoft webviews and Office versions. If you don't already have a Microsoft 365 subscription to use for development, you might qualify for a Microsoft 365 E5 developer subscription through the [Microsoft 365 Developer Program](https://aka.ms/m365devprogram) ; for details, see the [FAQ](https://learn.microsoft.com/en-us/office/developer-program/microsoft-365-developer-program-faq#who-qualifies-for-a-microsoft-365-e5-developer-subscription-). Alternatively, you can [sign up for a 1-month free trial](https://www.microsoft.com/microsoft-365/try) or [purchase a Microsoft 365 plan](https://www.microsoft.com/microsoft-365/business/compare-all-microsoft-365-business-products-g) .

- The automatic npm install step Yo Office performs may fail. If you see errors when trying to run npm start , navigate to the newly created project folder in a command prompt and manually run npm install . For more information about Yo Office, see Create Office Add-in projects using the Yeoman Generator.
### **Code samples**

- [PowerPoint "Hello world" add-in](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/hello-world/powerpoint-hello-world) : Learn how to build a simple Office Add-in with only a manifest, HTML web page, and a logo.
### **See also**

- Office Add-ins platform overview
- Develop Office Add-ins
- Using Visual Studio Code to publish

# **Build your first PowerPoint task pane addin with Visual Studio**

06/20/2025

In this article, you'll walk through the process of building a PowerPoint task pane add-in.

### **Prerequisites**

- [Visual Studio 2019 or later](https://www.visualstudio.com/vs/) with the **Office/SharePoint development** workload installed.
7 **Note**

If you've previously installed Visual Studio, use the Visual Studio Installer to ensure that the **Office/SharePoint development** workload is installed.

- Office connected to a Microsoft 365 subscription (including Office on the web).
### **Create the add-in project**

- 1. In Visual Studio, choose **Create a new project**.
- 2. Using the search box, enter **add-in**. Choose **PowerPoint Web Add-in**, then select **Next**.
- 3. Name your project and select **Create**.
- 4. In the **Create Office Add-in** dialog window, choose **Add new functionalities to PowerPoint**, and then choose **Finish** to create the project.
- 5. Visual Studio creates a solution and its two projects appear in **Solution Explorer**. The **Home.html** file opens in Visual Studio.

# **Explore the Visual Studio solution**

When you've completed the wizard, Visual Studio creates a solution that contains two projects.

ノ **Expand table**


| Project                       | Description                                                                                                                                                                                                                                                                                                                                                                                                                                         |
|-------------------------------|-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| Add-in<br>project             | Contains only an XML-formatted add-in only manifest file, which contains all the settings<br>that describe your add-in. These settings help the Office application determine when your<br>add-in should be activated and where the add-in should appear. Visual Studio generates<br>the contents of this file for you so that you can run the project and use your add-in<br>immediately. Change these settings any time by modifying the XML file. |
| Web<br>application<br>project | Contains the content pages of your add-in, including all the files and file references that<br>you need to develop Office-aware HTML and JavaScript pages. While you develop your<br>add-in, Visual Studio hosts the web application on your local IIS server. When you're<br>ready to publish the add-in, you'll need to deploy this web application project to a web<br>server.                                                                   |

### **Update the code**

- 1. **Home.html** specifies the HTML that will be rendered in the add-in's task pane. In **Home.html**, replace the <body> element with the following markup and save the file.

```
HTML
<body class="ms-font-m ms-welcome">
 <div id="content-header">
 <div class="padding">
 <h1>Welcome</h1>
 </div>
 </div>
 <div id="content-main">
 <div class="padding">
 <p>Select a slide and then choose the buttons to below to add
content to it.</p>
 <br />
 <h3>Try it out</h3>
 <button class="ms-Button" id="insert-image">Insert Image</button>
 <br/><br/>
 <button class="ms-Button" id="insert-text">Insert Text</button>
 </div>
 </div>
</body>
```
- 2. Open the file **Home.js** in the root of the web application project. This file specifies the script for the add-in. Replace the entire contents with the following code and save the file.
JavaScript 'use strict'; (function () {


```
 Office.onReady(function() {
 // Office is ready
 $(document).ready(function () {
 // The document is ready
 $('#insert-image').on("click", insertImage);
 $('#insert-text').on("click", insertText);
 });
 });
 function insertImage() {

Office.context.document.setSelectedDataAsync(getImageAsBase64String(), {
 coercionType: Office.CoercionType.Image,
 imageLeft: 50,
 imageTop: 50,
 imageWidth: 400
 },
 function (asyncResult) {
 if (asyncResult.status === Office.AsyncResultStatus.Failed) {
 console.log(asyncResult.error.message);
 }
 });
 } 
 function insertText() {
 Office.context.document.setSelectedDataAsync("Hello World!",
 function (asyncResult) {
 if (asyncResult.status === Office.AsyncResultStatus.Failed) {
 console.log(asyncResult.error.message);
 }
 });
 }
 function getImageAsBase64String() {
 return
'iVBORw0KGgoAAAANSUhEUgAAAZAAAAEFCAIAAABCdiZrAAAACXBIWXMAAAsSAAALEgHS3X78AAAb
X0lEQVR42u2da2xb53nH/xIpmpRMkZQs2mZkkb7UV3lifFnmNYnorO3SLYUVpFjQYoloYA3SoZjVZ
Ri2AVtptF+GNTUzbGiwDQu9deg2pCg9FE3aYQ3lDssw2zGNKc5lUUr6ItuULZKiJUoyJe2DFFsXXs
6VOpf/D/kS6/Ac6T2Hv/M8z3nf5zTMz8+DEEL0QCOHgBBCYRFCCIVFCKGwCCGEwiKEEAqLEEJhEUI
IhUUIIRQWIYTCIoQQCosQQigsQgiFRQghFBYhhFBYhBAKixBC1hArh2CBwtlYaTRV6ac2f7Cx2Q3A
sTfEsSKEwlprYQ3Gpt4bFLixfU+vpdltCwTte0JNHQFrR4ADSEgdaGCL5AVGvhkSLqyV1t/gd+wN2
feGHHtClBchFJbq3Hq5b+LCGfn7sfl7nI+HWw710VyEUFhqkf1BJPuDkwrusOXgsfW94ZZDfRxb8o
BCEpn4yn90BmF1ozUIq5sjVCOb4RCoxMSFMxMXzlg3+D1fjDgfD3NAzE4ph6EwMlWjeKsLziDaQvC
E0BbimDHCquyX8/Fb33lapZ3b9/RueD5q8wc5zuYl2VfDVqvx9MLbB28fHCwvUFjLmUknr/3xw6oe
wvPMNzzPRDjUZmQsgfNHpX/cewzePvgYp1NYS/j4yw1qH8K+p3fTS/GFKV3ERLw/gCuvyN2Jww9fG
P4BM5e6ONP9ATZ/j9qHmHpvcOSbobnJHEfbXBSSCuykmMbwSZwNYDiCkkkvIQpryQ1sT6guueclOo
tIp5Rf1NZIjMIyNfZ6LbuZSV8a/W6YA05kaWvoOM6FlIndKCxdRlh1XCc4ceFM/o0ox9wsqDRHITu
Itx9G2kQXEoW1ZCya3S0Hj9XtcNkfRJgYmgVfGFaXWjv/4Os4FzJJVYvCWkbz4fpNTJ+bzDPIMk30
HsDuqIrOyg7i7aAZ0kNOa1ghkVzqdzx1jOlcgb9jkGUaiimkow+0UkiilFdy/1YXdkeNPV2LwlrJ6
KvhwtnT5f1iQYsbdifWNcPmkH2k/SK3X5j37B/gOTIaYwlMpTCeRDaBwiW5e+t+zcDOorBWUnbKu9
UGjw/OdkWPtF/SpzY9C18YG57kmTImpRwycWTiotfxmMNZFFYZlvbGarTA44PLq8Jh9sv4rMOPfTG
ujzW4ua7HcCWKYprOorCqlhouJ2586ygAWzO8ASWyP8WFtUDXCexm2d7w988YhiNStGVEZ1FYFYOs
ufSgbycaLeocwA58Son9eHrxcJx9lIzPcATpqOgi/ZGLcBqqRwiFVZ7ZD37ccOY31bIVgBZgm0K7c
vbgSJKnzASRfwpDYWTFNPK2uvB4ykj3M87DKsd0znL2d1W0FQAF08zCJQyFedKMjyOAwwnsOiXiI6
```


U8zoWMNAYUVjnifRhPq3uIJmUz2NNlGu8SQ+IfwJGLIuagFi5hOEJhGZcLUVwbVP0oihfyh8KmbTl iOpxBHEnCKbgb0vBJjCUoLGMmg3i7LrejFqV3WMqbahEs00McTohw1rsGKRpQWKvCq+m86kdpUWe3 FJapsLpFOKuYNkZiSGGtCK9O1uNArerstpRnJcuMzhJYz0pHUUxRWMYKr+qDDGEVpiwXPnZe+NhZ/ scUFp1V5X6m/yCL87CW8FfueuSDMqaMJi67I68H7k5ZAGx2z7z83PDOzZPLtuCcLHMyEsPQcUFbPv YLXb80jBHWJ7wbq4etAMjoXnPfVgBu5Gwv/eP2VQHYJZ5JM+ILwyus96TOgywK6xM+qlcyJVVYH95 ovm+r+87ieSOLdMcEJYYjp3U9/YWvqgcATOfw0Zl6HMgDSJ1AvzL7A9bbZ8ts9/OAkIWyh/7kYJWf bt68+eWXX965cycvDf18ld3YHRWUGKaj2K7XOIsRFgDgaqJOB5LXpuapA3eW/u+XP50ps5GwZf3lZ Xc/drtx44UXXvjwww95aegsMfT0CgiyYkwJmQ8KC6/k5XAvPXX1qQN3DmwtHNha+MYXUy/82ojkXa 2O11Zw9+7db3/727w0dIaQ0KmY1u/TZKaEdYywZHcBdNpnI19MKfK7HNp2951fOKtv88477/DS0Bl tIXh6a3d0yMTh7dPj38cICxhPqb7UGcAGueGVshzcWuCZNyZCuv7rNsKisICM+hOXLAqEVwoLa1uh ehmL6BVvHxz+GtuU8jp9JxiFVRdhdUp/OKiqs3jyjYmQzsj6DLIoLPULWK2qLR6UR2gv29GYWFj6b DhDYQHjKRV33gR0avTv/sKBO8wKjYkjUDsrZEqoW2GpVnG3AAEtJoP3KT+TixiAjloPAUt5PTZvML 2wVC1gbQbsmv7rv/TpWwyyjImQd1bqMMgyvbCmVavjdMpa51wfnPZZBlnGRMjbvSgsRlgPYiuPPga AQZYxEdJD5p7+nrqYfqa7GhFWp25stRBkJf6MLbSMSM0p74ywiL5sRQiFpSsUnIRlAfy0FdEMQuru TAlNih3YovVngoToPSWksJRgA+DV9HwrQspQyuvuV6aw5NEEbFHtPYOEEApLGda54MpjIweC6BbhL 47WDHxKKIl9/fhKirYimqbmNCurmxGW0aOqAwM4OIB1bg4G0ToFA06vo7CEseMY9oWxo48jQQiFpU k6erAlhC0heorokpqd3XU4Ucv0wvIuWSNqd6MjiHVueIPYEuIFX7unEtEsQlrHsIalP45GeW1XxEN r6ze8StTeRkhHB43Bp4SkMkI67RJtIqQDciuFRYwTXvUacjEaI6wH+b4OU0IKi1RAyDuEiTYpJFGs 1fhbh/kghUUq0HWC4ZWOuR4TEEHr8vxSWGQVvn7s5rMIPTMq4J2DfFU9MYitumMcBh2TiQvIB3sE9 VDWHpw4Sj7B4ceuqE5vvOQBaQHRsW4nrFBYpsfqQlsI3j5OYjACY4naE9wB+AcoLCKD/1ViJ/uBz8 1zLE3NcKT2NrrNB8EallZgb2Uin5GYscMrCksz2DgERB6lHN4XYCKrS9e5P4WlDdhkmchkKCyoR7u ewysKSzOs5xAQeclg5oyg8IrCIgpgB5o4CkQShSSGjgva0j+gx/WDFJYmcXEIiHiKKZwLCdrS4TfA +lBOa9AMHuC27J38tIEDWSccftgDcAbhCMAZXJull6Uckn1CXy+4ywjLrSgsLWWFdmCKA6GX0CaNY nrZNALvMXhC8PbVaZZTKYdzIRQuCdrYe8wYaxga5uc51VAz4UwWuCbj4/t5SjQSLPfiobC6swcKSZ wLCY2trC48ntJ79WoB1rA0lhWy9G4AsoMYOo6fuTEcQSmn/P4zcRG2AtAdM4atKCzt4eMQGIVSHsM ncTagsLbeH0DyaRG26jphpAXtTAm1lBIu8DEwwZTQWCjSCWMsgQ8GhBatFnD24IihXqfKCEt7dAIW joKxKKaRfBrJPomhVjGFoTDOHxVtq8MJgw0kIyztRVgAxoG0+E8xwtI+VheCcRFzIIopjMSQjorIA e8f6HBCp43bKSy9CQvALSBDYRmUXadqL5HJxJGJY+S0RC0a0VbgPCztshG4B2Q5EEbkg6+jkCzTir qUw1gCmTiyidptjs1nKwpL23QCduAGB8KIjJzG9E1s/SOUcigkUUyhkBRXojKfrZgSajglvM84cA2 YZUpIarFQZTfKlKuy8Cmh5mkFdgEeDgSpiq/f8LaisGSH3/XBAnRSW6Qyu04ZaTo7haVSBF7fSoEN 6AT2Ap1AK+dqkU/SwCMX9d6WT0SQwDOuMyyA55NQaxaYAmaBIsfFlAG+f8AALa4orDpGWEJeUqKqv BaawbfyZJgMXz+2R/T7ti4Kay1ocnMMSL3x9GJ7ZG36BVJYuo+wCGFURWHpA3NfOqRe98Ue+MJ4KG yGh4AUFiMsok/q3HCZwjLL3U/+cgpCltL+6zj4JoehLJyHxSCLaIw7P1GlsTKFRSgsogrXYxwDCov CIjrhSpRjQGGpgFmnwxB1KaZRSHIYKCwV8PRyDIjyZOIcAwpLBQz0DiVCYVFYzAoJEU/hEp8VUlgq 4AzC4ecwEOUZZxmLwlIDD4MsogLZBMeAwlIBlrGIKlkhIywKSyVh1a1dMjEP91jDorBUwhfmGBClU 8JBjgGFpQ4PUViEUFh6wRmEs4fDQAiFpRNM8+YSQigs/cPSOyEUlm6wull6J4TCYlZICKGwFMcRgK +fw0AIhaUTTPYmXkIoLAZZhBAKi0EWIRQWWRlkdZ3gMBBCYemEHRHOySJEcfgiVcFM5/BRHFcTyCQ xKuDlqRuAzRw1QiisOnM1gXei+OiMuE/dBjyAncNHCIVVHzJJvDWAa1K7fIwA2ziIhFBYdeC/Inj7 pKw9TAC3gQ0cSkIoLPWYzuFfQoIKVbVjNMAFNHFMCVEAPiUslwYqZSsAs0CKY0oIIyw1GE/hX0OYz iu5zykgA3g5uIQwwlI2E4z3KWyrBW4BExxfQigsBXkzrFgmuJo0MMshJoTCUoQL4mdaiWIW+JijTA iFpUgy+HZE9aNMAdc41oRQWDJ5a0CV0tVqssBtDjchFJZkxlN493T9DncDyHLQCaGwpDEUq/cRr/G hISFS4Dws4N3YGhw0DWzj0mgVmZvBzBhmZzAzVuFebcO6NljXw7qeo0Vh6YVMEuPpNTjuwkNDOktR pm6ieBMzYyjexNyMiA86NsHWBvsmODah0caBpLA0y0fxNTs0naUQE1cweQUTV8RJainFmyjeRP7yo ryau9DSxchLi5i+hnU1sZZHX3AW61nSaN6J7tfwRLbpsxfh65dsq9XyuvM/uPI6bv0ME1c4ytqiYX 5+3tQD8LeBtUkJV9AJeHg1CsayHr/0fXQ8tfTf5iZz+Tei+Teic5NKzlCxrocnCOeONfpLPzfPs01 hLeHlBq38JpvZOUsYzh4cTsDqLvtDo2mLwmJKKCMO6lVx5zeAIoe4Fr5+HElWshWAxma355lI11+m nI8r+XbI0l2M/ieu/RumbvIcMMLSS4T10jwAZJIYTSKTRCYpvXtymTQH2MsLsirdr8EXFr558XJi9 NVw6bbCKb9rLzzBej1MZIRFYckV1gqU8pcH6OQFWQGHH8E4nEGxn5ubzGVfj+TffEXZX8e6Ht5HYd 9EYVFYuhOWUv7yA628IMvhPYbuWJU0sCYT5+Ojr4aVrWoB8AThCar8t1NYFJa6wpLsr72AhRfkikj Ghd1RUWlgJUqjqZvf6ZtJK9zvrKULHY+qmR5SWBRWXYUl0F+tgJ9X44oAphfdMTgCSu1vbjI3+t3w xAWFu57Z2rDpCdVmmVJYFNZaCquSvyb+GbZbvBwfBFbdMXj71Nj36KvhwlmFm3M02uB7ErY2Ckt1O K1hTfEGsS+Mo1E0T3EwFuk6gcdTKtkKQMeLMWVnPACYm8HIm4Czh2ePwjIBhSRKeQ4DPL147BfYHZ VTX19DZ+Fwgs6isEzA9RhVhUNv4XBCwYpV/Z0Fq5vOorBMwGjcvH/7fVW1hep8ZBWd5eADFArLwPl gMW3GP9zXjyMX10RV92l/PmrzKx0QWd0IxmF18dKmsIzIWEKZ/ez8C30kIw4/dp3CE1l0xyTMXFf4 6m92+/400distFycQQTjvLQpLCMyElNgJ95jCPwBjiRx5CK6TmgxJbG64OvHobfwWAr+AbXL6mKdp


ciuZtLJB//TFsL2bygTgBMKSysUUygoMfe6a+DBvX13FI+ltGIuhx++fgR/iCdy6I6tYfZXBZs/2P 7cKfn7mZ3ILfv/7RF4j8nd6b0cvyXLbnwcgrUkE1dGCqtFsGCu3VEUU8jEkU1gLFGnyRNWF9pC8IT QFlrzpE8grs8PTF1OKD4JHt0xnA1wzgqFRWEtv5NX01kA/gH4BxYDumwC40kUkgpP/vL0whlEaxDO oF4ktYKOr8aKvxdQeIH0QgH+/FFe6RSW/inlkJXdTsvqEjEp3BGAIwzfkl9gPIlSbrFQcr/8X8qVS VQdftgDD8K3JjccAdgDaA1qpyAlqzjS7O54MXbrO08rvN+2ELpO4MorvN4pLIZXkFXAtroXc8kF5W 03+wlpOdTXcvCY5MSwdDtV/gc7IhiNS5y8MpXiF2XZfYVDoG9hKdF6hSxNDCXPciiNpireGHZFJf5 CRQqLwtKKsGSXeH39dVvLYpbvQ7Pb80xEYWEtxLAeSS8EoLAoLIZXpAquzw9YN0iZDnJvtKpcumMU FoVlYmF5erU5rckIieGLUuSybOLoahwBdJ0QvVNOHF2RXnMIRCDhJYb7+rFveRzkDWKdWwFhPcTwS i0ce0P2Pb1T74l7hjs3mZ+bzDU2V34GsiOCkZi42SSlPIopJv6MsOp2sw5iS2jZfwu2kjkHyuFnPq gq0ipZxcuJquGBe3FCHIMsCksi61ReVe8tN4tS/oJn2qouQZbCWSEWpqGIvOSUWh5PYRmBr6Rw4IS K+y8rLJkNsKwuKTdqon6QVSPCkhZkZSksCutBhOXG0Sieu4gOFXqzdPRgnbtMhC+zAZa3zxiTy40X ZE29Nzg3WWu5stggq3CJzwoprFVx0LMJ7FO6/+TBcvfS6zG5u90e4RmrD5u2lnw74fHB2Q5bs3JBl tiMPsPuWhTW6lDryRhCpxTbYat/5fNBRSJ8Ty+fGdXv67HtM3YnPJvREUDnHgSC2LgdznZYK785df K8ALmIzQoVaZpGYRmQgwN48jVldnW03GoM+Q2wGF7V+XpY+m2xoMWNjgC69qNzL1xeNK56WffE+Xj trNARENcqi1khhVWRfWEF4qx9/djRp3xsX7b1FVEz7p7f9aWyP7E50L5lMeZqWVJRnJvMTwgJssRm hQyyKKxq91U59ayOnvLhlfzLjuFV3Wn45T+svkGLGxu3o2s/nO2LAVdhUMBZ9vaJ6webjvJcUFhVE zppzw07evBsoszDQfn5oKjWV0QpvMF5187aJ8e2mCp6fJj5cLD2hCyxQVYpzyCLwqqWC+BJ8dfHgR N4PlneVpBdbtfSuxvMFWQ9/FWhXycLPJuxaTvybwgIiMRmhcOMrymsqrdWEXNKO3vxW29VzAQXkFn A4uz2taJb3MjbnShdPF2q3rwBgCMgrudMMc3EkMKqyq9Gaqzd6ezFgRN47iKeTWBLqGpIn5PVAIut r9Y03J7f8llRn2jfgsyrAjQndgX7cAQlU79Hh90aaiWGX1Po+mB4peuscN+XcfXfhW9vc6ApNVi8n HDsrXob8/YBx0X8HqU8Popgt3njLEZY9ULOEla2vlpzdoh+3NG+BXf+5vkac7KsbtHvLrzyipmXQ1 NY9UJOhMXWVzrMChst8DivZl+P1NhOwpPfd8OmTQwprHrZSnIDLLa+0k5WKJIWN+bOvVJjdaFHfOx cTGPIpJcEhaX58Iq20m1WuJAYZv/6C9WeGDoCcIqf8Zc5Y85ZDhSWtoXF1ldaygrnOg6K/oJZsMl/ 93b0N6oVs6TNBx4+acKppBSW+sh5KTxbX2nq27Lvt6V8yoI2x3s3v/VYRWdJXsAwdNxszqKw1Oe6j EuKiwc1RfWpdpWxObCheejOqc+Ud5YzKLpvslmdRWGpj+SGyGx9pTW8wXmLXbKz2psuZP/84fLOkj NtZei4eWrwFJb6+aDkhsgMr7TH/MZfkf5ls8A+lypfgJfwrHApI6cpLKIEkuf4sfWVNr8wO4/J+fh METZ/UOEIa+FqobCIAkiuLzC80iZSy1gLzLZVmMEgp4y18HEKi8hFcgMstr7SLDLKWAAaHzqiinQo LKIAkqdfsfWVhpl37ZL2wdIMmrY+UvHHcrJCmSUwCovIEhZnt2v5O7NLYvBbmsY6f1CVKKmVERaRS SmH7KAkW7H1ldazQokVgrsVKu4yIyyH3zzxOIXF8IrUSVhzLVXvQ1a3xId9pilgUVjaE5azh7MZtE 5rQGLdfWOtpYh2SZE1hUWUEJakhshc6qwHJNTd52Zh3fpojY2k3atMU3GnsDQWXrH1lU5o2Pak2I/ MTFYtYMmJlVoZYZE1ERZtpRdhbdgt9iPFu6jR3x2Q8rDFTBV3CktjwmI+qBdaRZulRsVdcoRlpgIW haUOYwkpDbB8/ZwsqhvEL9Bp3LhfaMREYVFYOgivuHhQV8w3rBP3gY0HBG0m9kGhmSruFJY6SGiAx dZXuhOWR0QZa6oA+x5hZhEbMbUywiJykNYAi9Ur3eHsEr7tdBHrAsLM0iSmLGCyijuFpQLXY6I/4v CzN4P+vjm+A8I3LlnaG5uFmUVUiidtoimFRR6QTYj+CKtXekTUAp32/ar8DuZbFEFhKYqEBlhsfaV T1onIxRoDvao4yGSPCCkspZHwfNAX5mwGXeISmo6JqLiLhcIispDQEJnldp0ieO7ovRnBFfdFDQl7 EbTVZcInyxSWcpRyovNBtr4yAffu2YVW3BdNJGxj84VXFJYG8kGiW+Zc+wRt5hHZ2kHgzAZTtiGis NZOWGx9pXcsVkFbiW2pLDB0YoRFZOWDYhtgsXqld5rW19yk4osI5UNhkfqFV2x9pX/mW2svVJ6erP riiTK3vThy/117S1NW3AFYedkpg9g3PNNWumXuys9LP33JmnvH0jBbc+N79+zOjoCg6+dKVESQbsr wisJauwiL+aAeQ6r/+9HsT79mnUrbADQIs1vNinshifcHRL9gyazVTwpLIVuJaoDF1le6Yzp374fP NV3/kdgvjPVTVZspD0cwfFLK78MIi1QL1zNxZBMS3zu/Gi4e1BfjqdLfH26avS32c6UZWDZW6EJTy uFin8Q3V1JYpKKqPhhQzFMLsPWVzsLn5Pz3HrHOz0gJyypV3Es5nAtJv67MWnGnsKrcHHMYCkt8VV d1WL3SVWw1/71HGiTZCsBMES1lhTUUlnUXNGt4RWFVtpWcG2AV2PpKR0znZv/h0xaptgIw21ZuVeB wRO6N0MTzjTkPq462AqtXemLuP37fMj0i69v10JGV/1RISqyyL7vtmbekQGGtQj1bsfWVrpLBxvde k3Xjm0HT1kdW/uv7ShQETJwSUlirwnWVbAW2vtITsz9+UW6kPr2q4j6WkP5YkMKisMokg+moivtnu V0/4ZXl+k9k7qN4d9UqQgnt0lbj6TXzmaGwlpCOSnkBqtDwiq2vdMP85e/L30mZVz1Le2ElwysKqz yK3ACr5INEL8Ia+icF9rLx4Mp8UJHbYSuFRSD1fYJC74psfaWrb0X+Xbnh1SysWx9d9k8SXqfECIv CqshYQsWds3qlI8ZT8vcxM7mqgFVIUlgUlnKUcmrtma2v9EVeAWEV78Kxd3lMfU+JC8zcFXcKqy7Q VuajTMWd4ZUSNMzPz/PyIoQwwiKEEAqLEEJhEUIIhUUIIRQWIYTCIoQQCosQQigsQgiFRQghFBYhh FBYhBAKixBCKCxCCKGwCCGG4/8BAjn5LoppTCkAAAAASUVORK5CYII=';

}

})();


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

- 1. Open the add-in only manifest file in the add-in project. This file defines the add-in's settings and capabilities.
- 2. The ProviderName element has a placeholder value. Replace it with your name.
- 3. The DefaultValue attribute of the DisplayName element has a placeholder. Replace it with **My Office Add-in**.
- 4. The DefaultValue attribute of the Description element has a placeholder. Replace it with **A task pane add-in for PowerPoint**.
- 5. Save the file.

XML


```
...
<ProviderName>John Doe</ProviderName>
<DefaultLocale>en-US</DefaultLocale>
<!-- The display name of your add-in. Used on the store and various places of
the Office UI such as the add-ins dialog. -->
<DisplayName DefaultValue="My Office Add-in" />
<Description DefaultValue="A task pane add-in for PowerPoint"/>
...
```
# **Try it out**

- 1. Using Visual Studio, test the newly created PowerPoint add-in by pressing F5 or choosing the **Start** button to launch PowerPoint with the **Show Taskpane** add-in button displayed on the ribbon. The add-in will be hosted locally on IIS.
- 2. In PowerPoint, insert a new blank slide, choose the **Home** tab, and then choose the **Show Taskpane** button on the ribbon to open the add-in task pane.

- 3. In the task pane, choose the **Insert Image** button to add an image to the selected slide.


- 4. In the task pane, choose the **Insert Text** button to add text to the selected slide.

| AutoSave @ of                                  | F                                        |                |             |               | Presentation1 - PowerPoint |        |            |         |                                                                                                                                   |                                         | 图       |                                    | ■       | × |
|------------------------------------------------|------------------------------------------|----------------|-------------|---------------|----------------------------|--------|------------|---------|-----------------------------------------------------------------------------------------------------------------------------------|-----------------------------------------|---------|------------------------------------|---------|---|
| File<br>Home                                   | Insert                                   | Design<br>Draw | Transitions | Animations    | Shide Show                 | Review | View       | Help    | Storyboarding                                                                                                                     | Script Lab                              | Tell me |                                    | 120     | 0 |
| ರ್ಕೆ<br>Paste<br>New<br>Slide -<br>Clipboard G | Layout .<br>Reset<br>Section .<br>Slides |                | Font        |               | Paragraph                  |        | Protection |         | Shapes Arrange<br>SIVIES<br>Drawing                                                                                               | Firsd<br>Replace<br>Select .<br>Editing |         | Show<br>Taskpane<br>Commands Group |         | A |
| r                                              |                                          |                |             |               |                            |        |            |         | My Office Add-in                                                                                                                  |                                         |         |                                    |         | × |
|                                                |                                          |                |             | Hello Warld I |                            |        |            |         | Welcome<br>Select a slide and then choose the buttons<br>below to add content to it.<br>Try it out<br>Insert Image<br>Insert Text |                                         |         |                                    |         |   |
| Slide 1 of 1 []3                               |                                          |                |             |               |                            |        |            | = Notes |                                                                                                                                   |                                         |         |                                    | + 38% + |   |

#### 7 **Note**

To see the console.log output, you'll need a separate set of developer tools for a JavaScript console. To learn more about F12 tools and the Microsoft Edge DevTools, visit 


**Debug add-ins using developer tools for Internet Explorer**, **Debug add-ins using developer tools for Edge Legacy**, or **Debug add-ins using developer tools in Microsoft Edge (Chromium-based)**.

### **Next steps**

Congratulations, you've successfully created a PowerPoint task pane add-in! Next, learn more about the capabilities of a PowerPoint add-in and build a more complex add-in by following along with the PowerPoint add-in tutorial.

### **Troubleshooting**

- Ensure your environment is ready for Office development by following the instructions in Set up your development environment.
- Some of the sample code uses ES6 JavaScript. This isn't compatible with older versions of Office that use the Trident (Internet Explorer 11) browser engine. For information on how to support those platforms in your add-in, see Support older Microsoft webviews and Office versions. If you don't already have a Microsoft 365 subscription to use for development, you might qualify for a Microsoft 365 E5 developer subscription through the [Microsoft 365 Developer Program](https://aka.ms/m365devprogram) ; for details, see the [FAQ.](https://learn.microsoft.com/en-us/office/developer-program/microsoft-365-developer-program-faq#who-qualifies-for-a-microsoft-365-e5-developer-subscription-) Alternatively, you can [sign up for a 1-month free trial](https://www.microsoft.com/microsoft-365/try) or [purchase a Microsoft 365 plan](https://www.microsoft.com/microsoft-365/business/compare-all-microsoft-365-business-products-g) .
- If your add-in shows an error (for example, "This add-in could not be started. Close this dialog to ignore the problem or click "Restart" to try again.") when you press F5 or choose **Debug** > **Start Debugging** in Visual Studio, see Debug Office Add-ins in Visual Studio for other debugging options.

# **Code samples**

- [PowerPoint "Hello world" add-in](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/hello-world/powerpoint-hello-world) : Learn how to build a simple Office Add-in with only a manifest, HTML web page, and a logo.
# **See also**

- Office Add-ins platform overview
- Develop Office Add-ins
- Publish your add-in using Visual Studio


# **Build your first PowerPoint content addin**

Article • 08/27/2024

In this article, you'll walk through the process of building a PowerPoint content add-in using Visual Studio.

### **Prerequisites**

- [Visual Studio 2019 or later](https://www.visualstudio.com/vs/) with the **Office/SharePoint development** workload installed.
7 **Note**

If you've previously installed Visual Studio, use the Visual Studio Installer to ensure that the **Office/SharePoint development** workload is installed.

- Office connected to a Microsoft 365 subscription (including Office on the web).
### **Create the add-in project**

- 1. In Visual Studio, choose **Create a new project**.
- 2. Using the search box, enter **add-in**. Choose **PowerPoint Web Add-in**, then select **Next**.
- 3. Name your project and select **Create**.
- 4. In the **Create Office Add-in** dialog window, choose **Insert content into PowerPoint slides**, and then choose **Finish** to create the project.
- 5. Visual Studio creates a solution and its two projects appear in **Solution Explorer**. The **Home.html** file opens in Visual Studio.

## **Explore the Visual Studio solution**

When you've completed the wizard, Visual Studio creates a solution that contains two projects.


| Project                       | Description                                                                                                                                                                                                                                                                                                                                                                                                                                            |
|-------------------------------|--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| Add-in<br>project             | Contains only an XML-formatted add-in only manifest file, which contains all the<br>settings that describe your add-in. These settings help the Office application<br>determine when your add-in should be activated and where the add-in should<br>appear. Visual Studio generates the contents of this file for you so that you can<br>run the project and use your add-in immediately. Change these settings any time<br>by modifying the XML file. |
| Web<br>application<br>project | Contains the content pages of your add-in, including all the files and file<br>references that you need to develop Office-aware HTML and JavaScript pages.<br>While you develop your add-in, Visual Studio hosts the web application on your<br>local IIS server. When you're ready to publish the add-in, you'll need to deploy<br>this web application project to a web server.                                                                      |

### **Update the code**

- 1. **Home.html** specifies the HTML that will be rendered in the add-in's task pane. In **Home.html**, find the <p> element that contains the text "This example will read the current document selection." and the <button> element where the id is "get-datafrom-selection". Replace these entire elements with the following markup then save the file.

```
HTML
<p class="ms-font-m-plus">This example will get some details about the
current slide.</p>
<button class="Button Button--primary" id="get-data-from-selection">
 <span class="Button-icon"><i class="ms-Icon ms-Icon--plus"></i>
</span>
 <span class="Button-label">Get slide details</span>
 <span class="Button-description">Gets and displays the current
slide's details.</span>
</button>
```
- 2. Open the file **Home.js** in the root of the web application project. This file specifies the script for the add-in. Find the getDataFromSelection function and replace the entire function with the following code then save the file.
JavaScript // Gets some details about the current slide and displays them in a notification.


```
function getDataFromSelection() {
 if (Office.context.document.getSelectedDataAsync) {

Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideR
ange,
 function (result) {
 if (result.status ===
Office.AsyncResultStatus.Succeeded) {
 showNotification('Some slide details are:', '"' + 
JSON.stringify(result.value) + '"');
 } else {
 showNotification('Error:', result.error.message);
 }
 }
 );
 } else {
 app.showNotification('Error:', 'Reading selection data is not
supported by this host application.');
 }
}
```
### **Update the manifest**

- 1. Open the add-in only manifest file in the add-in project. This file defines the addin's settings and capabilities.
- 2. The ProviderName element has a placeholder value. Replace it with your name.
- 3. The DefaultValue attribute of the DisplayName element has a placeholder. Replace it with **My Office Add-in**.
- 4. The DefaultValue attribute of the Description element has a placeholder. Replace it with **A content add-in for PowerPoint.**.
- 5. Save the file. The updated lines should look like the following code sample.

```
XML
...
<ProviderName>John Doe</ProviderName>
<DefaultLocale>en-US</DefaultLocale>
<!-- The display name of your add-in. Used on the store and various
places of the Office UI such as the add-ins dialog. -->
<DisplayName DefaultValue="My Office Add-in" />
<Description DefaultValue="A content add-in for PowerPoint."/>
...
```


# **Try it out**

- 1. Using Visual Studio, test the newly created PowerPoint add-in by pressing F5 or choosing the **Start** button to launch PowerPoint with the content add-in displayed over the slide.
- 2. In PowerPoint, choose the **Get slide details** button in the content add-in to get details about the current slide.

| Welcome<br>This example will get some details about the current |
|-----------------------------------------------------------------|
| slide.<br>Get slide details                                     |
| Find more samples online                                        |
|                                                                 |
| `lick tc                                                        |
|                                                                 |

#### 7 **Note**

To see the console.log output, you'll need a separate set of developer tools for a JavaScript console. To learn more about F12 tools and the Microsoft Edge DevTools, visit **Debug add-ins using developer tools for Internet Explorer**, **Debug add-ins using developer tools for Edge Legacy**, or **Debug add-ins using developer tools in Microsoft Edge (Chromium-based)**.

# **Next steps**

Congratulations, you've successfully created a PowerPoint content add-in! Next, learn more about developing Office Add-ins with Visual Studio.

# **Troubleshooting**


- Ensure your environment is ready for Office development by following the instructions in Set up your development environment.
- Some of the sample code uses ES6 JavaScript. This isn't compatible with older versions of Office that use the Trident (Internet Explorer 11) browser engine. For information on how to support those platforms in your add-in, see Support older Microsoft webviews and Office versions. If you don't already have a Microsoft 365 subscription to use for development, you might qualify for a Microsoft 365 E5 developer subscription through the [Microsoft 365 Developer Program](https://aka.ms/m365devprogram) ; for details, see the [FAQ](https://learn.microsoft.com/en-us/office/developer-program/microsoft-365-developer-program-faq#who-qualifies-for-a-microsoft-365-e5-developer-subscription-). Alternatively, you can [sign up for a 1-month free trial](https://www.microsoft.com/microsoft-365/try) or [purchase a Microsoft 365 plan](https://www.microsoft.com/microsoft-365/business/compare-all-microsoft-365-business-products-g) .
- If your add-in shows an error (for example, "This add-in could not be started. Close this dialog to ignore the problem or click "Restart" to try again.") when you press F5 or choose **Debug** > **Start Debugging** in Visual Studio, see Debug Office Addins in Visual Studio for other debugging options.

### **See also**

- Office Add-ins platform overview
- Develop Office Add-ins
- Using Visual Studio Code to publish