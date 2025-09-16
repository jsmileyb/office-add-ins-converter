
# **Design the UI of Office Add-ins**

Article • 04/04/2023

Office Add-ins extend the Office experience by providing contextual functionality that users can access within Office clients. Add-ins empower users to get more done by enabling them to access external functionality within Office, without costly context switches.

Your add-in UI design must integrate seamlessly with Office to provide an efficient, natural interaction for your users. Take advantage of add-in commands to provide access to your add-in and apply the best practices that we recommend when you create a custom HTML-based UI.

## **Office design principles**

Office applications follow a general set of interaction guidelines. The applications share content and have elements that look and behave similarly. This commonality is built on a set of design principles. The principles help the Office team create interfaces that support customers' tasks. Understanding and following them will help you support your customers' goals inside of Office.

Follow the Office design principles to create positive add-in experiences.

- **Design explicitly for Office.** The functionality, as well as the look and feel, of an add-in must harmoniously complement the Office experience. Add-ins should feel native. They should fit seamlessly into Word on an iPad or PowerPoint on the web. A well-designed add-in will be an appropriate blend of your experience, the platform, and the Office application. Apply document and UI theming where appropriate. Consider using Fluent UI for the web as your design language and tool set. The Fluent UI for the web has two flavors.
	- **For non-React UIs:** Use **Fabric Core**, an open-source collection of CSS classes and Sass mixins that give you access to colors, animations, fonts, icons, and grids. (It's called "Fabric Core" instead of "Fluent Core" for historical reasons.) To get started, see Fabric Core in Office Add-ins.

#### 7 **Note**

While Fabric Core is the recommended library to design non-React add-ins, the team is working on **[Fluent UI Web Components](https://learn.microsoft.com/en-us/fluent-ui/web-components/)** to provide a newer solution. Built on **[FAST](https://www.fast.design/)** , the Fluent UI Web Components library allows you to 


use, customize, and build Web Components to create a more modern and [standards-based UI. We invite you to test this library by](https://learn.microsoft.com/en-us/fluent-ui/web-components/getting-started/quick-start) **completing the quick start** and welcome feedback on your experience through **[GitHub](https://github.com/microsoft/fluentui/issues?q=is%3Aopen+is%3Aissue+label%3Aweb-components)** .

- **For React UIs:** use **Fluent UI React**, a React front-end framework designed to build experiences that fit seamlessly into a broad range of Microsoft products. It provides robust, up-to-date, accessible React-based components which are highly customizable using CSS-in-JS. To get started, see Fluent UI React in Office Add-ins.
- **Favor content over chrome.** Allow customers' page, slide, or spreadsheet to remain the focus of the experience. An add-in is an auxiliary interface. No accessory chrome should interfere with the add-in's content and functionality. Brand your experience wisely. We know it's important to provide users with a unique, recognizable experience but avoid distraction. Strive to keep the focus on content and task completion, not brand attention.
- **Make it enjoyable and keep users in control.** People enjoy using products that are both functional and visually appealing. Craft your experience carefully. Get the details right by considering every interaction and visual detail. Allow users to control their experience. The necessary steps to complete a task must be clear and relevant. Important decisions should be easy to understand. Actions should be easily reversible. An add-in is not a destination – it's an enhancement to Office functionality.
- **Design for all platforms and input methods**. Add-ins are designed to work on all the platforms that Office supports, and your add-in UX should be optimized to work across platforms and form factors. Support mouse/keyboard and touch input devices, and ensure that your custom HTML UI is responsive to adapt to different form factors. For more information, see Touch.

### **See also**

- Add-in development best practices


# **Use Fluent UI React in Office Add-ins**

Article • 02/12/2025

[Fluent UI React](https://react.fluentui.dev/) is the official open-source JavaScript front-end framework designed to build experiences that fit seamlessly into a broad range of Microsoft products, including Microsoft 365 applications. It provides robust, up-to-date, accessible React-based components which are highly customizable using CSS-in-JS.

#### 7 **Note**

This article describes the use of Fluent UI React in the context of Office Add-ins. However, it's also used in a wide range of Microsoft 365 apps and extensions. For more information, see **[Fluent UI React](https://react.fluentui.dev/)** and the **[Fluent UI Web](https://github.com/microsoft/fluentui)** open source repository.

This article describes how to create an add-in that's built with React and that uses Fluent UI React components.

### **Create an add-in project**

You'll use the Yeoman generator for Office Add-ins to create an add-in project that uses React.

#### 7 **Note**

The React-based Add-ins created with the generator use Fluent UI React V9. This version doesn't support the Trident (IE) webview. If your add-in's users have versions of Office that require Trident, use one of the samples in **[Office-Add-ins-](https://github.com/OfficeDev/Office-Add-ins-Fluent-React-version-8)[Fluent-React-version-8](https://github.com/OfficeDev/Office-Add-ins-Fluent-React-version-8)** instead of this article. For information about which versions of Office use Trident, see **Browsers and webview controls used by Office Add-ins**.

### **Install the prerequisites**

- Node.js (the latest LTS version). Visit the [Node.js site](https://nodejs.org/) to download and install the right version for your operating system.
- The latest version of Yeoman and the Yeoman generator for Office Add-ins. To install these tools globally, run the following command via the command prompt.


```
command line
```
npm install -g yo generator-office

#### 7 **Note**

Even if you've previously installed the Yeoman generator, we recommend you update your package to the latest version from npm.

- Office connected to a Microsoft 365 subscription (including Office on the web).
#### 7 **Note**

If you don't already have Office, you might qualify for a Microsoft 365 E5 developer subscription through the **[Microsoft 365 Developer Program](https://aka.ms/m365devprogram)** ; for details, see the **[FAQ](https://learn.microsoft.com/en-us/office/developer-program/microsoft-365-developer-program-faq#who-qualifies-for-a-microsoft-365-e5-developer-subscription-)**. Alternatively, you can **[sign up for a 1-month free trial](https://www.microsoft.com/microsoft-365/try)** or **[purchase a Microsoft 365 plan](https://www.microsoft.com/microsoft-365/business/compare-all-microsoft-365-business-products-g)** .

### **Create the project**

Run the following command to create an add-in project using the Yeoman generator. A folder that contains the project will be added to the current directory.

command line

yo office

#### 7 **Note**

When you run the yo office command, you may receive prompts about the data collection policies of Yeoman and the Office Add-in CLI tools. Use the information that's provided to respond to the prompts as you see fit.

When prompted, provide the following information to create your add-in project.

- **Choose a project type:** Specify Office Add-in Task Pane project using React framework .
- **Choose a script type:** Specify either TypeScript or JavaScript .
- **What do you want to name your add-in?** Specify My Office Add-in .


- **Which Office client application would you like to support?** Specify one of the hosts. (The screenshots in this article use Word . Running the project for the first time is easier if you select Excel , PowerPoint , or Word . See Try it out.)
The following is an example.

After you complete the wizard, the generator creates the project and installs supporting Node components.

#### 7 **Note**

Fluent UI React v9 or later isn't supported with the Trident (IE) or EdgeHTML (Edge Legacy) webview controls. If your version of Office is using either of those, the task pane of the add-in generated by Yo Office simply contains a message to upgrade your version of Office. For more information, see **Browsers and webview controls used by Office Add-ins**.

### **Explore the project**

The add-in project that you've created with the Yeoman generator contains sample code for a basic task pane add-in. If you'd like to explore the components of your add-in project, open the project in your code editor and review the following files. The file name extensions depend on which language you choose. TypeScript extensions are in parentheses. When you're ready to try out your add-in, proceed to the next section.

- The **./manifest.xml** or **./manifest.json** file in the root directory of the project defines the settings and capabilities of the add-in. To learn more about the **manifest.xml** file, see Office Add-ins with the add-in only manifest. To learn more


about the **manifest.json** file, see Office Add-ins with the unified app manifest for Microsoft 365.

#### 7 **Note**

The **unified manifest for Microsoft 365** can be used in production Outlook add-ins. It's available only as a preview for Excel, PowerPoint, and Word addins.

- The **./src/taskpane/taskpane.html** file contains the HTML markup for the task pane and loads the Office JavaScript Library. It also tests whether the webview control supports Fluent UI React v9 and displays a special message if it doesn't.
- The **./src/taskpane/index.jsx (tsx)** file is the React root component. It loads React and Fluent UI React, ensures that the Office JavaScript library has been loaded, and applies the Fluent-defined theme.
- The **./src/taskpane/office-document.js (ts)** file contains the Office JavaScript API code that facilitates interaction between the task pane and the Office client application.
- The **./src/taskpane/components/** folder contains the React component *.jss (tsx) files that create the UI.

### **Try it out**

- 1. Navigate to the root folder of the project.
command line cd "My Office Add-in"

- 2. Complete the following steps to start the local web server and sideload your addin.
#### 7 **Note**

- Office Add-ins should use HTTPS, not HTTP, even while you're developing. If you're prompted to install a certificate after you run one of the following commands, accept the prompt to install the certificate that the Yeoman generator provides. You may also have to run your


command prompt or terminal as an administrator for the changes to be made.

- If this is your first time developing an Office Add-in on your machine, you may be prompted in the command line to grant Microsoft Edge WebView a loopback exemption ("Allow localhost loopback for Microsoft Edge WebView?"). When prompted, enter Y to allow the exemption. Note that you'll need administrator privileges to allow the exemption. Once allowed, you shouldn't be prompted for an exemption when you sideload Office Add-ins in the future (unless you remove the exemption from your machine). To learn more, see **["We can't open this](https://learn.microsoft.com/en-us/office/troubleshoot/office-suite-issues/cannot-open-add-in-from-localhost) [add-in from localhost" when loading an Office Add-in or using Fiddler](https://learn.microsoft.com/en-us/office/troubleshoot/office-suite-issues/cannot-open-add-in-from-localhost)**.
#### **Tip**

If you're testing your add-in on Mac, run the following command before proceeding. When you run this command, the local web server starts.

command line

npm run dev-server

- To test your add-in, run the following command in the root directory of your project. This starts the local web server and opens the Office host application with your add-in loaded.
command line npm start 7 **Note**


If you're testing your add-in in Outlook, npm start sideloads the add-in to both the Outlook desktop and web clients. For more information on how to sideload add-ins in Outlook, see **Sideload Outlook add-ins for testing**.

- To test your add-in in Excel, Word, or PowerPoint on the web, run the following command in the root directory of your project. When you run this command, the local web server starts. Replace "{url}" with the URL of a Word document on your OneDrive or a SharePoint library to which you have permissions.
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

#### 7 **Note**

If this is the first time that you have sideloaded an Office add-in on your computer (or the first time in over a month), you're prompted first to delete an old certificate and then to install a new one. Agree to both prompts.


- 3. A **WebView Stop On Load** prompt appears. Select **OK**.
- 4. If the "My Office Add-in" task pane isn't already open, choose the **Home** tab, and then choose the **Show Taskpane** button on the ribbon to open the add-in task pane.

#### 7 **Note**

If you're testing your add-in in Outlook, create a new message. Then, navigate to the **Message** tab and choose **Show Taskpane** from the ribbon to open the add-in task pane.

- 5. Enter text into the text box and then select **Insert text**.
- 6. When you're ready to stop the dev server and uninstall the add-in, run the following command.

| command line |  |  |  |  |  |  |
|--------------|--|--|--|--|--|--|
| npm stop     |  |  |  |  |  |  |


### **Migrate to Fluent UI React v9**

If you have an existing add-in that implements an older version of Fluent UI React, we recommend migrating to Fluent UI v9. For guidance on the migration process, see [Getting started migrating to v9](https://react.fluentui.dev/?path=/docs/concepts-migration-getting-started--page) .

### **Troubleshooting**

- Ensure your environment is ready for Office development by following the instructions in Set up your development environment.
- Some of the sample code uses ES6 JavaScript. This isn't compatible with older versions of Office that use the Trident (Internet Explorer 11) browser engine. For information on how to support those platforms in your add-in, see Support older Microsoft webviews and Office versions. If you don't already have a Microsoft 365 subscription to use for development, you might qualify for a Microsoft 365 E5 developer subscription through the [Microsoft 365 Developer Program](https://aka.ms/m365devprogram) ; for details, see the [FAQ](https://learn.microsoft.com/en-us/office/developer-program/microsoft-365-developer-program-faq#who-qualifies-for-a-microsoft-365-e5-developer-subscription-). Alternatively, you can [sign up for a 1-month free trial](https://www.microsoft.com/microsoft-365/try) or [purchase a Microsoft 365 plan](https://www.microsoft.com/microsoft-365/business/compare-all-microsoft-365-business-products-g) .
- The automatic npm install step Yo Office performs may fail. If you see errors when trying to run npm start , navigate to the newly created project folder in a command prompt and manually run npm install . For more information about Yo Office, see Create Office Add-in projects using the Yeoman Generator.

### **See also**

- UX design patterns for Office Add-ins
- [Fluent UI React](https://react.fluentui.dev/)
- [Fluent UI GitHub repository](https://github.com/microsoft/fluentui)


# **Fabric Core in Office Add-ins**

Article • 04/04/2023

Fabric Core is an open-source collection of CSS classes and Sass mixins that's *intended for use in non-React* Office Add-ins. Fabric Core contains basic elements of the Fluent UI design language such as icons, colors, typefaces, and grids. Fabric Core is framework independent, so it can be used with any single-page application or any server-side web UI framework. (It's called "Fabric Core" instead of "Fluent Core" for historical reasons.)

If your add-in's UI isn't React-based, you can also make use of a set of non-React components. See Use Office UI Fabric JS components.

#### 7 **Note**

This article describes the use of Fabric Core in the context of Office Add-ins, but it's also used in a wide range of Microsoft 365 apps and extensions. For more information, see **[Fabric Core](https://developer.microsoft.com/fluentui#/get-started/web#fabric-core)** and the open source repo **[Office UI Fabric Core](https://github.com/OfficeDev/office-ui-fabric-core)** .

#### 7 **Note**

While Fabric Core is the recommended library to design non-React add-ins, the team is working on **[Fluent UI Web Components](https://learn.microsoft.com/en-us/fluent-ui/web-components/)** to provide a newer solution. Built on **[FAST](https://www.fast.design/)** , the Fluent UI Web Components library allows you to use, customize, and build Web Components to create a more modern and standards-based UI. We invite you to test this library by **[completing the quick start](https://learn.microsoft.com/en-us/fluent-ui/web-components/getting-started/quick-start)** and welcome feedback on your experience through **[GitHub](https://github.com/microsoft/fluentui/issues?q=is%3Aopen+is%3Aissue+label%3Aweb-components)** .

### **Use Fabric Core: icons, fonts, colors**

- 1. Add the content delivery network (CDN) reference to the HTML on your page.

```
HTML
<link rel="stylesheet" href="https://res-1.cdn.office.net/files/fabric-
cdn-prod_20230815.002/office-ui-fabric-core/11.0.0/css/fabric.min.css">
```
- 2. Use Fabric Core icons and fonts.


To use a Fabric Core icon, include the "i" element on your page, and then reference the appropriate classes. You can control the size of the icon by changing the font size. For example, the following code shows how to make an extra-large table icon that uses the themePrimary (#0078d7) color.

```
HTML
<i class="ms-Icon ms-font-xl ms-Icon--Table ms-fontColor-themePrimary">
</i>
```
For more detailed instructions, see [Fluent UI Icons.](https://developer.microsoft.com/fluentui#/styles/web/icons) To find more icons that are available in Fabric Core, use the search feature on that page. When you find an icon to use in your add-in, be sure to prefix the icon name with ms-Icon-- .

For information about font sizes and colors that are available in Fabric Core, see [Typography](https://developer.microsoft.com/fluentui#/styles/web/typography) and the **Colors** table of contents at [Colors.](https://developer.microsoft.com/fluentui#/styles/web/colors)

Examples are included in the Samples later in this article.

### **Use Office UI Fabric JS components**

Add-ins with non-React UIs can also use any of the many components from [Office UI](https://github.com/OfficeDev/office-ui-fabric-js) [Fabric JS](https://github.com/OfficeDev/office-ui-fabric-js) , including buttons, dialogs, pickers, and more. See the readme of the repo for instructions.

Examples are included in the Samples later in this article.

### **Samples**

The following sample add-ins use Fabric Core and/or Office UI Fabric JS components. Some of these repos are archived, meaning that they are no longer being updated with bug or security fixes, but you can still use them to learn how to use Fabric Core and Fabric UI components.

- [Excel Add-in JavaScript SalesTracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker)
- [Excel Add-in SalesLeads](https://github.com/OfficeDev/Excel-Add-in-SalesLeads)
- [Excel Add-in WoodGrove Expense Trends](https://github.com/OfficeDev/Excel-Add-in-WoodGrove-Expense-Trends)
- [Excel Content Add-in Humongous Insurance](https://github.com/OfficeDev/Excel-Content-Add-in-Humongous-Insurance)
- [Office Add-in Fabric UI Sample](https://github.com/OfficeDev/Office-Add-in-Fabric-UI-Sample)
- [Office-Add-in-UX-Design-Patterns-Code](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)
- [Outlook Add-in GifMe](https://github.com/OfficeDev/Outlook-Add-in-GifMe)
- [PowerPoint Add-in Microsoft Graph ASPNET InsertChart](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)


- [Word Add-in Angular2 StyleChecker](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker)
- [Word Add-in JS Redact](https://github.com/OfficeDev/Word-Add-in-JS-Redact)
- [Word Add-in MarkdownConversion](https://github.com/OfficeDev/Word-Add-in-MarkdownConversion)


# **Accessibility guidelines**

Article • 12/03/2024

As you design and develop your Office Add-ins, you'll want to ensure that all potential users and customers are able to use your add-in successfully. Engineering and implementing inclusive experiences provide better usability and customer satisfaction, as well as a larger market for your solutions. We recommend you become familiar with the Web Content Accessibility Guidelines (WCAG), international web standards that define what's needed for your add-in to be accessible.

- [Explore the WCAG standards and resources](https://learn.microsoft.com/en-us/compliance/regulatory/offering-wcag-2-1)
- [Explore the WCAG tutorials](https://www.w3.org/WAI/tutorials/)

Apply the following guidelines to ensure that your solution is accessible to all audiences.

## **Design for multiple input methods**

- Ensure that users can perform operations by using only the keyboard. Users should be able to move to all actionable elements on the page by using a combination of the Tab and arrow keys.
- On a mobile device, when users operate a control by touch, the device should provide useful audio feedback.
- Provide helpful labels for all interactive controls.
- [Explore more design and UI resources.](https://learn.microsoft.com/en-us/windows/apps/design/accessibility/accessibility)

### **Make your add-in easy to use**

- Don't rely on a single attribute, such as color, size, shape, location, orientation, or sound, to convey meaning in your UI.
- Avoid unexpected changes of context, such as moving the focus to a different UI element without user action.
- Provide a way to verify, confirm, or reverse all binding actions.
- Provide a way to pause or stop media, such as audio and video.
- Don't impose a time limit for user action.

## **Make your add-in easy to see**

- Avoid unexpected color changes.


- Provide meaningful and timely information to describe UI elements, titles and headings, inputs, and errors. Ensure that names of controls adequately describe the intent of the control.
- Verify you UI elements render correctly in the Windows high-contrast themes.
- Follow [standard guidelines](https://www.w3.org/TR/UNDERSTANDING-WCAG20/visual-audio-contrast-contrast.html) for color contrast.

### **Account for assistive technologies**

- Avoid using features that interfere with assistive technologies, including visual, audio, or other interactions.
- Don't provide text in an image format. Screen readers can't read text within images.
- Provide a way for users to adjust or mute all audio sources.
- Provide a way for users to turn on captions or audio description with audio sources.
- Provide alternatives to sound as a means to alert users, such as visual cues or vibrations.

## **Test your add-in**

- Always use accessibility verification and testing tools like [Accessibility Insights](https://accessibilityinsights.io/) on your add-in to catch and resolve issues before you ship.
- Verify the screen reading experience using [Windows Narrator](https://support.microsoft.com/windows/e4397a0d-ef4f-b386-d8ae-c172f109bdb1) , [JAWS](https://support.freedomscientific.com/Downloads/JAWS) , or [NVDA](https://www.nvaccess.org/download/) .
- Periodically run the tools to keep up with changes to the international accessibility guidelines. For more information, see [Accessibility testing.](https://learn.microsoft.com/en-us/windows/apps/design/accessibility/accessibility-testing)

## **See also**

- [Accessibility in the Store](https://learn.microsoft.com/en-us/windows/apps/design/accessibility/accessibility-in-the-store)
- [Web Content Accessibility Guidelines (WCAG) 2.2](https://www.w3.org/TR/WCAG22/)
- [Developing for Web Accessibility](https://www.w3.org/WAI/tips/developing/)
- [Accessibility Fundamentals Learning Path](https://learn.microsoft.com/en-us/training/paths/accessibility-fundamental/)
- [European Accessibility Act (EAA)](https://www.deque.com/blog/european-accessibility-act-eaa-top-20-key-questions-answered/)

# **Data visualization style guidelines for Office Add-ins**

Article • 06/29/2023

Good data visualizations help users find insights in their data. They can use those insights to tell stories that inform and persuade. This article provides guidelines to help you design effective data visualizations in your add-ins for Excel and other Office apps.

We recommend that you use Fluent UI to create the chrome for your data visualizations. Fluent UI includes styles and components that integrate seamlessly with the Office look and feel.

## **Data visualization elements**

Data visualizations share a general framework and common visual and interactive elements, including titles, labels, and data plots, as shown in the following figure.

### **Chart titles**

Follow these guidelines for chart titles.

- Make your chart titles easily readable. Position them to create a clear visual hierarchy in relation to the rest of the chart.


- In general, use sentence capitalization (capitalize the first word). To create contrast or to reinforce hierarchies, you can use all caps, but all caps should be used sparingly.
- Incorporate the Fluent UI type ramp to make your charts consistent with the Office UI, which uses Segoe. You can also use a different typeface to differentiate chart content from the UI.
	- [Fluent UI React typography styles](https://react.fluentui.dev/?path=/docs/theme-typography--page)
	- [Fabric Core typography styles](https://developer.microsoft.com/fluentui#/styles/web/typography)
- Use sans-serif typefaces with large counters.

#### **Axis labels**

Make your axis labels dark enough to read clearly, with adequate contrast ratios between the text and background colors. Make sure that they are not so dark that they compete with data ink.

Light grays are most effective for axis labels. Explore the following Fluent UI neutral color palettes.

- [Fluent UI React color schemes](https://react.fluentui.dev/?path=/docs/theme-colors--page)
- [Fabric Core color schemes](https://developer.microsoft.com/fluentui#/styles/web/colors/neutrals)

#### **Data ink**

The pixels that represent the actual data in a chart are referred to as data ink. This should be the central focus of the visualization. Avoid the use of drop shadows, heavy outlines, or unnecessary design elements that distort or compete with the data. Use gradients only when data values are tied to color values. Avoid three-dimensional charts unless a measurable, objective value is bound to a third dimension.

#### **Color**

Choose colors that follow operating system or application themes rather than hardcoded colors. At the same time, make sure that the colors you apply don't distort the data. Misuse of color in data visualizations can result in data distortion and incorrect reading of information.

For best practices for use of color in data visualizations, see the following:

- [Why rainbow colors aren't the best option for data visualizations](https://www.poynter.org/2013/why-rainbow-colors-arent-always-the-best-options-for-data-visualizations/224413/)
- [Color Brewer 2.0: Color Advice for Cartography](https://colorbrewer2.org/)


- [I Want Hue](https://tools.medialab.sciences-po.fr/iwanthue/)
#### **Gridlines**

Gridlines are often necessary for accurately reading a chart, but should be presented as a secondary visual element, enhancing the data ink, not competing with it. Make static gridlines thin and light, unless they are designed specifically for high contrast. You can also use interaction to create dynamic, just-in-time gridlines that appear in context when a user interacts with a chart.

Light grays are most effective for gridlines. Explore the following Fluent UI neutral color palettes.

- [Fluent UI React color schemes](https://react.fluentui.dev/?path=/docs/theme-colors--page)
- [Fabric Core color schemes](https://developer.microsoft.com/fluentui#/styles/web/colors/neutrals)

The following image shows a data visualization with gridlines.

### **Legends**

Add legends if necessary to:

- Distinguish between series.
- Present scale or value changes.

Make sure that your legends enhance the data ink and don't compete with it. Place legends:

- Flush left above the plot area by default, if all legend items fit above the chart.


- On the upper right side of the plot area, if all legend items don't fit above the chart, and make it scrollable, if necessary.
To optimize for readability and accessibility, map legend markers to the relevant chart shape. For example, use circle legend markers for scatter plot and bubble chart legends. Use line segment legend markers for line charts.

#### **Data labels and tooltips**

Ensure that data labels and tooltips have adequate white space and type variation. Use algorithms to minimize occlusion and collision. For example, a tooltip might surface to the right of a data point by default, but surface to the left if right edges are detected.

## **Design principles**

The Office Design team created the following set of design principles, which we use when designing new data visualizations for the Office product suite.

#### **Visual design principles**

- Visualizations should honor and enhance the data, making it easy to understand. Highlight the data, adding supporting elements only as needed to provide context. Avoid unnecessary embellishments, such as drop shadows and outlines, chart junk, or data distortion.
- Visualizations should encourage exploration by providing rich visual feedback. Use well-established interaction patterns, interface controls, and clear system feedback.
- Embody time-honored design principles. Use established typographic and visual communication design principles to enhance form, readability, and meaning.

#### **Interaction design principles**

- Design to allow for exploration.
- Allow for direct interactions with objects that reveal new insights (sorting via drag, for example).
- Use simple, direct, familiar interaction models.

For more information about how to design user-friendly interactive data visualizations, see [UI Tenets and Traps](https://uitraps.com/) .

## **Motion design principles**


Motion follows stimulus. Visual elements should move in the same direction at the same rate. This applies to:

- Chart creation
- Transition from one chart type to another chart type
- Filtering
- Sorting
- Adding or subtracting data
- Brushing or slicing data
- Resizing a chart

Create a perception of causality. When staging animations:

- Stage one thing at a time.
- Stage changes to axes before changes to data ink.
- Stage and animate objects as a group if they are moving at the same speed in the same direction.
- Stage data elements in groups of no more than 4-5 objects. Viewers have difficulty tracking more than 4-5 objects independently.

Motion adds meaning.

- Animations increase user comprehension of changes to the data, provide context, and act as a non-verbal annotation layer.
- Motion should occur in a meaningful coordinate space of the visualization.
- Tailor the animation to the visual.
- Avoid gratuitous animations.

Motion follows data.

- Preserve data mappings. If an area is tied to a measure, maintain that area in transition.
- Maintain a consistent animation design language. Where possible, map data visualization animation to existing Office motion design language. Use similar animations for similar chart types.

## **Accessibility in data visualizations**

- Don't use color as the only way to communicate information. People who are color blind will not be able to interpret the results. Use shape, size and texture in addition to color when possible to communicate information.
- Make all interactive elements, such as push buttons or pick lists, accessible from a keyboard.


- Send accessibility events to screen readers to announce focus changes, tooltips, and so on.
## **See also**

- [The Five Best Libraries for Building Data Visualizations](https://www.fastcompany.com/3029760/the-five-best-libraries-for-building-data-vizualizations)
- [The Visual Display of Quantitative Information](https://www.edwardtufte.com/book/the-visual-display-of-quantitative-information/)

# **Office Add-in design language**

Article • 02/11/2025

The Office design language is a clean and simple visual system that ensures consistency across experiences. It contains a set of visual elements that define Office interfaces, including:

- A standard typeface
- A common color palette
- A set of typographic sizes and weights
- Icon guidelines
- Shared icon assets
- Animation definitions
- Common components

Fluent UI is the official front-end framework for building with the Office design language. Using Fluent UI is optional, but it's the fastest way to ensure that your add-ins feel like a natural extension of Office. Take advantage of Fluent UI to design and build add-ins that complement Office.

Many Office Add-ins are associated with a preexisting brand. You can retain a strong brand and its visual or component language in your add-in. Look for opportunities to retain your own visual language while integrating with Office. Consider ways to swap out Office colors, typography, icons, or other stylistic elements with elements of your own brand. Consider ways to follow common add-in layouts or UX design patterns while inserting controls and components that are familiar to your customers.

Inserting a heavily branded HTML-based UI inside of Office can create dissonance for customers. Find a balance that fits seamlessly in Office but also clearly aligns with your service or parent brand. When an add-in doesn't fit with Office, it's often because stylistic elements conflict. For example, typography is too large and off grid, colors are contrasting or particularly loud, or animations are superfluous and behave differently than Office. The appearance and behavior of controls or components veer too far from Office standards.

# **Review guidelines for visual elements**

Learn about each visual element that makes up an Office Add-in, including guidelines and recommended practices.

- Color guidelines for Office Add-ins


- Icon guidelines for Office Add-ins
- Layout guidelines for Office Add-ins
- Using motion in Office Add-ins
- Typography guidelines for Office Add-ins

# **Design toolkits for download**

To help you get started, we've created toolkits for use with either the [Sketch](https://www.sketch.com/) application for Mac or the [Adobe XD](https://www.adobe.com/products/xd/features.html) application for Windows or Mac. The following downloads include all of our available patterns, along with brief descriptions and layout recommendations.

- [Fluent UI Design Sketch Toolkit](https://aka.ms/fabric-sketch-toolkit)
- [Fluent UI Design Adobe XD Toolkit](https://aka.ms/fabric-toolkit)
- [Add-in Sketch Toolkit](https://aka.ms/addins_sketch_toolkit)
- [Add-in Adobe XD Toolkit](https://aka.ms/addins_toolkit)
- [Segoe UI and Fabric MDL2 icon font](https://res-1.cdn.office.net/files/fabric-cdn-prod_20230815.002/fabric-website/files/segoeui_fabricmdl2_icon_fonts.zip)


# **Color guidelines for Office Add-ins**

Article • 06/20/2024

Color is often used to emphasize brand and reinforce visual hierarchy. It helps identify an interface as well as guide customers through an experience. Inside Office, color is used for the same goals but it's applied purposefully and minimally. At no point does it overwhelm customer content. Even when each Office app is branded with its own dominant color, it's used sparingly.

Fluent UI React and Fabric Core include a set of default theme colors. When these libraries are applied to the components or layouts of an Office Add-in, the same goals apply. Color should communicate hierarchy, purposefully guiding customers to action without interfering with content. Theme colors from Fluent UI React or Fabric Core can introduce a new accent color to the overall interface. These accent colors can conflict with Office app branding and the hierarchy. Consider ways to avoid conflicts and interference. Use neutral accents or overwrite theme colors to match Office app branding or your own brand colors.

In Office applications, customers personalize their interfaces by applying an [Office UI](https://support.microsoft.com/office/365-63e65e1c-08d4-4dea-820e-335f54672310) [theme](https://support.microsoft.com/office/365-63e65e1c-08d4-4dea-820e-335f54672310) . Customers choose between four UI themes to vary styling of backgrounds and buttons in Excel, Outlook, PowerPoint, Word, and other apps in the Office suite. To make your add-ins feel like a natural part of Office and respond to personalization, use our [Theming APIs.](https://learn.microsoft.com/en-us/javascript/api/office/office.officetheme) For example, task pane background colors switch to a dark gray in some themes. With the Theming APIs, follow suit and adjust foreground text to ensure accessibility.


- To ensure that your add-in applies the correct color combinations in its UI, test it with all four Office themes, including the **Use system setting** option.
- For guidance on how to dynamically match the design of your PowerPoint add-in with the current document or Office theme, see **Use document themes in your PowerPoint add-ins**.

Apply the following general guidelines for color.

- Use color sparingly to communicate hierarchy and reinforce brand.
- Overuse of a single accent color applied to both interactive and non-interactive elements can lead to confusion. For example, avoid using the same color for selected and unselected items in a navigation menu.
- Avoid unnecessary conflicts with Office branded app colors.
- Use your own brand colors to build association with your service or company.
- Ensure that all text is accessible. Be sure that there is a 4.5:1 contrast ratio between foreground text and background.
- Be aware of color blindness. Use more than just color to indicate interactivity and hierarchy.
- To learn more about designing add-in command icons with the Office icon color palette, see icon guidelines.


# **Icons**

Article • 04/30/2024

Icons are the visual representation of a behavior or concept. They are often used to add meaning to controls and commands. Visuals, either realistic or symbolic, enable the user to navigate the UI the same way signs help users navigate their environment. They should be simple, clear, and contain only the necessary details to enable customers to quickly parse what action will occur when they choose a control.

Office app ribbon interfaces have a standard visual style. This ensures consistency and familiarity across Office apps. The guidelines will help you design a set of PNG assets for your solution that fit in as a natural part of Office.

Many HTML containers contain controls with iconography. Use Fabric Core's custom font to render Office styled icons in your add-in. The icon font provided by Fabric Core contains many glyphs for common Office metaphors that you can scale, color, and style to suit your needs. If you have an existing visual language with your own set of icons, feel free to use it in your HTML canvases. Building continuity with your own brand with a standard set of icons is an important part of any design language. Be careful to avoid creating confusion for customers by conflicting with Office metaphors.

# **Design icons for add-in commands**

Add-in commands add buttons, text, and icons to the Office UI. Your add-in command buttons should provide meaningful icons and labels that clearly identify the action the user is taking when they use a command. The following articles provide stylistic and production guidelines to help you design icons that integrate seamlessly with Office.

- For the Monoline style of Microsoft 365, see Monoline style icon guidelines for Office Add-ins.
- For the Fresh style of perpetual Office 2016 and later, see Fresh style icon guidelines for Office Add-ins.

#### 7 **Note**

You must choose one style or the other and your add-in will use the same icons whether it's running in Microsoft 365 or perpetual Office.

**See also**


- Add-in development best practices
- Add-in commands for Excel, Word, and PowerPoint

#### 6 **Collaborate with us on GitHub**

The source for this content can be found on GitHub, where you can also create and review issues and pull requests. For more information, see [our](https://github.com/OfficeDev/office-js-docs-pr/blob/main/Contributing.md) [contributor guide](https://github.com/OfficeDev/office-js-docs-pr/blob/main/Contributing.md).

#### **Office Add-ins feedback**

Office Add-ins is an open source project. Select a link to provide feedback:

- [Open a documentation issue](https://github.com/OfficeDev/office-js-docs-pr/issues/new?template=3-customer-feedback.yml&pageUrl=https%3A%2F%2Flearn.microsoft.com%2Fen-us%2Foffice%2Fdev%2Fadd-ins%2Fdesign%2Fadd-in-icons&pageQueryParams=&contentSourceUrl=https%3A%2F%2Fgithub.com%2FOfficeDev%2Foffice-js-docs-pr%2Fblob%2Fmain%2Fdocs%2Fdesign%2Fadd-in-icons.md&documentVersionIndependentId=19466ab5-cb52-e947-e4a1-0757d6675d51&feedback=%0A%0A%5BEnter+feedback+here%5D%0A&author=%40o365devx&metadata=*+ID%3A+5107ee19-7eee-237c-d18b-324f26cdece4+%0A*+Service%3A+**microsoft-365**%0A*+Sub-service%3A+**add-ins**)
- [Provide product feedback](https://aka.ms/office-addins-dev-questions)


# **Fresh style icon guidelines for Office Addins**

08/25/2025

Perpetual Office 2016 and later use Microsoft's Fresh style iconography. If you would prefer that your icons match the Monoline style of Microsoft 365, see Monoline style icon guidelines for Office Add-ins.

# **Office Fresh visual style**

The Fresh icons include only essential communicative elements. Non-essential elements including perspective, gradients, and light source are removed. The simplified icons support faster parsing of commands and controls. Follow this style to best fit with Office perpetual clients.

## **Best practices**

Follow these guidelines when you create your icons.

|  | ノ | Expand table |
|--|---|--------------|

| Do                                                                                      | Don't                                                                                                                                                                     |
|-----------------------------------------------------------------------------------------|---------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| Keep visuals simple and clear, focusing<br>on the key elements of the<br>communication. | Don't use artifacts that make your icon look messy.                                                                                                                       |
| Use the Office icon language to<br>represent behaviors or concepts.                     | Don't repurpose Fabric Core glyphs for add-in commands in<br>the Office app ribbon or contextual menus. Fabric Core icons<br>are stylistically different and won't match. |
| Reuse common Office visual metaphors                                                    | Don't reuse visual metaphors for different commands. Using                                                                                                                |
| such as paintbrush for format or                                                        | the same icon for different behaviors and concepts can cause                                                                                                              |
| magnifying glass for find.                                                              | confusion.                                                                                                                                                                |
| Redraw your icons to make them small                                                    | Don't resize your icons by shrinking or enlarging in size. This                                                                                                           |
| or larger. Take the time to redraw                                                      | can lead to poor visual quality and unclear actions. Complex                                                                                                              |
| cutouts, corners, and rounded edges to                                                  | icons created at a larger size may lose clarity if resized to be                                                                                                          |
| maximize line clarity.                                                                  | smaller without redraw.                                                                                                                                                   |


| Do                                                                                                                                                                        | Don't                                                                                                                                                                                                                                                                                                                  |
|---------------------------------------------------------------------------------------------------------------------------------------------------------------------------|------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| Use a white fill for accessibility. Most<br>objects in your icons will require a white<br>background to be legible across Office<br>UI themes and in high-contrast modes. | Avoid relying on your logo or brand to communicate what an<br>add-in command does. Brand marks aren't always recognizable<br>at smaller icon sizes and when modifiers are applied. Brand<br>marks often conflict with Office app ribbon icon styles, and can<br>compete for user attention in a saturated environment. |
| Use the PNG format with a transparent<br>background.                                                                                                                      | None                                                                                                                                                                                                                                                                                                                   |
| Avoid localizable content in your icons,<br>including typographic characters,<br>indications of paragraph rags, and<br>question marks.                                    | None                                                                                                                                                                                                                                                                                                                   |

## **Icon size recommendations and requirements**

Office desktop icons are bitmap images. Different sizes will render depending on the user's DPI setting and touch mode. Include all eight supported sizes to create the best experience in all supported resolutions and contexts. The following are the supported sizes - three are required.

- 16 px (Required)
- 20 px
- 24 px
- 32 px (Required)
- 40 px
- 48 px
- 64 px (Recommended, best for Mac)
- 80 px (Required)

#### ) **Important**

For an image that is your add-in's representative icon, see **[Create effective listings in](https://learn.microsoft.com/en-us/partner-center/marketplace-offers/create-effective-office-store-listings#create-an-icon-for-your-add-in) [AppSource and within Office](https://learn.microsoft.com/en-us/partner-center/marketplace-offers/create-effective-office-store-listings#create-an-icon-for-your-add-in)** for size and other requirements.

Make sure to redraw your icons for each size rather than shrink them to fit.


| Do<br>Redraw for each size | Don't<br>×<br>Simply shrink icons to fit at every size |  |  |
|----------------------------|--------------------------------------------------------|--|--|
| 16px                       | 16px                                                   |  |  |
| 32px                       | 32px                                                   |  |  |
| 80px                       | 80px                                                   |  |  |
| D                          | Add-in ribbon command design. @2016 Microsoft          |  |  |

## **Icon anatomy and layout**

Office icons are typically comprised of a base element with action and conceptual modifiers overlaid. Action modifiers represent concepts such as add, open, new, or close. Conceptual modifiers represent status, alteration, or a description of the icon.

To create commands that align with the Office UI, follow layout guidelines for the base element and modifiers. This ensures that your commands look professional and that your customers will trust your add-in. If you make exceptions to these guidelines, do so intentionally.

The following image shows the layout of base elements and modifiers in an Office icon.

- Center base elements in the pixel frame with empty padding all around.
- Place action modifiers on the top left.
- Place conceptual modifiers on the bottom right.
- Limit the number of elements in your icons. At 32 px, limit the number of modifiers to a maximum of two. At 16 px, limit the number of modifiers to one.


### **Base element padding**

Place base elements consistently across sizes. If base elements can't be centered in the frame, align them to the top left, leaving the extra pixels on the bottom right. For best results, apply the padding guidelines listed in the table in the following section.

### **Modifiers**

All modifiers should have a 1 px transparent cutout between each element, including the background. Elements shouldn't directly overlap. Create whitespace between rules and edges. Modifiers can vary slightly in size, but use these dimensions as a starting point.

ノ **Expand table**

| Icon size | Padding around base element | Modifier size |
|-----------|-----------------------------|---------------|
| 16 px     | 0                           | 9 px          |
| 20 px     | 1px                         | 10 px         |
| 24 px     | 1px                         | 12 px         |
| 32 px     | 2px                         | 14 px         |
| 40 px     | 2px                         | 20 px         |
| 48 px     | 3px                         | 22 px         |
| 64 px     | 5px                         | 29 px         |
| 80 px     | 5px                         | 38 px         |

## **Icon colors**

#### 7 **Note**

These color guidelines are for ribbon icons used in **Add-in commands**. These icons aren't rendered with Fluent UI.

Office icons have a limited color palette. Use the colors listed in the following table to guarantee seamless integration with the Office UI. Apply the following guidelines to the use of color.


- Use color to communicate meaning rather than for embellishment. It should highlight or emphasize an action, status, or an element that explicitly differentiates the mark.
- If possible, use only one additional color beyond gray. Limit additional colors to two at the most.
- Colors should have a consistent appearance in all icon sizes. Office icons have slightly different color palettes for different icon sizes. 16 px and smaller icons are slightly darker and more vibrant than 32 px and larger icons. Without these subtle adjustments, colors appear to vary across sizes.

ノ **Expand table**

| Color name      | RGB           | Hex     | Color | Category        |
|-----------------|---------------|---------|-------|-----------------|
| Text Gray (80)  | 80, 80, 80    | #505050 |       | Text            |
| Text Gray (95)  | 95, 95, 95    | #5F5F5F |       | Text            |
| Text Gray (105) | 105, 105, 105 | #696969 |       | Text            |
| Dark Gray 32    | 128, 128, 128 | #808080 |       | 32 px and above |
| Medium Gray 32  | 158, 158, 158 | #9E9E9E |       | 32 px and above |
| Light Gray ALL  | 179, 179, 179 | #B3B3B3 |       | All sizes       |
| Dark Gray 16    | 114, 114, 114 | #727272 |       | 16 px and below |
| Medium Gray 16  | 144, 144, 144 | #909090 |       | 16 and below    |
| Blue 32         | 77, 130, 184  | #4d82B8 |       | 32 px and above |


| Color name | RGB           | Hex     | Color | Category        |
|------------|---------------|---------|-------|-----------------|
| Blue 16    | 74, 125, 177  | #4A7DB1 |       | 16 px and below |
| Yellow ALL | 234, 194, 130 | #EAC282 |       | All sizes       |
| Orange 32  | 231, 142, 70  | #E78E46 |       | 32 px and above |
| Orange 16  | 227, 142, 70  | #E3751C |       | 16 px and below |
| Pink ALL   | 230, 132, 151 | #E68497 |       | All sizes       |
| Green 32   | 118, 167, 151 | #76A797 |       | 32 px and above |
| Green 16   | 104, 164, 144 | #68A490 |       | 16 px and below |
| Red 32     | 216, 99, 68   | #D86344 |       | 32 px and above |
| Red 16     | 214, 85, 50   | #D65532 |       | 16 px and below |
| Purple 32  | 152, 104, 185 | #9868B9 |       | 32 px and above |
| Purple 16  | 137, 89, 171  | #8959AB |       | 16 px and below |

# **Icons in high contrast modes**


Office icons are designed to render well in high contrast modes. Foreground elements are well differentiated from backgrounds to maximize legibility and enable recoloring. In high contrast modes, Office will recolor any pixel of your icon with a red, green, or blue value less than 190 to full black. All other pixels will be white. In other words, each RGB channel is assessed where 0-189 values are black and 190-255 values are white. Other high-contrast themes recolor using the same 190 value threshold but with different rules. For example, the high-contrast white theme will recolor all pixels greater than 190 opaque but all other pixels as transparent. Apply the following guidelines to maximize legibility in high-contrast settings.

- Aim to differentiate foreground and background elements along the 190 value threshold.
- Follow Office icon visual styles.
- Use colors from our icon palette.
- Avoid the use of gradients.
- Avoid large blocks of color with similar values.

# **See also**

### **Unified manifest reference**

- ["extensions.ribbons" array](https://learn.microsoft.com/en-us/microsoft-365/extensibility/schema/extension-ribbons-array)
### **Add-in only manifest reference**

- [Icon manifest element](https://learn.microsoft.com/en-us/javascript/api/manifest/icon)
- [IconUrl manifest element](https://learn.microsoft.com/en-us/javascript/api/manifest/iconurl)
- [HighResolutionIconUrl manifest element](https://learn.microsoft.com/en-us/javascript/api/manifest/highresolutioniconurl)
- [Create an icon for your add-in](https://learn.microsoft.com/en-us/partner-center/marketplace-offers/create-effective-office-store-listings#create-an-icon-for-your-add-in)


# **Monoline style icon guidelines for Office Add-ins**

Article • 02/12/2025

Monoline style iconography are used in Office apps. If you'd prefer that your icons match the Fresh style of perpetual Office 2016 and later, see Fresh style icon guidelines for Office Add-ins.

# **Office Monoline visual style**

The goal of the Monoline style to have consistent, clear, and accessible iconography to communicate action and features with simple visuals, ensure the icons are accessible to all users, and have a style that is consistent with those used elsewhere in Windows.

The following guidelines are for 3rd party developers who want to create icons for features that will be consistent with the icons already present Office products.

### **Design principles**

- Simple, clean, clear.
- Contain only necessary elements.
- Inspired by Windows icon style.
- Accessible to all users.

#### **Convey meaning**

- Use descriptive elements such as a page to represent a document or an envelope to represent mail.
- Use the same element to represent the same concept. For example, mail is always represented by an envelope, not a stamp.
- Use a core metaphor during concept development.

#### **Reduction of elements**

- Reduce the icon to its core meaning, using only elements that are essential to the metaphor.
- Limit the number of elements in an icon to two, regardless of icon size.


#### **Consistency**

Sizes, arrangement, and color of icons should be consistent.

#### **Styling**

#### **Perspective**

Monoline icons are forward-facing by default. Certain elements that require perspective and/or rotation, such as a cube, are allowed, but exceptions should be kept to a minimum.

#### **Embellishment**

Monoline is a clean minimal style. Everything uses flat color, which means there are no gradients, textures, or light sources.

# **Designing**

### **Sizes**

We recommend that you produce each icon in all these sizes to support high DPI devices. The absolutely *required* sizes are 16 px, 20 px, and 32 px, as those are the 100% sizes.

**16 px, 20 px, 24 px, 32 px, 40 px, 48 px, 64 px, 80 px, 96 px**

#### ) **Important**

For an image that is your add-in's representative icon, see **[Create effective listings](https://learn.microsoft.com/en-us/partner-center/marketplace-offers/create-effective-office-store-listings#create-an-icon-for-your-add-in) [in AppSource and within Office](https://learn.microsoft.com/en-us/partner-center/marketplace-offers/create-effective-office-store-listings#create-an-icon-for-your-add-in)** for size and other requirements.

### **Layout**

The following is an example of icon layout with a modifier.


#### **Elements**

- **Base**: The main concept that the icon represents. This is usually the only visual needed for the icon, but sometimes the main concept can be enhanced with a secondary element, a modifier.
- **Modifier** Any element that overlays the base; that is, a modifier that typically represents an action or a status. It modifies the base element by acting as an addition, alteration, or a descriptor.

### **Construction**

### **Element placement**

Base elements are placed in the center of the icon within the padding. If it can't be placed perfectly centered, then the base should err to the top right. In the following example, the icon is perfectly centered.


In the following example, the icon is erring to the left.

Modifiers are almost always placed in the bottom right corner of the icon canvas. In some rare cases, modifiers are placed in a different corner. For example, if the base element would be unrecognizable with the modifier in the bottom right corner, then consider placing it in the upper left corner.

#### **Padding**

Each size icon has a specified amount of padding around the icon. The base element stays within the padding, but the modifier should butt up to the edge of the canvas, extending outside of the padding to the edge of the icon border. The following images show the recommended padding to use for each of the icon sizes.

### **Line weights**

Monoline is a style dominated by line and outlined shapes. Depending on what size you are producing the icon should use the following line weights.

ノ **Expand table**


| Icon Size:       | 16px | 20px | 24px | 32px | 40px | 48px | 64px | 80px | 96px |
|------------------|------|------|------|------|------|------|------|------|------|
| Line<br>Weight:  | 1px  | 1px  | 1px  | 1px  | 2px  | 2px  | 2px  | 2px  | 3px  |
| Example<br>icon: |      |      |      |      |      |      |      |      |      |

#### **Cutouts**

When an icon element is placed on top of another element, a cutout (of the bottom element) is used to provide space between the two elements, mainly for readability purposes. This usually happens when a modifier is placed on top of a base element, but there are also cases where neither of the elements is a modifier. These cutouts between the two elements is sometimes referred to as a "gap".

The size of the gap should be the same width as the line weight used on that size. If making a 16 px icon, the gap width would be 1px and if it's a 48 px icon then the gap should be 2px. The following example shows a 32 px icon with a gap of 1px between the modifier and the underlying base.

In some cases, the gap can be increase by a 1/2 px if the modifier has a diagonal or curved edge and the standard gap doesn't provide enough separation. This will likely only affect the icons with 1px line weight: 16 px, 20 px, 24 px, and 32 px.

### **Background fills**

Most icons in the Monoline icon set require background fills. However, there are cases where the object would not naturally have a fill, so no fill should be applied. The following icons have a white fill.


The following icons have no fill. (The gear icon is included to show that the center hole isn't filled.)

#### **Best practices for fills**

**Do**

- Fill any element that has a defined boundary, and would naturally have a fill.
- Use a separate shape to create the background fill.
- Use **Background Fill** from the color palette.
- Maintain the pixel separation between overlapping elements.
- Fill between multiple objects.

#### **Don't**

- Don't fill objects that would not naturally be filled; for example, a paperclip.
- Don't fill brackets.
- Don't fill behind numbers or alpha characters.

### **Color**

The color palette has been designed for simplicity and accessibility. It contains 4 neutral colors and two variations for blue, green, yellow, red, and purple. Orange is intentionally not included in the Monoline icon color palette. Each color is intended to be used in specific ways as outlined in this section.

### **Palette**

|                               | Dark Gray - Standalone/Outline |  | Background Fill   |
|-------------------------------|--------------------------------|--|-------------------|
|                               | 58,58,56                       |  | 250,250,250       |
|                               | #3A3A38                        |  | #FAFAFA           |
| Medium Gray - Outline/Content |                                |  | Light Gray - Fill |
|                               | 121,119,116                    |  | 200,198,196       |
|                               | #797774                        |  | #C8C6C4           |


|  | Blue - Standalone   |  | Blue - Outline   |  | Blue - Fill   |
|--|---------------------|--|------------------|--|---------------|
|  | 30,139,205          |  | 0,99,177         |  | 131, 190, 236 |
|  | # 1E8BCD            |  | #0063B1          |  | #83BEEC       |
|  | Green - Standalone  |  | Green - Outline  |  | Green - Fill  |
|  | 24,171,80           |  | 48,144,72        |  | 161, 221, 170 |
|  | #18AB50             |  | #309048          |  | #A1DDAA       |
|  | Yellow - Standalone |  | Yellow - Outline |  | Yellow - Fill |
|  | 251,152,59          |  | 237, 135, 51     |  | 248, 219, 143 |
|  | #FB983B             |  | #ED8733          |  | #F8DB8F       |
|  | Red - Standalone    |  | Red - Outline    |  | Red - Fill    |
|  | 237,61,59           |  | 212, 35, 20      |  | 255, 145, 152 |
|  | #ED3D3B             |  | #D42314          |  | #FF9198       |
|  | Purple - Standalone |  | Purple - Outline |  | Purple - Fill |
|  | 168,70,178          |  | 146, 46, 155     |  | 212, 146, 216 |
|  | #A846B2             |  | #922E9B          |  | #D492D8       |

#### **How to use color**

In the Monoline color palette, all colors have Standalone, Outline, and Fill variations. Generally, elements are constructed with a fill and a border. The colors are applied in one of the following patterns.

- The Standalone color alone for objects that have no fill.
- The border uses the Outline color and the fill uses the Fill color.
- The border uses the Standalone color and the fill uses the Background Fill color.

The following are examples of using color.

The most common situation will be to have an element use Dark Gray Standalone with Background Fill.

When using a colored Fill, it should always be with its corresponding Outline color. For example, Blue Fill should only be used with Blue Outline. But there are two exceptions to this general rule.

- Background Fill can be used with any color Standalone.
- Light Gray Fill can be used with two different Outline colors: Dark Gray or Medium Gray.

### **When to use color**


Color should be used to convey the meaning of the icon rather than for embellishment. It should **highlight the action** to the user. When a modifier is added to a base element that has color, the base element is typically turned into Dark Gray and Background Fill so that the modifier can be the element of color, such as the case below with the "X" modifier being added to the picture base in the leftmost icon of the following set.

You should limit your icons to **one** additional color, other than the Outline and Fill mentioned above. However, more colors can be used if it's vital for its metaphor, with a limit of two additional colors other than gray. In rare cases, there are exceptions when more colors are needed. The following are good examples of icons that use just one color.

But the following icons use too many colors.

Use **Medium Gray** for interior "content", such as grid lines in an icon of a spreadsheet. Additional interior colors are used when the content needs to show the behavior of the control.

### **Text lines**

When text lines are in a "container" (for example, text on a document), use medium gray. Text lines not in a container should be **Dark Gray**.

### **Text**

Avoid using text characters in icons. Since Office products are used around the world, we want to keep icons as language neutral as possible.


# **Production**

## **Icon file format**

The final icons should be saved as .png image files. Use PNG format with a transparent background and have 32-bit depth.

# **See also**

## **Unified manifest reference**

- ["extensions.ribbons" array](https://learn.microsoft.com/en-us/microsoftteams/platform/resources/schema/manifest-schema#extensionsribbons)
## **Add-in only manifest reference**

- [Icon manifest element](https://learn.microsoft.com/en-us/javascript/api/manifest/icon)
- [IconUrl manifest element](https://learn.microsoft.com/en-us/javascript/api/manifest/iconurl)
- [HighResolutionIconUrl manifest element](https://learn.microsoft.com/en-us/javascript/api/manifest/highresolutioniconurl)
- [Create an icon for your add-in](https://learn.microsoft.com/en-us/partner-center/marketplace-offers/create-effective-office-store-listings#create-an-icon-for-your-add-in)


# **Layout**

Article • 08/18/2023

Each HTML container embedded in Office will have a layout. These layouts are the main screens of your add-in. In them you'll create experiences that enable customers to initiate actions, modify settings, view, scroll, or navigate content. Design your add-in with a consistent layouts across screens to guarantee continuity of experience. If you have an existing website that your customers are familiar with using, consider reusing layouts from your existing web pages. Adapt them to fit harmoniously within Office HTML containers.

For guidelines on layout, see Task pane, Content. For more information about how to assemble Fluent UI React, or Office UI Fabric JS, components into common layouts and user experience flows, see UX design patterns templates.

Apply the following general guidelines for layouts.

- Avoid narrow or wide margins on your HTML containers. 20 pixels is a great default.
- Align elements intentionally. Extra indents and new points of alignment should aid visual hierarchy.
- Office interfaces are on a 4px grid. Aim to keep your padding between elements at multiples of 4.
- Overcrowding your interface can lead to confusion and inhibit ease of use with touch interactions.
- Keep layouts consistent across screens. Unexpected layout changes look like visual bugs that contribute to a lack of confidence and trust with your solution.
- Follow common layout patterns. Conventions help users understand how to use an interface.
- Avoid redundant elements like branding or commands.
- Consolidate controls and views to avoid requiring too much mouse movement.
- Create responsive experiences that adapt to HTML container widths and heights.


# **Using motion in Office Add-ins**

Article • 06/29/2023

When you design an Office Add-in, you can use motion to enhance the user experience. UI elements, controls, and components often have interactive behaviors that require transitions, motion, or animation. Common characteristics of motion across UI elements define the animation aspects of a design language.

Because Office is focused on productivity, the animation language supports the goal of helping customers get things done. It strikes a balance between performant response, reliable choreography, and detailed delight. Office Add-ins sit within this existing animation language. Given this context, it's important to consider the following guidelines when applying motion.

## **Create motion with a purpose**

Motion should have a purpose that communicates additional value to the user. Consider the tone and purpose of your content when choosing animations. Handle critical messages differently than exploratory navigation.

Standard elements used in an add-in can incorporate motion to help focus the user, show how elements relate to each other, and validate user actions. Choreograph elements to reinforce hierarchy and mental models.

### **Best practices**

| Do                                                                                                                                                                                     | Don't                                                                                           |
|----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|-------------------------------------------------------------------------------------------------|
| Identify key elements in the add-in that should                                                                                                                                        | Don't overwhelm the user by animating                                                           |
| have motion. Commonly animated elements in an                                                                                                                                          | every element. Avoid applying multiple                                                          |
| add-in are panels, overlays, modals, tool tips,                                                                                                                                        | motions that attempt to lead or focus the                                                       |
| menus, and teaching call outs.                                                                                                                                                         | user on many elements at once.                                                                  |
| Use simple, subtle motion that behaves in<br>expected ways. Consider the origin of your<br>triggering element. Use motion to create a link<br>between the action and the resulting UI. | Don't create wait time for a motion. Motion<br>in add-ins should not hinder task<br>completion. |


|  | A | ー X<br>Add-in name here |
|--|---|-------------------------|
|  |   | 三                       |
|  |   |                         |
|  |   | ili<br>문                |
|  |   | A                       |
|  |   | ர்<br>11                |
|  |   |                         |
|  |   |                         |
|  |   | JI<br>ji i              |
|  |   |                         |
|  |   |                         |
|  |   | V                       |
|  | V |                         |

## **Use expected motions**

We recommend using Fluent UI to create a visual connection with the Office platform.

- [Fluent UI React motion patterns](https://react.fluentui.dev/?path=/docs/theme-motion--page)
- [Fabric Core motion and animation patterns](https://developer.microsoft.com/fluentui#/styles/web/motion)

Use it to fit seamlessly in your add-in. It will help you create experiences that are more felt than observed. The animation CSS classes provide directionality, enter/exit, and duration specifics that reinforce Office mental models and provide opportunities for customers to learn how to interact with your add-in.

## **Best practices**

| Do                                                                                                                                     | Don't                                                                                        |
|----------------------------------------------------------------------------------------------------------------------------------------|----------------------------------------------------------------------------------------------|
| Use motion that aligns with behaviors in Fluent UI.                                                                                    | Don't create motions that interfere or<br>conflict with common motion patterns<br>in Office. |
| Ensure that there is a consistent application of motion<br>across like elements.                                                       | Don't use different motions to<br>animate the same component or<br>object.                   |
| Create consistency with use of direction in animation.<br>For example, a panel that opens from the right should<br>close to the right. | Don't animate an element using<br>multiple directions.                                       |


# **Avoid out of character motion for an element**

Consider the size of the HTML canvas (task pane, dialog box, or content add-in) when implementing motion. Avoid overloading in constrained spaces. Moving elements should be in tune with Office. The character of add-in motion should be performant, reliable, and fluid. Instead of impeding productivity, aim to inform and direct.

### **Best practices**

| Do                                      | Don't                                                                                                                                                   |
|-----------------------------------------|---------------------------------------------------------------------------------------------------------------------------------------------------------|
| Use<br>recommended<br>motion durations. | Don't use exaggerated animations. Avoid creating experiences that embellish<br>and distract your customers.                                             |
| Follow<br>recommended<br>easing curves. | Don't move elements in a jerky or disjointed manner. Avoid anticipations,<br>bounces, rubber band, or other effects that emulate natural world physics. |


## **See also**

- [Fluent UI React motion patterns](https://react.fluentui.dev/?path=/docs/theme-motion--page)
- [Fabric Core animation guidelines](https://developer.microsoft.com/fluentui#/styles/web/motion)
- [Motion for Universal Windows Platform apps](https://learn.microsoft.com/en-us/windows/uwp/design/motion)


# **Typography**

Article • 08/23/2023

Segoe is the standard typeface for Office. Use it in your add-in to align with Office task panes, dialog boxes, and content objects. Fabric Core gives you access to Segoe. It provides a full type ramp of Segoe with many variations - across font weight and size in convenient CSS classes. Not all Fabric Core sizes and weights will look great in an Office Add-in. To fit harmoniously or avoid conflicts, consider using a subset of the Fabric Core type ramp. The following table lists Fabric Core's base classes that we recommend for use in Office Add-ins.

#### 7 **Note**

Text color isn't included in these base classes. Use Fabric Core's "neutral primary" for most text on white backgrounds.

To learn more about available typography, see **[Web Typography](https://developer.microsoft.com/fluentui#/styles/web/typography)**.

| Type     | Class              | Size     | Weight             | Recommended Usage                                                                                                                                                                                                                                                                                                                        |
|----------|--------------------|----------|--------------------|------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| Hero     | .ms<br>font<br>xxl | 28<br>px | Segoe<br>Light     | This class is larger than all other typographic<br>elements in Office. Use it sparingly to avoid<br>unseating visual hierarchy.<br>Avoid use on long strings in constrained<br>spaces.<br>Provide ample whitespace around text using<br>this class.<br>Commonly used for first-run messages, hero<br>elements, or other calls to action. |
| Title    | .ms<br>font-xl     | 21<br>px | Segoe<br>Light     | This class matches the task pane title of Office<br>applications.<br>Use it sparingly to avoid a flat typographic<br>hierarchy.<br>Commonly used as the top-level element such<br>as dialog box, page, or content titles.                                                                                                                |
| Subtitle | .ms<br>font-l      | 17<br>px | Segoe<br>Semilight | This class is the first stop below titles.<br>Commonly used as a subtitle, navigation<br>element, or group header.                                                                                                                                                                                                                       |
| Body     | .ms-               | 14       | Segoe              | Commonly used as body text within add-ins.                                                                                                                                                                                                                                                                                               |


| Type       | font-m<br>Class   | px<br>Size | Regular<br>Weight | Recommended Usage                                                                                                                 |
|------------|-------------------|------------|-------------------|-----------------------------------------------------------------------------------------------------------------------------------|
| Caption    | .ms<br>font-xs    | 11<br>px   | Segoe<br>Regular  | Commonly used for secondary or tertiary text<br>such as timestamps, by lines, captions, or field<br>labels.                       |
| Annotation | .ms<br>font<br>mi | 10<br>px   | Segoe<br>Semibold | The smallest step in the type ramp should be<br>used rarely. It's available for circumstances<br>where legibility isn't required. |


# **UX design patterns for Office Add-ins**

Article • 02/11/2025

Designing the user experience for Office Add-ins should provide a compelling experience for Office users and extend the overall Office experience by fitting seamlessly within the default Office UI.

Our UX patterns are composed of components. Components are controls that help your customers interact with elements of your software or service. Buttons, navigation, and menus are examples of common components that often have consistent styles and behaviors.

Fluent UI React components look and behave like a part of Office, as do the frameworkneutral components of Office UI Fabric JS. Take advantage of either set of components to integrate with Office. Alternatively, if your add-in has its own preexisting component language, you don't need to discard it. Look for opportunities to retain it while integrating with Office. Consider ways to swap out stylistic elements, remove conflicts, or adopt styles and behaviors that remove user confusion.

The provided patterns are best practice solutions based on common customer scenarios and user experience research. They are meant to provide both a quick entry point to designing and developing add-ins as well as guidance to achieve balance between Microsoft brand elements and your own. Providing a clean, modern user experience that balances design elements from Microsoft's Fluent UI design language and the partner's unique brand identity may help increase user retention and adoption of your add-in.

Use the UX pattern templates to:

- Apply solutions to common customer scenarios.
- Apply design best practices.
- Incorporate Fluent UI components and styles.
- Build add-ins that visually integrate with the default Office UI.
- Ideate and visualize UX.

# **Getting started**

The patterns are organized by key actions or experiences that are common in an add-in. The main groups are:

- First-run experience (FRE)
- Authentication


- Navigation
- Branding Design

Browse each grouping to get an idea of how you can design your add-in using best practices.

#### 7 **Note**

The example screens shown throughout this documentation are designed and displayed at a resolution of **1366x768**.

# **See also**

- Office Add-in design language
- Best practices for developing Office Add-ins
- Fluent UI React in Office Add-ins