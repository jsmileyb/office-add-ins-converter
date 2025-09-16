{0}------------------------------------------------

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

{1}------------------------------------------------

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

{2}------------------------------------------------

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

{3}------------------------------------------------

U8zoWMNAYUVjnifRhPq3uIJmUz2NNlGu8SQ+IfwJGLIuagFi5hOEJhGZcLUVwbVP0oihfyh8KmbTl iOpxBHEnCKbgb0vBJjCUoLGMmg3i7LrejFqV3WMqbahEs00McTohw1rsGKRpQWKvCq+m86kdpUWe3 FJapsLpFOKuYNkZiSGGtCK9O1uNArerstpRnJcuMzhJYz0pHUUxRWMYKr+qDDGEVpiwXPnZe+NhZ/ scUFp1V5X6m/yCL87CW8FfueuSDMqaMJi67I68H7k5ZAGx2z7z83PDOzZPLtuCcLHMyEsPQcUFbPv YLXb80jBHWJ7wbq4etAMjoXnPfVgBu5Gwv/eP2VQHYJZ5JM+ILwyus96TOgywK6xM+qlcyJVVYH95 ovm+r+87ieSOLdMcEJYYjp3U9/YWvqgcATOfw0Zl6HMgDSJ1AvzL7A9bbZ8ts9/OAkIWyh/7kYJWf bt68+eWXX965cycvDf18ld3YHRWUGKaj2K7XOIsRFgDgaqJOB5LXpuapA3eW/u+XP50ps5GwZf3lZ Xc/drtx44UXXvjwww95aegsMfT0CgiyYkwJmQ8KC6/k5XAvPXX1qQN3DmwtHNha+MYXUy/82ojkXa 2O11Zw9+7db3/727w0dIaQ0KmY1u/TZKaEdYywZHcBdNpnI19MKfK7HNp2951fOKtv88477/DS0Bl tIXh6a3d0yMTh7dPj38cICxhPqb7UGcAGueGVshzcWuCZNyZCuv7rNsKisICM+hOXLAqEVwoLa1uh ehmL6BVvHxz+GtuU8jp9JxiFVRdhdUp/OKiqs3jyjYmQzsj6DLIoLPULWK2qLR6UR2gv29GYWFj6b DhDYQHjKRV33gR0avTv/sKBO8wKjYkjUDsrZEqoW2GpVnG3AAEtJoP3KT+TixiAjloPAUt5PTZvML 2wVC1gbQbsmv7rv/TpWwyyjImQd1bqMMgyvbCmVavjdMpa51wfnPZZBlnGRMjbvSgsRlgPYiuPPga AQZYxEdJD5p7+nrqYfqa7GhFWp25stRBkJf6MLbSMSM0p74ywiL5sRQiFpSsUnIRlAfy0FdEMQuru TAlNih3YovVngoToPSWksJRgA+DV9HwrQspQyuvuV6aw5NEEbFHtPYOEEApLGda54MpjIweC6BbhL 47WDHxKKIl9/fhKirYimqbmNCurmxGW0aOqAwM4OIB1bg4G0ToFA06vo7CEseMY9oWxo48jQQiFpU k6erAlhC0heorokpqd3XU4Ucv0wvIuWSNqd6MjiHVueIPYEuIFX7unEtEsQlrHsIalP45GeW1XxEN r6ze8StTeRkhHB43Bp4SkMkI67RJtIqQDciuFRYwTXvUacjEaI6wH+b4OU0IKi1RAyDuEiTYpJFGs 1fhbh/kghUUq0HWC4ZWOuR4TEEHr8vxSWGQVvn7s5rMIPTMq4J2DfFU9MYitumMcBh2TiQvIB3sE9 VDWHpw4Sj7B4ceuqE5vvOQBaQHRsW4nrFBYpsfqQlsI3j5OYjACY4naE9wB+AcoLCKD/1ViJ/uBz8 1zLE3NcKT2NrrNB8EallZgb2Uin5GYscMrCksz2DgERB6lHN4XYCKrS9e5P4WlDdhkmchkKCyoR7u ewysKSzOs5xAQeclg5oyg8IrCIgpgB5o4CkQShSSGjgva0j+gx/WDFJYmcXEIiHiKKZwLCdrS4TfA +lBOa9AMHuC27J38tIEDWSccftgDcAbhCMAZXJull6Uckn1CXy+4ywjLrSgsLWWFdmCKA6GX0CaNY nrZNALvMXhC8PbVaZZTKYdzIRQuCdrYe8wYaxga5uc51VAz4UwWuCbj4/t5SjQSLPfiobC6swcKSZ wLCY2trC48ntJ79WoB1rA0lhWy9G4AsoMYOo6fuTEcQSmn/P4zcRG2AtAdM4atKCzt4eMQGIVSHsM ncTagsLbeH0DyaRG26jphpAXtTAm1lBIu8DEwwZTQWCjSCWMsgQ8GhBatFnD24IihXqfKCEt7dAIW joKxKKaRfBrJPomhVjGFoTDOHxVtq8MJgw0kIyztRVgAxoG0+E8xwtI+VheCcRFzIIopjMSQjorIA e8f6HBCp43bKSy9CQvALSBDYRmUXadqL5HJxJGJY+S0RC0a0VbgPCztshG4B2Q5EEbkg6+jkCzTir qUw1gCmTiyidptjs1nKwpL23QCduAGB8KIjJzG9E1s/SOUcigkUUyhkBRXojKfrZgSajglvM84cA2 YZUpIarFQZTfKlKuy8Cmh5mkFdgEeDgSpiq/f8LaisGSH3/XBAnRSW6Qyu04ZaTo7haVSBF7fSoEN 6AT2Ap1AK+dqkU/SwCMX9d6WT0SQwDOuMyyA55NQaxaYAmaBIsfFlAG+f8AALa4orDpGWEJeUqKqv BaawbfyZJgMXz+2R/T7ti4Kay1ocnMMSL3x9GJ7ZG36BVJYuo+wCGFURWHpA3NfOqRe98Ue+MJ4KG yGh4AUFiMsok/q3HCZwjLL3U/+cgpCltL+6zj4JoehLJyHxSCLaIw7P1GlsTKFRSgsogrXYxwDCov CIjrhSpRjQGGpgFmnwxB1KaZRSHIYKCwV8PRyDIjyZOIcAwpLBQz0DiVCYVFYzAoJEU/hEp8VUlgq 4AzC4ecwEOUZZxmLwlIDD4MsogLZBMeAwlIBlrGIKlkhIywKSyVh1a1dMjEP91jDorBUwhfmGBClU 8JBjgGFpQ4PUViEUFh6wRmEs4fDQAiFpRNM8+YSQigs/cPSOyEUlm6wull6J4TCYlZICKGwFMcRgK +fw0AIhaUTTPYmXkIoLAZZhBAKi0EWIRQWWRlkdZ3gMBBCYemEHRHOySJEcfgiVcFM5/BRHFcTyCQ xKuDlqRuAzRw1QiisOnM1gXei+OiMuE/dBjyAncNHCIVVHzJJvDWAa1K7fIwA2ziIhFBYdeC/Inj7 pKw9TAC3gQ0cSkIoLPWYzuFfQoIKVbVjNMAFNHFMCVEAPiUslwYqZSsAs0CKY0oIIyw1GE/hX0OYz iu5zykgA3g5uIQwwlI2E4z3KWyrBW4BExxfQigsBXkzrFgmuJo0MMshJoTCUoQL4mdaiWIW+JijTA iFpUgy+HZE9aNMAdc41oRQWDJ5a0CV0tVqssBtDjchFJZkxlN493T9DncDyHLQCaGwpDEUq/cRr/G hISFS4Dws4N3YGhw0DWzj0mgVmZvBzBhmZzAzVuFebcO6NljXw7qeo0Vh6YVMEuPpNTjuwkNDOktR pm6ieBMzYyjexNyMiA86NsHWBvsmODah0caBpLA0y0fxNTs0naUQE1cweQUTV8RJainFmyjeRP7yo ryau9DSxchLi5i+hnU1sZZHX3AW61nSaN6J7tfwRLbpsxfh65dsq9XyuvM/uPI6bv0ME1c4ytqiYX 5+3tQD8LeBtUkJV9AJeHg1CsayHr/0fXQ8tfTf5iZz+Tei+Teic5NKzlCxrocnCOeONfpLPzfPs01 hLeHlBq38JpvZOUsYzh4cTsDqLvtDo2mLwmJKKCMO6lVx5zeAIoe4Fr5+HElWshWAxma355lI11+m nI8r+XbI0l2M/ieu/RumbvIcMMLSS4T10jwAZJIYTSKTRCYpvXtymTQH2MsLsirdr8EXFr558XJi9 NVw6bbCKb9rLzzBej1MZIRFYckV1gqU8pcH6OQFWQGHH8E4nEGxn5ubzGVfj+TffEXZX8e6Ht5HYd 9EYVFYuhOWUv7yA628IMvhPYbuWJU0sCYT5+Ojr4aVrWoB8AThCar8t1NYFJa6wpLsr72AhRfkikj Ghd1RUWlgJUqjqZvf6ZtJK9zvrKULHY+qmR5SWBRWXYUl0F+tgJ9X44oAphfdMTgCSu1vbjI3+t3w xAWFu57Z2rDpCdVmmVJYFNZaCquSvyb+GbZbvBwfBFbdMXj71Nj36KvhwlmFm3M02uB7ErY2Ckt1O K1hTfEGsS+Mo1E0T3EwFuk6gcdTKtkKQMeLMWVnPACYm8HIm4Czh2ePwjIBhSRKeQ4DPL147BfYHZ VTX19DZ+Fwgs6isEzA9RhVhUNv4XBCwYpV/Z0Fq5vOorBMwGjcvH/7fVW1hep8ZBWd5eADFArLwPl gMW3GP9zXjyMX10RV92l/PmrzKx0QWd0IxmF18dKmsIzIWEKZ/ez8C30kIw4/dp3CE1l0xyTMXFf4 6m92+/400distFycQQTjvLQpLCMyElNgJ95jCPwBjiRx5CK6TmgxJbG64OvHobfwWAr+AbXL6mKdp

{4}------------------------------------------------

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';

}

})();

{5}------------------------------------------------

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

{6}------------------------------------------------

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

{7}------------------------------------------------

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

{8}------------------------------------------------

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

{9}------------------------------------------------

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

{10}------------------------------------------------

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

{11}------------------------------------------------

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

{12}------------------------------------------------

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

{13}------------------------------------------------

- Ensure your environment is ready for Office development by following the instructions in Set up your development environment.
- Some of the sample code uses ES6 JavaScript. This isn't compatible with older versions of Office that use the Trident (Internet Explorer 11) browser engine. For information on how to support those platforms in your add-in, see Support older Microsoft webviews and Office versions. If you don't already have a Microsoft 365 subscription to use for development, you might qualify for a Microsoft 365 E5 developer subscription through the [Microsoft 365 Developer Program](https://aka.ms/m365devprogram) ; for details, see the [FAQ](https://learn.microsoft.com/en-us/office/developer-program/microsoft-365-developer-program-faq#who-qualifies-for-a-microsoft-365-e5-developer-subscription-). Alternatively, you can [sign up for a 1-month free trial](https://www.microsoft.com/microsoft-365/try) or [purchase a Microsoft 365 plan](https://www.microsoft.com/microsoft-365/business/compare-all-microsoft-365-business-products-g) .
- If your add-in shows an error (for example, "This add-in could not be started. Close this dialog to ignore the problem or click "Restart" to try again.") when you press F5 or choose **Debug** > **Start Debugging** in Visual Studio, see Debug Office Addins in Visual Studio for other debugging options.

### **See also**

- Office Add-ins platform overview
- Develop Office Add-ins
- Using Visual Studio Code to publish